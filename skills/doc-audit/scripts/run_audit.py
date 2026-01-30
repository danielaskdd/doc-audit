#!/usr/bin/env python3
"""
ABOUTME: Executes LLM-based audit on document text blocks
ABOUTME: Sends each block with context and rules to LLM, saves results to manifest
ABOUTME: Supports parallel processing with configurable concurrency
"""

import argparse
import asyncio
import json
import os
import random
import sys
from datetime import datetime
from pathlib import Path

# Add script directory to path for local module imports
_SCRIPT_DIR = Path(__file__).resolve().parent
if str(_SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(_SCRIPT_DIR))

from utils import (  # noqa: E402
    HAS_GEMINI,
    HAS_OPENAI,
    audit_block_gemini_async,
    audit_block_openai_async,
    create_gemini_client,
    create_openai_client,
    estimate_tokens,
    get_gemini_provider_name,
    get_openai_provider_name,
    global_extract_gemini_async,
    global_extract_openai_async,
    global_verify_gemini_async,
    global_verify_openai_async,
    is_gemini_retryable,
    is_openai_reasoning_model,
    is_openai_retryable,
    is_vertex_ai_mode,
)

from prompt import (  # noqa: E402
    build_block_audit_system_prompt,
    build_block_audit_user_prompt,
    build_global_extract_system_prompt,
    build_global_extract_user_prompt,
    build_global_verify_system_prompt,
    build_global_verify_user_prompt,
)

# Maximum number of concurrent LLM API calls
MAX_PARALLEL_WORKERS = 8

# Retry configuration
DEFAULT_MAX_RETRIES = 3
INITIAL_BACKOFF = 1.0  # Initial backoff time in seconds
MAX_BACKOFF = 60.0     # Maximum backoff time in seconds
BACKOFF_MULTIPLIER = 2.0  # Backoff multiplier for exponential growth

# Global audit context limit (tokens)
DEFAULT_GLOBAL_AUDIT_MAX_TOKENS = 120000



def get_global_audit_max_tokens() -> int:
    """
    Resolve global audit context limit from environment.

    Returns:
        int: Maximum allowed tokens for a global audit context chunk
    """
    raw = os.getenv("GLOBAL_AUDIT_MAX_TOKENS")
    if not raw:
        return DEFAULT_GLOBAL_AUDIT_MAX_TOKENS
    try:
        value = int(raw)
    except ValueError:
        print(
            f"Warning: GLOBAL_AUDIT_MAX_TOKENS must be an integer, got '{raw}'. "
            f"Using default {DEFAULT_GLOBAL_AUDIT_MAX_TOKENS}.",
            file=sys.stderr
        )
        return DEFAULT_GLOBAL_AUDIT_MAX_TOKENS
    if value <= 0:
        print(
            f"Warning: GLOBAL_AUDIT_MAX_TOKENS must be positive, got '{value}'. "
            f"Using default {DEFAULT_GLOBAL_AUDIT_MAX_TOKENS}.",
            file=sys.stderr
        )
        return DEFAULT_GLOBAL_AUDIT_MAX_TOKENS
    return value


class NonRetryableError(Exception):
    """
    Exception for errors that should not be retried.
    
    These include authentication errors, invalid API keys, permission errors,
    and other permanent failures that won't be resolved by retrying.
    """
    pass


def validate_thinking_config(thinking_level: str = None, thinking_budget: int = None,
                            thinking_level_source: str = None, thinking_budget_source: str = None):
    """
    Validate that thinking_level and thinking_budget are not both set.
    
    thinking_level is for Gemini 3 models, thinking_budget is for Gemini 2.5 models.
    Using both simultaneously would cause the API to receive incompatible parameters.
    
    Args:
        thinking_level: Thinking level value (if set)
        thinking_budget: Thinking budget value (if set)
        thinking_level_source: Source description for error message (e.g., "env GEMINI_THINKING_LEVEL")
        thinking_budget_source: Source description for error message (e.g., "--thinking-budget")
        
    Raises:
        SystemExit: If both parameters are set
    """
    if thinking_level and thinking_budget is not None:
        print("Error: Both thinking_level and thinking_budget are set.", file=sys.stderr)
        if thinking_level_source:
            print(f"  thinking_level: {thinking_level.upper()} (from {thinking_level_source})", file=sys.stderr)
        if thinking_budget_source:
            print(f"  thinking_budget: {thinking_budget} (from {thinking_budget_source})", file=sys.stderr)
        print("", file=sys.stderr)
        print("thinking_level is for Gemini 3 models, thinking_budget is for Gemini 2.5 models.", file=sys.stderr)
        print("Please use only one:", file=sys.stderr)
        print("  - For Gemini 3: Use --thinking-level or GEMINI_THINKING_LEVEL", file=sys.stderr)
        print("  - For Gemini 2.5: Use --thinking-budget or GEMINI_THINKING_BUDGET", file=sys.stderr)
        sys.exit(1)


async def audit_block_with_retry(
    block: dict,
    system_prompt: str,
    model_name: str,
    client,
    use_gemini: bool,
    max_retries: int = DEFAULT_MAX_RETRIES,
    block_idx: int = 0,
    total_blocks: int = 1,
    thinking_level: str = None,
    thinking_budget: int = None,
    reasoning_effort: str = None
) -> dict:
    """
    Audit a block with automatic retry on transient errors.
    
    Uses exponential backoff with jitter to handle rate limits and transient
    server errors. Non-retryable errors (authentication, invalid request) are
    raised immediately without retry.
    
    Args:
        block: Text block to audit
        system_prompt: Cached system prompt with rules
        model_name: LLM model name
        client: Async LLM client instance
        use_gemini: Whether to use Gemini (True) or OpenAI (False)
        max_retries: Maximum number of retry attempts
        block_idx: Block index for logging
        total_blocks: Total blocks for logging
        thinking_level: Thinking level for Gemini 3 models
        thinking_budget: Thinking budget for Gemini 2.5 models
        reasoning_effort: Reasoning effort for OpenAI o-series models
        
    Returns:
        Audit result dictionary
        
    Raises:
        NonRetryableError: For permanent errors that should not be retried
        Exception: For errors that exceeded retry attempts
    """
    last_error = None
    
    for attempt in range(max_retries + 1):
        try:
            user_prompt = build_block_audit_user_prompt(block)
            if use_gemini:
                return await audit_block_gemini_async(
                    user_prompt, system_prompt, model_name, client,
                    thinking_level=thinking_level,
                    thinking_budget=thinking_budget
                )
            else:
                return await audit_block_openai_async(
                    user_prompt, system_prompt, model_name, client,
                    reasoning_effort=reasoning_effort
                )
                
        except Exception as e:
            last_error = e
            
            # Check if error is retryable
            if use_gemini:
                retryable = is_gemini_retryable(e)
            else:
                retryable = is_openai_retryable(e)
            
            if not retryable:
                raise NonRetryableError(f"Non-retryable error: {e}") from e
            
            # Check if we have retries left
            if attempt >= max_retries:
                raise
            
            # Calculate backoff with jitter
            backoff = min(INITIAL_BACKOFF * (BACKOFF_MULTIPLIER ** attempt), MAX_BACKOFF)
            jitter = backoff * 0.1 * (2 * random.random() - 1)  # ±10% jitter
            wait_time = backoff + jitter
            
            print(f"  [{block_idx+1}/{total_blocks}] Retry {attempt + 1}/{max_retries} "
                  f"after {wait_time:.1f}s: {type(e).__name__}: {str(e)[:100]}")
            await asyncio.sleep(wait_time)
    
    # Should not reach here, but just in case
    raise last_error


async def global_extract_with_retry(
    block: dict,
    system_prompt: str,
    model_name: str,
    client,
    use_gemini: bool,
    max_retries: int = DEFAULT_MAX_RETRIES,
    block_idx: int = 0,
    total_blocks: int = 0,
    thinking_level: str = None,
    thinking_budget: int = None,
    reasoning_effort: str = None
) -> dict:
    last_error = None
    for attempt in range(max_retries + 1):
        try:
            user_prompt = build_global_extract_user_prompt(block)
            if use_gemini:
                return await global_extract_gemini_async(
                    user_prompt, system_prompt, model_name, client,
                    thinking_level=thinking_level, thinking_budget=thinking_budget
                )
            return await global_extract_openai_async(
                user_prompt, system_prompt, model_name, client,
                reasoning_effort=reasoning_effort
            )
        except Exception as e:
            last_error = e
            retryable = is_gemini_retryable(e) if use_gemini else is_openai_retryable(e)
            if not retryable:
                raise NonRetryableError(f"Non-retryable error: {e}") from e
            if attempt >= max_retries:
                raise
            backoff = min(INITIAL_BACKOFF * (BACKOFF_MULTIPLIER ** attempt), MAX_BACKOFF)
            jitter = backoff * 0.1 * (2 * random.random() - 1)
            wait_time = backoff + jitter
            print(
                f"  [Extract {block_idx+1}/{total_blocks}] Retry {attempt + 1}/{max_retries} "
                f"after {wait_time:.1f}s: {type(e).__name__}: {str(e)[:100]}",
                file=sys.stderr
            )
            await asyncio.sleep(wait_time)
    raise last_error


async def global_verify_with_retry(
    rule: dict,
    items: list,
    system_prompt: str,
    model_name: str,
    client,
    use_gemini: bool,
    max_retries: int = DEFAULT_MAX_RETRIES,
    rule_idx: int = 0,
    total_rules: int = 0,
    thinking_level: str = None,
    thinking_budget: int = None,
    reasoning_effort: str = None
) -> dict:
    last_error = None
    for attempt in range(max_retries + 1):
        try:
            user_prompt = build_global_verify_user_prompt(rule, items)
            if use_gemini:
                return await global_verify_gemini_async(
                    user_prompt, system_prompt, model_name, client,
                    thinking_level=thinking_level, thinking_budget=thinking_budget
                )
            return await global_verify_openai_async(
                user_prompt, system_prompt, model_name, client,
                reasoning_effort=reasoning_effort
            )
        except Exception as e:
            last_error = e
            retryable = is_gemini_retryable(e) if use_gemini else is_openai_retryable(e)
            if not retryable:
                raise NonRetryableError(f"Non-retryable error: {e}") from e
            if attempt >= max_retries:
                raise
            backoff = min(INITIAL_BACKOFF * (BACKOFF_MULTIPLIER ** attempt), MAX_BACKOFF)
            jitter = backoff * 0.1 * (2 * random.random() - 1)
            wait_time = backoff + jitter
            print(
                f"  [Verify {rule_idx+1}/{total_rules}] Retry {attempt + 1}/{max_retries} "
                f"after {wait_time:.1f}s: {type(e).__name__}: {str(e)[:100]}",
                file=sys.stderr
            )
            await asyncio.sleep(wait_time)
    raise last_error


def chunk_items_by_token_limit(rule: dict, items: list, max_tokens: int) -> list:
    """
    Chunk items so that each chunk's total prompt (system + user) stays within max_tokens.
    
    The system prompt includes the rule's topic and verification text, which can be
    substantial for rules with detailed descriptions or many fields. This overhead
    must be accounted for to avoid exceeding the token limit when the actual API
    call includes both system and user prompts.
    
    Args:
        rule: Global audit rule
        items: List of extracted items to chunk
        max_tokens: Maximum allowed tokens per chunk (total context)
        
    Returns:
        List of item chunks, each staying within the token budget
    """
    # Calculate system prompt token overhead for this rule
    system_prompt = build_global_verify_system_prompt(rule)
    system_tokens = estimate_tokens(system_prompt)
    available_tokens = max_tokens - system_tokens
    
    # Handle edge case where system prompt alone exceeds max_tokens
    if available_tokens <= 0:
        print(
            f"Warning: System prompt for rule {rule.get('id','')} exceeds GLOBAL_AUDIT_MAX_TOKENS "
            f"({system_tokens} tokens > {max_tokens}). Using fallback budget of {max_tokens // 2} tokens.",
            file=sys.stderr
        )
        available_tokens = max_tokens // 2  # Fallback: use half for user prompt
    
    chunks = []
    current = []
    for item in items:
        trial = current + [item]
        prompt_text = build_global_verify_user_prompt(rule, trial)
        user_tokens = estimate_tokens(prompt_text)
        
        if user_tokens <= available_tokens or not current:
            current = trial
            # Warn if a single item exceeds available budget
            if user_tokens > available_tokens and len(trial) == 1:
                total_tokens = user_tokens + system_tokens
                print(
                    f"Warning: Single item exceeds token budget for rule {rule.get('id','')}. "
                    f"Estimated {user_tokens} (user) + {system_tokens} (system) = {total_tokens} > {max_tokens}.",
                    file=sys.stderr
                )
        else:
            chunks.append(current)
            current = [item]
    
    if current:
        chunks.append(current)
    
    return chunks


def load_blocks(file_path: str) -> tuple:
    """
    Load text blocks and metadata from JSONL or JSON file.

    Args:
        file_path: Path to blocks file

    Returns:
        Tuple of (metadata dict, list of block dictionaries)
    """
    blocks = []
    metadata = {}
    path = Path(file_path)

    if path.suffix == '.jsonl':
        with open(path, 'r', encoding='utf-8') as f:
            for line in f:
                line = line.strip()
                if line:
                    entry = json.loads(line)
                    # Check if this is metadata
                    if entry.get('type') == 'meta':
                        metadata = entry
                    else:
                        blocks.append(entry)
    else:
        with open(path, 'r', encoding='utf-8') as f:
            data = json.load(f)
            if isinstance(data, list):
                blocks = data
            elif 'blocks' in data:
                blocks = data['blocks']
                # Extract metadata if present
                if 'meta' in data:
                    metadata = data['meta']
            else:
                raise ValueError(f"Unknown JSON format in {file_path}")

    return metadata, blocks


def load_rules(file_path: str) -> list:
    """
    Load audit rules from JSON file.

    Args:
        file_path: Path to rules JSON file

    Returns:
        List of rule dictionaries
    """
    with open(file_path, 'r', encoding='utf-8') as f:
        data = json.load(f)

    if isinstance(data, list):
        rules = data
    elif 'rules' in data:
        rules = data['rules']
        default_type = data.get('type')
        if default_type:
            for rule in rules:
                rule.setdefault('type', default_type)
    else:
        raise ValueError(f"Unknown rules format in {file_path}")
    
    # Normalize type values for backward compatibility
    # 1. Set default type for rules missing the field
    # 2. Normalize legacy "block_level" → "block"
    for rule in rules:
        if 'type' not in rule:
            rule['type'] = 'block'  # Set default
        elif rule['type'] == 'block_level':
            rule['type'] = 'block'  # Normalize legacy
    
    return rules


def merge_rules(rule_files: list) -> list:
    """
    Load and merge rules from multiple JSON files.
    Checks for duplicate rule IDs and exits if found.
    
    Args:
        rule_files: List of paths to rules JSON files
        
    Returns:
        Merged list of rule dictionaries
        
    Exits:
        If duplicate rule IDs are found
    """
    merged_rules = []
    seen_ids = {}  # rule_id -> source_file
    
    for file_path in rule_files:
        rules = load_rules(file_path)
        for rule in rules:
            rule_id = rule.get('id', '')
            if not rule_id:
                print(f"Warning: Rule without ID found in {file_path}, skipping", file=sys.stderr)
                continue
            
            if rule_id in seen_ids:
                print(f"Error: Duplicate rule ID '{rule_id}' found.", file=sys.stderr)
                print(f"  First occurrence: {seen_ids[rule_id]}", file=sys.stderr)
                print(f"  Duplicate in: {file_path}", file=sys.stderr)
                sys.exit(1)
            
            seen_ids[rule_id] = file_path
            merged_rules.append(rule)
    
    return merged_rules


def build_rule_category_map(rules: list) -> dict:
    """
    Build a mapping from rule_id to category.

    Args:
        rules: List of rule dictionaries

    Returns:
        Dictionary mapping rule_id to category
    """
    return {rule['id']: rule.get('category', 'other') for rule in rules}


async def save_manifest_entry_async(manifest_path: str, entry: dict, lock: asyncio.Lock):
    """
    Append an entry to the manifest JSONL file (async with lock).

    Args:
        manifest_path: Path to manifest file
        entry: Entry dictionary to append
        lock: asyncio.Lock for thread-safe writing
    """
    async with lock:
        with open(manifest_path, 'a', encoding='utf-8') as f:
            f.write(json.dumps(entry, ensure_ascii=False) + '\n')


def save_manifest_entry(manifest_path: str, entry: dict):
    """
    Append an entry to the manifest JSONL file.

    Args:
        manifest_path: Path to manifest file
        entry: Entry dictionary to append
    """
    with open(manifest_path, 'a', encoding='utf-8') as f:
        f.write(json.dumps(entry, ensure_ascii=False) + '\n')


def derive_extraction_path(manifest_path: str) -> str:
    """
    Derive extraction file path from manifest path.
    
    Args:
        manifest_path: Path to manifest file
    
    Returns:
        Path to extraction file (e.g., manifest.jsonl → manifest_extractions.jsonl)
    """
    path = Path(manifest_path)
    stem = path.stem  # e.g., "manifest"
    parent = path.parent
    return str(parent / f"{stem}_extractions.jsonl")


async def save_extraction_entry_async(extraction_path: str, entry: dict, lock: asyncio.Lock):
    """
    Append an extraction entry to the extraction JSONL file (async with lock).
    
    Args:
        extraction_path: Path to extraction file
        entry: Entry dictionary to append
        lock: asyncio.Lock for thread-safe writing
    """
    async with lock:
        with open(extraction_path, 'a', encoding='utf-8') as f:
            f.write(json.dumps(entry, ensure_ascii=False) + '\n')


def load_completed_extraction_uuids(extraction_path: str) -> set:
    """
    Load UUIDs of already-extracted blocks from extraction file.
    
    Args:
        extraction_path: Path to extraction file
    
    Returns:
        Set of completed UUIDs
    """
    completed = set()
    if not Path(extraction_path).exists():
        return completed
    for entry in iter_manifest_entries(extraction_path, ignore_errors=True):
        # Skip metadata entries
        if entry.get('type') == 'meta':
            continue
        uuid = entry.get('uuid', '')
        if uuid:
            completed.add(uuid)
    return completed


def load_extraction_buckets(extraction_path: str, global_rules: list) -> dict:
    """
    Rebuild rule_buckets from extraction file.
    
    Args:
        extraction_path: Path to extraction file
        global_rules: List of global audit rules
    
    Returns:
        Dict mapping rule_id -> list of extracted items
    """
    rule_buckets = {rule.get('id', ''): [] for rule in global_rules}
    if not Path(extraction_path).exists():
        return rule_buckets
    
    for entry in iter_manifest_entries(extraction_path, ignore_errors=True):
        # Skip metadata entries
        if entry.get('type') == 'meta':
            continue
        
        uuid = entry.get('uuid', '')
        uuid_end = entry.get('uuid_end', uuid)
        p_heading = entry.get('p_heading', '')
        
        for rule_result in entry.get('results', []):
            rule_id = rule_result.get('rule_id', '')
            if rule_id not in rule_buckets:
                continue
            
            for extracted in rule_result.get('extracted_results', []):
                item = {
                    "uuid": uuid,
                    "uuid_end": uuid_end,
                    "p_heading": p_heading,
                    "entity": extracted.get('entity', ''),
                    "fields": extracted.get('fields', [])
                }
                rule_buckets[rule_id].append(item)
    
    return rule_buckets


def iter_manifest_entries(manifest_path: str, ignore_errors: bool = False):
    """
    Iterate over JSONL entries in a manifest file with optional error tolerance.

    Args:
        manifest_path: Path to manifest JSONL file
        ignore_errors: If True, skip malformed JSON lines with a warning
    """
    path = Path(manifest_path)
    if not path.exists():
        return

    with open(path, 'r', encoding='utf-8') as f:
        for lineno, line in enumerate(f, start=1):
            line = line.strip()
            if not line:
                continue
            try:
                yield json.loads(line)
            except json.JSONDecodeError as e:
                if ignore_errors:
                    print(
                        f"Warning: Skipping malformed JSONL line {lineno} in {manifest_path}: {e}",
                        file=sys.stderr
                    )
                    continue
                raise


def load_completed_uuids(manifest_path: str) -> set:
    """
    Load UUIDs of already-processed blocks from manifest.

    Args:
        manifest_path: Path to manifest file

    Returns:
        Set of completed UUIDs
    """
    completed = set()
    for entry in iter_manifest_entries(manifest_path, ignore_errors=True):
        # Skip metadata entries
        if 'audited_at' in entry or entry.get('type') == 'meta':
            continue
        uuid = entry.get('uuid', '')
        if uuid:
            completed.add(uuid)

    return completed


def load_manifest_metadata(manifest_path: str):
    """
    Load metadata entry from a manifest JSONL file if present.

    Args:
        manifest_path: Path to manifest JSONL file

    Returns:
        Metadata entry dict or None
    """
    for entry in iter_manifest_entries(manifest_path, ignore_errors=True):
        if 'audited_at' in entry or entry.get('type') == 'meta':
            return entry
    return None


def load_existing_entries_with_block_idx(manifest_path: str, uuid_to_block_idx: dict) -> list:
    """
    Load existing manifest entries with block_idx looked up from UUID mapping.
    Used in resume mode to preserve previously processed entries.
    
    Args:
        manifest_path: Path to manifest JSONL file
        uuid_to_block_idx: Mapping from uuid to block index
    
    Returns:
        List of (block_idx, entry) tuples for existing entries
    """
    entries = []
    path = Path(manifest_path)
    
    if not path.exists():
        return entries

    for entry in iter_manifest_entries(manifest_path, ignore_errors=True):
        # Skip metadata entry
        if 'audited_at' in entry or entry.get('type') == 'meta':
            continue

        # Look up block_idx from uuid
        uuid = entry.get('uuid', '')
        if not uuid:
            continue
        block_idx = uuid_to_block_idx.get(uuid, -1)

        if block_idx >= 0:
            entries.append((block_idx, entry))
        else:
            # UUID not found in blocks - might be from different document
            print(f"Warning: UUID {uuid} not found in blocks, skipping",
                  file=sys.stderr)
    
    return entries


def rewrite_manifest_sorted(manifest_path: str, metadata: dict, results: list):
    """
    Sort results by block index and rule ID, then rewrite manifest file.
    Uses safe rewrite: backup old file, write new, delete backup on success.
    
    Args:
        manifest_path: Path to manifest JSONL file
        metadata: Metadata entry to write first (can be None)
        results: List of (block_idx, entry) tuples from audit
    """
    path = Path(manifest_path)
    backup_path = Path(str(manifest_path) + '.bak')
    
    # Sort by block_idx (primary sort key)
    results.sort(key=lambda x: x[0])
    
    # Sort violations within each entry by rule_id (secondary sort)
    for _, entry in results:
        if 'violations' in entry:
            entry['violations'].sort(key=lambda v: v.get('rule_id', ''))
    
    # Safe rewrite with backup
    try:
        # Step 1: Rename original to backup
        if path.exists():
            path.rename(backup_path)
        
        # Step 2: Write sorted content to new file
        with open(path, 'w', encoding='utf-8') as f:
            if metadata:
                f.write(json.dumps(metadata, ensure_ascii=False) + '\n')
            for _, entry in results:
                f.write(json.dumps(entry, ensure_ascii=False) + '\n')
        
        # Step 3: Delete backup on success
        if backup_path.exists():
            backup_path.unlink()
        
    except Exception as e:
        # Restore from backup if write failed
        if backup_path.exists() and not path.exists():
            backup_path.rename(path)
        raise e


async def process_block_async(
    block: dict,
    block_idx: int,
    total_blocks: int,
    system_prompt: str,
    model_name: str,
    use_gemini: bool,
    client,
    rule_category_map: dict,
    manifest_path: str,
    manifest_lock: asyncio.Lock,
    rate_limit: float,
    semaphore: asyncio.Semaphore,
    max_retries: int = DEFAULT_MAX_RETRIES,
    thinking_level: str = None,
    thinking_budget: int = None,
    reasoning_effort: str = None
) -> tuple:
    """
    Process a single block asynchronously with concurrency control and retry.

    Args:
        block: Text block to audit
        block_idx: Index of the block (0-based)
        total_blocks: Total number of blocks
        system_prompt: Cached system prompt
        model_name: LLM model name
        use_gemini: Whether to use Gemini (True) or OpenAI (False)
        client: LLM client instance
        rule_category_map: Mapping from rule_id to category
        manifest_path: Path to manifest file
        manifest_lock: asyncio.Lock for thread-safe manifest writing
        rate_limit: Seconds to wait between API calls
        semaphore: asyncio.Semaphore for concurrency control
        max_retries: Maximum number of retry attempts for transient errors

    Returns:
        Tuple of (block_idx, success, violation_count, heading, entry)
        entry is the manifest entry dict (None if failed)
    """
    async with semaphore:
        block_uuid = block.get('uuid', str(block_idx))
        heading = block.get('heading', 'Unknown')[:50]
        
        try:
            # Rate limiting (applied per concurrent task)
            if rate_limit > 0:
                await asyncio.sleep(rate_limit)
            
            # Call LLM with retry mechanism
            result = await audit_block_with_retry(
                block=block,
                system_prompt=system_prompt,
                model_name=model_name,
                client=client,
                use_gemini=use_gemini,
                max_retries=max_retries,
                block_idx=block_idx,
                total_blocks=total_blocks,
                thinking_level=thinking_level,
                thinking_budget=thinking_budget,
                reasoning_effort=reasoning_effort
            )

            # Get UUID range from block for injection into violations
            block_uuid_start = block.get('uuid', '')
            block_uuid_end = block.get('uuid_end', block_uuid_start)
            
            # Add category and UUID range to each violation based on rule_id
            violations_with_metadata = []
            for violation in result.get('violations', []):
                rule_id = violation.get('rule_id', '')
                category = rule_category_map.get(rule_id, 'other')
                violation_with_metadata = {
                    **violation,
                    "category": category,
                    "uuid": block_uuid_start,
                    "uuid_end": block_uuid_end
                }
                violations_with_metadata.append(violation_with_metadata)

            # Normalize is_violation based on actual violations
            has_violations = len(violations_with_metadata) > 0
            is_violation = has_violations

            # Build manifest entry
            entry = {
                "uuid": block_uuid,
                "uuid_end": block_uuid_end,
                "p_heading": block.get('heading', ''),
                "p_content": block.get('content', '') if isinstance(block.get('content'), str) else json.dumps(block.get('content', ''), ensure_ascii=False),
                "is_violation": is_violation,
                "violations": violations_with_metadata
            }

            # Save to manifest (thread-safe) - ensures resume capability
            await save_manifest_entry_async(manifest_path, entry, manifest_lock)

            violation_count = len(violations_with_metadata)
            return (block_idx, True, violation_count, heading, entry)

        except json.JSONDecodeError as e:
            print(f"[{block_idx+1}/{total_blocks}] Error: Failed to parse LLM response: {e}", file=sys.stderr)
            return (block_idx, False, 0, heading, None)
        except Exception as e:
            print(f"[{block_idx+1}/{total_blocks}] Error: {e}", file=sys.stderr)
            return (block_idx, False, 0, heading, None)


async def run_global_extraction_async(
    blocks: list,
    global_rules: list,
    use_gemini: bool,
    model_name: str,
    client,
    max_retries: int,
    rate_limit: float,
    workers: int,
    extraction_path: str = None,
    resume: bool = False,
    metadata: dict = None,
    thinking_level: str = None,
    thinking_budget: int = None,
    reasoning_effort: str = None,
    base_index: int = 0
) -> dict:
    """
    Extract global rule information from each block.
    
    Args:
        blocks: List of text blocks
        global_rules: List of global audit rules
        use_gemini: Whether to use Gemini
        model_name: LLM model name
        client: LLM client instance
        max_retries: Maximum retries for transient errors
        rate_limit: Seconds to wait between API calls
        workers: Number of parallel workers
        extraction_path: Path to extraction file (for persistence)
        resume: Whether to resume from previous run
        metadata: Document metadata (for writing to extraction file)
        thinking_level: Thinking level for Gemini 3 models
        thinking_budget: Thinking budget for Gemini 2.5 models
        reasoning_effort: Reasoning effort for OpenAI o-series models
        base_index: Base index offset for UUID fallback (default: 0)

    Returns:
        Dict mapping rule_id -> list of extracted items
    """
    system_prompt = build_global_extract_system_prompt(global_rules)
    semaphore = asyncio.Semaphore(workers)
    extraction_lock = asyncio.Lock()

    # Build rule metadata mapping (rule_id -> {topic, category})
    rule_meta_map = {
        rule.get("id", ""): {
            "topic": rule.get("topic", ""),
            "category": rule.get("category", "other")
        }
        for rule in global_rules
    }

    # Initialize rule_buckets and completed_uuids
    rule_buckets = {rule.get("id", ""): [] for rule in global_rules}
    completed_uuids = set()
    
    # Handle resume mode
    if extraction_path and resume and Path(extraction_path).exists():
        completed_uuids = load_completed_extraction_uuids(extraction_path)
        rule_buckets = load_extraction_buckets(extraction_path, global_rules)
        print(f"Resuming extraction: {len(completed_uuids)} blocks already extracted")
    elif extraction_path and not resume:
        # Initialize new extraction file with metadata
        extraction_metadata = {
            "type": "meta",
            "extraction_started_at": datetime.now().isoformat()
        }
        if metadata:
            extraction_metadata.update({
                "source_file": metadata.get("source_file", ""),
                "source_hash": metadata.get("source_hash", "")
            })
        with open(extraction_path, 'w', encoding='utf-8') as f:
            f.write(json.dumps(extraction_metadata, ensure_ascii=False) + '\n')

    # Filter out already-extracted blocks
    blocks_to_extract = [
        (idx, block) for idx, block in enumerate(blocks)
        if block.get('uuid', str(base_index + idx)) not in completed_uuids
    ]

    # Progress tracking
    completed_count = len(completed_uuids)  # Start from already completed
    total_entities = sum(len(items) for items in rule_buckets.values())  # Count entities already extracted
    progress_lock = asyncio.Lock()
    total_blocks = len(blocks)

    # Print header
    print(f"\nGlobal extraction: {len(global_rules)} rules, {total_blocks} blocks")
    if completed_uuids:
        print(f"Skipping {len(completed_uuids)} already extracted, processing {len(blocks_to_extract)} remaining")
    print("-" * 50)

    if not blocks_to_extract:
        print("All blocks already extracted!")
        print("-" * 50)
        entity_str = "entity" if total_entities == 1 else "entities"
        print(f"Extraction complete: {total_blocks} blocks, {total_entities} {entity_str} extracted")
        return rule_buckets

    async def extract_block(block_idx: int, block: dict):
        nonlocal completed_count, total_entities
        async with semaphore:
            if rate_limit > 0:
                await asyncio.sleep(rate_limit)
            
            result = await global_extract_with_retry(
                block=block,
                system_prompt=system_prompt,
                model_name=model_name,
                client=client,
                use_gemini=use_gemini,
                max_retries=max_retries,
                block_idx=block_idx,
                total_blocks=total_blocks,
                thinking_level=thinking_level,
                thinking_budget=thinking_budget,
                reasoning_effort=reasoning_effort
            )

            # Count entities extracted from this block
            entity_count = sum(
                len(r.get("extracted_results", []))
                for r in result.get("results", [])
            ) if result and isinstance(result, dict) else 0

            # Save to extraction file if path provided
            if extraction_path and result:
                # Enrich results with topic and category from rule metadata
                enriched_results = []
                for r in result.get("results", []):
                    rule_id = r.get("rule_id", "")
                    meta = rule_meta_map.get(rule_id, {})
                    enriched_r = {
                        "rule_id": rule_id,
                        "topic": meta.get("topic", ""),
                        "category": meta.get("category", "other"),
                        "extracted_results": r.get("extracted_results", [])
                    }
                    enriched_results.append(enriched_r)

                extraction_entry = {
                    "uuid": block.get("uuid", str(base_index + block_idx)),
                    "uuid_end": block.get("uuid_end", block.get("uuid", str(base_index + block_idx))),
                    "p_heading": block.get("heading", ""),
                    "results": enriched_results
                }
                await save_extraction_entry_async(extraction_path, extraction_entry, extraction_lock)

            # Progress output
            heading = block.get("heading", "Unknown")[:40]
            async with progress_lock:
                completed_count += 1
                total_entities += entity_count
                entity_str = "entity" if entity_count == 1 else "entities"
                print(f"[Extract {completed_count}/{total_blocks}] {heading}... {entity_count} {entity_str}")

            return block_idx, result

    tasks = [
        extract_block(idx, block)
        for idx, block in blocks_to_extract
    ]

    for coro in asyncio.as_completed(tasks):
        block_idx, result = await coro
        block = blocks[block_idx]
        if not result or not isinstance(result, dict):
            continue
        for rule_result in result.get("results", []):
            rule_id = rule_result.get("rule_id", "")
            if rule_id not in rule_buckets:
                continue
            for extracted in rule_result.get("extracted_results", []):
                item = {
                    "uuid": block.get("uuid", str(base_index + block_idx)),
                    "uuid_end": block.get("uuid_end", block.get("uuid", str(base_index + block_idx))),
                    "p_heading": block.get("heading", ""),
                    "entity": extracted.get("entity", ""),
                    "fields": extracted.get("fields", [])
                }
                rule_buckets[rule_id].append(item)

    # Print summary
    print("-" * 50)
    entity_str = "entity" if total_entities == 1 else "entities"
    print(f"Extraction complete: {total_blocks} blocks, {total_entities} {entity_str} extracted")

    return rule_buckets


async def run_global_verification_async(
    rule_buckets: dict,
    global_rules: list,
    use_gemini: bool,
    model_name: str,
    client,
    max_retries: int,
    thinking_level: str = None,
    thinking_budget: int = None,
    reasoning_effort: str = None,
    max_tokens: int = DEFAULT_GLOBAL_AUDIT_MAX_TOKENS
) -> list:
    """
    Verify global consistency per rule bucket.

    Returns:
        List of violation dicts
    """
    # Print header
    print(f"\nGlobal verification: {len(global_rules)} rules")
    print("-" * 50)

    violations = []
    for idx, rule in enumerate(global_rules):
        rule_id = rule.get("id", "")
        topic = rule.get("topic", "Unknown")[:30]
        items = rule_buckets.get(rule_id, [])

        if not items:
            print(f"[Verify {rule_id}] {topic} (0 items)... skipped")
            continue

        system_prompt = build_global_verify_system_prompt(rule)
        chunks = chunk_items_by_token_limit(rule, items, max_tokens)

        rule_violations = []
        for chunk in chunks:
            result = await global_verify_with_retry(
                rule=rule,
                items=chunk,
                system_prompt=system_prompt,
                model_name=model_name,
                client=client,
                use_gemini=use_gemini,
                max_retries=max_retries,
                rule_idx=idx,
                total_rules=len(global_rules),
                thinking_level=thinking_level,
                thinking_budget=thinking_budget,
                reasoning_effort=reasoning_effort
            )
            for v in result.get("violations", []):
                if not v.get("rule_id"):
                    v["rule_id"] = rule_id
                rule_violations.append(v)

        # Progress output for this rule
        chunk_str = "chunk" if len(chunks) == 1 else "chunks"
        viol_str = "violation" if len(rule_violations) == 1 else "violations"
        print(f"[Verify {rule_id}] {topic} ({len(items)} items, {len(chunks)} {chunk_str})... {len(rule_violations)} {viol_str}")

        violations.extend(rule_violations)

    # Print summary
    print("-" * 50)
    viol_str = "violation" if len(violations) == 1 else "violations"
    print(f"Verification complete: {len(violations)} {viol_str} found")

    return violations


def merge_global_violations(
    entries_with_idx: list,
    global_violations: list,
    blocks: list,
    uuid_to_block_idx: dict,
    rule_category_map: dict
) -> list:
    """
    Merge global violations into existing manifest entries.

    Args:
        entries_with_idx: List of (block_idx, entry) tuples
        global_violations: List of violation dicts from global verification
        blocks: Original blocks list
        uuid_to_block_idx: Mapping from uuid to block index
        rule_category_map: Mapping from rule_id to category

    Returns:
        Updated list of (block_idx, entry) tuples
    """
    entry_by_uuid = {}
    for block_idx, entry in entries_with_idx:
        entry_by_uuid[entry.get("uuid", "")] = (block_idx, entry)

    for violation in global_violations:
        uuid = violation.get("uuid", "")
        if not uuid:
            print("Warning: Global violation missing uuid, skipping.", file=sys.stderr)
            continue
        block_idx = uuid_to_block_idx.get(uuid)
        if block_idx is None:
            print(f"Warning: Global violation uuid not found in blocks: {uuid}", file=sys.stderr)
            continue

        entry_tuple = entry_by_uuid.get(uuid)
        if entry_tuple:
            _, entry = entry_tuple
        else:
            block = blocks[block_idx]
            entry = {
                "uuid": block.get("uuid", uuid),
                "uuid_end": block.get("uuid_end", block.get("uuid", uuid)),
                "p_heading": block.get("heading", ""),
                "p_content": block.get("content", "") if isinstance(block.get("content"), str)
                else json.dumps(block.get("content", ""), ensure_ascii=False),
                "is_violation": False,
                "violations": []
            }
            entry_by_uuid[uuid] = (block_idx, entry)

        # Deduplicate by rule_id + violation_text + uuid_end
        existing_keys = {
            (v.get("rule_id", ""), v.get("violation_text", ""), v.get("uuid_end", ""))
            for v in entry.get("violations", [])
        }

        rule_id = violation.get("rule_id", "")
        violation_with_meta = {
            **violation,
            "category": rule_category_map.get(rule_id, "other")
        }
        key = (violation_with_meta.get("rule_id", ""),
               violation_with_meta.get("violation_text", ""),
               violation_with_meta.get("uuid_end", ""))
        if key not in existing_keys:
            entry.setdefault("violations", []).append(violation_with_meta)
        entry["is_violation"] = len(entry.get("violations", [])) > 0

    return list(entry_by_uuid.values())


async def run_audit_async(args, blocks, rules, metadata, use_gemini, model_name, provider_name, client, completed_uuids,
                         thinking_level=None, thinking_budget=None, reasoning_effort=None):
    """
    Run the audit process asynchronously with parallel block processing.

    Args:
        args: Parsed command-line arguments
        blocks: List of text blocks to audit
        rules: List of audit rules
        metadata: Document metadata
        use_gemini: Whether to use Gemini
        model_name: LLM model name
        provider_name: LLM provider name (e.g., "Google Gemini", "OpenAI")
        client: LLM client instance
        completed_uuids: Set of already-processed UUIDs
        thinking_level: Thinking level for Gemini 3 models (resolved in main)
        thinking_budget: Thinking budget for Gemini 2.5 models (resolved in main)
        reasoning_effort: Reasoning effort for OpenAI o-series models (resolved in main)
    """
    # Build rule category mapping
    rule_category_map = build_rule_category_map(rules)

    # Build system prompt once (will be cached by LLM)
    system_prompt = build_block_audit_system_prompt(rules)

    # Build uuid → block_idx mapping for all blocks
    uuid_to_block_idx = {}
    for idx, block in enumerate(blocks):
        uuid = block.get('uuid', str(idx))
        uuid_to_block_idx[uuid] = idx

    # In resume mode, load existing entries to preserve them during rewrite
    all_results = []  # Collect results for final sorting
    if completed_uuids:
        all_results = load_existing_entries_with_block_idx(args.output, uuid_to_block_idx)
        print(f"Loaded {len(all_results)} existing entries for merge")

    # Determine block range
    start_idx = args.start_block
    end_idx = args.end_block if args.end_block >= 0 else len(blocks) - 1
    blocks_to_process = blocks[start_idx:end_idx + 1]

    # Filter out already-processed blocks
    blocks_with_indices = [
        (start_idx + i, block) 
        for i, block in enumerate(blocks_to_process)
        if block.get('uuid', str(start_idx + i)) not in completed_uuids
    ]

    if not blocks_with_indices:
        print("\nAll blocks already processed!")
        return

    print(f"\nProcessing: blocks {start_idx}-{end_idx} ({len(blocks_with_indices)} pending), {args.workers} workers")
    print(f"Output: {args.output}")
    print("-" * 50)

    # Create concurrency controls
    semaphore = asyncio.Semaphore(args.workers)
    manifest_lock = asyncio.Lock()

    # Configuration is already resolved and validated in main()
    # No need to re-parse or validate here

    # Create tasks for all blocks
    tasks = [
        process_block_async(
            block=block,
            block_idx=block_idx,
            total_blocks=len(blocks),
            system_prompt=system_prompt,
            model_name=model_name,
            use_gemini=use_gemini,
            client=client,
            rule_category_map=rule_category_map,
            manifest_path=args.output,
            manifest_lock=manifest_lock,
            rate_limit=args.rate_limit,
            semaphore=semaphore,
            max_retries=args.max_retries,
            thinking_level=thinking_level,
            thinking_budget=thinking_budget,
            reasoning_effort=reasoning_effort
        )
        for block_idx, block in blocks_with_indices
    ]

    # Process all blocks in parallel with progress reporting
    total_violations = 0
    blocks_processed = 0
    blocks_failed = 0
    # Note: all_results already initialized above (may contain existing entries in resume mode)

    # Use asyncio.as_completed for real-time progress updates
    for coro in asyncio.as_completed(tasks):
        block_idx, success, violation_count, heading, entry = await coro
        
        if success:
            blocks_processed += 1
            total_violations += violation_count
            all_results.append((block_idx, entry))  # Collect for sorting
            if violation_count > 0:
                print(f"[{block_idx+1}/{len(blocks)}] {heading}... Found {violation_count} violation(s)")
            else:
                print(f"[{block_idx+1}/{len(blocks)}] {heading}... OK")
        else:
            blocks_failed += 1

    # Rewrite manifest with sorted results
    if all_results:
        # Prepare metadata for rewrite
        audit_metadata = None
        if metadata:
            audit_metadata = {
                **metadata,
                "llm_provider": provider_name,
                "llm_model": model_name,
                "audited_at": datetime.now().isoformat()
            }
            # Add thinking/reasoning config if set
            if thinking_level:
                audit_metadata["thinking_level"] = thinking_level
            if thinking_budget is not None:
                audit_metadata["thinking_budget"] = thinking_budget
            if reasoning_effort:
                audit_metadata["reasoning_effort"] = reasoning_effort
        
        print("Sorting and rewriting manifest...")
        rewrite_manifest_sorted(args.output, audit_metadata, all_results)

    # Summary
    print("\n" + "=" * 50)
    print("Audit Complete")
    print(f"Blocks processed: {blocks_processed}")
    print(f"Blocks failed: {blocks_failed}")
    print(f"Total violations: {total_violations}")
    print(f"Manifest saved to: {args.output} (sorted by block order)")


async def run_full_audit_async(
    args,
    blocks: list,
    block_rules: list,
    global_rules: list,
    metadata: dict,
    use_gemini: bool,
    model_name: str,
    provider_name: str,
    client,
    completed_uuids: set,
    thinking_level: str = None,
    thinking_budget: int = None,
    reasoning_effort: str = None
):
    """
    Unified async entry point for both block-level and global audits.

    This function runs both audit phases in sequence within a single event loop
    for better efficiency.
    """
    # Phase 1: Block-level audit
    if block_rules:
        await run_audit_async(
            args, blocks, block_rules, metadata, use_gemini, model_name,
            provider_name, client, completed_uuids,
            thinking_level=thinking_level,
            thinking_budget=thinking_budget,
            reasoning_effort=reasoning_effort
        )
    else:
        print("\nNo block rules provided; skipping block-level audit.")

    # Phase 2: Global extraction + verification
    if global_rules:
        max_tokens = get_global_audit_max_tokens()

        start_idx = args.start_block
        end_idx = args.end_block if args.end_block >= 0 else len(blocks) - 1
        blocks_for_global = blocks[start_idx:end_idx + 1]

        # Derive extraction file path and handle resume
        extraction_path = derive_extraction_path(args.output)
        
        rule_buckets = await run_global_extraction_async(
            blocks=blocks_for_global,
            global_rules=global_rules,
            use_gemini=use_gemini,
            model_name=model_name,
            client=client,
            max_retries=args.max_retries,
            rate_limit=args.rate_limit,
            workers=args.workers,
            extraction_path=extraction_path,
            resume=args.resume,
            metadata=metadata,
            thinking_level=thinking_level,
            thinking_budget=thinking_budget,
            reasoning_effort=reasoning_effort,
            base_index=start_idx
        )

        global_violations = await run_global_verification_async(
            rule_buckets=rule_buckets,
            global_rules=global_rules,
            use_gemini=use_gemini,
            model_name=model_name,
            client=client,
            max_retries=args.max_retries,
            thinking_level=thinking_level,
            thinking_budget=thinking_budget,
            reasoning_effort=reasoning_effort,
            max_tokens=max_tokens
        )

        if global_violations:
            uuid_to_block_idx = {
                block.get('uuid', str(idx)): idx for idx, block in enumerate(blocks)
            }
            existing_entries = load_existing_entries_with_block_idx(args.output, uuid_to_block_idx)
            global_rule_category_map = build_rule_category_map(global_rules)
            merged_entries = merge_global_violations(
                entries_with_idx=existing_entries,
                global_violations=global_violations,
                blocks=blocks,
                uuid_to_block_idx=uuid_to_block_idx,
                rule_category_map=global_rule_category_map
            )

            audit_metadata = load_manifest_metadata(args.output)
            if not audit_metadata and metadata:
                audit_metadata = {
                    **metadata,
                    "llm_provider": provider_name,
                    "llm_model": model_name,
                    "audited_at": datetime.now().isoformat()
                }
                if thinking_level:
                    audit_metadata["thinking_level"] = thinking_level
                if thinking_budget is not None:
                    audit_metadata["thinking_budget"] = thinking_budget
                if reasoning_effort:
                    audit_metadata["reasoning_effort"] = reasoning_effort

            rewrite_manifest_sorted(args.output, audit_metadata, merged_entries)
            print(f"Global violations merged: {len(global_violations)}")
        else:
            print("Global audit completed: no violations found.")
    else:
        print("\nNo global rules provided; skipping global audit.")


def main():
    parser = argparse.ArgumentParser(
        description="Run LLM-based audit on document text blocks"
    )
    parser.add_argument(
        "--document", "-d",
        type=str,
        required=True,
        help="Path to document blocks file (JSONL or JSON)"
    )
    parser.add_argument(
        "--rules", "-r",
        type=str,
        action='extend',
        nargs='+',
        required=True,
        help="Path to audit rules JSON file(s). Multiple files will be merged."
    )
    parser.add_argument(
        "--output", "-o",
        type=str,
        default="manifest.jsonl",
        help="Output manifest file path (default: manifest.jsonl)"
    )
    parser.add_argument(
        "--provider",
        type=str,
        choices=["auto", "gemini", "openai"],
        default="auto",
        help="Force LLM provider: gemini, openai, or auto (default: auto)"
    )
    parser.add_argument(
        "--model",
        type=str,
        default="auto",
        help="LLM model to use: gemini-3-flash, gpt-5.2, or auto (default: auto)"
    )
    parser.add_argument(
        "--workers",
        type=int,
        default=MAX_PARALLEL_WORKERS,
        help=f"Number of parallel workers (default: {MAX_PARALLEL_WORKERS})"
    )
    parser.add_argument(
        "--rate-limit",
        type=float,
        default=0.05,
        help="Seconds to wait between API calls per worker (default: 0.05)"
    )
    parser.add_argument(
        "--start-block",
        type=int,
        default=0,
        help="Start from this block index (for resuming)"
    )
    parser.add_argument(
        "--end-block",
        type=int,
        default=-1,
        help="End at this block index (inclusive, default: all blocks)"
    )
    parser.add_argument(
        "--resume",
        action="store_true",
        help="Resume from previous run (skip already-processed blocks)"
    )
    parser.add_argument(
        "--max-retries",
        type=int,
        default=DEFAULT_MAX_RETRIES,
        help=f"Maximum retries for transient errors (default: {DEFAULT_MAX_RETRIES})"
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Print prompts without calling LLM"
    )
    parser.add_argument(
        "--thinking-level",
        type=str,
        choices=["minimal", "low", "medium", "high"],
        default=None,
        help="Gemini 3 thinking level: minimal, low, medium, high (env: GEMINI_THINKING_LEVEL)"
    )
    parser.add_argument(
        "--thinking-budget",
        type=int,
        default=None,
        help="Gemini 2.5 thinking token budget, 0 to disable (env: GEMINI_THINKING_BUDGET)"
    )
    parser.add_argument(
        "--reasoning-effort",
        type=str,
        choices=["low", "medium", "high"],
        default=None,
        help="OpenAI o-series reasoning effort: low, medium, high (env: OPENAI_REASONING_EFFORT)"
    )

    args = parser.parse_args()


    # Validate input format (JSON/JSONL blocks only)
    doc_path = Path(args.document)
    if doc_path.suffix.lower() not in {'.json', '.jsonl'}:
        print("Error: --document must be a JSON/JSONL blocks file (use parse_document.py first).", file=sys.stderr)
        sys.exit(1)

    # Check for LLM availability
    if not args.dry_run:
        if not HAS_GEMINI and not HAS_OPENAI:
            print("Error: No LLM library installed.", file=sys.stderr)
            print("Install one of:", file=sys.stderr)
            print("  pip install google-genai", file=sys.stderr)
            print("  pip install openai", file=sys.stderr)
            sys.exit(1)

    # Determine which model to use
    use_gemini = False
    client = None
    model_name = args.model

    # Check if Vertex AI mode is enabled
    use_vertex = is_vertex_ai_mode()

    # Validate credentials when provider is explicitly specified
    if args.provider == "gemini":
        if not HAS_GEMINI:
            print("Error: google-genai library not installed", file=sys.stderr)
            print("Install with: pip install google-genai", file=sys.stderr)
            sys.exit(1)
        
        # Check credentials based on Vertex AI mode
        if use_vertex:
            if not os.getenv("GOOGLE_CLOUD_PROJECT"):
                print("Error: --provider=gemini requires GOOGLE_CLOUD_PROJECT for Vertex AI mode", file=sys.stderr)
                print("Hint: Set GOOGLE_GENAI_USE_VERTEXAI=false to use AI Studio mode", file=sys.stderr)
                sys.exit(1)
        else:
            if not os.getenv("GOOGLE_API_KEY"):
                print("Error: --provider=gemini requires GOOGLE_API_KEY for AI Studio mode", file=sys.stderr)
                print("Hint: Set GOOGLE_GENAI_USE_VERTEXAI=true to use Vertex AI mode", file=sys.stderr)
                sys.exit(1)
        
        # Force Gemini usage
        use_gemini = True
        if model_name == "auto":
            model_name = os.getenv("DOC_AUDIT_GEMINI_MODEL", "gemini-2.5-flash")
        try:
            client = create_gemini_client(use_async=True)
        except ValueError as e:
            print(f"Error: {e}", file=sys.stderr)
            sys.exit(1)
    
    elif args.provider == "openai":
        if not HAS_OPENAI:
            print("Error: openai library not installed", file=sys.stderr)
            print("Install with: pip install openai", file=sys.stderr)
            sys.exit(1)
        
        if not os.getenv("OPENAI_API_KEY"):
            print("Error: --provider=openai requires OPENAI_API_KEY", file=sys.stderr)
            sys.exit(1)
        
        # Force OpenAI usage
        use_gemini = False
        if model_name == "auto":
            model_name = os.getenv("DOC_AUDIT_OPENAI_MODEL", "gpt-4.1")
        client = create_openai_client(use_async=True)
    
    elif model_name == "auto":
        # Auto-detect: check Gemini credentials first (AI Studio or Vertex AI)
        gemini_available = False
        if HAS_GEMINI:
            if use_vertex:
                # Vertex AI mode explicitly enabled - require project configuration
                # Do NOT silently fall back to OpenAI as user may have compliance/data residency requirements
                if not os.getenv("GOOGLE_CLOUD_PROJECT"):
                    print("Error: Vertex AI mode is enabled (GOOGLE_GENAI_USE_VERTEXAI=true) "
                          "but GOOGLE_CLOUD_PROJECT is not set.", file=sys.stderr)
                    print("Either:", file=sys.stderr)
                    print("  1. Set GOOGLE_CLOUD_PROJECT to your GCP project ID", file=sys.stderr)
                    print("  2. Unset GOOGLE_GENAI_USE_VERTEXAI to use AI Studio mode", file=sys.stderr)
                    sys.exit(1)
                gemini_available = True
            else:
                # AI Studio mode - check for API key
                gemini_available = bool(os.getenv("GOOGLE_API_KEY"))
        
        if gemini_available:
            use_gemini = True
            model_name = os.getenv("DOC_AUDIT_GEMINI_MODEL", "gemini-2.5-flash")
            try:
                client = create_gemini_client(use_async=True)
            except ValueError as e:
                print(f"Error: {e}", file=sys.stderr)
                sys.exit(1)
        elif HAS_OPENAI and os.getenv("OPENAI_API_KEY"):
            model_name = os.getenv("DOC_AUDIT_OPENAI_MODEL", "gpt-4.1")
            client = create_openai_client(use_async=True)
        else:
            print("Error: No LLM credentials found.", file=sys.stderr)
            print("For AI Studio: Set GOOGLE_API_KEY", file=sys.stderr)
            print("For Vertex AI: Set GOOGLE_GENAI_USE_VERTEXAI=true and GOOGLE_CLOUD_PROJECT", file=sys.stderr)
            print("For OpenAI: Set OPENAI_API_KEY", file=sys.stderr)
            sys.exit(1)
    elif "gemini" in model_name.lower():
        if not HAS_GEMINI:
            print("Error: google-genai not installed", file=sys.stderr)
            sys.exit(1)
        
        # Validate credentials based on mode
        if use_vertex:
            if not os.getenv("GOOGLE_CLOUD_PROJECT"):
                print("Error: GOOGLE_CLOUD_PROJECT not set for Vertex AI mode", file=sys.stderr)
                print("Hint: Set GOOGLE_GENAI_USE_VERTEXAI=false to use AI Studio mode", file=sys.stderr)
                sys.exit(1)
        else:
            if not os.getenv("GOOGLE_API_KEY"):
                print("Error: GOOGLE_API_KEY not set for AI Studio mode", file=sys.stderr)
                print("Hint: Set GOOGLE_GENAI_USE_VERTEXAI=true for Vertex AI mode", file=sys.stderr)
                sys.exit(1)
        
        use_gemini = True
        try:
            client = create_gemini_client(use_async=True)
        except ValueError as e:
            print(f"Error: {e}", file=sys.stderr)
            sys.exit(1)
    else:
        # Treat all other models as OpenAI (gpt-4.1, o1-mini, o3-mini, etc.)
        if not HAS_OPENAI:
            print("Error: openai not installed", file=sys.stderr)
            sys.exit(1)
        if not os.getenv("OPENAI_API_KEY"):
            print("Error: OPENAI_API_KEY not set", file=sys.stderr)
            sys.exit(1)
        client = create_openai_client(use_async=True)

    # Determine and print LLM provider name
    provider_name = get_gemini_provider_name() if use_gemini else get_openai_provider_name()
    
    # Display LLM configuration
    print(f"\nLLM: {provider_name} / {model_name}")
    
    # Resolve thinking/reasoning configuration (command line args > environment variables)
    # CLI parameters override and clear environment-derived values
    thinking_level = None
    thinking_budget = None
    reasoning_effort = None
    
    if use_gemini:
        # CLI takes precedence - if CLI provides one type, ignore env for the other type
        if args.thinking_level is not None:
            # CLI --thinking-level provided, use it exclusively
            thinking_level = args.thinking_level
            thinking_budget = None  # Clear any env-derived budget
        elif args.thinking_budget is not None:
            # CLI --thinking-budget provided, use it exclusively
            thinking_budget = args.thinking_budget
            thinking_level = None  # Clear any env-derived level
        else:
            # Neither CLI arg provided, fall back to env vars
            thinking_level = os.getenv("GEMINI_THINKING_LEVEL")
            
            # Parse thinking_budget from env with error handling
            thinking_budget_str = os.getenv("GEMINI_THINKING_BUDGET")
            if thinking_budget_str:
                try:
                    thinking_budget = int(thinking_budget_str)
                except ValueError:
                    print(f"Error: GEMINI_THINKING_BUDGET must be an integer, got: '{thinking_budget_str}'", file=sys.stderr)
                    print("Set to a valid integer (e.g., 1024) or unset the variable.", file=sys.stderr)
                    sys.exit(1)
            
            # Validate only when both come from env (conflict situation)
            if thinking_level and thinking_budget is not None:
                validate_thinking_config(
                    thinking_level=thinking_level,
                    thinking_budget=thinking_budget,
                    thinking_level_source="env GEMINI_THINKING_LEVEL",
                    thinking_budget_source="env GEMINI_THINKING_BUDGET"
                )
        
        # Display thinking configuration if set
        if thinking_level:
            print(f"Thinking: {thinking_level.upper()}")
        elif thinking_budget is not None:
            print(f"Thinking: Budget {thinking_budget} tokens")
    else:
        # For OpenAI, only resolve reasoning_effort
        reasoning_effort = args.reasoning_effort or os.getenv("OPENAI_REASONING_EFFORT")
        
        # Display reasoning configuration if set
        if reasoning_effort:
            if is_openai_reasoning_model(model_name):
                print(f"Reasoning: {reasoning_effort.upper()}")
            else:
                print(f"Note: reasoning_effort set but ignored (model {model_name} does not support it)")
    
    # Load inputs
    print(f"\nLoading: {args.document}")
    metadata, blocks = load_blocks(args.document)
    print(f"  → {len(blocks)} blocks")
    if metadata:
        print(f"  Source: {metadata.get('source_file', 'Unknown')}")
        print(f"  Hash: {metadata.get('source_hash', 'Unknown')[:16]}...")

    # Load and merge rules from multiple files
    print(f"Rules: {len(args.rules)} file(s)")
    for rule_file in args.rules:
        print(f"  - {rule_file}")
    rules = merge_rules(args.rules)
    print(f"  → {len(rules)} rules total")

    block_rules = [r for r in rules if r.get("type", "block") == "block"]
    global_rules = [r for r in rules if r.get("type") == "global"]
    print(f"  → {len(block_rules)} block rules, {len(global_rules)} global rules")

    # Handle resume
    completed_uuids = set()
    output_path = Path(args.output)
    backup_path = Path(str(args.output) + '.bak')
    if args.resume:
        if output_path.exists():
            completed_uuids = load_completed_uuids(output_path)
            print(f"Resuming: {len(completed_uuids)} blocks already processed")
        elif backup_path.exists():
            print(f"Resume requested but output missing; found backup: {backup_path}")
            completed_uuids = load_completed_uuids(backup_path)

            start_idx = args.start_block
            end_idx = args.end_block if args.end_block >= 0 else len(blocks) - 1
            blocks_to_process = blocks[start_idx:end_idx + 1]
            target_uuids = {
                block.get('uuid', str(start_idx + i))
                for i, block in enumerate(blocks_to_process)
            }

            if target_uuids and target_uuids.issubset(completed_uuids):
                print("Backup appears complete for the requested range. Restoring sorted manifest and exiting.")
                uuid_to_block_idx = {
                    block.get('uuid', str(idx)): idx for idx, block in enumerate(blocks)
                }
                existing_entries = load_existing_entries_with_block_idx(backup_path, uuid_to_block_idx)
                manifest_metadata = load_manifest_metadata(backup_path)
                rewrite_manifest_sorted(args.output, manifest_metadata, existing_entries)
                # Clean up backup after successful restore
                if backup_path.exists():
                    backup_path.unlink()
                    print(f"Cleaned up backup file: {backup_path}")
                print(f"Manifest restored from backup: {args.output}")
                return

            print("Backup is incomplete for the requested range. Continuing resume from backup.")
            try:
                backup_path.rename(output_path)
            except OSError as e:
                print(f"Warning: Could not rename backup to output ({e}); copying instead.", file=sys.stderr)
                with open(backup_path, 'r', encoding='utf-8') as src, \
                        open(output_path, 'w', encoding='utf-8') as dst:
                    dst.write(src.read())
                # Clean up backup after successful copy
                if backup_path.exists():
                    backup_path.unlink()
            completed_uuids = load_completed_uuids(output_path)
            print(f"Resuming: {len(completed_uuids)} blocks already processed")
        else:
            print("Resume requested but no existing output or backup found; starting fresh.")
    else:
        # Write metadata as first line for new manifest
        if metadata:
            audit_metadata = {
                **metadata,
                "llm_provider": provider_name,
                "llm_model": model_name,
                "audited_at": datetime.now().isoformat()
            }
            # Add thinking/reasoning config if set
            if thinking_level:
                audit_metadata["thinking_level"] = thinking_level
            if thinking_budget is not None:
                audit_metadata["thinking_budget"] = thinking_budget
            if reasoning_effort:
                audit_metadata["reasoning_effort"] = reasoning_effort
            
            with open(args.output, 'w', encoding='utf-8') as f:
                f.write(json.dumps(audit_metadata, ensure_ascii=False) + '\n')
            print(f"Created new manifest with source file metadata")

    # Handle dry-run mode (no async needed)
    if args.dry_run:
        if not block_rules:
            print("Dry-run: No block rules provided; skipping block audit prompts.")
        system_prompt = build_block_audit_system_prompt(block_rules)
        start_idx = args.start_block
        end_idx = args.end_block if args.end_block >= 0 else len(blocks) - 1
        blocks_to_process = blocks[start_idx:end_idx + 1]
        
        if block_rules:
            for i, block in enumerate(blocks_to_process):
                block_idx = start_idx + i
                block_uuid = block.get('uuid', str(block_idx))

                if block_uuid in completed_uuids:
                    print(f"[{block_idx+1}/{len(blocks)}] Skipping (already processed)")
                    continue

                print(f"[{block_idx+1}/{len(blocks)}] Auditing: {block.get('heading', 'Unknown')[:50]}...")
                user_prompt = build_block_audit_user_prompt(block)
                print(f"\n--- System Prompt ---\n{system_prompt[:300]}...\n")
                print(f"--- User Prompt ---\n{user_prompt[:300]}...")

        if global_rules:
            print("Dry-run: Global rules detected. Global extraction/verification prompts are skipped.")
        return

    # Run full audit (block-level + global) in a single event loop
    asyncio.run(run_full_audit_async(
        args=args,
        blocks=blocks,
        block_rules=block_rules,
        global_rules=global_rules,
        metadata=metadata,
        use_gemini=use_gemini,
        model_name=model_name,
        provider_name=provider_name,
        client=client,
        completed_uuids=completed_uuids,
        thinking_level=thinking_level,
        thinking_budget=thinking_budget,
        reasoning_effort=reasoning_effort
    ))


if __name__ == "__main__":
    main()
