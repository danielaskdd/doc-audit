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
import sys
from pathlib import Path

# Maximum number of concurrent LLM API calls
MAX_PARALLEL_WORKERS = 4

# Try to import LLM libraries
HAS_GEMINI = False
HAS_OPENAI = False

try:
    from google import genai
    from google.genai import types
    HAS_GEMINI = True
except ImportError:
    pass

try:
    import openai
    HAS_OPENAI = True
except ImportError:
    pass


# JSON Schema for LLM structured output
AUDIT_RESULT_SCHEMA = {
    "type": "object",
    "properties": {
        "is_violation": {
            "type": "boolean",
            "description": "Whether any violations were found"
        },
        "violations": {
            "type": "array",
            "description": "List of violations found",
            "items": {
                "type": "object",
                "properties": {
                    "rule_id": {
                        "type": "string",
                        "description": "ID of the violated rule (e.g., R001)"
                    },
                    "violation_text": {
                        "type": "string",
                        "description": "The problematic text directly verbatim quote from the source content, and not span multiple cells"
                    },
                    "violation_reason": {
                        "type": "string",
                        "description": "Explanation of why this violates the rule"
                    },
                    "fix_action": {
                        "type": "string",
                        "enum": ["delete", "replace", "manual"],
                        "description": "Action type: delete removes the text, replace substitutes it, manual requires human review"
                    },
                    "revised_text": {
                        "type": "string",
                        "description": "For replace: complete replacement text. For delete: empty string. For manual: additional guidance for human reviewer"
                    }
                },
                "required": ["rule_id", "violation_text", "violation_reason", "fix_action", "revised_text"]
            }
        }
    },
    "required": ["is_violation", "violations"]
}


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
        return data
    elif 'rules' in data:
        return data['rules']
    else:
        raise ValueError(f"Unknown rules format in {file_path}")


def build_rule_category_map(rules: list) -> dict:
    """
    Build a mapping from rule_id to category.

    Args:
        rules: List of rule dictionaries

    Returns:
        Dictionary mapping rule_id to category
    """
    return {rule['id']: rule.get('category', 'other') for rule in rules}


def format_block_for_prompt(block: dict) -> str:
    """
    Format a text block for inclusion in the audit prompt.

    Args:
        block: Block dictionary with heading, content, type

    Returns:
        Formatted string
    """
    heading = block.get('heading', 'Unknown')
    content = block.get('content', '')
    block_type = block.get('type', 'text')
    parent_headings = block.get('parent_headings', [])

    ### Context format
    # Context hierarchy: 1  header1 → 1.2  header2 → 1.2.2  header3
    context = ""
    if parent_headings:
        context = f"Context hierarchy: {' → '.join(parent_headings)}\n"

    if block_type == 'table':
        # Format table as readable text
        if isinstance(content, list):
            rows = []
            for row in content:
                rows.append(" | ".join(str(cell) for cell in row))
            content = "\n".join(rows)

    return f"""Section: {heading}
{context}

Content:
{content}"""


def format_rules_for_prompt(rules: list) -> str:
    """
    Format audit rules for inclusion in the prompt.

    Args:
        rules: List of rule dictionaries

    Returns:
        Formatted string
    """
    lines = ["Audit Rules:"]
    for rule in rules:
        severity = rule.get('severity', 'medium').upper()
        lines.append(f"- [{rule['id']}] ({severity}) {rule['description']}")

    return "\n".join(lines)


def build_system_prompt(rules: list) -> str:
    """
    Build the system prompt containing static instructions and rules.
    This can be cached by the LLM across multiple block audits.

    Args:
        rules: Audit rules to apply

    Returns:
        System prompt string
    """
    # Get output language from environment variable
    output_language = os.getenv("AUDIT_LANGUAGE", "Chinese")
    
    rules_text = format_rules_for_prompt(rules)

    system_prompt = f"""You are a professional document auditor. Your task is to analyze text blocks and check for violations of audit rules.

{rules_text}

---

Instructions:
1. Check if the provided text block violates ANY of the rules above
2. Report each violation as a separate item. Do not merge multiple instances of the same violation category into one entry.
3. For each violation found, provide:
   - The rule ID that was violated
   - The violation text with enough surrounding context for unique string matching
   - Why it's a violation
   - The fix action: "delete", "replace", or "manual"
   - The revised text based on fix_action

violation_text guidelines:
- The extracted text must be a direct verbatim quote from the source content, include line breaks, tabs, and other whitespace characters
- Do not use ellipses to replace or omit any part of the original text
- Exclude chapter/heading numbers, list markers, and bullet points from the violation_text
- If the violating content is excessively long (e.g., spanning multiple sentences), extract only the leading portion, ensuring it is sufficient to uniquely locate via string search
- If an entire section is in violation, select the corresponding heading as the violation_text (excluding `Section:` and the following heading number) 
- For violations spanning multiple table cells, select text from one of the most relevant cell only; do not consolidate multiple cells into a single violation_text entry

fix_action guidelines:
- "delete": Use when the problematic text should be completely removed
- "replace": Use when the text can be corrected with a specific replacement
- "manual": Use when the fix requires human judgment or complex restructuring

revised_text guidelines:
- For "delete": Set to empty string ""
- For "replace": Provide the complete replacement text that can directly substitute violation_text
- For "manual": Provide guidance for the human reviewer

If the violation_text is truncated due to excessive length or fails to achieve an exact match with the source material, the fix_action must be set to "manual"

Return your analysis as a JSON object with this structure:
{{
  "is_violation": true/false,
  "violations": [
    {{
      "rule_id": "R001",
      "violation_text": "the specific problematic text with sufficient context",
      "violation_reason": "explanation of why this violates the rule written in {output_language}",
      "fix_action": "delete|replace|manual",
      "revised_text": "corrected text in original language if fix_action is 'replace', otherwise guidance for human reviewer written in {output_language}"
    }}
  ]
}}

If there are no violations, return:
{{
  "is_violation": false,
  "violations": []
}}

Return ONLY the JSON object, no other text."""

    return system_prompt


def build_user_prompt(block: dict) -> str:
    """
    Build the user prompt containing the dynamic block content to audit.

    Args:
        block: Text block to audit

    Returns:
        User prompt string
    """
    block_text = format_block_for_prompt(block)
    return f"""Analyze the following text block for rule violations:

{block_text}"""


async def audit_block_gemini_async(block: dict, system_prompt: str, model_name: str, async_client) -> dict:
    """
    Audit a text block using Google Gemini with strict JSON mode (async version).

    Args:
        block: Text block to audit
        system_prompt: Cached system prompt with rules and instructions
        model_name: Gemini model to use
        async_client: Gemini async client instance (client.aio)

    Returns:
        Audit result dictionary
    """
    user_prompt = build_user_prompt(block)

    response = await async_client.models.generate_content(
        model=model_name,
        contents=user_prompt,
        config=types.GenerateContentConfig(
            system_instruction=system_prompt,
            response_mime_type="application/json",
            response_schema=AUDIT_RESULT_SCHEMA
        )
    )
    
    # With structured output, response is guaranteed to be valid JSON
    result = json.loads(response.text)
    return result


async def audit_block_openai_async(block: dict, system_prompt: str, model_name: str, client) -> dict:
    """
    Audit a text block using OpenAI with strict JSON mode (async version).

    Args:
        block: Text block to audit
        system_prompt: Cached system prompt with rules and instructions
        model_name: OpenAI model to use
        client: AsyncOpenAI client instance

    Returns:
        Audit result dictionary
    """
    user_prompt = build_user_prompt(block)

    response = await client.chat.completions.create(
        model=model_name,
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt}
        ],
        temperature=0.2,
        response_format={
            "type": "json_schema",
            "json_schema": {
                "name": "audit_result",
                "strict": True,
                "schema": AUDIT_RESULT_SCHEMA
            }
        }
    )

    # With structured output, response is guaranteed to be valid JSON
    result = json.loads(response.choices[0].message.content)
    return result


def audit_block_gemini(block: dict, system_prompt: str, model_name: str = None, client = None) -> dict:
    """
    Audit a text block using Google Gemini with strict JSON mode (sync version).

    Args:
        block: Text block to audit
        system_prompt: Cached system prompt with rules and instructions
        model_name: Gemini model to use (uses DOC_AUDIT_GEMINI_MODEL env var if None)
        client: Gemini client instance (uses DOC_AUDIT_GEMINI_MODEL env var if None)

    Returns:
        Audit result dictionary
    """
    if model_name is None:
        model_name = os.getenv("DOC_AUDIT_GEMINI_MODEL", "gemini-3-flash")
    
    if client is None:
        client = genai.Client(api_key=os.getenv("GOOGLE_API_KEY"))
    
    user_prompt = build_user_prompt(block)

    response = client.models.generate_content(
        model=model_name,
        contents=user_prompt,
        config=types.GenerateContentConfig(
            system_instruction=system_prompt,
            response_mime_type="application/json",
            response_schema=AUDIT_RESULT_SCHEMA
        )
    )
    
    # With structured output, response is guaranteed to be valid JSON
    result = json.loads(response.text)
    return result


def audit_block_openai(block: dict, system_prompt: str, model_name: str = None) -> dict:
    """
    Audit a text block using OpenAI with strict JSON mode (sync version).

    Args:
        block: Text block to audit
        system_prompt: Cached system prompt with rules and instructions
        model_name: OpenAI model to use (uses DOC_AUDIT_OPENAI_MODEL env var if None)

    Returns:
        Audit result dictionary
    """
    if model_name is None:
        model_name = os.getenv("DOC_AUDIT_OPENAI_MODEL", "gpt-5.2")
    
    user_prompt = build_user_prompt(block)

    client = openai.OpenAI()
    response = client.chat.completions.create(
        model=model_name,
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt}
        ],
        temperature=0.2,
        response_format={
            "type": "json_schema",
            "json_schema": {
                "name": "audit_result",
                "strict": True,
                "schema": AUDIT_RESULT_SCHEMA
            }
        }
    )

    # With structured output, response is guaranteed to be valid JSON
    result = json.loads(response.choices[0].message.content)
    return result


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


def load_completed_uuids(manifest_path: str) -> set:
    """
    Load UUIDs of already-processed blocks from manifest.

    Args:
        manifest_path: Path to manifest file

    Returns:
        Set of completed UUIDs
    """
    completed = set()
    path = Path(manifest_path)

    if path.exists():
        with open(path, 'r', encoding='utf-8') as f:
            for line in f:
                line = line.strip()
                if line:
                    entry = json.loads(line)
                    completed.add(entry.get('uuid', ''))

    return completed


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
    semaphore: asyncio.Semaphore
) -> tuple:
    """
    Process a single block asynchronously with concurrency control.

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

    Returns:
        Tuple of (block_idx, success, violation_count, heading)
    """
    async with semaphore:
        block_uuid = block.get('uuid', str(block_idx))
        heading = block.get('heading', 'Unknown')[:50]
        
        try:
            # Rate limiting (applied per concurrent task)
            if rate_limit > 0:
                await asyncio.sleep(rate_limit)
            
            # Call LLM with cached system prompt
            if use_gemini:
                result = await audit_block_gemini_async(block, system_prompt, model_name, client)
            else:
                result = await audit_block_openai_async(block, system_prompt, model_name, client)

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

            # Save to manifest (thread-safe)
            await save_manifest_entry_async(manifest_path, entry, manifest_lock)

            violation_count = len(violations_with_metadata)
            return (block_idx, True, violation_count, heading)

        except json.JSONDecodeError as e:
            print(f"[{block_idx+1}/{total_blocks}] Error: Failed to parse LLM response: {e}", file=sys.stderr)
            return (block_idx, False, 0, heading)
        except Exception as e:
            print(f"[{block_idx+1}/{total_blocks}] Error: {e}", file=sys.stderr)
            return (block_idx, False, 0, heading)


async def run_audit_async(args, blocks, rules, metadata, use_gemini, model_name, client, completed_uuids):
    """
    Run the audit process asynchronously with parallel block processing.

    Args:
        args: Parsed command-line arguments
        blocks: List of text blocks to audit
        rules: List of audit rules
        metadata: Document metadata
        use_gemini: Whether to use Gemini
        model_name: LLM model name
        client: LLM client instance
        completed_uuids: Set of already-processed UUIDs
    """
    # Build rule category mapping
    rule_category_map = build_rule_category_map(rules)

    # Build system prompt once (will be cached by LLM)
    system_prompt = build_system_prompt(rules)
    print(f"System prompt built ({len(system_prompt)} chars, will be cached)")

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
        print("All blocks already processed!")
        return

    print(f"\nUsing model: {model_name}")
    print(f"Processing blocks {start_idx} to {end_idx} ({len(blocks_with_indices)} blocks to process)")
    print(f"Parallel workers: {args.workers}")
    print(f"Output: {args.output}")
    print("-" * 50)

    # Create concurrency controls
    semaphore = asyncio.Semaphore(args.workers)
    manifest_lock = asyncio.Lock()

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
            semaphore=semaphore
        )
        for block_idx, block in blocks_with_indices
    ]

    # Process all blocks in parallel with progress reporting
    total_violations = 0
    blocks_processed = 0
    blocks_failed = 0

    # Use asyncio.as_completed for real-time progress updates
    for coro in asyncio.as_completed(tasks):
        block_idx, success, violation_count, heading = await coro
        
        if success:
            blocks_processed += 1
            total_violations += violation_count
            if violation_count > 0:
                print(f"[{block_idx+1}/{len(blocks)}] {heading}... Found {violation_count} violation(s)")
            else:
                print(f"[{block_idx+1}/{len(blocks)}] {heading}... OK")
        else:
            blocks_failed += 1

    # Summary
    print("\n" + "=" * 50)
    print("Audit Complete")
    print(f"Blocks processed: {blocks_processed}")
    print(f"Blocks failed: {blocks_failed}")
    print(f"Total violations: {total_violations}")
    print(f"Manifest saved to: {args.output}")


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
        required=True,
        help="Path to audit rules JSON file"
    )
    parser.add_argument(
        "--output", "-o",
        type=str,
        default="manifest.jsonl",
        help="Output manifest file path (default: manifest.jsonl)"
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
        "--dry-run",
        action="store_true",
        help="Print prompts without calling LLM"
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

    if model_name == "auto":
        if HAS_GEMINI and os.getenv("GOOGLE_API_KEY"):
            use_gemini = True
            model_name = os.getenv("DOC_AUDIT_GEMINI_MODEL", "gemini-3-flash")
            # Use .aio for async client access
            client = genai.Client(api_key=os.getenv("GOOGLE_API_KEY")).aio
        elif HAS_OPENAI and os.getenv("OPENAI_API_KEY"):
            model_name = os.getenv("DOC_AUDIT_OPENAI_MODEL", "gpt-5.2")
            client = openai.AsyncOpenAI()
        else:
            print("Error: No API key found. Set GOOGLE_API_KEY or OPENAI_API_KEY", file=sys.stderr)
            sys.exit(1)
    elif "gemini" in model_name.lower():
        if not HAS_GEMINI:
            print("Error: google-genai not installed", file=sys.stderr)
            sys.exit(1)
        if not os.getenv("GOOGLE_API_KEY"):
            print("Error: GOOGLE_API_KEY not set", file=sys.stderr)
            sys.exit(1)
        use_gemini = True
        # Use .aio for async client access
        client = genai.Client(api_key=os.getenv("GOOGLE_API_KEY")).aio
    else:
        # Treat all other models as OpenAI (gpt-5.2, o1-mini, o3-mini, etc.)
        if not HAS_OPENAI:
            print("Error: openai not installed", file=sys.stderr)
            sys.exit(1)
        if not os.getenv("OPENAI_API_KEY"):
            print("Error: OPENAI_API_KEY not set", file=sys.stderr)
            sys.exit(1)
        client = openai.AsyncOpenAI()

    # Load inputs
    print(f"Loading blocks from: {args.document}")
    metadata, blocks = load_blocks(args.document)
    print(f"Loaded {len(blocks)} blocks")
    if metadata:
        print(f"Source file: {metadata.get('source_file', 'Unknown')}")
        print(f"File hash: {metadata.get('source_hash', 'Unknown')[:20]}...")

    print(f"Loading rules from: {args.rules}")
    rules = load_rules(args.rules)
    print(f"Loaded {len(rules)} rules")

    # Handle resume
    completed_uuids = set()
    if args.resume and Path(args.output).exists():
        completed_uuids = load_completed_uuids(args.output)
        print(f"Resuming: {len(completed_uuids)} blocks already processed")
    else:
        # Write metadata as first line for new manifest
        if metadata:
            from datetime import datetime
            audit_metadata = {
                **metadata,
                "audited_at": datetime.now().isoformat()
            }
            with open(args.output, 'w', encoding='utf-8') as f:
                f.write(json.dumps(audit_metadata, ensure_ascii=False) + '\n')
            print(f"Created new manifest with source file metadata")

    # Handle dry-run mode (no async needed)
    if args.dry_run:
        system_prompt = build_system_prompt(rules)
        start_idx = args.start_block
        end_idx = args.end_block if args.end_block >= 0 else len(blocks) - 1
        blocks_to_process = blocks[start_idx:end_idx + 1]
        
        for i, block in enumerate(blocks_to_process):
            block_idx = start_idx + i
            block_uuid = block.get('uuid', str(block_idx))
            
            if block_uuid in completed_uuids:
                print(f"[{block_idx+1}/{len(blocks)}] Skipping (already processed)")
                continue
            
            print(f"[{block_idx+1}/{len(blocks)}] Auditing: {block.get('heading', 'Unknown')[:50]}...")
            user_prompt = build_user_prompt(block)
            print(f"\n--- System Prompt ---\n{system_prompt[:300]}...\n")
            print(f"--- User Prompt ---\n{user_prompt[:300]}...")
        return

    # Run async audit
    asyncio.run(run_audit_async(args, blocks, rules, metadata, use_gemini, model_name, client, completed_uuids))


if __name__ == "__main__":
    main()
