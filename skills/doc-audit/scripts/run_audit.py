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
from pathlib import Path

# Maximum number of concurrent LLM API calls
MAX_PARALLEL_WORKERS = 8

# Retry configuration
DEFAULT_MAX_RETRIES = 3
INITIAL_BACKOFF = 1.0  # Initial backoff time in seconds
MAX_BACKOFF = 60.0     # Maximum backoff time in seconds
BACKOFF_MULTIPLIER = 2.0  # Backoff multiplier for exponential growth

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


def is_vertex_ai_mode() -> bool:
    """
    Check if Vertex AI mode is enabled via environment variable.
    
    Returns:
        True if GOOGLE_GENAI_USE_VERTEXAI is set to 'true', False otherwise
    """
    return os.getenv("GOOGLE_GENAI_USE_VERTEXAI", "").lower() == "true"


def create_gemini_client(use_async: bool = False):
    """
    Create Gemini client for AI Studio or Vertex AI.
    
    Supports two modes:
    - AI Studio (default): Uses GOOGLE_API_KEY for authentication
    - Vertex AI: Uses ADC (GOOGLE_APPLICATION_CREDENTIALS or gcloud auth)
    
    Environment variables for Vertex AI mode:
    - GOOGLE_GENAI_USE_VERTEXAI: Set to 'true' to enable Vertex AI mode
    - GOOGLE_CLOUD_PROJECT: Required GCP project ID
    - GOOGLE_CLOUD_LOCATION: Optional region (default: us-central1)
    - GOOGLE_VERTEX_BASE_URL: Optional custom API endpoint (for API gateway proxies)
    - GOOGLE_APPLICATION_CREDENTIALS: Path to service account JSON (or use gcloud auth)
    
    Args:
        use_async: If True, return the async client (.aio), otherwise return sync client
        
    Returns:
        Gemini client instance (sync or async based on use_async parameter)
        
    Raises:
        ValueError: If required environment variables are not set
    """
    use_vertex = is_vertex_ai_mode()
    
    if use_vertex:
        # Vertex AI mode - uses ADC (GOOGLE_APPLICATION_CREDENTIALS or gcloud auth)
        project = os.getenv("GOOGLE_CLOUD_PROJECT")
        location = os.getenv("GOOGLE_CLOUD_LOCATION", "us-central1")
        base_url = os.getenv("GOOGLE_VERTEX_BASE_URL")
        
        if not project:
            raise ValueError(
                "GOOGLE_CLOUD_PROJECT is required for Vertex AI mode. "
                "Set GOOGLE_GENAI_USE_VERTEXAI=false to use AI Studio mode instead."
            )
        
        # Build http_options only if custom base_url is specified
        http_options = None
        if base_url:
            http_options = {"base_url": base_url}
        
        # Note: ADC handles authentication automatically
        # via GOOGLE_APPLICATION_CREDENTIALS env var or gcloud auth
        client = genai.Client(
            vertexai=True,
            project=project,
            location=location,
            http_options=http_options
        )
    else:
        # AI Studio mode - requires API key
        api_key = os.getenv("GOOGLE_API_KEY")
        if not api_key:
            raise ValueError(
                "GOOGLE_API_KEY is required for AI Studio mode. "
                "Set GOOGLE_GENAI_USE_VERTEXAI=true and configure GCP credentials for Vertex AI mode."
            )
        
        client = genai.Client(api_key=api_key)
    
    # Return async or sync client based on parameter
    return client.aio if use_async else client


def get_gemini_provider_name() -> str:
    """
    Get the Gemini provider name based on current mode.
    
    Returns:
        Provider name string for display purposes
    """
    if is_vertex_ai_mode():
        project = os.getenv("GOOGLE_CLOUD_PROJECT", "unknown")
        location = os.getenv("GOOGLE_CLOUD_LOCATION", "us-central1")
        return f"Google Gemini (Vertex AI: {project}/{location})"
    else:
        return "Google Gemini (AI Studio)"


def create_openai_client(use_async: bool = True):
    """
    Create OpenAI client with optional custom base URL.
    
    Environment variables:
    - OPENAI_API_KEY: Required API key
    - OPENAI_BASE_URL: Optional custom API endpoint (for proxies, Azure, etc.)
    
    Args:
        use_async: If True, return AsyncOpenAI, otherwise return OpenAI
        
    Returns:
        OpenAI client instance (async or sync based on use_async parameter)
        
    Raises:
        ValueError: If OPENAI_API_KEY is not set
    """
    api_key = os.getenv("OPENAI_API_KEY")
    if not api_key:
        raise ValueError("OPENAI_API_KEY is required for OpenAI mode.")
    
    base_url = os.getenv("OPENAI_BASE_URL")
    
    if use_async:
        return openai.AsyncOpenAI(base_url=base_url)
    else:
        return openai.OpenAI(base_url=base_url)


def get_openai_provider_name() -> str:
    """
    Get the OpenAI provider name, including custom endpoint if configured.
    
    Returns:
        Provider name string for display purposes
    """
    base_url = os.getenv("OPENAI_BASE_URL")
    if base_url:
        return f"OpenAI (Custom: {base_url})"
    else:
        return "OpenAI"


class NonRetryableError(Exception):
    """
    Exception for errors that should not be retried.
    
    These include authentication errors, invalid API keys, permission errors,
    and other permanent failures that won't be resolved by retrying.
    """
    pass


def is_openai_retryable(error: Exception) -> bool:
    """
    Determine if an OpenAI error should be retried.
    
    Non-retryable errors:
    - AuthenticationError (401): Invalid API key
    - PermissionDeniedError (403): No access to resource
    - BadRequestError (400): Invalid request format
    - NotFoundError (404): Model or resource not found
    
    Retryable errors:
    - RateLimitError (429): Rate limit exceeded
    - APIConnectionError: Network issues
    - InternalServerError (500): Server errors
    - APIStatusError with 502, 503, 504: Gateway/service errors
    
    Args:
        error: The exception from OpenAI API call
        
    Returns:
        True if the error should be retried, False otherwise
    """
    if not HAS_OPENAI:
        return True
    
    # Authentication error - invalid API key (401)
    if isinstance(error, openai.AuthenticationError):
        return False
    
    # Permission denied - no access to resource (403)
    if isinstance(error, openai.PermissionDeniedError):
        return False
    
    # Bad request - invalid request format (400)
    if isinstance(error, openai.BadRequestError):
        return False
    
    # Not found - model or resource doesn't exist (404)
    if isinstance(error, openai.NotFoundError):
        return False
    
    # Rate limit exceeded - should retry with backoff (429)
    if isinstance(error, openai.RateLimitError):
        return True
    
    # API connection error - network issues, should retry
    if isinstance(error, openai.APIConnectionError):
        return True
    
    # Internal server error - should retry (500)
    if isinstance(error, openai.InternalServerError):
        return True
    
    # For other APIStatusError, check HTTP status code
    if isinstance(error, openai.APIStatusError):
        # Retryable server-side errors
        return error.status_code in (429, 500, 502, 503, 504)
    
    # For unknown errors, default to retry (network issues, timeouts, etc.)
    return True


def is_gemini_retryable(error: Exception) -> bool:
    """
    Determine if a Gemini error should be retried.
    
    Uses string matching on error messages since google-genai may not have
    well-defined exception types for all error cases.
    
    Non-retryable errors:
    - API key errors
    - Authentication/permission errors
    - Invalid request errors
    - Model not found errors
    - Billing/quota permanently exceeded
    
    Retryable errors:
    - Rate limit (429)
    - Server errors (500, 502, 503, 504)
    - Timeout/connection errors
    
    Args:
        error: The exception from Gemini API call
        
    Returns:
        True if the error should be retried, False otherwise
    """
    error_str = str(error).lower()
    
    # API key / authentication errors - do not retry
    if 'api_key' in error_str or 'api key' in error_str:
        return False
    if 'authentication' in error_str or 'authenticate' in error_str:
        return False
    if 'invalid_api_key' in error_str or 'invalid api key' in error_str:
        return False
    
    # Permission / forbidden errors - do not retry
    if 'permission' in error_str and 'denied' in error_str:
        return False
    if 'forbidden' in error_str or '403' in error_str:
        return False
    
    # Invalid request errors - do not retry
    if 'invalid' in error_str and ('request' in error_str or 'argument' in error_str):
        return False
    if '400' in error_str and 'bad request' in error_str:
        return False
    
    # Model not found - do not retry
    if 'model' in error_str and ('not found' in error_str or 'not exist' in error_str):
        return False
    if '404' in error_str:
        return False
    
    # Billing / permanent quota errors - do not retry
    if 'billing' in error_str:
        return False
    if 'quota' in error_str and ('exceeded' in error_str or 'exhausted' in error_str):
        # Check if it mentions billing which indicates permanent quota issue
        if 'billing' in error_str or 'payment' in error_str:
            return False
        # Temporary quota (rate limit) - should retry
        return True
    
    # Rate limit errors - should retry (429)
    if 'rate' in error_str and 'limit' in error_str:
        return True
    if '429' in error_str or 'resource_exhausted' in error_str:
        return True
    
    # Server errors - should retry (500, 502, 503, 504)
    if any(code in error_str for code in ['500', '502', '503', '504']):
        return True
    if 'internal' in error_str and ('error' in error_str or 'server' in error_str):
        return True
    if 'service' in error_str and 'unavailable' in error_str:
        return True
    if 'gateway' in error_str:
        return True
    
    # Timeout / connection errors - should retry
    if 'timeout' in error_str or 'timed out' in error_str:
        return True
    if 'connection' in error_str:
        return True
    if 'network' in error_str:
        return True
    
    # Unknown errors - default to retry with limited attempts
    return True


async def audit_block_with_retry(
    block: dict,
    system_prompt: str,
    model_name: str,
    client,
    use_gemini: bool,
    max_retries: int = DEFAULT_MAX_RETRIES,
    block_idx: int = 0,
    total_blocks: int = 1
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
        
    Returns:
        Audit result dictionary
        
    Raises:
        NonRetryableError: For permanent errors that should not be retried
        Exception: For errors that exceeded retry attempts
    """
    last_error = None
    
    for attempt in range(max_retries + 1):
        try:
            if use_gemini:
                return await audit_block_gemini_async(block, system_prompt, model_name, client)
            else:
                return await audit_block_openai_async(block, system_prompt, model_name, client)
                
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
    
    Supports both AI Studio and Vertex AI modes based on environment configuration.
    See create_gemini_client() for details on environment variables.

    Args:
        block: Text block to audit
        system_prompt: Cached system prompt with rules and instructions
        model_name: Gemini model to use (uses DOC_AUDIT_GEMINI_MODEL env var if None)
        client: Gemini client instance (creates one using create_gemini_client if None)

    Returns:
        Audit result dictionary
    """
    if model_name is None:
        model_name = os.getenv("DOC_AUDIT_GEMINI_MODEL", "gemini-2.5-flash")
    
    if client is None:
        client = create_gemini_client(use_async=False)
    
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


def audit_block_openai(block: dict, system_prompt: str, model_name: str = None, client = None) -> dict:
    """
    Audit a text block using OpenAI with strict JSON mode (sync version).
    
    Supports custom base URL via OPENAI_BASE_URL environment variable.
    See create_openai_client() for details.

    Args:
        block: Text block to audit
        system_prompt: Cached system prompt with rules and instructions
        model_name: OpenAI model to use (uses DOC_AUDIT_OPENAI_MODEL env var if None)
        client: OpenAI client instance (creates one using create_openai_client if None)

    Returns:
        Audit result dictionary
    """
    if model_name is None:
        model_name = os.getenv("DOC_AUDIT_OPENAI_MODEL", "gpt-5.2")
    
    if client is None:
        client = create_openai_client(use_async=False)
    
    user_prompt = build_user_prompt(block)

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
    
    with open(path, 'r', encoding='utf-8') as f:
        for line in f:
            line = line.strip()
            if line:
                entry = json.loads(line)
                # Skip metadata entry
                if 'audited_at' in entry or entry.get('type') == 'meta':
                    continue
                
                # Look up block_idx from uuid
                uuid = entry.get('uuid', '')
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
    max_retries: int = DEFAULT_MAX_RETRIES
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
                total_blocks=total_blocks
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


async def run_audit_async(args, blocks, rules, metadata, use_gemini, model_name, provider_name, client, completed_uuids):
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
    """
    # Build rule category mapping
    rule_category_map = build_rule_category_map(rules)

    # Build system prompt once (will be cached by LLM)
    system_prompt = build_system_prompt(rules)
    print(f"System prompt built ({len(system_prompt)} chars, will be cached)")

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
            semaphore=semaphore,
            max_retries=args.max_retries
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
            from datetime import datetime
            audit_metadata = {
                **metadata,
                "llm_provider": provider_name,
                "llm_model": model_name,
                "audited_at": datetime.now().isoformat()
            }
        
        print("Sorting and rewriting manifest...")
        rewrite_manifest_sorted(args.output, audit_metadata, all_results)

    # Summary
    print("\n" + "=" * 50)
    print("Audit Complete")
    print(f"Blocks processed: {blocks_processed}")
    print(f"Blocks failed: {blocks_failed}")
    print(f"Total violations: {total_violations}")
    print(f"Manifest saved to: {args.output} (sorted by block order)")


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

    if model_name == "auto":
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
    print(f"\nLLM Provider: {provider_name}")
    print(f"Model: {model_name}")

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
                "llm_provider": provider_name,
                "llm_model": model_name,
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
    asyncio.run(run_audit_async(args, blocks, rules, metadata, use_gemini, model_name, provider_name, client, completed_uuids))


if __name__ == "__main__":
    main()
