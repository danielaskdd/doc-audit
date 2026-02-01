#!/usr/bin/env python3
"""
ABOUTME: Parses natural language audit criteria into structured JSON rules
ABOUTME: Supports LLM-based merging of base rules with user requirements
"""

import argparse
import json
import os
import sys
from pathlib import Path
from typing import Optional

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


def create_gemini_client():
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
    
    Returns:
        Gemini client instance (sync)
        
    Raises:
        ValueError: If required environment variables are not set
    """
    if not HAS_GEMINI:
        raise ImportError("google-genai not installed")
    
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
    
    return client


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


# JSON Schema for rule validation and LLM structured output
RULE_SCHEMA = {
    "type": "object",
    "properties": {
        "id": {"type": "string", "description": "Unique identifier (R001, R002, ...)"},
        "description": {"type": "string", "description": "Clear description of what to check"},
        "severity": {"type": "string", "enum": ["high", "medium", "low"]},
        "category": {
            "type": "string",
            "description": "Rule category"
        },
        "examples": {
            "type": "array",
            "description": "Array of example violation/correction pairs",
            "items": {
                "type": "object",
                "properties": {
                    "violation": {"type": "string", "description": "Example text that violates the rule"},
                    "correction": {"type": "string", "description": "Corrected version of the violation"}
                },
                "required": ["violation", "correction"],
                "additionalProperties": False
            }
        }
    },
    "required": ["id", "description", "severity", "category"]
}

# Wrapper schema for structured output (used by both Gemini and OpenAI)
# Note: additionalProperties is required by OpenAI strict mode, Gemini tolerates it
RULES_RESPONSE_SCHEMA = {
    "type": "object",
    "properties": {
        "rules": {
            "type": "array",
            "items": RULE_SCHEMA
        }
    },
    "required": ["rules"],
    "additionalProperties": False
}


def renumber_rules(rules: list) -> list:
    """
    Renumber rules with sequential IDs in format RXXX.
    
    This ensures no duplicate IDs and consistent formatting
    regardless of what the LLM returns.
    
    Args:
        rules: List of rule dictionaries
    
    Returns:
        List of rules with renumbered IDs (R001, R002, ...)
    """
    renumbered = []
    for idx, rule in enumerate(rules, start=1):
        new_rule = rule.copy()
        new_rule['id'] = f"R{idx:03d}"
        renumbered.append(new_rule)
    return renumbered


def load_base_rules(base_rules_path: Optional[str] = None) -> list:
    """
    Load base rules from file.
    
    Args:
        base_rules_path: Path to base rules JSON file. If None, auto-detects default_rules.json
    
    Returns:
        List of rule dictionaries
    """
    if base_rules_path is None:
        # Use default_rules.json from skill's assets directory
        script_dir = Path(__file__).parent
        base_rules_path = script_dir.parent / "assets" / "default_rules.json"
    
    path = Path(base_rules_path)
    if not path.exists():
        print(f"Warning: Base rules file not found: {base_rules_path}", file=sys.stderr)
        return []
    
    with open(path, 'r', encoding='utf-8') as f:
        data = json.load(f)
    
    # Handle both direct array and wrapped format
    if isinstance(data, list):
        return data
    elif isinstance(data, dict) and 'rules' in data:
        return data['rules']
    else:
        print(f"Warning: Unknown rules format in {base_rules_path}", file=sys.stderr)
        return []


def _build_prompt(base_rules: list, input_text: str, output_language: str) -> str:
    """
    Build the prompt for LLM rule merging.
    
    Args:
        base_rules: Existing base rules (empty list if starting from scratch)
        input_text: User's requirements or modification requests
        output_language: Language for rule descriptions
    
    Returns:
        Formatted prompt string
    """
    has_base_rules = bool(base_rules)
    
    if has_base_rules:
        return f"""You are an audit rule expert. Your task is to ADD new rules to an existing ruleset while PRESERVING the existing rules.

EXISTING BASE RULES (PROTECTED - DO NOT MODIFY unless explicitly requested):
{json.dumps(base_rules, indent=2, ensure_ascii=False)}

USER'S REQUIREMENTS:
{input_text}

CRITICAL INSTRUCTIONS:
1. **PRESERVE ALL EXISTING RULES AS-IS**: Copy all existing base rules to the output WITHOUT any modifications to their description, severity, category, or examples - unless the user EXPLICITLY requests a change (e.g., "modify R003", "change the severity of R005", "delete R007").

2. **ADD NEW RULES ONLY**: If the user's requirement describes something NOT already covered by existing rules, add it as a new rule.

3. **DETECT EXPLICIT MODIFICATION REQUESTS**: Only modify an existing rule if the user clearly and explicitly requests it, such as:
   - "Change R003 severity to high"
   - "Modify the description of R005 to..."
   - "Delete/Remove R007"
   - "Update rule about XXX to..."

4. **DO NOT "merge" or "improve" existing rules**: Even if a user requirement seems similar to an existing rule, do NOT modify the existing rule. Instead, either:
   - Skip adding if the existing rule already covers it sufficiently
   - Add as a separate new rule if there's meaningful difference

5. Renumber all rules sequentially starting from R001.

Each rule must have:
- id: Unique identifier (R001, R002, ...)
- description: Clear description of what to check
- severity: "high", "medium", or "low"
- category: Suggested values: "grammar", "clarity", "logic", "compliance", "format", "semantic", "other"
  You may also use custom categories if they better fit the rule type.
- examples: Optional array of example objects. Each object has "violation" (problematic text) and "correction" (corrected text). Multiple examples help illustrate the rule.

IMPORTANT: All rule descriptions, violation examples, correction examples, and any other textual content MUST be written in {output_language}.

Return a valid JSON object with a "rules" property containing an array of the complete rules."""
    else:
        return f"""You are an audit rule expert. Create structured audit rules based on user's requirements.

USER'S REQUIREMENTS:
{input_text}

Task: Create a comprehensive ruleset based on the requirements:
1. Parse and structure each requirement as a separate rule
2. Assign sequential rule IDs starting from R001
3. Determine appropriate severity levels based on the requirement's importance
4. Categorize rules appropriately
5. Add helpful examples where relevant

Each rule must have:
- id: Unique identifier (R001, R002, ...)
- description: Clear description of what to check
- severity: "high", "medium", or "low"
- category: Suggested values: "grammar", "clarity", "logic", "compliance", "format", "semantic", "other"
  You may also use custom categories if they better fit the rule type.
- examples: Optional array of example objects. Each object has "violation" (problematic text) and "correction" (corrected text). Multiple examples help illustrate the rule.

IMPORTANT: All rule descriptions, violation examples, correction examples, and any other textual content MUST be written in {output_language}.

Return a valid JSON object with a "rules" property containing an array of the complete rules."""


def merge_rules_with_llm(base_rules: list, input_text: str, api_key: Optional[str] = None) -> list:
    """
    Use LLM to intelligently merge base rules with user requirements.
    
    Supports both AI Studio and Vertex AI modes for Gemini.
    See create_gemini_client() for details on environment variables.

    Args:
        base_rules: Existing base rules (from default or previous iteration)
        input_text: User's new requirements or modification requests
        api_key: API key for LLM service (Gemini AI Studio or OpenAI only, 
                 Vertex AI uses ADC)

    Returns:
        Complete merged list of structured rule dictionaries
    """
    # Get output language from environment variable
    output_language = os.getenv("AUDIT_LANGUAGE", "Chinese")
    
    # Build unified prompt
    prompt = _build_prompt(base_rules, input_text, output_language)
    
    # Determine Gemini availability based on mode
    use_vertex = is_vertex_ai_mode()
    gemini_available = False
    
    if HAS_GEMINI:
        if use_vertex:
            # Vertex AI mode - check for project configuration
            gemini_available = bool(os.getenv("GOOGLE_CLOUD_PROJECT"))
        else:
            # AI Studio mode - check for API key
            gemini_available = bool(api_key or os.getenv("GOOGLE_API_KEY"))
    
    # Try Gemini first
    if gemini_available:
        try:
            client = create_gemini_client()
            # Use environment variable for model name, fallback to default
            model_name = os.getenv("DOC_AUDIT_GEMINI_MODEL", "gemini-2.5-flash")
            
            provider_name = get_gemini_provider_name()
            print(f"Using LLM: {provider_name} ({model_name})")

            response = client.models.generate_content(
                model=model_name,
                contents=prompt,
                config=types.GenerateContentConfig(
                    response_mime_type="application/json",
                    response_schema=RULES_RESPONSE_SCHEMA
                )
            )
            # With structured output, response is guaranteed to be valid JSON
            response_data = json.loads(response.text)
            merged_rules = response_data["rules"]
            return merged_rules

        except ImportError:
            print("Warning: google-genai not installed. Trying OpenAI instead.", file=sys.stderr)
        except ValueError as e:
            # Configuration error from create_gemini_client
            print(f"Warning: Gemini configuration error: {e}. Trying OpenAI instead.", file=sys.stderr)
        except Exception as e:
            print(f"Warning: Gemini merging failed: {e}. Trying OpenAI instead.", file=sys.stderr)

    # Try OpenAI as fallback
    openai_key = api_key or os.getenv("OPENAI_API_KEY")
    if HAS_OPENAI and openai_key:
        try:
            base_url = os.getenv("OPENAI_BASE_URL")
            client = openai.OpenAI(api_key=openai_key, base_url=base_url)
            # Use environment variable for model name, fallback to default
            model_name = os.getenv("DOC_AUDIT_OPENAI_MODEL", "gpt-4.1")
            
            provider_name = f"OpenAI ({base_url})" if base_url else "OpenAI"
            print(f"Using LLM: {provider_name} ({model_name})")

            response = client.chat.completions.create(
                model=model_name,
                messages=[{"role": "user", "content": prompt}],
                response_format={
                    "type": "json_schema",
                    "json_schema": {
                        "name": "audit_rules_response",
                        "strict": True,
                        "schema": RULES_RESPONSE_SCHEMA
                    }
                }
            )
            # With structured output, response is guaranteed to be valid JSON
            response_data = json.loads(response.choices[0].message.content)
            merged_rules = response_data["rules"]
            return merged_rules

        except ImportError:
            print("Error: openai not installed.", file=sys.stderr)
        except Exception as e:
            print(f"Error: OpenAI merging failed: {e}", file=sys.stderr)

    # No fallback - LLM is required
    print("Error: Unable to use LLM for rule merging. Please ensure LLM dependencies are installed.", file=sys.stderr)
    sys.exit(1)


def main():
    parser = argparse.ArgumentParser(
        description="Parse natural language audit criteria into structured JSON rules"
    )
    parser.add_argument(
        "--input", "-i",
        type=str,
        help="Natural language audit criteria text or modification requests"
    )
    parser.add_argument(
        "--file", "-f",
        type=str,
        help="File containing audit criteria (one per line or paragraph)"
    )
    parser.add_argument(
        "--base-rules",
        type=str,
        default=None,
        help="Base rules file to merge with. If not specified, starts from scratch."
    )
    parser.add_argument(
        "--output", "-o",
        type=str,
        default="rules.json",
        help="Output JSON file path (default: rules.json)"
    )
    parser.add_argument(
        "--api-key",
        type=str,
        help="API key for LLM service (optional, uses environment variables by default)"
    )

    args = parser.parse_args()

    # Validate LLM setup (always required now)
    # Check for Gemini availability (AI Studio or Vertex AI)
    use_vertex = is_vertex_ai_mode()
    usable_gemini = False
    
    if HAS_GEMINI:
        if use_vertex:
            # Vertex AI mode - requires project configuration (ADC handles auth)
            if os.getenv("GOOGLE_CLOUD_PROJECT"):
                usable_gemini = True
            elif not os.getenv("OPENAI_API_KEY"):
                # Only error if no fallback available
                print("Error: Vertex AI mode is enabled (GOOGLE_GENAI_USE_VERTEXAI=true) "
                      "but GOOGLE_CLOUD_PROJECT is not set.", file=sys.stderr)
                print("Either:", file=sys.stderr)
                print("  1. Set GOOGLE_CLOUD_PROJECT to your GCP project ID", file=sys.stderr)
                print("  2. Unset GOOGLE_GENAI_USE_VERTEXAI to use AI Studio mode", file=sys.stderr)
                print("  3. Set OPENAI_API_KEY to use OpenAI instead", file=sys.stderr)
                sys.exit(1)
        else:
            # AI Studio mode - requires API key
            google_key = args.api_key or os.getenv("GOOGLE_API_KEY")
            usable_gemini = bool(google_key)
    
    # Check for OpenAI availability
    openai_key = args.api_key or os.getenv("OPENAI_API_KEY")
    usable_openai = bool(HAS_OPENAI and openai_key)

    if not usable_gemini and not usable_openai:
        if not HAS_GEMINI and not HAS_OPENAI:
            print("Error: No supported LLM client installed.", file=sys.stderr)
            print("Install one of:", file=sys.stderr)
            print("  pip install google-genai", file=sys.stderr)
            print("  pip install openai", file=sys.stderr)
        else:
            print("Error: No LLM credentials found.", file=sys.stderr)
            print("For AI Studio: Set GOOGLE_API_KEY", file=sys.stderr)
            print("For Vertex AI: Set GOOGLE_GENAI_USE_VERTEXAI=true and GOOGLE_CLOUD_PROJECT", file=sys.stderr)
            print("For OpenAI: Set OPENAI_API_KEY", file=sys.stderr)
        sys.exit(1)

    # Load base rules only if explicitly specified
    base_rules = []
    if args.base_rules:
        base_rules = load_base_rules(args.base_rules)
        if base_rules:
            print(f"Loaded {len(base_rules)} base rules from {args.base_rules}")

    # Get user input text
    input_text = ""
    if args.input:
        input_text = args.input
    elif args.file:
        file_path = Path(args.file)
        if not file_path.exists():
            print(f"Error: File not found: {args.file}", file=sys.stderr)
            sys.exit(1)
        input_text = file_path.read_text(encoding="utf-8")

    # Generate or merge rules (always uses LLM)
    if not input_text:
        # No input text, just use base rules
        if not base_rules:
            print("Error: No base rules and no input provided. Nothing to generate.", file=sys.stderr)
            print("Either provide input text (--input or --file) or ensure base rules are available.", file=sys.stderr)
            sys.exit(1)
        all_rules = base_rules
    else:
        # Use LLM to merge base rules with user requirements
        all_rules = merge_rules_with_llm(base_rules, input_text, args.api_key)

    # Renumber rules to ensure consistent RXXX format and avoid duplicates
    all_rules = renumber_rules(all_rules)

    # Output
    output_data = {
        "version": "1.0",
        "description": "Customized audit rules generated by LLM requested by user",
        "type": "block",
        "rules": all_rules
    }

    output_path = Path(args.output)
    output_path.write_text(json.dumps(output_data, indent=2, ensure_ascii=False), encoding="utf-8")

    print(f"\nGenerated {len(all_rules)} rules:")
    for rule in all_rules:
        print(f"  [{rule['id']}] ({rule['severity']}) {rule['description'][:60]}...")
    print(f"\nSaved to: {output_path}")


if __name__ == "__main__":
    main()
