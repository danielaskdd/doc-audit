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
            "type": "object",
            "properties": {
                "violation": {"type": "string"},
                "correction": {"type": "string"}
            }
        }
    },
    "required": ["id", "description", "severity", "category"]
}

RULES_ARRAY_SCHEMA = {
    "type": "array",
    "items": RULE_SCHEMA
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
        # Auto-detect default_rules.json
        # Priority: 1) .claude-work/doc-audit/default_rules.json, 2) skill's assets/default_rules.json
        work_dir_rules = Path.cwd() / ".claude-work" / "doc-audit" / "default_rules.json"
        script_dir = Path(__file__).parent
        skill_rules = script_dir.parent / "assets" / "default_rules.json"
        
        if work_dir_rules.exists():
            base_rules_path = work_dir_rules
        else:
            base_rules_path = skill_rules
    
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
- examples: Optional object with "violation" and "correction" examples

IMPORTANT: All rule descriptions, violation examples, correction examples, and any other textual content MUST be written in {output_language}.

Return a valid JSON array of the complete rules."""
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
- examples: Optional object with "violation" and "correction" examples

IMPORTANT: All rule descriptions, violation examples, correction examples, and any other textual content MUST be written in {output_language}.

Return a valid JSON array of the complete rules."""


def merge_rules_with_llm(base_rules: list, input_text: str, api_key: Optional[str] = None) -> list:
    """
    Use LLM to intelligently merge base rules with user requirements.

    Args:
        base_rules: Existing base rules (from default or previous iteration)
        input_text: User's new requirements or modification requests
        api_key: API key for LLM service (Gemini or OpenAI)

    Returns:
        Complete merged list of structured rule dictionaries
    """
    # Get output language from environment variable
    output_language = os.getenv("AUDIT_LANGUAGE", "Chinese")
    
    # Build unified prompt
    prompt = _build_prompt(base_rules, input_text, output_language)
    
    # Try Gemini first
    google_key = api_key or os.getenv("GOOGLE_API_KEY")
    if google_key:
        try:
            from google import genai
            from google.genai import types
            
            client = genai.Client(api_key=google_key)
            # Use environment variable for model name, fallback to default
            model_name = os.getenv("DOC_AUDIT_GEMINI_MODEL", "gemini-3-flash")

            response = client.models.generate_content(
                model=model_name,
                contents=prompt,
                config=types.GenerateContentConfig(
                    response_mime_type="application/json",
                    response_schema=RULES_ARRAY_SCHEMA
                )
            )
            # With structured output, response is guaranteed to be valid JSON
            merged_rules = json.loads(response.text)
            return merged_rules

        except ImportError:
            print("Warning: google-genai not installed. Trying OpenAI instead.", file=sys.stderr)
        except Exception as e:
            print(f"Warning: LLM merging failed: {e}. Trying fallback.", file=sys.stderr)

    # Try OpenAI as fallback
    openai_key = api_key or os.getenv("OPENAI_API_KEY")
    if openai_key:
        try:
            import openai
            client = openai.OpenAI(api_key=openai_key)
            # Use environment variable for model name, fallback to default
            model_name = os.getenv("DOC_AUDIT_OPENAI_MODEL", "gpt-5.2")

            response = client.chat.completions.create(
                model=model_name,
                messages=[{"role": "user", "content": prompt}],
                response_format={
                    "type": "json_schema",
                    "json_schema": {
                        "name": "audit_rules_array",
                        "strict": True,
                        "schema": {
                            "type": "object",
                            "properties": {
                                "rules": RULES_ARRAY_SCHEMA
                            },
                            "required": ["rules"],
                            "additionalProperties": False
                        }
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
            print(f"Error: LLM merging failed: {e}", file=sys.stderr)

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
        help="Base rules file to merge with (default: auto-detect assets/default_rules.json)"
    )
    parser.add_argument(
        "--output", "-o",
        type=str,
        default="rules.json",
        help="Output JSON file path (default: rules.json)"
    )
    parser.add_argument(
        "--no-base",
        action="store_true",
        help="Don't load base rules (start from scratch)"
    )
    parser.add_argument(
        "--api-key",
        type=str,
        help="API key for LLM service (optional, uses environment variables by default)"
    )

    args = parser.parse_args()

    # Validate LLM setup (always required now)
    google_key = args.api_key or os.getenv("GOOGLE_API_KEY")
    openai_key = args.api_key or os.getenv("OPENAI_API_KEY")
    has_gemini = False
    has_openai = False
    try:
        from google import genai  # noqa: F401
        has_gemini = True
    except ImportError:
        pass
    try:
        import openai  # noqa: F401
        has_openai = True
    except ImportError:
        pass

    usable_gemini = bool(google_key and has_gemini)
    usable_openai = bool(openai_key and has_openai)

    if not usable_gemini and not usable_openai:
        if not (google_key or openai_key):
            print("Error: LLM API key required. Set GOOGLE_API_KEY or OPENAI_API_KEY (or use --api-key).", file=sys.stderr)
        else:
            print("Error: No supported LLM client installed.", file=sys.stderr)
        print("Install one of:", file=sys.stderr)
        print("  pip install google-genai", file=sys.stderr)
        print("  pip install openai", file=sys.stderr)
        sys.exit(1)

    # Load base rules
    base_rules = []
    if not args.no_base:
        base_rules = load_base_rules(args.base_rules)
        if base_rules:
            print(f"Loaded {len(base_rules)} base rules")

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
        "total_rules": len(all_rules),
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
