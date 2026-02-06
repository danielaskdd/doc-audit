#!/usr/bin/env python3
"""
ABOUTME: Centralized prompt management for document audit LLM calls
ABOUTME: Contains system and user prompts for block audit, global extraction, and global verification
"""

import json
import os


# ============================================================
# Prompt Templates
# ============================================================

PROMPT_TEMPLATES = {
    # Block Audit System Prompt
    "block_audit_system": """You are a professional document auditor. Your task is to analyze text blocks and check for violations of audit rules.

{rules_text}

---

Instructions:
1. Check if the provided text block violates ANY of the rules above.
2. Use the "Section hierarchy context" to understand where this block sits in the document structure. This context helps you:
   - Understand the semantic scope of the content (e.g., a clause under "Penalty Terms" vs "Payment Terms")
   - Apply rules appropriately based on document section type
   - Identify context-dependent violations (e.g., vague terms acceptable in summaries but not in binding clauses)
3. Report each violation as a separate item. Do not merge multiple instances of the same audit rule into one entry. If one violation satisfies multiple rules, report only the first rule ID that applies.
4. CRITICAL: The rule_id MUST exactly match the ID of the specific rule violated. If content violates rule [R005], report "rule_id": "R005". Never default to R001 - always use the actual rule ID from the list above.
5. For each violation found, provide:
   - The rule ID that was violated
   - The violation text with enough surrounding context for unique string matching
   - Why it's a violation
   - The fix action: "delete", "replace", or "manual"
   - The revised text based on fix_action

violation_text guidelines:
- The violation_text field is used to locate the original text. It must be a direct verbatim quote from the content of text block. All punctuation, line breaks, and whitespace must be strictly preserved to ensure an exact match
    - Example 1: Keep `Line 1\nLine 2` as is, do not convert to `Line 1 Line 2`
    - Example 2: Keep `Word1\tWord2` as is, do not convert to `Word1 Word2`
- Preserve the subscript and superscript formatting by keeping all `<sub>` and `<sup>` tags intact; do not simplify them to markdown format or plain text.
    - Example 1: Keep chemical equation `H<sub>2</sub>O` with subscript as is, **do not** convert it to markdown format like `H_2O`
    - Example 2: Keep math equation `65×12×10<sup>-6</sup>/h = 7.8×10<sup>-4</sup>/h` with supscript as is, **do not** convert it to markdown format like`65×12×10^-6/h = 7.8×10^-4/h`
- Preserve equation formatting by keeping all `<equation>` tags intact; content inside is LaTeX format. Do not simplify them to plain text.
    - Example: Keep `<equation>\\frac{{1}}{{2}}mv^2</equation>` as is, do not convert to `1/2 mv²`
    - Example: Keep `<equation>E = mc^2</equation>` as is, do not strip the equation tags
- Do not use ellipses to replace or omit any part of the original text
- If the violating content is excessively long (e.g., spanning multiple sentences), extract only the leading portion, ensuring it is sufficient to uniquely locate via string search
- If an entire section is in violation, select the first paragraph as the violation_text
- For violation_text in table content of JSON format
    - Report violation_text at cell level without JSON format if cell content is long enough for unique matching. If the cell content is too brief (e.g., just numbers or short phrases), expand the violation_text to encompass the entire row in JSON format.
    - Report violation_text at row level with JSON format if cell-level is not feasible, violation_text should start from the first cell in the row, i.e., starting with '['; The revise_text field should be aligned at the row level, consistent with violation_text.
- Exclude chapter/heading numbers, list markers, and bullet points from the violation_text
    - Example 1: For the candidate violation text `(B) Component model is CT41-1210-X7R`, the leading index should be removed. The corrected violation_text should be `Component model is CT41-1210-X7R`    
    - Example 2: For the candidate violation text `表16 软件配置项目表"`, the leading table number should be removed. The corrected violation_text should be `软件配置项目表`

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
      "rule_id": "<exact_id_of_violated_rule>",
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

Return ONLY the JSON object, no other text.""",

    # Block Audit User Prompt
    "block_audit_user": """Perform a careful and contextual analysis of the text block content to detect rule violations:

{block_text}""",

    # Global Extraction System Prompt
    "global_extract_system": """You are a precise information extractor. Your task is to identify and extract structured data from document text according to predefined schemas.

Global Extraction Rules:

{rules_text}

Understanding the Rules:
- Each rule defines a TOPIC (what kind of information to look for)
- EXTRACTION describes the scope and types of data to capture
- ENTITY is the main identifier that groups related fields together
- FIELDS are specific attributes to extract for each entity instance
- EVIDENCE is a verbatim text snippet that justifies each extracted value

Key Principles:
1. One entity = one complete set of fields (even if some fields are empty)
2. Multiple instances of the same entity type should be separate results
3. Value should be a concise summary; Evidence should be the exact source text (no more than one sentence)
4. If the same entity appears multiple times with different information, create separate entries

Extraction Guidelines:
1. Extract ONLY information explicitly stated in the text - never infer or assume.
2. Each unique entity instance should be a separate extraction result.
3. Fields must use the exact field names defined in the rule.
4. If a field value cannot be found, use empty string for both value and evidence.
5. Evidence must be a verbatim quote from origin content that directly supports the extracted value.
6. For tabular data, extract one entity per row where applicable.
7. If the same entity appears multiple times with different information, create separate entries.
8. CRITICAL: The rule_id MUST exactly match the ID of the rule being extracted. If extracting for rule [G003], use "rule_id": "G003". Never default to G001.

Return JSON only with this structure:
{{
  "results": [
    {{
      "rule_id": "<exact_id_of_matching_rule>",
      "extracted_results": [
        {{
          "entity": "entity identifier or empty string if unknown",
          "fields": [
            {{"name": "field_name", "value": "extracted value", "evidence": "verbatim quote"}}
          ]
        }}
      ]
    }}
  ]
}}

If no relevant information is found, return:
{{
  "results": []
}}

Return ONLY the JSON object, no other text. Output language for value summaries should be {output_language}.""",

    # Global Extraction User Prompt
    "global_extract_user": """Extract information from the following content of text block according to the provided global extraction rules:

{block_text}""",

    # Global Verification System Prompt
    "global_verify_system": """You are a cross-reference auditor specializing in document consistency verification.

Rule: [{rule_id}] {topic}
Extraction Scope: {extraction}
Entity Name: {extracted_entity}
Verification Criteria: {verification}

Understanding the Input Data:
- Each item represents an extracted entity instance from a document block
- "entity" field contains the entity name (e.g., {extracted_entity})
- "fields" contain extracted attributes with values and evidence
- "uuid"/"uuid_end" mark the source location in the original document

Comparison Scope:
IMPORTANT: Only compare items that refer to THE SAME ENTITY.
- Items with similar or identical "entity" names represent the same entity
- Different entities (e.g., "Server A" vs "Server B") should NOT be compared against each other
- Focus on finding conflicts where the SAME entity has inconsistent information across different document sections

Your Primary Task:
Follow the VERIFICATION CRITERIA above to identify inconsistencies WITHIN the same entity.
Specifically check if the same entity has conflicting attribute values in different parts of the document.

Types of inconsistencies to check (only for same entity):
- VALUE CONFLICTS: Same entity has different values for the same attribute
- COMPLETENESS GAPS: Same entity has information in some sections but missing in related sections

Instructions:
1. Group items by entity name (similar names = same entity)
2. Within each entity group, cross-reference all fields based on the verification criteria
3. Report conflicts only when the SAME entity shows inconsistent information; An empty value should not be treated as an inconsistency unless specifically required by the verification criteria
4. Use the uuid and uuid_end from the INPUT items when reporting violations
5. violation_text MUST be a verbatim evidence quote from the input items
6. Mark fix_action as "manual" since resolution requires human judgment

violation_text guidelines:
- The violation_text field is used to locate the original text. It must be a direct verbatim quote from the evidence. All punctuation, line breaks, and whitespace must be strictly preserved to ensure an exact match
    - Example 1: Keep `Line 1\nLine 2` as is, do not convert to `Line 1 Line 2`
    - Example 2: Keep `Word1\tWord2` as is, do not convert to `Word1 Word2`
- Preserve the subscript and superscript formatting by keeping all `<sub>` and `<sup>` tags intact; do not simplify them to markdown format or plain text.
    - Example 1: Keep chemical equation `H<sub>2</sub>O` with subscript as is, **do not** convert it to markdown format like `H_2O`
    - Example 2: Keep math equation `65×12×10<sup>-6</sup>/h = 7.8×10<sup>-4</sup>/h` with supscript as is, **do not** convert it to markdown format like`65×12×10^-6/h = 7.8×10^-4/h`
- Preserve equation formatting by keeping all `<equation>` tags intact; content inside is LaTeX format. Do not simplify them to plain text.
    - Example: Keep `<equation>\\frac{{1}}{{2}}mv^2</equation>` as is, do not convert to `1/2 mv²`
    - Example: Keep `<equation>E = mc^2</equation>` as is, do not strip the equation tags
- Do not use ellipses to replace or omit any part of the original text
- If the violating content is excessively long (e.g., spanning multiple sentences), extract only conflicting content, ensuring it is sufficient to uniquely locate via string search
- For violation_text in table content of JSON format
    - Report violation_text at cell level without JSON format if cell content is long enough for unique matching. If the cell content is too brief (e.g., just numbers or short phrases), expand the violation_text to encompass the entire row in JSON format.
    - Report violation_text at row level with JSON format if cell-level is not feasible, violation_text should start from the first cell in the row, i.e., starting with '['; The revise_text field should be aligned at the row level, consistent with violation_text.
- Exclude chapter/heading numbers, list markers, and bullet points from the violation_text
    - Example 1: For the candidate violation text `(B) Component model is CT41-1210-X7R`, the leading index should be removed. The corrected violation_text should be `Component model is CT41-1210-X7R`    
    - Example 2: For the candidate violation text `表16 软件配置项目表"`, the leading table number should be removed. The corrected violation_text should be `软件配置项目表`

Return JSON only:
{{
  "violations": [
    {{
      "rule_id": "{rule_id}",
      "uuid": "<uuid from conflicting item>",
      "uuid_end": "<uuid_end from conflicting item>",
      "violation_text": "<verbatim evidence from input>",
      "violation_reason": "<explanation in {output_language}>",
      "fix_action": "manual",
      "revised_text": "<resolution guidance in {output_language}>"
    }}
  ]
}}

If no violations found, return:
{{ "violations": [] }}

Return ONLY the JSON object, no other text.""",

    # Global Verification User Prompt
    "global_verify_user": """Check consistency or violation for the following extracted items according to the provided verification criteria:

{payload_text}""",
}


# ============================================================
# Helper Functions: Rule and Block Formatting
# ============================================================

def normalize_extracted_fields(rule: dict) -> list:
    """
    Normalize extracted_fields definitions into a list of {name, desc, evidence_desc}.

    Supports two formats:
    1) New: {"name": "...", "desc": "...", "evidence": "..."}
    2) Legacy: {"field_name": "...", "evidence": "..."}
    """
    normalized = []
    for field in rule.get("extracted_fields", []):
        if not isinstance(field, dict):
            continue
        if "name" in field:
            normalized.append({
                "name": field.get("name", "").strip(),
                "desc": field.get("desc", "").strip(),
                "evidence_desc": field.get("evidence", "").strip()
            })
            continue
        evidence_desc = field.get("evidence", "").strip()
        for key, value in field.items():
            if key == "evidence":
                continue
            normalized.append({
                "name": str(key).strip(),
                "desc": str(value).strip(),
                "evidence_desc": evidence_desc
            })
    return normalized


def format_block_for_prompt(block: dict) -> str:
    """
    Format a text block for inclusion in the audit prompt.

    Args:
        block: Block dictionary with heading, content, type

    Returns:
        Formatted string
    """
    heading = block.get('heading', 'Unknown').strip()
    content = block.get('content', '').strip()
    block_type = block.get('type', 'text')
    parent_headings = block.get('parent_headings', [])

    ### Context format
    # Context hierarchy: 1  header1 → 1.2  header2 → 1.2.2  header3
    context = ""
    if parent_headings:
        context = f"Section hierarchy context: {' → '.join(h.strip() for h in parent_headings)}  → {heading}"

    if block_type == 'table':
        # Format table as readable text
        if isinstance(content, list):
            rows = []
            for row in content:
                rows.append(" | ".join(str(cell) for cell in row))
            content = "\n".join(rows)

    return f"""{context}

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
        lines.append(f"- [{rule['id']}] {rule['description']}")
        # Include examples if available for better rule understanding
        examples = rule.get('examples', [])
        # Support legacy dict format (auto-convert to list)
        if isinstance(examples, dict):
            examples = [examples]
        
        for i, example in enumerate(examples):
            violation = example.get('violation', '')
            correction = example.get('correction', '')
            
            if violation or correction:
                lines.append(f"  - Example {i + 1}:")
                if violation:
                    lines.append(f"       ✗ violation: {violation}")
                if correction:
                    lines.append(f"       ✓ correction: {correction}")

    return "\n".join(lines)


def format_global_rules_for_extraction(rules: list) -> str:
    """
    Format global rules for the extraction prompt.

    Args:
        rules: List of global rule dictionaries

    Returns:
        Formatted string
    """
    lines = ["Global Extraction Rules:"]
    for rule in rules:
        fields = normalize_extracted_fields(rule)
        lines.append(f"- [{rule.get('id', '')}] {rule.get('topic', '')}")
        extraction = rule.get("extraction", "")
        if extraction:
            lines.append(f"  Extraction: {extraction}")
        entity_label = rule.get("extracted_entity", "")
        if entity_label:
            lines.append(f"  Entity: {entity_label}")
        if fields:
            lines.append("  Fields:")
            for f in fields:
                desc = f.get("desc", "")
                evidence_desc = f.get("evidence_desc", "")
                lines.append(f"    - {f.get('name','')}: {desc}")
                if evidence_desc:
                    lines.append(f"      Evidence: {evidence_desc}")
    return "\n".join(lines)


# ============================================================
# Block Audit Prompts
# ============================================================

def build_block_audit_system_prompt(rules: list) -> str:
    """
    Build the system prompt for block-level audit.

    This prompt instructs the LLM to check a single text block against
    all provided rules and report any violations found.

    Args:
        rules: Audit rules to apply

    Returns:
        System prompt string
    """
    output_language = os.getenv("AUDIT_LANGUAGE", "Chinese")
    rules_text = format_rules_for_prompt(rules)
    return PROMPT_TEMPLATES["block_audit_system"].format(
        rules_text=rules_text,
        output_language=output_language
    )


def build_block_audit_user_prompt(block: dict) -> str:
    """
    Build the user prompt for block-level audit.

    Args:
        block: Text block to audit

    Returns:
        User prompt string
    """
    block_text = format_block_for_prompt(block)
    return PROMPT_TEMPLATES["block_audit_user"].format(block_text=block_text)


# ============================================================
# Global Extraction Prompts
# ============================================================

def build_global_extract_system_prompt(rules: list) -> str:
    """
    Build the system prompt for global information extraction.

    This prompt instructs the LLM to extract structured data from a text block
    according to the global rule schemas (entities and fields).

    Args:
        rules: List of global rules with extraction definitions

    Returns:
        System prompt string
    """
    output_language = os.getenv("AUDIT_LANGUAGE", "Chinese")
    rules_text = format_global_rules_for_extraction(rules)
    return PROMPT_TEMPLATES["global_extract_system"].format(
        rules_text=rules_text,
        output_language=output_language
    )


def build_global_extract_user_prompt(block: dict) -> str:
    """
    Build the user prompt for global information extraction.

    Args:
        block: Text block to extract information from

    Returns:
        User prompt string
    """
    block_text = format_block_for_prompt(block)
    return PROMPT_TEMPLATES["global_extract_user"].format(block_text=block_text)


# ============================================================
# Global Verification Prompts
# ============================================================

def build_global_verify_system_prompt(rule: dict) -> str:
    """
    Build the system prompt for global consistency verification.

    This prompt instructs the LLM to compare extracted items and identify
    any inconsistencies, conflicts, or contradictions based on the rule.

    Args:
        rule: Single global rule with verification criteria

    Returns:
        System prompt string
    """
    output_language = os.getenv("AUDIT_LANGUAGE", "Chinese")
    rule_id = rule.get("id", "")
    topic = rule.get("topic", "")
    extraction = rule.get("extraction", "")
    extracted_entity = rule.get("extracted_entity", "")
    verification = rule.get("verification", "")
    return PROMPT_TEMPLATES["global_verify_system"].format(
        rule_id=rule_id,
        topic=topic,
        extraction=extraction,
        extracted_entity=extracted_entity,
        verification=verification,
        output_language=output_language
    )


def build_global_verify_user_prompt(rule: dict, items: list) -> str:
    """
    Build the user prompt for global consistency verification.

    Args:
        rule: Global rule being verified
        items: List of extracted items to check for consistency

    Returns:
        User prompt string
    """
    payload = {
        "rule_id": rule.get("id", ""),
        "items": items
    }
    payload_text = json.dumps(payload, ensure_ascii=False, indent=2)
    return PROMPT_TEMPLATES["global_verify_user"].format(payload_text=payload_text)
