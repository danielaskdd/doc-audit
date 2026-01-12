---
name: doc-audit
description: Intelligent document audit system for compliance review, legal or technical document verification, and engineering document validation using LLM
type: active
version: 1.0.0
---

# Document Audit Skill

**This is an ACTIVE skill** - Uses Python scripts with python-docx to parse DOCX documents and LLM to perform intelligent auditing.

## When to Use This Skill

Use this skill when you need to:
- Audit Word documents (.docx) for compliance with specific rules
- Verify legal or technical documents for language accuracy and consistency
- Review engineering specifications for technical correctness
- Check documents for typos, grammar errors, unclear references, and logical inconsistencies
- Generate detailed audit reports with issue tracing

## Core Workflow

The doc-audit skill supports two workflow paths depending on whether user has specific audit requirements:

### Phase 0: Environment Setup (First Time Only)

Before running any audit, set up the project environment:

```bash
bash skills/doc-audit/scripts/setup_project_env.sh
source .claude-work/doc-audit/env.sh
```

This creates:
- `.claude-work/env.sh` - Script for setting up python virtual environment and environment variables
- `.claude-work/doc-audit/` - Directory for all doc-audit files (env, scripts, intermediate files)
- `.claude-work/venv/` - Python virtual environment (shared across skills)
- `.claude-work/logs/` - Operation logs (shared across skills)

### Phase 1: Rule Selection

**Decision Point:** Does user specify custom audit requirements?

**Path A: Use Default Rules (Simple)**

- User only requests "audit [filename]" without specific requirements
- **Skip rule generation** - use `.claude-work/doc-audit/default_rules.json` (copied from `assets/default_rules.json` during enviroment setup)
- Proceed immediately to Phase 2

**Path B: Custom Rules (Iterative)**

1. **Analyze Requirements** - Agent converts user's needs into clear criteria

2. **Generate Rules** - Invoke `parse_rules.py` to generate customized rules by merging them with the default rules

   ⚠️ **CRITICAL**: Do NOT use the `--no-base` flag unless the user explicitly requests to exclude default rules. The default behavior is to merge user requirements WITH base rules.

3. **User Confirmation** - ⚠️ **MANDATORY STEP - DO NOT SKIP**:

   After generating rules, you **MUST**:
   - Use `read_file` to read the generated rules file (`.claude-work/doc-audit/<docname>_custom_rules.json`)
   - Present ALL rules to user in the following simplified format:
     ```
     [R001] Rule description...
     [R002] Rule description...
     [R003] Rule description...
     ...
     Total: N rules
     ```
   - Ask user explicitly: "请审阅以上规则。是否批准继续审计？或需要修改规则？" (Please review the rules above. Approve to continue audit? Or need modifications?)
   - **DO NOT proceed to Phase 2 until user explicitly confirms approval**

4. **Iterate if Needed** - If user requests changes, refine rules using `parse_rules.py` again, then return to step 3 for re-confirmation

### Phase 2: Parse and Audit

5. **Parse Document** - Extract text blocks from .docx with proper numbering (python-docx)
   - Output: `.claude-work/doc-audit/<docname>_blocks.jsonl` (with document name prefix)
   - ⚠️ **Error handling**: If `parse_document.py` fails (e.g., missing paraId error), **stop the workflow immediately** and inform the user. Do NOT proceed to step 6.
6. **Execute Audit Work Flow** - LLM audits each text block against rules by `workflow.sh` (created by enviroment setup)
   - Intermediate: `.claude-work/doc-audit/<docname>_manifest.jsonl`
   - Output: `<document_directory>/<document_name>_audit_report.html` (same directory as source document)

### Phase 3: Review and Apply Audit Results

After Phase 2 generates the HTML audit report, user can review the results and apply them to the source document. Tell user to check the final report location or directly apply the edits without review.

**Path A: Direct Apply (Skip Review)**

- User doesn't need to review or filter audit results
- Apply directly from Phase 2 manifest file:
  ```bash
  python $DOC_AUDIT_SKILL_PATH/scripts/apply_audit_edits.py .claude-work/doc-audit/<docname>_manifest.jsonl
  ```
- Output: `<docname>_edited.docx` with track changes and comments

**Path B: Review and Apply (Recommended)**

7. **Review Results** - User opens HTML audit report in browser
   - Location: `<document_directory>/<document_name>_audit_report.html`
   - User can review each issue with source context

8. **Block Unwanted Results** - User marks unreasonable or unwanted items as "blocked" (屏蔽)
   - Click "屏蔽本条" checkbox on each item to exclude
   - Blocked items are visually dimmed and excluded from export

9. **Export Control File** - User clicks "导出结果" button
   - Output: `<document_name>_audit_export.jsonl` (downloaded to browser's download folder)
   - First line contains metadata with `source_file` and `source_hash`
   - Only non-blocked items are included

10. **Apply Edits** - Use the exported control file to apply changes:
    ```bash
    python $DOC_AUDIT_SKILL_PATH/scripts/apply_audit_edits.py <document_name>_audit_export.jsonl
    ```
    - The control file's metadata line contains the source document path, so only the control file path is needed
    - Output: `<original_document>_edited.docx` with track changes and comments

```
Phase 0 (Setup - First Time Only):
Environment Setup → [User sets API key] → Ready to Audit

Path A (Default Rules):
User: "Audit file.docx" → Parse Document → Audit (default rules) → Report

Path B (Custom Rules):
User: "Check for X, Y, Z" → Generate Rules → Present ───┐
                              ↑                         │
                              └─ (Modify) ←─ Review ────┘ (User confirms)
                                               │
                                          (User Approves)
                                               ↓
                                          Parse Document → Execute Audit
                                          
Phase 3 (Apply Results):

Path A (Direct Apply - Skip Review):
Manifest.jsonl → apply_audit_edits.py → <origin_file_name>_edited.docx

Path B (Reviewed Apply - Recommended):
Open HTML Report → Block unwanted → Export JSONL → apply_audit_edits.py → _edited.docx

Final Report Location: Same directory as source document (<filename>_audit_report.html)
```

## Available Tools

> **Note:** All script examples below use `$DOC_AUDIT_SKILL_PATH` environment variable, which is automatically set by `source .claude-work/doc-audit/env.sh`. Always run `source .claude-work/doc-audit/env.sh` before executing any scripts.

### 1. Environment Setup (First Time Only)

Setup the project environment before running any audit:

```bash
bash skills/doc-audit/scripts/setup_project_env.sh
source ./.claude-work/doc-audit/env.sh
```

**What it creates:**

- `.claude-work/venv/` - Python virtual environment (shared across skills)
- `.claude-work/logs/` - Operation logs (shared across skills)
- `.claude-work/doc-audit/` - Document audit working directory
- `.claude-work/doc-audit/env.sh` - Environment activation script
- `.claude-work/doc-audit/workflow.sh` - Convenience workflow script
- `.claude-work/doc-audit/default_rules.json` - Default audit rules (copied from assets)
- `.claude-work/doc-audit/report_template.html` - Report template (copied from assets)

**Installed packages:**
- `python-docx` - DOCX parsing
- `lxml` - XML parsing
- `defusedxml` - Defused XML parsing
- `jinja2` - HTML templating
- `openpyxl` - Excel report generation
- `google-genai` - Google Gemini LLM
- `openai` - OpenAI LLM

**Note:**
- User must set `GOOGLE_API_KEY` or `GOOGLE_APPLICATION_CREDENTIALS`or `OPENAI_API_KEY` environment variable before running audits.
- Always run `source ./.claude-work/doc-audit/env.sh` before executing any scripts.

### 2. Generate Customized Rules (Iterative)

Intelligently merge base rules with user requirements using LLM.

⚠️ **CRITICAL**: Do NOT use `--no-base` flag unless user explicitly requests to exclude default rules. Default behavior merges with base rules.

```bash
# Initial generation (merges with default rules)
python $DOC_AUDIT_SKILL_PATH/scripts/parse_rules.py \
  --input "Check for ambiguous payment terms and missing signatures" \
  --output .claude-work/doc-audit/mydoc_custom_rules.json

# Iterative refinement (continues from previous output)
python $DOC_AUDIT_SKILL_PATH/scripts/parse_rules.py \
  --base-rules .claude-work/doc-audit/mydoc_custom_rules.json \
  --input "Remove R009, make signature rule more specific" \
  --output .claude-work/doc-audit/mydoc_custom_rules.json
```

📖 **Detailed parameters, decision guide, LLM configuration, and output format**: See [TOOLS.md - Parse Rules](TOOLS.md#Parse-Rules)

### 3. Workflow Script (Recommended for Normal Audit Workflow)

`workflow.sh` runs the complete audit pipeline: parse → audit → report. **This is the recommended way to perform audits instead of involving each tool separately  .**

```bash
# Use default rules
./.claude-work/doc-audit/workflow.sh document.docx

# Use custom rules
./.claude-work/doc-audit/workflow.sh document.docx custom_rules.json
```

**What it does**:

1. Parse document → `<docname>_blocks.jsonl`
2. Run audit → `<docname>_manifest.jsonl`
3. Generate reports → `<document_name>_audit_report.html` and `<document_name>_audit_report.xlsx` (saved alongside source document)

**Note**: If workflow fails, use individual tools (Parse, Audit, Report) to debug or continue manually.

📖 **Internal process details**: See [TOOLS.md - Workflow Script](TOOLS.md#Workflow-Script)

### 4. Parse

Extract text blocks from a Word document. **Use independently only when workflow.sh cannot be applied.**

```bash
# Basic usage
python $DOC_AUDIT_SKILL_PATH/scripts/parse_document.py document.docx
```

⚠️ **Error handling**: If parsing fails due to missing `w14:paraId` (Word 2013+ required), stop workflow and inform user.

📖 **Detailed parameters, features, and output format**: See [TOOLS.md - Parse Document](TOOLS.md#Parse-Document)

### 5. Audit

Execute LLM-based audit on each text block against audit rules. **Use this audit script independently only in cases where the workflow script cannot be applied.**

**Independent use cases**:

- Debugging audit behavior with `--dry-run`
- Processing large documents in chunks (`--start-block`, `--end-block`)
- Resuming interrupted runs (`--resume`)
- Custom model selection (`--model`)
- Adjusting parallelism (`--workers`)

```bash
# Basic usage
python $DOC_AUDIT_SKILL_PATH/scripts/run_audit.py \
  --document blocks.jsonl \
  --rules rules.json

# Resume from interruption
python $DOC_AUDIT_SKILL_PATH/scripts/run_audit.py \
  --document blocks.jsonl \
  --rules rules.json \
  --resume

# Increase parallelism for faster processing
python $DOC_AUDIT_SKILL_PATH/scripts/run_audit.py \
  --document blocks.jsonl \
  --rules rules.json \
  --workers 12
```

📖 **Detailed parameters, parallel processing, resume functionality, and advanced use cases**: See [TOOLS.md - Run Audit](TOOLS.md#Run-Audit)

### 6. Report

Generate interactive HTML audit report from manifest. **Use this report script independently only in cases where the workflow script cannot be applied.** (see tool #6 below).

**Independent use cases**:

- Re-generating reports after template modifications
- Custom output locations
- JSON export for further processing (`--json`)
- Excel export for spreadsheet review (`--excel`)

```bash
# Basic usage
python $DOC_AUDIT_SKILL_PATH/scripts/generate_report.py manifest.jsonl \
  --template .claude-work/doc-audit/report_template.html \
  --rules rules.json \
  --output audit_report.html
```

**Key features**: Interactive filters, issue blocking, export to JSONL, rule details in modals.

📖 **Detailed parameters and features**: See [TOOLS.md - Generate Report](TOOLS.md#Generate-Report)

### 7. Edits

Apply audit results to the original word document using track changes and comments. This post-processing should be performed only upon user request.

**Quick Usage:**

```bash
# From manifest (skip review)
python $DOC_AUDIT_SKILL_PATH/scripts/apply_audit_edits.py .claude-work/doc-audit/<docname>_manifest.jsonl

# From exported control file (after review)
python $DOC_AUDIT_SKILL_PATH/scripts/apply_audit_edits.py <docname>_audit_export.jsonl
```

**Output:** `<source>_edited.docx` with track changes and comments

📖 **Detailed parameters, JSONL format, and error handling**: See [TOOLS.md - Apply Audit Edits](TOOLS.md#Apply-Audit-Edits)

## Technical Requirements

### Dependencies

**Core Libraries:**

- `python-docx`: DOCX parsing with lxml for XML manipulation
- `jinja2`: HTML report templating
- `openpyxl`: Excel report generation
- `google-genai` / `openai`: LLM API access

Enviroment Setup `setup_project_env.sh` will create Python venv and install all dependencies automatically.

### Environment Variables

```bash
# API Keys (required)
# For Gemini (If both Gemini and OpenAI are set, Gemini is used by default)
export GOOGLE_API_KEY=your_api_key

# For OpenAI
export OPENAI_API_KEY=your_api_key

# Model Configuration (optional - set in env.sh automatically)
# Override these to use different models across all scripts
export DOC_AUDIT_GEMINI_MODEL=gemini-3-flash-preview    # Default Gemini model
export DOC_AUDIT_OPENAI_MODEL=gpt-5.2           # Default OpenAI model

# Output Language Configuration (optional - set in env.sh automatically)
# Specifies the language for LLM-generated rules and audit results
# Examples: "Chinese", "English", "Japanese", "Korean", etc.
export AUDIT_LANGUAGE=Chinese                   # Default: Chinese
```

**⚠️ OpenAI Model Compatibility:**

When using OpenAI, the scripts use Structured Outputs (`json_schema` response format), which requires:
- ✅ `gpt-4o-2024-08-06` or later
- ✅ `gpt-4o-mini` or later
- ✅ `gpt-4o` (latest)
- ✅ `gpt-5.x` series (e.g., `gpt-5.2`)

Older models are **NOT supported** and will cause API errors. If you encounter errors like "json_schema is not supported", ensure you're using a compatible model.

**Model Configuration:**
The default models for all scripts are centralized in `.claude-work/doc-audit/env.sh`:
- **Gemini**: `gemini-3-flash-preview` (changeable via `DOC_AUDIT_GEMINI_MODEL`)
- **OpenAI**: `gpt-5.2` (changeable via `DOC_AUDIT_OPENAI_MODEL`)

```bash
# Example: Use a different model across all scripts
export DOC_AUDIT_GEMINI_MODEL="gemini-2.5-flash"
export DOC_AUDIT_OPENAI_MODEL="gpt-4o"  # or gpt-5.2, gpt-4o-mini, etc.
```

### Failure handling

If a required package or API key is missing, do not proceed with the workflow. Provide the exact `uv pip install ...` command(s) and the `export ...` command(s) needed to prepare the environment.

**Missing paraId Error:**

If the document is missing `w14:paraId` attributes on paragraphs, `parse_document.py` will display a user-friendly error message and exit with code 1. This typically occurs with documents created by older versions of Microsoft Word (before Office 2013), or generated programmatically. When this error occurs , the agent must stop the workflow and inform the user immediately.

## Data Structures

### Audit Rule Format

```json
{
  "id": "R001",
  "description": "Check for vague or ambiguous monetary amounts",
  "severity": "high",
  "category": "semantic",
  "examples": {
    "violation": "Party B shall pay approximately 10% of the total amount.",
    "correction": "Party B shall pay exactly 10% of the total contract amount (RMB)."
  }
}
```

### Text Block Format

```json
{
  "uuid": "12AB34CD",
  "uuid_end": "56EF78AB",
  "heading": "2.1 Penalty Clause",
  "content": "If Party B delays payment, they shall pay approximately 1% of the total amount as compensation.",
  "type": "text",
  "parent_headings": ["Chapter 2 Contract Terms"]
}
```

**UUID Range Fields:**
- `uuid`: 8-character hex ID (from `w14:paraId`) of the first paragraph in the block
- `uuid_end`: 8-character hex ID of the last paragraph in the block (for tables, uses the last cell's last paragraph)

**Note:** `parent_headings` contains only the ancestor headings hierarchy, not the current heading (which is in the `heading` field).

### Audit Result Format

The manifest entry written by `run_audit.py` contains audit results with actionable fix information:

```json
{
  "uuid": "12AB34CD",
  "uuid_end": "56EF78AB",
  "p_heading": "2.1 Penalty Clause",
  "p_content": "If Party B delays payment, they shall pay approximately 1% of the total amount as compensation.",
  "is_violation": true,
  "violations": [
    {
      "rule_id": "R002",
      "category": "semantic",
      "uuid": "12AB34CD",
      "uuid_end": "56EF78AB",
      "violation_text": "approximately 1% of the total amount",
      "violation_reason": "Contains vague term 'approximately' and does not specify currency",
      "fix_action": "replace",
      "revised_text": "1% of the contract total amount as penalty (settled in CNY)"
    }
  ]
}
```

**Entry-Level Fields:**
- `uuid`: 8-character hex ID of the first paragraph in the source block
- `uuid_end`: 8-character hex ID of the last paragraph in the source block

**Violation Fields:**
- `rule_id`: ID of the violated rule (e.g., "R002")
- `category`: Automatically populated by script from rule's category
- `uuid`: Injected by script - same as entry-level uuid (for apply_audit_edits.py)
- `uuid_end`: Injected by script - same as entry-level uuid_end (for range-restricted search)
- `violation_text`: Problematic text with sufficient context for unique string matching
- `violation_reason`: Explanation of why this violates the rule
- `fix_action`: Action to take - `"delete"`, `"replace"`, or `"manual"`
- `revised_text`:
  - For `"replace"`: Complete replacement text
  - For `"delete"`: Empty string
  - For `"manual"`: Guidance for human reviewer

**LLM Output:** The LLM outputs `rule_id`, `violation_text`, `violation_reason`, `fix_action`, and `revised_text` for each violation. The script adds `category`, `uuid`, and `uuid_end` by lookup and injection.

When no violations are found:
```json
{
  "uuid": "12AB34CD",
  "uuid_end": "56EF78AB",
  "p_heading": "2.1 Penalty Clause",
  "p_content": "Party B shall pay 1% of the contract amount within 30 days.",
  "is_violation": false,
  "violations": []
}
```

## Acceptance Criteria

1. **Numbering Accuracy**: All heading numbers must match the Word document display (including multi-level lists)
2. **Table Integrity**: Tables must preserve row/column relationships in JSON format
3. **Block Independence**: Each block is audited independently without cross-block interference
4. **Traceability**: Every issue can be traced back to its source heading and content

## File Structure

```
doc-audit/
├── SKILL.md                    # This file
├── scripts/
│   ├── setup_project_env.sh    # Environment setup script
│   ├── parse_rules.py          # Rule parsing
│   ├── parse_document.py       # DOCX parsing (python-docx)
│   ├── run_audit.py            # LLM audit execution
│   ├── generate_report.py      # Report generation
│   └── apply_audit_edits.py    # Apply audit edits to Word document
└── assets/
    ├── default_rules.json      # Default audit rules (source)
    └── report_template.html    # Jinja2 report template (source)

# Working directory (created by setup script - all work happens here)
.claude-work/
├── venv/                                 # Python virtual environment (shared across skills)
├── logs/                                 # Operation logs (shared across skills)
└── doc-audit/                            # Document audit working directory
    ├── env.sh                            # Environment activation script
    ├── workflow.sh                       # Convenience workflow script
    ├── README.md                         # Working directory documentation
    ├── default_rules.json                # Default rules (copied from assets)
    ├── report_template.html              # Report template (copied from assets)
    ├── <docname>_blocks.jsonl            # Parsed document blocks (per document)
    ├── <docname>_manifest.jsonl          # Audit results (per document)
    └── <docname>_custom_rules.json       # Custom rules (optional, per document)

# Output files (generated by workflow)
<document_directory>/
├── <docname>_audit_report.html           # HTML audit report (Phase 2 output)
├── <docname>_audit_report.xlsx           # Excel audit report (Phase 2 output)
└── <docname>_edited.docx                 # Edited document with track changes (Phase 3 output)

# User downloaded files (from browser)
<browser_download_folder>/
└── <docname>_audit_export.jsonl          # Exported control file (from HTML report)
```

## Limitations

- Only supports .docx format (not .doc, .pdf, or other formats)
- Each text block is audited independently - no cross-reference validation
- LLM quality depends on chosen model and rule clarity
