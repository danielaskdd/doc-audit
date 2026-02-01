---
name: doc-audit
description: Document Audit Skill - Proficient semantic verification and customizable rules for Microsoft Word (.docx) files.
type: active
version: 1.0.0
---

# Document Audit Skill

**This is an ACTIVE skill** - Uses Python scripts with python-docx to parse DOCX documents and LLM to perform intelligent auditing. Only supports .docx format (not .doc, .pdf, or other formats)

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

Run setup to create working directory and install dependencies:

```bash
bash skills/doc-audit/scripts/setup_project_env.sh
source .claude-work/doc-audit/env.sh
```

This creates:
- `.claude-work/doc-audit/env.sh` - Environment activation script
- `.claude-work/venv/` - Python virtual environment (shared across skills)
- `.claude-work/logs/` - Operation logs (shared across skills)

> **Note:** `workflow.sh` can auto-initialize if `.claude-work/doc-audit/` doesn't exist, but explicit setup is recommended for first-time use to verify environment configuration.

### Phase 1: Rule Selection

**Decision Point:** What audit rules does the user want?

**Path A: Use Default Rules (Simple)**

1. User only requests "audit [filename]" without specific requirements
2. **Skip rule generation** - use `$DOC_AUDIT_SKILL_PATH/assets/default_rules.json`
3. Proceed immediately to Phase 2

**Path B: Custom Rules (Iterative)**

1. **Analyze Requirements** - Agent converts user's needs into clear criteria

2. **Generate Rules** - Invoke `parse_rules.py` to generate customized rules from user requirements

3. **User Confirmation** - After generating rules, you **MUST**:
   - Use `read_file` to read the generated rules file (`.claude-work/doc-audit/<docname>_custom_rules.json`)
   - Present ALL rules to user in the following simplified format:
     ```
     [R001] Rule description...
     [R002] Rule description...
     [R003] Rule description...
     ...
     Total: N rules
     ```
   - Ask user explicitly: Please review the rules above. Approve to continue audit? Or need modifications?
   - **DO NOT proceed to Phase 2 until user explicitly confirms approval**

4. **Iterate if Needed** - Upon receiving a user request to amend any rules, invoke parse_rules.py with the user's input and specify the generated rules file using the `--base-rules` flag.

5. Once rules are confirmed, proceed to Phase 2

**Path C: Use Additional Rule Sets (Multi-Rules)**

When user requests using specific rule file(s) (e.g., "use bidding_rules to audit"):

1. **Find Rule Files** - Search in this order:
   - `skills/doc-audit/assets/<filename>.json` - Predefined rule sets
   - `.claude-work/doc-audit/<filename>.json` - Working directory
   - Current directory - User-provided files
   - Absolute/relative paths - As specified by user

2. **Determine Merge Mode**:
   - **Default behavior (Merge)**: Automatically include `default_rules.json` + user-specified rules
     - User says: "use bidding rules", "add 招标规则", "also check with X"
   - **Exclude default rules**: Only when user explicitly says "only/just/仅用/只使用"
     - User says: "only use bidding rules", "just use X, no default rules", "仅使用招标规则"

3. **Verify and Proceed**:
   - Confirm all rule files are found
   - Show user which rules will be used
   - Proceed to Phase 2

### Phase 2: Execute Audit

After rules are determined in Phase 1, execute the audit using `workflow.sh`:

```bash
# Path A: Default rules only
$DOC_AUDIT_SKILL_PATH/scripts/workflow.sh document.docx

# Path B: With custom rules
$DOC_AUDIT_SKILL_PATH/scripts/workflow.sh document.docx -r .claude-work/doc-audit/<docname>_custom_rules.json

# Path C: With additional rule sets (auto-merge with defaults)
$DOC_AUDIT_SKILL_PATH/scripts/workflow.sh document.docx -r $DOC_AUDIT_SKILL_PATH/assets/bidding_rules.json

# Path C: Exclude default rules (only specified rules)
$DOC_AUDIT_SKILL_PATH/scripts/workflow.sh document.docx --rules-only -r custom_rules.json
```

`workflow.sh` automatically performs:
1. **Parse document** → `.claude-work/doc-audit/<docname>_blocks.jsonl`
2. **Run audit** → `.claude-work/doc-audit/<docname>_manifest.jsonl`
3. **Generate report** → `<document_directory>/<docname>_audit_report.html`

⚠️ **Error handling**: If parsing fails (e.g., missing paraId error), **stop the workflow immediately** and inform the user.

**Fallback (Manual Execution)**: If workflow.sh fails or you need finer control, run individual scripts:

```bash
# Step 1: Parse document
python $DOC_AUDIT_SKILL_PATH/scripts/parse_document.py document.docx

# Step 2: Run audit (use the same rules as workflow.sh would)
python $DOC_AUDIT_SKILL_PATH/scripts/run_audit.py \
  --document .claude-work/doc-audit/<docname>_blocks.jsonl \
  --rules $DOC_AUDIT_SKILL_PATH/assets/default_rules.json \
  --output .claude-work/doc-audit/<docname>_manifest.jsonl

# Step 3: Generate report
python $DOC_AUDIT_SKILL_PATH/scripts/generate_report.py \
  --manifest .claude-work/doc-audit/<docname>_manifest.jsonl \
  --template $DOC_AUDIT_SKILL_PATH/assets/report_template.html \
  --rules $DOC_AUDIT_SKILL_PATH/assets/default_rules.json \
  --output <document_directory>/<docname>_audit_report.html
```

📖 For detailed parameters (resume, parallel workers, etc.), see [Available Tools](#available-tools) section below.

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

1. **Review Results** - User opens HTML audit report in browser
   - Location: `<document_directory>/<document_name>_audit_report.html`
   - User can review each issue with source context

2. **Block Unwanted Results** - User marks unreasonable or unwanted items as "blocked" (屏蔽)
   - Click "屏蔽本条" checkbox on each item to exclude
   - Blocked items are visually dimmed and excluded from export

3. **Export Control File** - User clicks "导出结果" button
   - Output: `<document_name>_audit_export.jsonl` (downloaded to browser's download folder)
   - First line contains metadata with `source_file` and `source_hash`
   - Only non-blocked items are included

4. **Apply Edits** - Use the exported control file to apply changes:
   ```bash
   python $DOC_AUDIT_SKILL_PATH/scripts/apply_audit_edits.py <document_name>_audit_export.jsonl
   ```
   - The control file's metadata line contains the source document path, so only the control file path is needed
   - Output: `<original_document>_edited.docx` with track changes and comments

```
Phase 0 (Setup - First Time Only):
  setup_project_env.sh → source env.sh → [User sets API key] → Ready

Phase 1 (Rule Selection):
  Path A: User: "Audit file.docx" → Use default_rules.json → Phase 2
  Path B: User: "Check for X, Y" → parse_rules.py → User confirms → Phase 2
  Path C: User: "Use bidding_rules" → Find rule files → Phase 2

Phase 2 (Execute Audit):
  workflow.sh document.docx [-r rules.json]
    → Parse → Audit → Report
    → Output: <docname>_audit_report.html (same directory as source)

Phase 3 (Apply Results):
  Path A (Direct): manifest.jsonl → apply_audit_edits.py → _edited.docx
  Path B (Review): HTML Report → Block/Export → apply_audit_edits.py → _edited.docx
```

## File Location Rules

**Working Directory** (`.claude-work/doc-audit/`):
- All intermediate files generated during audit process
- Custom rules generated by `parse_rules.py` → `<docname>_custom_rules.json`
- Parsed document blocks → `<docname>_blocks.jsonl`
- Audit manifest → `<docname>_manifest.jsonl`

**Source Document Directory**:
- Final HTML report → `<docname>_audit_report.html`
- Final Excel report → `<docname>_audit_report.xlsx`
- Edited document with track changes → `<docname>_edited.docx`

**Assets Directory** (`$DOC_AUDIT_SKILL_PATH/assets/`) - Read-only:
- Report template → `report_template.html`
- Default audit rules → `default_rules.json`
- Additional audit rulesets → `bidding_rules.json`, `contract_rules.json`, etc.

## Available Tools

> **Note:** `workflow.sh` handles setup automatically. The following manual setup is only needed when using individual scripts (parse_document.py, run_audit.py, etc.) directly.

### 1. Environment Setup (Required for Individual Scripts)

If using individual scripts instead of `workflow.sh`, setup the environment first:

```bash
bash skills/doc-audit/scripts/setup_project_env.sh
source ./.claude-work/doc-audit/env.sh
# $DOC_AUDIT_SKILL_PATH is now set, can use individual scripts
```

**What it creates:**

- `.claude-work/venv/` - Python virtual environment (shared across skills)
- `.claude-work/logs/` - Operation logs (shared across skills)
- `.claude-work/doc-audit/` - Document audit working directory
- `.claude-work/doc-audit/env.sh` - Environment activation script

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

Use LLM to generate audit rules from user requirements. Can start from scratch or build upon existing rulesets.

**Default Behavior**: Creates new rules without loading default rules. To extend default rules, explicitly specify `--base-rules`.

```bash
# Generate rules from scratch (most common)
python $DOC_AUDIT_SKILL_PATH/scripts/parse_rules.py \
  --input "Check for ambiguous payment terms and missing signatures" \
  --output .claude-work/doc-audit/mydoc_custom_rules.json

# Extend default rules (when user wants to add to defaults)
python $DOC_AUDIT_SKILL_PATH/scripts/parse_rules.py \
  --base-rules $DOC_AUDIT_SKILL_PATH/assets/default_rules.json \
  --input "Add rule for checking ambiguous references" \
  --output .claude-work/doc-audit/mydoc_custom_rules.json

# Iterative refinement (modify existing custom rules)
python $DOC_AUDIT_SKILL_PATH/scripts/parse_rules.py \
  --base-rules .claude-work/doc-audit/mydoc_custom_rules.json \
  --input "Remove R009, make signature rule more specific" \
  --output .claude-work/doc-audit/mydoc_custom_rules.json
```

📖 **Detailed parameters, decision guide, LLM configuration, and output format**: See [TOOLS.md - Parse Rules](TOOLS.md#Parse-Rules)

### 3. Workflow Script (Recommended for Phase 2)

`workflow.sh` runs the complete audit pipeline: parse → audit → report. **Use this after Phase 1 (rule selection) is complete.**

> **Important:** `workflow.sh` combines parse + audit + report into a single command. It should be used in Phase 2 after determining which rules to use.

**Auto-initialization**: If `.claude-work/doc-audit/` doesn't exist, workflow.sh automatically runs setup. However, explicit setup (Phase 0) is recommended for first-time use.

```bash
# Standard usage (after Phase 0 setup and Phase 1 rule selection)
source .claude-work/doc-audit/env.sh
$DOC_AUDIT_SKILL_PATH/scripts/workflow.sh document.docx

# With custom rules (from Phase 1 Path B)
$DOC_AUDIT_SKILL_PATH/scripts/workflow.sh document.docx -r .claude-work/doc-audit/<docname>_custom_rules.json

# With additional rule set (Phase 1 Path C - auto-merge with defaults)
$DOC_AUDIT_SKILL_PATH/scripts/workflow.sh document.docx -r $DOC_AUDIT_SKILL_PATH/assets/bidding_rules.json

# With multiple additional rule sets (auto-merge)
$DOC_AUDIT_SKILL_PATH/scripts/workflow.sh document.docx -r $DOC_AUDIT_SKILL_PATH/assets/bidding_rules.json -r $DOC_AUDIT_SKILL_PATH/assets/contract_rules.json

# Exclude default rules (only specified rules)
$DOC_AUDIT_SKILL_PATH/scripts/workflow.sh document.docx --rules-only -r .claude-work/doc-audit/<docname>_custom_rules.json
```

**Rule File Search Order**:
1. `.claude-work/doc-audit/` - Working directory
2. `skills/doc-audit/assets/` - Predefined rule sets
3. Current directory
4. Absolute/relative paths as specified

**What it does**:

1. Parse document → `<docname>_blocks.jsonl`
2. Run audit with merged rules → `<docname>_manifest.jsonl`
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

⚠️ **Resume Critical Note**: When resuming an interrupted audit, you **MUST** specify the correct `--output` path pointing to the existing manifest file (e.g., `.claude-work/doc-audit/<docname>_manifest.jsonl`). If omitted, the script defaults to `manifest.jsonl` in the current directory, which effectively restarts the audit from scratch.

```bash
# Basic usage with default rules
python $DOC_AUDIT_SKILL_PATH/scripts/run_audit.py \
  --document .claude-work/doc-audit/<docname>_blocks.jsonl \
  --rules $DOC_AUDIT_SKILL_PATH/assets/default_rules.json

# Use multiple rule files (auto-merge)
python $DOC_AUDIT_SKILL_PATH/scripts/run_audit.py \
  --document .claude-work/doc-audit/<docname>_blocks.jsonl \
  -r $DOC_AUDIT_SKILL_PATH/assets/default_rules.json \
  -r $DOC_AUDIT_SKILL_PATH/assets/bidding_rules.json

# Resume from interruption (MUST specify the same output file as the original run)
python $DOC_AUDIT_SKILL_PATH/scripts/run_audit.py \
  --document .claude-work/doc-audit/<docname>_blocks.jsonl \
  --rules $DOC_AUDIT_SKILL_PATH/assets/default_rules.json \
  --output .claude-work/doc-audit/<docname>_manifest.jsonl \
  --resume

# Increase parallelism for faster processing
python $DOC_AUDIT_SKILL_PATH/scripts/run_audit.py \
  --document .claude-work/doc-audit/<docname>_blocks.jsonl \
  --rules $DOC_AUDIT_SKILL_PATH/assets/default_rules.json \
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
# Basic usage with single rule file
python $DOC_AUDIT_SKILL_PATH/scripts/generate_report.py \
  --manifest manifest.jsonl \
  --template $DOC_AUDIT_SKILL_PATH/assets/report_template.html \
  --rules rules.json \
  --output audit_report.html

# Use multiple rule files (auto-merge)
python $DOC_AUDIT_SKILL_PATH/scripts/generate_report.py \
  -m manifest.jsonl \
  -t $DOC_AUDIT_SKILL_PATH/assets/report_template.html \
  -r $DOC_AUDIT_SKILL_PATH/assets/default_rules.json -r $DOC_AUDIT_SKILL_PATH/assets/bidding_rules.json \
  -o audit_report.html
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
│   ├── workflow.sh             # Complete audit workflow script
│   ├── parse_rules.py          # Rule parsing
│   ├── parse_document.py       # DOCX parsing (python-docx)
│   ├── run_audit.py            # LLM audit execution
│   ├── generate_report.py      # Report generation
│   └── apply_audit_edits.py    # Apply audit edits to Word document
└── assets/
    ├── default_rules.json      # Default audit rules
    ├── bidding_rules.json      # Additional audit rules for bidding document
    └── report_template.html    # Jinja2 report template

# Working directory (created by setup script - intermediate files only)
.claude-work/
├── venv/                                 # Python virtual environment (shared across skills)
├── logs/                                 # Operation logs (shared across skills)
└── doc-audit/                            # Document audit working directory
    ├── env.sh                            # Environment activation script
    ├── README.md                         # Working directory documentation
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
