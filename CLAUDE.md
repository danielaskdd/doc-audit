# CLAUDE.md

This file provides guidance to Claude Code when working with the doc-audit repository.

## Repository Overview

This repository contains the **doc-audit** skill - an LLM-powered Word document auditing system that reviews `.docx` files for compliance violations, language accuracy, technical correctness, and logical inconsistencies. Built on the [Agent Skills](https://agentskills.io/specification) specification.

Key capabilities:
- Parse Word documents into auditable text blocks with heading-based splitting
- Generate and refine audit rules using LLM (supports custom + default rules)
- Execute parallel LLM audits against rule sets
- Generate interactive HTML reports with filtering and false-positive blocking
- Apply fixes to source documents with track changes

## Repository Structure

```
doc-audit/
├── CLAUDE.md                    # This file
├── README.md                    # User-facing documentation
├── LICENSE                      # MIT License
├── requirements.txt             # Python dependencies
├── skills/
│   └── doc-audit/
│       ├── SKILL.md             # Skill instructions (activated by Claude)
│       ├── TOOLS.md             # Detailed tool documentation
│       ├── LICENSE.txt          # MIT License
│       ├── scripts/             # Core audit scripts
│       │   ├── setup_project_env.sh    # Environment initialization
│       │   ├── workflow.sh             # Complete audit workflow script
│       │   ├── parse_rules.py          # LLM-based rule generation
│       │   ├── parse_document.py       # DOCX parsing with numbering
│       │   ├── run_audit.py            # LLM audit execution
│       │   ├── generate_report.py      # HTML/Excel report generation
│       │   ├── apply_audit_edits.py    # Track changes entry point
│       │   ├── numbering_resolver.py   # Word numbering resolution
│       │   ├── table_extractor.py      # Table extraction utilities
│       │   ├── docx_edit/              # DOCX editing package (used by apply_audit_edits)
│       │   │   ├── common.py           # Shared constants, XML helpers, data classes
│       │   │   ├── navigation_mixin.py # Paragraph lookup and range navigation
│       │   │   ├── revision_mixin.py   # Track-changes revision markup
│       │   │   ├── table_edit_mixin.py # Table cell editing operations
│       │   │   ├── comment_workflow_mixin.py  # Comment insertion workflow
│       │   │   └── workflow_mixin.py   # High-level edit orchestration
│       │   └── omml/                   # Office Math ML utilities
│       │       ├── ommlparser.py       # OMML-to-text parser
│       │       ├── cleaners.py         # Math text cleanup routines
│       │       └── utils.py            # Shared OMML helpers
│       └── assets/
│           ├── default_rules.json      # Default audit rules
│           └── report_template.html    # Jinja2 HTML template
├── spec/                        # Agent Skills specification reference
├── tests/                       # Test files
│   ├── test_apply_audit_edits.py
│   └── test_fault_recovery.py
└── memory-bank/                 # Project context files
```

**Note**: There is no `template/` or `.claude-plugin/` directory in this repository.

## Environment Setup

### Prerequisites

- **[uv](https://github.com/astral-sh/uv)** - Fast Python package installer
- Python 3.12+
- Google Gemini API key OR OpenAI API key
- Word documents created in Microsoft Word 2013+ (requires `w14:paraId` attributes)

### First-Time Setup

```bash
# 1. Set API key (choose one)
export GOOGLE_API_KEY="your_gemini_key"
# OR
export OPENAI_API_KEY="your_openai_key"

# 2. Initialize environment
bash skills/doc-audit/scripts/setup_project_env.sh

# 3. Activate environment (required before running scripts)
source .claude-work/doc-audit/env.sh
```

This creates:
- `.claude-work/venv/` - Python virtual environment
- `.claude-work/doc-audit/` - Working directory for audit files
- `.claude-work/doc-audit/env.sh` - Environment activation script

### Environment Variables

```bash
# API Keys (required - one of these)
export GOOGLE_API_KEY="your_key"      # Gemini (preferred)
export OPENAI_API_KEY="your_key"      # OpenAI (gpt-4o-2024-08-06+ required)

# Model Configuration (optional)
export DOC_AUDIT_GEMINI_MODEL="gemini-3-flash"
export DOC_AUDIT_OPENAI_MODEL="gpt-5.2"

# Output Language (optional)
export AUDIT_LANGUAGE="Chinese"       # Default: Chinese
```

## Common Commands

### Complete Audit Workflow (Recommended)

```bash
# First time: use relative path (auto-initializes working directory)
skills/doc-audit/scripts/workflow.sh document.docx

# After setup: can use $DOC_AUDIT_SKILL_PATH
source .claude-work/doc-audit/env.sh
$DOC_AUDIT_SKILL_PATH/scripts/workflow.sh document.docx

# With custom rules
$DOC_AUDIT_SKILL_PATH/scripts/workflow.sh document.docx -r custom_rules.json
```

**Note**: workflow.sh automatically runs setup if the working directory doesn't exist.

Output: `<document_directory>/<docname>_audit_report.html`

### Individual Scripts

Always run `source .claude-work/doc-audit/env.sh` first, then use `$DOC_AUDIT_SKILL_PATH`:

```bash
# Generate custom rules (merges with defaults)
python $DOC_AUDIT_SKILL_PATH/scripts/parse_rules.py \
  --input "Check for ambiguous payment terms" \
  --output .claude-work/doc-audit/mydoc_custom_rules.json

# Parse document into text blocks
python $DOC_AUDIT_SKILL_PATH/scripts/parse_document.py document.docx

# Run audit with resume support
python $DOC_AUDIT_SKILL_PATH/scripts/run_audit.py \
  --document blocks.jsonl \
  --rules rules.json \
  --resume

# Generate HTML report
python $DOC_AUDIT_SKILL_PATH/scripts/generate_report.py manifest.jsonl \
  --template $DOC_AUDIT_SKILL_PATH/assets/report_template.html \
  --rules rules.json \
  --output audit_report.html

# Apply edits with track changes
python $DOC_AUDIT_SKILL_PATH/scripts/apply_audit_edits.py manifest.jsonl
```

### Running Tests

```bash
# Activate development environment
source .venv/bin/activate

# Run all tests
pytest tests/

# Run specific test file
pytest tests/test_apply_audit_edits.py -v
pytest tests/test_fault_recovery.py -v
```

## Architecture

### Audit Workflow Phases

```
Phase 0 (Setup - First Time Only):
  setup_project_env.sh → env.sh activation → Ready

Phase 1 (Rule Selection):
  Path A: Default rules → Skip to Phase 2
  Path B: parse_rules.py → User confirms → Phase 2

Phase 2 (Parse and Audit):
  parse_document.py → <docname>_blocks.jsonl
  run_audit.py → <docname>_manifest.jsonl
  generate_report.py → <docname>_audit_report.html

Phase 3 (Apply Results):
  Path A: apply_audit_edits.py manifest.jsonl → _edited.docx
  Path B: Review HTML → Export JSONL → apply_audit_edits.py → _edited.docx
```

### Key Data Structures

**Text Block** (output of parse_document.py):
```json
{
  "uuid": "12AB34CD",
  "uuid_end": "56EF78AB",
  "heading": "2.1 Penalty Clause",
  "content": "If Party B delays payment...",
  "type": "text",
  "parent_headings": ["Chapter 2 Contract Terms"]
}
```

**Audit Rule** (input to run_audit.py):
```json
{
  "id": "R001",
  "description": "Check for vague monetary amounts",
  "severity": "high",
  "category": "semantic",
  "examples": {
    "violation": "approximately 10% of the total",
    "correction": "exactly 10% of the contract amount (RMB)"
  }
}
```

**Audit Result** (output of run_audit.py):
```json
{
  "uuid": "12AB34CD",
  "uuid_end": "56EF78AB",
  "is_violation": true,
  "violations": [{
    "rule_id": "R002",
    "violation_text": "approximately 1%",
    "violation_reason": "Vague term without currency",
    "fix_action": "replace",
    "revised_text": "1% of contract total (CNY)"
  }]
}
```

### Script Dependencies

```
parse_rules.py ─────────────────────────────────────┐
                                                     │
parse_document.py ──→ _blocks.jsonl ──┐             │
                                       ├──→ run_audit.py ──→ _manifest.jsonl
                default_rules.json ────┘                           │
                                                                   │
                                 report_template.html ─────────────┤
                                                                   │
                                      generate_report.py ←─────────┘
                                             │
                                    _audit_report.html
                                             │
                                    apply_audit_edits.py
                                             │
                                       _edited.docx
```

Helper modules (imported by main scripts):
- `numbering_resolver.py` - Resolves Word numbering definitions to labels
- `table_extractor.py` - Extracts tables with structure preservation
- `docx_edit/` - Modular DOCX editing package (mixins composed into `DocxEditor` class)
- `omml/` - Office Math Markup Language parsing and text conversion

## Important Conventions

### Script Usage

- All scripts support `--help` for usage information
- Always activate environment before running: `source .claude-work/doc-audit/env.sh`
- Use `$DOC_AUDIT_SKILL_PATH` to reference script paths
- Scripts use absolute paths internally

### Word Document Requirements

- Must be `.docx` format (not `.doc`, `.pdf`, etc.)
- Created/saved in Microsoft Word 2013+ (requires `w14:paraId` attributes)
- Use proper Heading Styles (Heading 1, Heading 2, etc.) for structure
- Documents from LibreOffice/Google Docs must be opened and saved in Word first

### File Naming

- Intermediate files: `.claude-work/doc-audit/<docname>_blocks.jsonl`, `<docname>_manifest.jsonl`
- Output files: Same directory as source document
- Custom rules: `<docname>_custom_rules.json`

### Error Handling

- If `parse_document.py` fails with missing `paraId`, stop workflow and inform user
- Resume interrupted audits with `--resume` flag
- Check `.claude-work/logs/` for operation logs

## Development Guidelines

### Modifying Scripts

1. Scripts use `defusedxml` for secure XML parsing (not `xml.etree`)
2. All LLM calls use structured outputs for reliable JSON responses
3. Test changes with both Gemini and OpenAI backends
4. Run `pytest tests/` before committing

### Adding Features

1. Keep SKILL.md concise - it's loaded into Claude's context
2. Document complex logic in TOOLS.md
3. Support `--help` in new scripts
4. Follow existing patterns for error messages

### Testing Approach

- `test_apply_audit_edits.py` - Tests document editing with track changes
- `test_fault_recovery.py` - Tests resume and error recovery
- Test with real Word documents when possible

### Commit Style

Follow conventional commits. Recent examples:
```
Fix numbering state reset between table cells in extractor
Simplify violation processing to use only new violations array format
refactor: move source content to modal with deduplicated storage
```
