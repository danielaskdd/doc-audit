#!/bin/bash
# Document Audit Project Environment Setup Script
# Creates hidden working directory and Python virtual environment in current project directory

set -e

# Configuration
WORK_DIR=".claude-work"
VENV_DIR="$WORK_DIR/venv"
DOC_AUDIT_DIR="$WORK_DIR/doc-audit"
SKILL_PATH="${DOC_AUDIT_SKILL_PATH:-$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)}"

echo "=========================================="
echo "Document Audit Environment Setup"
echo "Project Directory: $(pwd)"
echo "=========================================="
echo

# 1. Create working directory structure
echo "1. Creating working directory structure..."
mkdir -p "$DOC_AUDIT_DIR"
mkdir -p "$WORK_DIR/logs"
echo "   ✓ Directory created: $DOC_AUDIT_DIR/"

# Copy default resources to working directory
cp "$SKILL_PATH/assets/default_rules.json" "$DOC_AUDIT_DIR/"
echo "   ✓ Copied default_rules.json to working directory"
cp "$SKILL_PATH/assets/report_template.html" "$DOC_AUDIT_DIR/"
echo "   ✓ Copied report_template.html to working directory"
echo

# 2. Create Python virtual environment
if [ ! -d "$VENV_DIR" ]; then
    echo "2. Creating Python virtual environment..."
    python3 -m venv "$VENV_DIR"
    echo "   ✓ Virtual environment created: $VENV_DIR/"
else
    echo "2. Python virtual environment already exists"
fi
echo

# 3. Install dependencies
echo "3. Installing Python dependencies..."
source "$VENV_DIR/bin/activate"

pip install --quiet --upgrade pip
pip install --quiet python-docx lxml defusedxml jinja2 google-genai openai

echo "   ✓ Installed packages:"

# 4. Create environment setup script
echo "4. Creating environment configuration..."
cat > "$DOC_AUDIT_DIR/env.sh" << EOF
#!/bin/bash
# Activate virtual environment and set environment variables
source "$VENV_DIR/bin/activate"
export DOC_AUDIT_SKILL_PATH="$SKILL_PATH"
export PYTHONPATH="\$DOC_AUDIT_SKILL_PATH:\$PYTHONPATH"

# Default LLM Model Configuration
# Change these to use different models across all scripts
export DOC_AUDIT_GEMINI_MODEL="\${DOC_AUDIT_GEMINI_MODEL:-gemini-3-flash-preview}"

# OpenAI Model Requirement: Must use gpt-4o-2024-08-06 or later, gpt-4o-mini, or gpt-5.x
# Older models like gpt-4-turbo, gpt-4, gpt-3.5-turbo do NOT support json_schema response format
export DOC_AUDIT_OPENAI_MODEL="\${DOC_AUDIT_OPENAI_MODEL:-gpt-5.2}"

# Audit Output Language Configuration
# Specifies the language for LLM-generated rules and audit results
# Examples: "Chinese", "English", "Japanese", "Korean", etc.
export AUDIT_LANGUAGE="\${AUDIT_LANGUAGE:-Chinese}"

# Show current environment
echo "Doc-Audit Environment Activated"
echo "  Skill Path: \$DOC_AUDIT_SKILL_PATH"
echo "  Python: \$(which python3)"
echo "  Gemini Model: \$DOC_AUDIT_GEMINI_MODEL"
echo "  OpenAI Model: \$DOC_AUDIT_OPENAI_MODEL"
echo "  API Keys: \${GOOGLE_API_KEY:+GOOGLE_API_KEY=set} \${OPENAI_API_KEY:+OPENAI_API_KEY=set}"
EOF

chmod +x "$DOC_AUDIT_DIR/env.sh"
echo "   ✓ Environment script created: $DOC_AUDIT_DIR/env.sh"
echo

# 5. Create convenience workflow script
echo "5. Creating convenience workflow script..."
cat > "$DOC_AUDIT_DIR/workflow.sh" << 'EOF'
#!/bin/bash
# Complete document audit workflow
set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

# Activate environment
source "$SCRIPT_DIR/env.sh"

# Usage check
if [ $# -lt 1 ]; then
    echo "Usage: $0 <document.docx> [rules.json]"
    echo
    echo "Examples:"
    echo "  $0 contract.docx                    # Use default rules"
    echo "  $0 contract.docx custom_rules.json  # Use custom rules"
    echo
    echo "The workflow will:"
    echo "  1. Parse the document to .claude-work/doc-audit/blocks.jsonl"
    echo "  2. Run audit to .claude-work/doc-audit/manifest.jsonl"
    echo "  3. Generate report to <document>_audit_report.html (same directory as source)"
    exit 1
fi

DOCUMENT="$1"
RULES="${2:-$SCRIPT_DIR/default_rules.json}"

if [ ! -f "$DOCUMENT" ]; then
    echo "Error: Document not found: $DOCUMENT"
    exit 1
fi

if [ ! -f "$RULES" ]; then
    echo "Error: Rules file not found: $RULES"
    exit 1
fi

# Extract document info
DOC_DIR="$(cd "$(dirname "$DOCUMENT")" && pwd)"
DOC_NAME="$(basename "$DOCUMENT" .docx)"
OUTPUT_REPORT="$DOC_DIR/${DOC_NAME}_audit_report.html"

# Define intermediate files with document name prefix
BLOCKS_FILE="$SCRIPT_DIR/${DOC_NAME}_blocks.jsonl"
MANIFEST_FILE="$SCRIPT_DIR/${DOC_NAME}_manifest.jsonl"

echo "=========================================="
echo "Document Audit Workflow"
echo "=========================================="
echo "Document: $DOCUMENT"
echo "Rules: $RULES"
echo "Report: $OUTPUT_REPORT"
echo

# Clean previous intermediate files
rm -f "$BLOCKS_FILE"
rm -f "$MANIFEST_FILE"

# Step 1: Parse document
echo "Step 1: Parsing document..."
python3 "$DOC_AUDIT_SKILL_PATH/scripts/parse_document.py" \
    "$DOCUMENT" \
    --output "$BLOCKS_FILE"
echo

# Step 2: Run audit
echo "Step 2: Running audit..."
python3 "$DOC_AUDIT_SKILL_PATH/scripts/run_audit.py" \
    --document "$BLOCKS_FILE" \
    --rules "$RULES" \
    --output "$MANIFEST_FILE"
echo

# Step 3: Generate report
echo "Step 3: Generating report..."
python3 "$DOC_AUDIT_SKILL_PATH/scripts/generate_report.py" \
    "$MANIFEST_FILE" \
    --output "$OUTPUT_REPORT" \
    --template "$SCRIPT_DIR/report_template.html" \
    --rules "$RULES"
echo

echo "=========================================="
echo "✓ Audit Complete!"
echo "Intermediate files:"
echo "  - Blocks: $BLOCKS_FILE"
echo "  - Manifest: $MANIFEST_FILE"
echo "Report: $OUTPUT_REPORT"
echo "=========================================="
EOF

chmod +x "$DOC_AUDIT_DIR/workflow.sh"
echo "   ✓ Workflow script created: $DOC_AUDIT_DIR/workflow.sh"
echo

# 6. Create README
echo "6. Creating documentation..."
cat > "$DOC_AUDIT_DIR/README.md" << 'EOF'
# Document Audit Working Directory

This directory is automatically created by Claude for document audit work.

## Directory Structure

```
.claude-work/
├── venv/                           # Python virtual environment (shared)
├── logs/                           # Operation logs (shared)
└── doc-audit/                      # Document audit working directory
    ├── env.sh                      # Environment activation script
    ├── workflow.sh                 # Convenience workflow script
    ├── README.md                   # This file
    ├── default_rules.json          # Default audit rules (copied from skill)
    ├── report_template.html        # Report template (copied from skill)
    ├── <docname>_blocks.jsonl      # Parsed document blocks (per document)
    ├── <docname>_manifest.jsonl    # Audit results (per document)
    └── <docname>_custom_rules.json # Custom rules (optional, per document)
```

**Note:** Intermediate files use the document name as a prefix (e.g., `contract_blocks.jsonl`, `contract_manifest.jsonl`) to allow processing multiple documents simultaneously without file conflicts.

## Quick Start

### One-Step Workflow (Recommended)

```bash
# Use default rules
./.claude-work/doc-audit/workflow.sh document.docx

# Use custom rules
./.claude-work/doc-audit/workflow.sh document.docx custom_rules.json
```

The audit report will be saved as `<document>_audit_report.html` in the same directory as the source document.

### Step-by-Step Workflow

```bash
# 1. Activate environment
source .claude-work/doc-audit/env.sh

# 2. Parse document (use document name prefix for intermediate files)
python skills/doc-audit/scripts/parse_document.py document.docx \
  --output .claude-work/doc-audit/document_blocks.jsonl

# 3. Run audit (with default rules from working directory)
python skills/doc-audit/scripts/run_audit.py \
  --document .claude-work/doc-audit/document_blocks.jsonl \
  --rules .claude-work/doc-audit/default_rules.json \
  --output .claude-work/doc-audit/document_manifest.jsonl

# 4. Generate report (with template and rules from working directory)
python skills/doc-audit/scripts/generate_report.py \
  .claude-work/doc-audit/document_manifest.jsonl \
  --output document_audit_report.html \
  --template .claude-work/doc-audit/report_template.html \
  --rules .claude-work/doc-audit/default_rules.json
```

## Custom Rules Workflow

If you need custom audit rules:

```bash
source .claude-work/doc-audit/env.sh

# Generate custom rules (recommended: use document name prefix)
python skills/doc-audit/scripts/parse_rules.py \
  --input "Check for vague payment terms and missing signatures" \
  --output .claude-work/doc-audit/document_custom_rules.json

# Run audit with custom rules
./.claude-work/doc-audit/workflow.sh document.docx .claude-work/doc-audit/document_custom_rules.json
```

**Tip:** Use document name prefixes for custom rules (e.g., `contract_custom_rules.json`, `report_custom_rules.json`) when auditing multiple documents to avoid confusion.

## Environment Variables

The following environment variables can be set:

```bash
# API Keys (required - choose one or both)
# For Gemini (recommended - used by default if both are set)
export GOOGLE_API_KEY=your_api_key

# For OpenAI
export OPENAI_API_KEY=your_api_key

# Custom skill path (optional)
export DOC_AUDIT_SKILL_PATH=/path/to/skills/doc-audit

# Model Configuration (optional - already set in env.sh)
# Override these to use different models across all scripts
export DOC_AUDIT_GEMINI_MODEL=gemini-3-flash    # Default Gemini model
export DOC_AUDIT_OPENAI_MODEL=gpt-5.2           # Default OpenAI model
```

## Changing Default Models

The default LLM models are configured in `.claude-work/doc-audit/env.sh`. To use different models:

1. **Edit `.claude-work/doc-audit/env.sh`** - Change the model environment variables:
   ```bash
   export DOC_AUDIT_GEMINI_MODEL="gemini-2.5-flash"
   export DOC_AUDIT_OPENAI_MODEL="gpt-4o-mini"
   ```

2. **Or set before activating** - Export variables before sourcing env.sh:
   ```bash
   export DOC_AUDIT_GEMINI_MODEL="gemini-2.0-flash-exp"
   source .claude-work/doc-audit/env.sh
   ```

All scripts (`parse_rules.py` and `run_audit.py`) will automatically use the configured models.

## Output Files

- **Intermediate files** → `.claude-work/doc-audit/` (with document name prefix)
  - `<docname>_blocks.jsonl` - Parsed document structure
  - `<docname>_manifest.jsonl` - Detailed audit results
  - `<docname>_custom_rules.json` - Custom rules (if generated)
  
- **Final report** → Same directory as source document
  - `<docname>_audit_report.html` - HTML audit report

**Example:** For `contract.docx`:
- `.claude-work/doc-audit/contract_blocks.jsonl`
- `.claude-work/doc-audit/contract_manifest.jsonl`
- `contract_audit_report.html` (in same directory as source)

## Features

- ✅ Isolated working environment (virtual environment)
- ✅ Temporary files don't pollute project directory
- ✅ Resume capability for interrupted audits
- ✅ Automatic cleanup of intermediate files
- ✅ Final report saved next to source document
- ✅ Already added to .gitignore

## API Requirements

The audit process requires an LLM API. Supported providers:

1. **Google Gemini** (recommended)
   - Install: `pip install google-genai`
   - Set: `export GOOGLE_API_KEY=...`

2. **OpenAI**
   - Install: `pip install openai`
   - Set: `export OPENAI_API_KEY=...`
   - **Model Requirement:** Must use `gpt-4o-2024-08-06` or later, `gpt-4o-mini`, `gpt-4o`, or `gpt-5.x`
   - Older models (gpt-4-turbo, gpt-4, gpt-3.5-turbo) do NOT support the required `json_schema` format

### OpenAI Model Compatibility

The scripts use OpenAI's Structured Outputs feature, which requires specific models:

✅ **Supported:**
- `gpt-4o-2024-08-06` or later
- `gpt-4o-mini`
- `gpt-4o` (latest)
- `gpt-5.x` series (e.g., `gpt-5.2`)

❌ **NOT Supported:**
- `gpt-4-turbo`
- `gpt-4`
- `gpt-3.5-turbo`

If you see an error like "json_schema is not supported", ensure you're using a compatible model.

## Troubleshooting

**Error: API key not found**
```bash
# Set your API key before running
export GOOGLE_API_KEY=your_key_here
source .claude-work/doc-audit/env.sh
```

**Error: Package not installed**
```bash
# Reinstall dependencies
source .claude-work/venv/bin/activate
pip install aspose-words jinja2 google-genai openai
```

**Resume interrupted audit**
```bash
python skills/doc-audit/scripts/run_audit.py \
  --document .claude-work/doc-audit/document_blocks.jsonl \
  --rules .claude-work/doc-audit/default_rules.json \
  --output .claude-work/doc-audit/document_manifest.jsonl \
  --resume
```
(Replace `document` with your actual document name)

## Clean Up

To remove all intermediate files and start fresh:

```bash
rm -rf .claude-work/doc-audit/*
```

To completely remove the environment:

```bash
rm -rf .claude-work/
```
EOF
echo "   ✓ README.md created"
echo

echo "=========================================="
echo "✓ Environment setup complete!"
echo "=========================================="
echo
echo "Quick start:"
echo "1. Set API key (choose one):"
echo "   export GOOGLE_API_KEY=your_key_here"
echo "   export OPENAI_API_KEY=your_key_here"
echo
echo "2. Run audit in one step:"
echo "   ./.claude-work/doc-audit/workflow.sh document.docx"
echo
echo "Or activate environment manually:"
echo "   source ./.claude-work/doc-audit/env.sh"
echo
echo "For detailed instructions, see: .claude-work/doc-audit/README.md"
echo
