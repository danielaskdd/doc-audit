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
echo

# 2. Create Python virtual environment
if [ ! -d "$VENV_DIR" ]; then
    echo "2. Creating Python virtual environment with uv (Python 3.12)..."
    uv venv "$VENV_DIR" --python 3.12 --seed
    echo "   ✓ Virtual environment created: $VENV_DIR/"
else
    echo "2. Python virtual environment already exists"
fi
echo

# 3. Install dependencies
echo "3. Installing Python dependencies with uv..."
source "$VENV_DIR/bin/activate"

uv pip install python-docx lxml defusedxml jinja2 google-genai openai openpyxl

echo "   ✓ Installed packages:"

# 4. Create environment setup script
echo "4. Creating environment configuration..."
cat > "$DOC_AUDIT_DIR/env.sh" << EOF
#!/bin/bash
# Activate virtual environment and set environment variables
source "$VENV_DIR/bin/activate"
export DOC_AUDIT_SKILL_PATH="$SKILL_PATH"
export PYTHONPATH="\$DOC_AUDIT_SKILL_PATH:\$PYTHONPATH"

# =============================================================================
# API Keys (Required - choose one)
# =============================================================================
# export GOOGLE_API_KEY="your_key"     # Google AI Studio (recommended)
# export OPENAI_API_KEY="your_key"     # OpenAI API
# Note: If both are set, Gemini is used by default

# =============================================================================
# Google Vertex AI Mode (Alternative to AI Studio)
# =============================================================================
# export GOOGLE_GENAI_USE_VERTEXAI=true           # Enable Vertex AI mode
# export GOOGLE_CLOUD_PROJECT="your-project-id"   # GCP project ID (required)
# export GOOGLE_CLOUD_LOCATION="us-central1"      # GCP region (optional)
# export GOOGLE_VERTEX_BASE_URL=""                # Custom API endpoint for proxy
# export GOOGLE_APPLICATION_CREDENTIALS=""        # Path to service account JSON
#                                                 # (not required if using gcloud auth)

# =============================================================================
# Model Configuration (Optional)
# =============================================================================
# Change these to use different models across all scripts
export DOC_AUDIT_GEMINI_MODEL="\${DOC_AUDIT_GEMINI_MODEL:-gemini-3-flash-preview}"

# OpenAI Model Requirement: Must use gpt-4o-2024-08-06 or later, gpt-4o-mini, or gpt-5.x
# Older models like gpt-4-turbo, gpt-4, gpt-3.5-turbo do NOT support json_schema response format
export DOC_AUDIT_OPENAI_MODEL="\${DOC_AUDIT_OPENAI_MODEL:-gpt-5.2}"

# Audit Output Language Configuration
# Specifies the language for LLM-generated rules and audit results
# Examples: "Chinese", "English", "Japanese", "Korean", etc.
export AUDIT_LANGUAGE="\${AUDIT_LANGUAGE:-Chinese}"

# =============================================================================
# Custom Endpoints (Optional)
# =============================================================================
# export OPENAI_BASE_URL=""            # Custom OpenAI API endpoint (for proxies, Azure, etc.)

# =============================================================================
# Thinking/Reasoning Configuration (Optional)
# =============================================================================
# For models that support extended reasoning capabilities:
# export GEMINI_THINKING_LEVEL=""      # Gemini 3: minimal, low, medium, high
# export GEMINI_THINKING_BUDGET=""     # Gemini 2.5: Token count (0 to disable)
# export OPENAI_REASONING_EFFORT=""    # OpenAI o-series: low, medium, high

# Show current environment
echo "Doc-Audit Environment Activated"
echo "  Skill Path: \$DOC_AUDIT_SKILL_PATH"
echo "  Python: \$(which python3)"
echo
EOF

chmod +x "$DOC_AUDIT_DIR/env.sh"
echo "   ✓ Environment script created: $DOC_AUDIT_DIR/env.sh"
echo

# 5. Create README
echo "5. Creating documentation..."
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
    ├── README.md                   # This file
    ├── <docname>_blocks.jsonl      # Parsed document blocks (per document)
    ├── <docname>_manifest.jsonl    # Audit results (per document)
    └── <docname>_custom_rules.json # Custom rules (optional, per document)

# Read-only assets (from skill directory)
$DOC_AUDIT_SKILL_PATH/assets/
├── default_rules.json              # Default audit rules
├── bidding_rules.json              # Additional audit rules for bidding document
├── global_rules.json               # Cross-reference verification rules
└── report_template.html            # Report template
```

**Note:** Intermediate files use the document name as a prefix (e.g., `contract_blocks.jsonl`, `contract_manifest.jsonl`) to allow processing multiple documents simultaneously without file conflicts.

## Quick Start

### One-Step Workflow (Recommended)

```bash
# First time: use relative path (auto-initializes if needed)
skills/doc-audit/scripts/workflow.sh document.docx

# After setup: can use $DOC_AUDIT_SKILL_PATH
source .claude-work/doc-audit/env.sh
$DOC_AUDIT_SKILL_PATH/scripts/workflow.sh document.docx

# Use custom rules (with additional rule file)
$DOC_AUDIT_SKILL_PATH/scripts/workflow.sh document.docx -r custom_rules.json
```

The audit reports will be saved as `<document>_audit_report.html` and `<document>_audit_report.xlsx` in the same directory as the source document.

### Step-by-Step Workflow

```bash
# 1. Activate environment
source .claude-work/doc-audit/env.sh

# 2. Parse document (use document name prefix for intermediate files)
python $DOC_AUDIT_SKILL_PATH/scripts/parse_document.py document.docx \
  --output .claude-work/doc-audit/document_blocks.jsonl

# 3. Run audit (with default rules from assets directory)
python $DOC_AUDIT_SKILL_PATH/scripts/run_audit.py \
  --document .claude-work/doc-audit/document_blocks.jsonl \
  --rules $DOC_AUDIT_SKILL_PATH/assets/default_rules.json \
  --output .claude-work/doc-audit/document_manifest.jsonl

# 4. Generate report (with template and rules from assets directory)
python $DOC_AUDIT_SKILL_PATH/scripts/generate_report.py \
  .claude-work/doc-audit/document_manifest.jsonl \
  --output document_audit_report.html \
  --template $DOC_AUDIT_SKILL_PATH/assets/report_template.html \
  --rules $DOC_AUDIT_SKILL_PATH/assets/default_rules.json \
  --excel
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
$DOC_AUDIT_SKILL_PATH/scripts/workflow.sh document.docx -r .claude-work/doc-audit/document_custom_rules.json
```

**Tip:** Use document name prefixes for custom rules (e.g., `contract_custom_rules.json`, `report_custom_rules.json`) when auditing multiple documents to avoid confusion.

## Environment Variables

The following environment variables can be configured. Edit `.claude-work/doc-audit/env.sh` or export before sourcing.

### API Keys (Required - choose one)

| Variable | Description |
|----------|-------------|
| `GOOGLE_API_KEY` | Google AI Studio API key (recommended) |
| `OPENAI_API_KEY` | OpenAI API key |

> If both are set, Gemini is used by default.

### Google Vertex AI Mode (Alternative to AI Studio)

| Variable | Required | Description |
|----------|----------|-------------|
| `GOOGLE_GENAI_USE_VERTEXAI` | Yes | Set to \`true\` to enable Vertex AI |
| `GOOGLE_CLOUD_PROJECT` | Yes | GCP project ID |
| `GOOGLE_CLOUD_LOCATION` | No | GCP region (default: \`us-central1\`) |
| `GOOGLE_VERTEX_BASE_URL` | No | Custom API endpoint for proxy |
| `GOOGLE_APPLICATION_CREDENTIALS` | No | Path to service account JSON |

### Model Configuration (Optional)

| Variable | Default | Description |
|----------|---------|-------------|
| `DOC_AUDIT_GEMINI_MODEL` | \`gemini-3-flash-preview\` | Gemini model name |
| `DOC_AUDIT_OPENAI_MODEL` | \`gpt-5.2\` | OpenAI model name |
| `AUDIT_LANGUAGE` | \`Chinese\` | Output language for audit results |

### Custom Endpoints (Optional)

| Variable | Description |
|----------|-------------|
| `OPENAI_BASE_URL` | Custom OpenAI API endpoint (for proxies, Azure, etc.) |

### Thinking/Reasoning Configuration (Optional)

| Variable | Model | Values |
|----------|-------|--------|
| `GEMINI_THINKING_LEVEL` | Gemini 3 | minimal, low, medium, high |
| `GEMINI_THINKING_BUDGET` | Gemini 2.5 | Token count (0 to disable) |
| `OPENAI_REASONING_EFFORT` | OpenAI o-series | low, medium, high |

## Changing Default Models

The default LLM models are configured in `.claude-work/doc-audit/env.sh`. To use different models:

1. **Edit `.claude-work/doc-audit/env.sh`** - Change the model environment variables:
   ```bash
   export DOC_AUDIT_GEMINI_MODEL="gemini-3-flash"
   export DOC_AUDIT_OPENAI_MODEL="gpt-4o-mini"
   ```

2. **Or set before activating** - Export variables before sourcing env.sh:
   ```bash
   export DOC_AUDIT_GEMINI_MODEL="gemini-3-flash"
   source .claude-work/doc-audit/env.sh
   ```

All scripts (`parse_rules.py` and `run_audit.py`) will automatically use the configured models.

## Output Files

- **Intermediate files** → `.claude-work/doc-audit/` (with document name prefix)
  - `<docname>_blocks.jsonl` - Parsed document structure
  - `<docname>_manifest.jsonl` - Detailed audit results
  - `<docname>_custom_rules.json` - Custom rules (if generated)
  
- **Final reports** → Same directory as source document
  - `<docname>_audit_report.html` - HTML audit report (interactive, with charts)
  - `<docname>_audit_report.xlsx` - Excel audit report (tabular, for further analysis)

**Example:** For `contract.docx`:
- `.claude-work/doc-audit/contract_blocks.jsonl`
- `.claude-work/doc-audit/contract_manifest.jsonl`
- `contract_audit_report.html` (in same directory as source)
- `contract_audit_report.xlsx` (in same directory as source)

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
   - Install: `uv pip install google-genai`
   - Set: `export GOOGLE_API_KEY=...`

2. **OpenAI**
   - Install: `uv pip install openai`
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
# Reinstall dependencies with uv
source .claude-work/venv/bin/activate
uv pip install python-docx lxml defusedxml jinja2 google-genai openai
```

**Resume interrupted audit**
```bash
python $DOC_AUDIT_SKILL_PATH/scripts/run_audit.py \
  --document .claude-work/doc-audit/document_blocks.jsonl \
  --rules $DOC_AUDIT_SKILL_PATH/assets/default_rules.json \
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
echo "2. Activate environment:"
echo "   source ./.claude-work/doc-audit/env.sh"
echo
echo "3. Run audit:"
echo "   \$DOC_AUDIT_SKILL_PATH/scripts/workflow.sh document.docx"
echo
echo "For detailed instructions, see: .claude-work/doc-audit/README.md"
echo
