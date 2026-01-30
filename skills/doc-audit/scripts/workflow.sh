#!/bin/bash
# Complete document audit workflow
set -e

# Get the directory where this script is located
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

# Working directory for intermediate files and resources
WORK_DIR=".claude-work/doc-audit"

# Check if environment exists, if not run setup
if [ ! -f "$WORK_DIR/env.sh" ]; then
    echo "Working directory not found. Running setup..."
    bash "$SCRIPT_DIR/setup_project_env.sh"
    echo
fi

# Activate environment (sets DOC_AUDIT_SKILL_PATH)
source "$WORK_DIR/env.sh"

# Usage check
if [ $# -lt 1 ]; then
    echo "Usage: $0 <document.docx> [options]"
    echo
    echo "Options:"
    echo "  -r, --rules FILE       Audit rules file (can be specified multiple times)"
    echo "  --rules-only           Use only specified rules (exclude default rules)"
    echo
    echo "Examples:"
    echo "  $0 contract.docx                           # Use default rules"
    echo "  $0 contract.docx -r custom_rules.json     # Default + custom rules"
    echo "  $0 contract.docx -r r1.json -r r2.json     # Default + r1 + r2"
    echo "  $0 contract.docx --rules-only -r custom.json  # Only custom rules"
    echo
    echo "The workflow will:"
    echo "  1. Parse the document to .claude-work/doc-audit/<docname>_blocks.jsonl"
    echo "  2. Run audit to .claude-work/doc-audit/<docname>_manifest.jsonl"
    echo "  3. Generate report to <document>_audit_report.html (same directory as source)"
    exit 1
fi

DOCUMENT="$1"
shift

# Parse options
RULE_FILES=()
RULES_ONLY=false

while [[ $# -gt 0 ]]; do
    case $1 in
        -r|--rules)
            if [ -z "$2" ]; then
                echo "Error: --rules requires an argument"
                exit 1
            fi
            RULE_FILES+=("$2")
            shift 2
            ;;
        --rules-only)
            RULES_ONLY=true
            shift
            ;;
        *)
            echo "Error: Unknown option: $1"
            exit 1
            ;;
    esac
done

# Check document exists
if [ ! -f "$DOCUMENT" ]; then
    echo "Error: Document not found: $DOCUMENT"
    exit 1
fi

# Helper function to find rule file
find_rule_file() {
    local filename="$1"

    # If absolute path or relative path with /, use as-is
    if [[ "$filename" == /* ]] || [[ "$filename" == */* ]]; then
        if [ -f "$filename" ]; then
            echo "$filename"
            return 0
        fi
        return 1
    fi

    # Search in standard locations
    for dir in "$WORK_DIR" "$DOC_AUDIT_SKILL_PATH/assets" "."; do
        if [ -f "$dir/$filename" ]; then
            echo "$dir/$filename"
            return 0
        fi
    done

    return 1
}

# Build rule file list
RESOLVED_RULES=()

# Add default rules if not --rules-only
if [ "$RULES_ONLY" = false ]; then
    RESOLVED_RULES+=("$DOC_AUDIT_SKILL_PATH/assets/default_rules.json")
fi

# Add user-specified rules
for rule_file in "${RULE_FILES[@]}"; do
    resolved=$(find_rule_file "$rule_file")
    if [ $? -ne 0 ]; then
        echo "Error: Rules file not found: $rule_file"
        echo "Searched in:"
        echo "  - $WORK_DIR/"
        echo "  - $DOC_AUDIT_SKILL_PATH/assets/"
        echo "  - Current directory"
        exit 1
    fi
    RESOLVED_RULES+=("$resolved")
done

# If no rules specified at all
if [ ${#RESOLVED_RULES[@]} -eq 0 ]; then
    if [ "$RULES_ONLY" = true ]; then
        echo "Error: --rules-only specified but no rules files provided"
        echo "Use -r/--rules to specify at least one rules file"
        exit 1
    fi
    # Fallback to default rules only when not in rules-only mode
    RESOLVED_RULES+=("$DOC_AUDIT_SKILL_PATH/assets/default_rules.json")
fi

# Verify all rule files exist
for rule_file in "${RESOLVED_RULES[@]}"; do
    if [ ! -f "$rule_file" ]; then
        echo "Error: Rules file not found: $rule_file"
        exit 1
    fi
done

# Extract document info
DOC_DIR="$(cd "$(dirname "$DOCUMENT")" && pwd)"
DOC_NAME="$(basename "$DOCUMENT" .docx)"
OUTPUT_REPORT="$DOC_DIR/${DOC_NAME}_audit_report.html"
OUTPUT_EXCEL="$DOC_DIR/${DOC_NAME}_audit_report.xlsx"

# Define intermediate files with document name prefix
BLOCKS_FILE="$WORK_DIR/${DOC_NAME}_blocks.jsonl"
MANIFEST_FILE="$WORK_DIR/${DOC_NAME}_manifest.jsonl"

echo "=========================================="
echo "Document Audit Workflow"
echo "=========================================="
echo "Document: $DOCUMENT"
echo "Rules (${#RESOLVED_RULES[@]} file(s)):"
for rule_file in "${RESOLVED_RULES[@]}"; do
    echo "  - $rule_file"
done
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

# Step 2: Run audit with multiple rule files
echo "Step 2: Running audit..."
AUDIT_CMD=(python3 "$DOC_AUDIT_SKILL_PATH/scripts/run_audit.py" --document "$BLOCKS_FILE" --output "$MANIFEST_FILE")
for rule_file in "${RESOLVED_RULES[@]}"; do
    AUDIT_CMD+=(--rules "$rule_file")
done
"${AUDIT_CMD[@]}"
echo

# Step 3: Generate report (HTML + Excel)
# Use the first rule file for report generation (contains all merged rules)
echo "Step 3: Generating report..."
REPORT_CMD=(python3 "$DOC_AUDIT_SKILL_PATH/scripts/generate_report.py" --manifest "$MANIFEST_FILE" --output "$OUTPUT_REPORT" --template "$DOC_AUDIT_SKILL_PATH/assets/report_template.html" --excel)
for rule_file in "${RESOLVED_RULES[@]}"; do
    REPORT_CMD+=(--rules "$rule_file")
done
"${REPORT_CMD[@]}"
echo

echo "=========================================="
echo "✓ Audit Complete!"
echo "Intermediate files:"
echo "  - Blocks: $BLOCKS_FILE"
echo "  - Manifest: $MANIFEST_FILE"
echo "Reports:"
echo "  - HTML:  $OUTPUT_REPORT"
echo "  - Excel: $OUTPUT_EXCEL"
echo "=========================================="
