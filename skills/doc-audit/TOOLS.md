# Document Audit Tools - Detailed Reference

This document provides detailed documentation for advanced tools that are typically invoked through `workflow.sh` but can also be used independently for debugging, large document processing, or error recovery.

**Quick Navigation:**
- [4. Run Audit](#4-run-audit) - Execute LLM-based audit
- [5. Generate Report](#5-generate-report) - Create HTML report from audit results
- [6. Workflow Script](#6-workflow-script) - Complete automated workflow
- [7. Apply Audit Edits](#7-apply-audit-edits) - Apply exported edits to Word document

---

## 4. Run Audit

Execute LLM-based audit on each text block against audit rules. This tool is automatically invoked by `workflow.sh`, but can be used independently for:
- **Debugging**: Test audit with specific blocks using `--dry-run`
- **Large documents**: Process in chunks using `--start-block` and `--end-block`
- **Resume interrupted runs**: Continue from where it stopped using `--resume`
- **Custom model selection**: Override default model with `--model`

### Usage Examples

```bash
# Basic usage with auto model selection
python scripts/run_audit.py \
  --document .claude-work/doc-audit/blocks.jsonl \
  --rules .claude-work/doc-audit/default_rules.json

# Specify model explicitly
python scripts/run_audit.py \
  --document blocks.jsonl \
  --rules custom_rules.json \
  --model gemini-2.5-flash

# Process specific block range
python skills/doc-audit/scripts/run_audit.py \
  --document blocks.jsonl \
  --rules rules.json \
  --start-block 10 \
  --end-block 50

# Resume from previous interrupted run
python skills/doc-audit/scripts/run_audit.py \
  --document blocks.jsonl \
  --rules rules.json \
  --resume

# Dry run to preview prompts without calling LLM
python skills/doc-audit/scripts/run_audit.py \
  --document blocks.jsonl \
  --rules rules.json \
  --dry-run
```

### Key Parameters

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `--document` / `-d` | path | Yes | Path to document blocks file (JSONL or JSON from `parse_document.py`) |
| `--rules` / `-r` | path | Yes | Path to audit rules JSON file |
| `--output` / `-o` | path | No | Output manifest file path (default: `manifest.jsonl`) |
| `--model` | text | No | LLM model: `auto` (default), `gemini-2.5-flash`, `gpt-5.2`, etc. |
| `--rate-limit` | float | No | Seconds to wait between API calls (default: 0.5) |
| `--start-block` | int | No | Start from this block index (0-based, default: 0) |
| `--end-block` | int | No | End at this block index (inclusive, default: last block) |
| `--resume` | flag | No | Resume from previous run (skip already-processed blocks) |
| `--dry-run` | flag | No | Print prompts without calling LLM (for debugging) |

### Model Selection (`--model`)

- `auto` (default): Auto-select based on available API keys (Gemini preferred if both are set)
- `gemini-2.5-flash`, `gemini-3-flash`: Use Google Gemini (requires `GOOGLE_API_KEY`)
- `gpt-5.2`, `gpt-4o`, `gpt-4o-mini`: Use OpenAI (requires `OPENAI_API_KEY`)
- Model defaults are configured in `.claude-work/doc-audit/env.sh`

### Resume Functionality (Advanced)

The `--resume` flag enables recovery from interrupted audit runs by:

1. **Loading completed UUIDs**: Reads `manifest.jsonl` to get UUIDs of already-processed blocks
2. **Skipping processed blocks**: During iteration, skips blocks whose UUIDs are in the completed set
3. **Appending new results**: New audit results are appended to existing `manifest.jsonl`

#### Resume Use Cases

**Case 1: Simple Resume After Interruption**
```bash
# Initial run (interrupted at block 45/100)
python skills/doc-audit/scripts/run_audit.py \
  --document blocks.jsonl \
  --rules rules.json \
  --output manifest.jsonl
# ... interrupted (Ctrl+C, network error, etc.)

# Resume from where it left off
python skills/doc-audit/scripts/run_audit.py \
  --document blocks.jsonl \
  --rules rules.json \
  --output manifest.jsonl \
  --resume
# Automatically skips blocks 0-44, continues from block 45
```

**Case 2: Chunked Processing with Resume**

For large documents, process in chunks to avoid API rate limits or long-running sessions:

```bash
# Process first chunk (blocks 0-99)
python skills/doc-audit/scripts/run_audit.py \
  --document blocks.jsonl \
  --rules rules.json \
  --start-block 0 \
  --end-block 99

# Process second chunk (blocks 100-199) - interrupted at block 150
python skills/doc-audit/scripts/run_audit.py \
  --document blocks.jsonl \
  --rules rules.json \
  --start-block 100 \
  --end-block 199
# ... interrupted

# Resume second chunk (will skip 100-149, continue from 150)
python skills/doc-audit/scripts/run_audit.py \
  --document blocks.jsonl \
  --rules rules.json \
  --start-block 100 \
  --end-block 199 \
  --resume

# Process third chunk (blocks 200-299)
python skills/doc-audit/scripts/run_audit.py \
  --document blocks.jsonl \
  --rules rules.json \
  --start-block 200 \
  --end-block 299
```

**Case 3: Re-audit Specific Blocks (Without Resume)**

To re-audit specific blocks (e.g., after changing rules), **do NOT use `--resume`**:

```bash
# Re-audit blocks 10-20 (will overwrite those results in manifest)
python skills/doc-audit/scripts/run_audit.py \
  --document blocks.jsonl \
  --rules updated_rules.json \
  --start-block 10 \
  --end-block 20
# Without --resume, it processes all blocks 10-20 regardless of manifest
```

### Important Notes

- ⚠️ **UUID Consistency**: Resume relies on UUIDs. If you re-run `parse_document.py`, `manifest.jsonl` is automatically deleted by the script, a fresh audit is required.
- ✅ **Append-Only**: Resume appends to `manifest.jsonl`. If you want to start completely fresh, delete the manifest file first.
- ✅ **Block Range + Resume**: Combining `--start-block`/`--end-block` with `--resume` is valid - it will skip already-processed blocks within the specified range.

### Workflow

1. **Build system prompt**: Formats rules as structured instructions (cached by LLM across all blocks)
2. **Load completed UUIDs**: If `--resume` is set, loads already-processed block UUIDs from manifest
3. **Iterate blocks**: For each block in range:
   - Skip if UUID already processed (resume mode)
   - Build user prompt with heading context + content
   - Call LLM with structured output schema (Gemini or OpenAI)
   - Parse violations from LLM response
   - Add category to each violation (lookup from rule ID)
   - Save entry to manifest.jsonl (append mode)
   - Rate limit between requests
4. **Error handling**: Catches JSON parsing errors and API errors, continues to next block

### Output Format (manifest.jsonl)

Each line is an audit result:
```json
{
  "uuid": "550e8400-e29b-41d4-a716-446655440000",
  "p_heading": "2.1 Penalty Clause",
  "p_content": "If Party B delays payment, they shall pay approximately 1%...",
  "is_violation": true,
  "violations": [
    {
      "rule_id": "R002",
      "category": "semantic",
      "violation_text": "approximately 1% of the total amount",
      "violation_reason": "Contains vague term 'approximately' and does not specify currency",
      "fix_action": "replace",
      "revised_text": "1% of the contract total amount as penalty (settled in CNY)"
    }
  ]
}
```

---

## 5. Generate Report

Create HTML audit report from audit manifest with statistics and traceability. This tool is automatically invoked by `workflow.sh`, but can be used independently for:
- **Re-generating reports**: After modifying the template
- **Custom output paths**: Save report to specific location
- **JSON export**: Generate both HTML and JSON output
- **Testing templates**: Preview template changes

### Usage Examples

```bash
# Basic usage
python scripts/generate_report.py manifest.jsonl \
  --template .claude-work/doc-audit/report_template.html \
  --rules .claude-work/doc-audit/default_rules.json \
  --output audit_report.html


# No rule descriptions in report (not recommended)
python scripts/generate_report.py manifest.jsonl \
  --template .claude-work/doc-audit/report_template.html \
  --output audit_report.html

# Also output JSON data
python skills/doc-audit/scripts/generate_report.py manifest.jsonl \
  --template .claude-work/doc-audit/report_template.html \
  --rules rules.json \
  --output report.html \
  --json

# For trusted HTML content (disables escaping, not recommended)
python skills/doc-audit/scripts/generate_report.py manifest.jsonl \
  --template .claude-work/doc-audit/report_template.html \
  --output report.html \
  --trusted-html
```

### Key Parameters

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `manifest` | path | Yes | Path to audit manifest JSONL file (from `run_audit.py`) |
| `--output` / `-o` | path | No | Output HTML file path (default: `audit_report.html`) |
| `--template` / `-t` | path | Yes | Path to Jinja2 HTML template |
| `--rules` / `-r` | path | No | Path to rules JSON file (optional, recommended for displaying full rule details in modal popups) |
| `--trusted-html` | flag | No | Disable HTML escaping (only for trusted inputs) |
| `--json` | flag | No | Also output report data as JSON (same name with `.json` extension) |

### Features

- **File Information Header**: Displays source document filename and hash (from metadata)
- **Interactive Fixed Header**:
  - Problem count displayed in title (Valid: N | Blocked: M)
  - Category filter dropdown
  - Status filter buttons (All / Valid / Blocked)
  - Export audit results button
- **Dynamic Statistics**: Real-time updates of valid/blocked counts as users interact
- **Issue Management**:
  - Each issue can be marked as "blocked" (false positive)
  - Blocked issues shown with gray styling and strikethrough
  - Filter issues by category and blocked status
- **Export Functionality**:
  - Export non-blocked violations to JSONL format
  - Uses File System Access API for native save dialog
  - Includes metadata (source file, hash, export timestamp)
  - Falls back to traditional download for unsupported browsers
- **Issue Details**: Each violation with heading, content, reason, and suggestion
- **Source Tracing**: Expandable source text with click-to-expand/collapse
- **Rule Information**: Clickable rule badges (e.g., `[R001]`) that display full rule details in modal popups (when `--rules` is provided)
- **HTML Safety**: Escapes HTML by default; use `--trusted-html` only if all inputs are trusted

### Workflow

1. **Load manifest**: Parse `manifest.jsonl` to get all audit results
2. **Load rules** (optional): If `--rules` provided, load rule descriptions for modal popups
3. **Generate report data**:
   - Count total blocks and violations
   - Group violations by category
   - Collect unique rules used
4. **Render HTML**: Use Jinja2 template with report data
5. **Save output**: Write HTML file (and optionally JSON)

### Output Files

- `<output>.html` - HTML report (always generated)
- `<output>.json` - JSON report data (if `--json` flag is used)

---

## 6. Workflow Script

`workflow.sh` is a convenience script that runs all three stages (parse, audit, report) automatically. This is the **recommended way** to perform a complete audit workflow.

### Usage Examples

```bash
# Use default rules
./.claude-work/doc-audit/workflow.sh document.docx

# Use custom rules
./.claude-work/doc-audit/workflow.sh document.docx .claude-work/doc-audit/custom_rules.json
```

### Internal Process

The script executes these steps in sequence:

1. **Parse document** → `.claude-work/doc-audit/blocks.jsonl` (via `parse_document.py`)
2. **Run audit** → `.claude-work/doc-audit/manifest.jsonl` (via `run_audit.py`)
3. **Generate report** → `<document_directory>/<document_name>_audit_report.html` (via `generate_report.py`)

### Features

- ✅ Cleans previous intermediate files (`blocks.jsonl`, `manifest.jsonl`) before starting
- ✅ Final report saved in same directory as source document
- ✅ Uses working directory's default rules if no custom rules specified
- ✅ Automatically passes rules to report generation for full rule details

### When to Use Individual Tools Instead

If the workflow fails at any stage, you can run individual tools to debug or continue manually:

- **Parse failed**: Check document format, paraId presence
- **Audit interrupted**: Use `run_audit.py --resume` to continue
- **Report customization**: Use `generate_report.py` with custom templates
- **Large documents**: Use `run_audit.py` with `--start-block`/`--end-block` for chunked processing

---

## 7. Apply Audit Edits

Apply audit results exported from HTML report to Word document with track changes and comments. This is a **post-processing tool** used after manual review of audit results.

**Typical scenario**:
1. Generate audit report using workflow
2. Review issues in HTML report
3. Mark false positives as "blocked"
4. Export non-blocked issues to JSONL
5. Apply edits to Word document using this tool

### Usage Examples

```bash
# Basic usage (outputs to <source>_edited.docx)
python scripts/apply_audit_edits.py weekly-report_audit_export.jsonl

# Specify output path
python scripts/apply_audit_edits.py export.jsonl -o output.docx

# Skip hash verification (if document was modified after export)
python scripts/apply_audit_edits.py export.jsonl --skip-hash

# Verbose output (show each edit item processing)
python scripts/apply_audit_edits.py export.jsonl -v

# Dry run (validate without saving)
python scripts/apply_audit_edits.py export.jsonl --dry-run
```

### Key Parameters

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `jsonl_file` | path | Yes | Path to audit export JSONL file (from HTML report) |
| `--output` / `-o` | path | No | Output file path (default: `<source>_edited.docx`) |
| `--skip-hash` | flag | No | Skip document hash verification |
| `--dry-run` | flag | No | Validate without saving (preview mode) |
| `--verbose` / `-v` | flag | No | Show detailed processing output |

### Features

- **Three Edit Modes**:
  - `delete`: Remove text with Word track changes (deletion markup)
  - `replace`: Replace text with Word track changes (minimal diff-based editing)
  - `manual`: Add Word comment on text (for human review)
- **Hash Verification**: Ensures document hasn't been modified since audit export
- **Precise Location**: Uses paragraph ID (`w14:paraId`) as anchor point for text search
- **Cross-Run Text Handling**: Handles Word's text fragmentation across multiple runs
- **Format Preservation**: Maintains original text formatting (font, size, color)
- **ID Conflict Prevention**: Scans existing comments/revisions to assign unique IDs
- **Error Reporting**: Detailed success/failure statistics with error messages

### Export JSONL Format

The JSONL file exported from HTML report has the following structure:

**First line (metadata):**
```json
{"type":"meta","source_file":"/path/to/source.docx","source_hash":"sha256:abc123...","exported_at":"2026-01-06T18:46:42.625Z"}
```

**Subsequent lines (edit actions):**
```json
{"category":"grammar","uuid":"682A7C9F","violation_text":"本周的组要工作","violation_reason":"\"组要\"是错别字","fix_action":"replace","revised_text":"本周的主要工作","rule_id":"R001"}
{"category":"logic","uuid":"682A7C9F","violation_text":"文件列表如下：","violation_reason":"缺少列表内容","fix_action":"manual","revised_text":"请补充文件列表","rule_id":"R008"}
```

**Field Descriptions:**
- `uuid`: Anchor paragraph's `w14:paraId` (search starts from this paragraph)
- `violation_text`: Text to find and modify (must match exactly)
- `fix_action`: `delete` | `replace` | `manual`
- `revised_text`: Replacement text (for replace) or suggestion (for manual)
- `violation_reason`: Explanation (used in Word comments for manual action)

### Workflow

1. **Load JSONL**: Parse metadata and edit items
2. **Verify Hash**: Check document hash matches export (prevents applying to wrong version)
3. **Load Document**: Open Word document with python-docx
4. **Initialize IDs**: Scan existing comments/track changes to avoid ID conflicts
5. **Process Each Item**:
   - Find anchor paragraph by `w14:paraId` using XPath
   - Search for `violation_text` from anchor paragraph in document order
   - Apply edit based on `fix_action`:
     - `delete`: Wrap text in `<w:del>` element
     - `replace`: Calculate diff, wrap deletions in `<w:del>`, insertions in `<w:ins>`
     - `manual`: Add `<w:commentRangeStart/End>` and create comment
6. **Save Comments**: Create/update `comments.xml` via OPC API
7. **Save Document**: Write modified document

### Output Example

```
Source file: /path/to/source.docx
Output to: /path/to/source_edited.docx
Edit items: 14
--------------------------------------------------
--------------------------------------------------
Completed: 14 succeeded, 0 failed

Saved to: /path/to/source_edited.docx
```

### Word Document Effects

1. **Delete**: Text shows strikethrough, marked as "AI deleted" in track changes
2. **Replace**:
   - Deleted portions show strikethrough
   - Inserted portions show underline (default revision style)
   - Original formatting preserved
3. **Manual**:
   - Text highlighted (comment range)
   - Comment balloon shows reason and suggestion

### Error Handling

Common failures and solutions:

| Error | Cause | Solution |
|-------|-------|----------|
| Hash verification failed | Document modified after export | Use `--skip-hash` |
| Paragraph ID not found | Document regenerated paraId | Re-run audit workflow |
| Text not found | Text already modified or formatting issue | Manual edit required |

### Technology Stack

- `python-docx`: Word document manipulation
- `lxml`: XML DOM operations
- `docx.opc`: OOXML package structure (auto-manages Content Types and Relationships)
- `difflib`: Minimal diff calculation for replace operations

### Performance

- Processing speed: ~50-100 items/second
- Memory usage: < 100MB for typical documents
- File size increase: +5-15% (due to revision metadata)

### See Also

- `scripts/APPLY_EDITS_README.md` - Detailed implementation documentation
- Phase 2 workflow in SKILL.md for generating audit reports
- HTML report's export functionality for creating JSONL files
