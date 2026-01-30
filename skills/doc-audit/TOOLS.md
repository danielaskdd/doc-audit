# Document Audit Tools - Detailed Reference

This document provides detailed documentation for advanced tools that are typically invoked through `workflow.sh` but can also be used independently for debugging, large document processing, or error recovery.

> **Note:** All script examples below use `$DOC_AUDIT_SKILL_PATH` environment variable, which is automatically set by `source .claude-work/doc-audit/env.sh`. Always run `source .claude-work/doc-audit/env.sh` before executing any scripts.

## Table of Contents

1. [Parse Rules](#parse-rules)
2. [Workflow Script](#workflow-script)
3. [Parse Document](#parse-document)
4. [Run Audit](#run-audit)
5. [Generate Report](#generate-report)
6. [Apply Audit Edits](#apply-audit-edits)

---

## Parse Rules

Parse natural language audit criteria into structured JSON rules using LLM. This tool is typically used to:
- **Create custom rules**: Convert user requirements into structured audit rules from scratch
- **Extend base rules**: Add new rules to existing rulesets while preserving them
- **Modify rules**: Update specific rules by explicit request

**DEFAULT BEHAVIOR**: Starts from scratch unless `--base-rules` is specified. Default rules are NOT automatically loaded.

### Usage Examples

**Common Usage Patterns:**

```bash
# ✅ Create rules from scratch (generates only user-requested rules)
# Use when: User wants custom requirements WITHOUT default rules
python $DOC_AUDIT_SKILL_PATH/scripts/parse_rules.py \
  --input "Check for ambiguous payment terms and missing signatures" \
  --output .claude-work/doc-audit/mydoc_custom_rules.json

# ✅ Create rules based on default rules (adds to or modifies defaults)
# Use when: User wants to extend or modify default rules
python $DOC_AUDIT_SKILL_PATH/scripts/parse_rules.py \
  --base-rules $DOC_AUDIT_SKILL_PATH/assets/default_rules.json \
  --input "Add rule for checking ambiguous references" \
  --output .claude-work/doc-audit/mydoc_custom_rules.json

# ✅ Iterative refinement (continues from previous output)
# Use when: User wants to modify/add/remove specific rules from existing set
python $DOC_AUDIT_SKILL_PATH/scripts/parse_rules.py \
  --base-rules .claude-work/doc-audit/mydoc_custom_rules.json \
  --input "Remove R009, make signature rule more specific" \
  --output .claude-work/doc-audit/mydoc_custom_rules.json

# Read requirements from file
python $DOC_AUDIT_SKILL_PATH/scripts/parse_rules.py \
  --file requirements.txt \
  --output custom_rules.json

# Read requirements from file and extend defaults
python $DOC_AUDIT_SKILL_PATH/scripts/parse_rules.py \
  --base-rules $DOC_AUDIT_SKILL_PATH/assets/default_rules.json \
  --file requirements.txt \
  --output custom_rules.json
```

### Decision Guide

When to use each parameter:

| User Request | Recommended Usage |
|--------------|-------------------|
| "Check for A, B, C" | ✅ Default (starts from scratch) |
| "Add these to default rules" | ✅ Use `--base-rules $DOC_AUDIT_SKILL_PATH/assets/default_rules.json` |
| "Modify existing rules" | ✅ Use `--base-rules` with the existing rules file |

### Naming Best Practice

When auditing multiple documents, use document name prefixes for custom rules to avoid confusion:
- `mydoc_custom_rules.json` for mydoc.docx
- `contract_custom_rules.json` for contract.docx

### Key Parameters

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `--input` / `-i` | text | No* | Natural language audit criteria text |
| `--file` / `-f` | path | No* | File containing audit criteria |
| `--base-rules` | path | No | Base rules file to merge with. If not specified, starts from scratch. |
| `--output` / `-o` | path | No | Output JSON file path (default: `rules.json`) |
| `--api-key` | text | No | API key for LLM service (uses env vars by default) |

\* Either `--input` or `--file` is required (LLM will generate rules based on input)

### LLM Configuration

The tool supports both Google Gemini and OpenAI for rule generation, with Gemini preferred when available.

#### Google Gemini Configuration

##### AI Studio Mode (Default)

Uses Google AI Studio API with an API key.

```bash
export GOOGLE_API_KEY="your-api-key"
```

##### Vertex AI Mode

Uses Google Cloud Vertex AI with Application Default Credentials (ADC).

```bash
# Enable Vertex AI mode
export GOOGLE_GENAI_USE_VERTEXAI=true

# Required: GCP project ID
export GOOGLE_CLOUD_PROJECT="your-project-id"

# Optional: GCP region (default: us-central1)
export GOOGLE_CLOUD_LOCATION="us-central1"

# Optional: Custom API endpoint
export GOOGLE_VERTEX_BASE_URL="https://custom-api-gateway.example.com"

# Authentication (one of the following)
export GOOGLE_APPLICATION_CREDENTIALS="/path/to/service-account.json"
# Or use: gcloud auth application-default login
```

#### OpenAI Configuration

```bash
export OPENAI_API_KEY="sk-..."

# Optional: Custom endpoint
export OPENAI_BASE_URL="https://my-proxy.example.com/v1"
```

#### Environment Variable Summary

| Variable | Mode | Required | Description |
|----------|------|----------|-------------|
| `GOOGLE_API_KEY` | AI Studio | Yes | API key from Google AI Studio |
| `GOOGLE_GENAI_USE_VERTEXAI` | Vertex AI | Yes | Set to `true` to enable |
| `GOOGLE_CLOUD_PROJECT` | Vertex AI | Yes | GCP project ID |
| `GOOGLE_CLOUD_LOCATION` | Vertex AI | No | GCP region (default: `us-central1`) |
| `GOOGLE_VERTEX_BASE_URL` | Vertex AI | No | Custom API endpoint |
| `GOOGLE_APPLICATION_CREDENTIALS` | Vertex AI | No* | Path to service account JSON |
| `OPENAI_API_KEY` | OpenAI | Yes | OpenAI API key |
| `OPENAI_BASE_URL` | OpenAI | No | Custom API endpoint |
| `DOC_AUDIT_GEMINI_MODEL` | Both | No | Gemini model name (default: `gemini-3-flash-preview`) |
| `DOC_AUDIT_OPENAI_MODEL` | Both | No | OpenAI model name (default: `gpt-5.2`) |
| `AUDIT_LANGUAGE` | Both | No | Output language for rules (default: `Chinese`) |

\* Not required if using `gcloud auth application-default login` or running on GCP

### Output Format

The generated `rules.json` file has this structure:

```json
{
  "version": "1.0",
  "rules": [
    {
      "id": "R001",
      "description": "Check for spelling and typo errors",
      "severity": "high",
      "category": "grammar",
      "examples": {
        "violation": "本周的组要工作",
        "correction": "本周的主要工作"
      }
    },
    {
      "id": "R002",
      "description": "Check for unclear or ambiguous expressions",
      "severity": "medium",
      "category": "clarity"
    }
  ]
}
```

**Rule Fields:**
- `id`: Unique identifier (R001, R002, ...)
- `description`: Clear description of what to check
- `severity`: `high`, `medium`, or `low`
- `category`: Rule category (e.g., `grammar`, `clarity`, `logic`, `compliance`, `format`, `semantic`, `other`)
- `examples`: Optional object with `violation` and `correction` examples

### Workflow

1. **Load base rules**: Auto-detect from `$DOC_AUDIT_SKILL_PATH/assets/default_rules.json`
2. **Build prompt**: Create LLM prompt with base rules and user requirements
3. **Call LLM**: Use Gemini (preferred) or OpenAI with structured JSON output
4. **Renumber rules**: Ensure sequential IDs (R001, R002, ...)
5. **Save output**: Write JSON file with version and total count

### Notes

- ⚠️ **LLM Required**: This tool always requires an LLM for rule generation
- ✅ **Rule Preservation**: When merging, existing rules are preserved unless explicitly modified
- ✅ **Auto-numbering**: Rules are automatically renumbered to avoid ID conflicts
- ✅ **Language Control**: Set `AUDIT_LANGUAGE` to control output language (e.g., `English`, `Chinese`)

---

## Workflow Script

`workflow.sh` is a convenience script that runs all three stages (parse, audit, report) automatically. This is the **recommended way** to perform a complete audit workflow.

**Location**: `$DOC_AUDIT_SKILL_PATH/scripts/workflow.sh`

Like other scripts, workflow.sh is used directly from the skill's scripts directory. It uses the working directory (`.claude-work/doc-audit/`) for intermediate files and resources.

**Auto-initialization**: If the working directory doesn't exist, workflow.sh automatically runs `setup_project_env.sh` to create it. This means you can run workflow.sh directly without manual setup.

### Usage Examples

```bash
# First time: use relative path (auto-initializes if needed)
skills/doc-audit/scripts/workflow.sh document.docx

# After setup: can use $DOC_AUDIT_SKILL_PATH
source .claude-work/doc-audit/env.sh
$DOC_AUDIT_SKILL_PATH/scripts/workflow.sh document.docx

# Use custom rules (with additional rule file)
$DOC_AUDIT_SKILL_PATH/scripts/workflow.sh document.docx -r custom_rules.json
```

### Internal Process

The script executes these steps in sequence:

1. **Parse document** → `.claude-work/doc-audit/<docname>_blocks.jsonl` (via `parse_document.py`)
2. **Run audit** → `.claude-work/doc-audit/<docname>_manifest.jsonl` (via `run_audit.py`)
3. **Generate report** → `<document_directory>/<document_name>_audit_report.html` and `<document_name>_audit_report.xlsx` (via `generate_report.py`)

### Features

- ✅ Cleans previous intermediate files (`<docname>_blocks.jsonl`, `<docname>_manifest.jsonl`) before starting
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

## Parse Document

Extract text blocks from a Word document with proper heading hierarchy and numbering. This tool is automatically invoked by `workflow.sh`, but can be used independently for:
- **Custom output paths**: Save blocks to a specific location
- **Preview mode**: Preview extracted blocks without full processing
- **Statistics**: Get document statistics (headings, characters, etc.)
- **Format selection**: Output as JSONL (streaming) or JSON (single file)

### Usage Examples

```bash
# Basic usage (outputs to <document>_blocks.jsonl)
python $DOC_AUDIT_SKILL_PATH/scripts/parse_document.py document.docx

# Custom output path
python $DOC_AUDIT_SKILL_PATH/scripts/parse_document.py document.docx \
  --output .claude-work/doc-audit/blocks.jsonl

# With preview and statistics
python $DOC_AUDIT_SKILL_PATH/scripts/parse_document.py document.docx \
  --output blocks.jsonl \
  --preview \
  --stats

# Output as regular JSON instead of JSONL
python $DOC_AUDIT_SKILL_PATH/scripts/parse_document.py document.docx \
  --output blocks.json \
  --format json
```

### Key Parameters

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `document` | path | Yes | Path to the DOCX file to parse |
| `--output` / `-o` | path | No | Output file path (default: `<document>_blocks.jsonl`) |
| `--format` | choice | No | Output format: `jsonl` (default) or `json` |
| `--preview` | flag | No | Print preview of first 5 extracted blocks |
| `--stats` | flag | No | Print document statistics (headings, characters, etc.) |

### Features

- **File Metadata**: Includes source file path, SHA256 hash, and parse timestamp
  - JSONL: First line contains metadata (type: "meta")
  - JSON: Top-level "meta" field with metadata
- **Automatic numbering capture**: Extracts list labels (e.g., "1.1", "Chapter 1") via Word's numbering XML
- **Heading-based splitting**: Each heading starts a new text block
- **Table embedding**: Tables converted to `<table>JSON</table>` format and embedded in text blocks with surrounding paragraphs
- **Heading hierarchy**: Preserves parent headings context for each block
- **Stable UUIDs**: Uses `w14:paraId` from heading paragraphs as block UUID (8-character hex ID unique within document)
- **paraId validation**: Requires Word 2013+ documents with `w14:paraId` attributes (terminates with error if missing)

### Workflow

1. Load document with python-docx library
2. Parse styles.xml to extract outline levels for headings
3. Iterate through body nodes (paragraphs and tables)
4. For each paragraph:
   - Extract `w14:paraId` attribute (validates presence, errors if missing)
   - Check if it's a heading via outline level
   - If heading: save previous block with heading's paraId as UUID
   - If content: append to current block, track first paraId for Preface blocks
5. For each table: convert to 2D array and embed in content
6. Use heading's `w14:paraId` as block UUID (or first content paraId for Preface blocks)
7. Clean up old `manifest.jsonl` to prevent UUID mismatch in resume mode

### Output Format (JSONL)

Each line is a JSON object. Tables are embedded as `<table>JSON</table>` within text content:

```json
{"uuid": "12AB34CD", "uuid_end": "56EF78AB", "heading": "2.1 Penalty Clause", "content": "If Party B delays...\n<table>[[\"Header 1\",\"Header 2\"],[\"Cell 1\",\"Cell 2\"]]</table>\nSubsequent paragraph...", "type": "text", "parent_headings": ["Chapter 2 Contract Terms"]}
```

### Output Format (JSON)

```json
{
  "total_blocks": 42,
  "blocks": [
    {
      "uuid": "12AB34CD",
      "uuid_end": "56EF78AB",
      "heading": "2.1 Penalty Clause",
      "content": "If Party B delays payment...\n<table>[[\"Penalty Type\",\"Amount\"],[\"Late Payment\",\"1% per day\"]]</table>\nThe above table shows penalty structure.",
      "type": "text",
      "parent_headings": ["Chapter 2 Contract Terms"]
    }
  ]
}
```

### Error Handling

**Missing paraId Error:**

If the document is missing `w14:paraId` attributes on paragraphs, the script will display a user-friendly error message and exit with code 1. This typically occurs with:
- Documents created by older versions of Microsoft Word (before Office 2013)
- Documents generated programmatically without paraId attributes

When this error occurs, the workflow must stop immediately.

---

## Run Audit

Execute LLM-based audit on each text block against audit rules. This tool is automatically invoked by `workflow.sh`, but can be used independently for:
- **Debugging**: Test audit with specific blocks using `--dry-run`
- **Large documents**: Process in chunks using `--start-block` and `--end-block`
- **Resume interrupted runs**: Continue from where it stopped using `--resume`
- **Custom model selection**: Override default model with `--model`

### Usage Examples

```bash
# Basic usage with single rule file
python $DOC_AUDIT_SKILL_PATH/scripts/run_audit.py \
  --document .claude-work/doc-audit/blocks.jsonl \
  --rules $DOC_AUDIT_SKILL_PATH/assets/default_rules.json

# Use multiple rule files (auto-merge, checks for duplicate IDs)
python $DOC_AUDIT_SKILL_PATH/scripts/run_audit.py \
  --document .claude-work/doc-audit/blocks.jsonl \
  -r $DOC_AUDIT_SKILL_PATH/assets/default_rules.json \
  -r $DOC_AUDIT_SKILL_PATH/assets/bidding_rules.json

# Another way to specify multiple files
python $DOC_AUDIT_SKILL_PATH/scripts/run_audit.py \
  --document .claude-work/doc-audit/blocks.jsonl \
  --rules rules1.json rules2.json rules3.json

# Specify model explicitly
python $DOC_AUDIT_SKILL_PATH/scripts/run_audit.py \
  --document blocks.jsonl \
  --rules custom_rules.json \
  --model gemini-2.5-flash

# Force Gemini provider (even if OPENAI_API_KEY is also set)
python $DOC_AUDIT_SKILL_PATH/scripts/run_audit.py \
  --document blocks.jsonl \
  --rules rules.json \
  --provider gemini

# Force OpenAI provider (even if GOOGLE_API_KEY is also set)
python $DOC_AUDIT_SKILL_PATH/scripts/run_audit.py \
  --document blocks.jsonl \
  --rules rules.json \
  --provider openai

# Combine provider and model selection
python $DOC_AUDIT_SKILL_PATH/scripts/run_audit.py \
  --document blocks.jsonl \
  --rules rules.json \
  --provider openai \
  --model gpt-5.2

# Process specific block range
python $DOC_AUDIT_SKILL_PATH/scripts/run_audit.py \
  --document blocks.jsonl \
  --rules rules.json \
  --start-block 10 \
  --end-block 50

# Resume from previous interrupted run
python $DOC_AUDIT_SKILL_PATH/scripts/run_audit.py \
  --document blocks.jsonl \
  --rules rules.json \
  --resume

# Dry run to preview prompts without calling LLM
python $DOC_AUDIT_SKILL_PATH/scripts/run_audit.py \
  --document blocks.jsonl \
  --rules rules.json \
  --dry-run
```

### Key Parameters

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `--document` / `-d` | path | Yes | Path to document blocks file (JSONL or JSON from `parse_document.py`) |
| `--rules` / `-r` | path | Yes | Path to audit rules JSON file(s). Can be specified multiple times to merge rules. |
| `--output` / `-o` | path | No | Output manifest file path (default: `manifest.jsonl`) |
| `--provider` | choice | No | Force LLM provider: `auto` (default), `gemini`, `openai` |
| `--model` | text | No | LLM model: `auto` (default), `gemini-2.5-flash`, `gpt-5.2`, etc. |
| `--workers` | int | No | Number of parallel workers for concurrent API calls (default: 8) |
| `--rate-limit` | float | No | Seconds to wait between API calls per worker (default: 0.05) |
| `--start-block` | int | No | Start from this block index (0-based, default: 0) |
| `--end-block` | int | No | End at this block index (inclusive, default: last block) |
| `--resume` | flag | No | Resume from previous run (skip already-processed blocks) |
| `--dry-run` | flag | No | Print prompts without calling LLM (for debugging) |

### Provider Selection (`--provider`)

The `--provider` parameter allows you to explicitly specify which LLM provider to use, which is useful when you have both Gemini and OpenAI credentials configured.

**Behavior:**

| `--provider` | `--model` | Result |
|--------------|-----------|--------|
| `auto` (default) | `auto` | Auto-detect: Gemini (if configured) > OpenAI |
| `auto` | `gemini-2.5-flash` | Use Gemini with specified model |
| `auto` | `gpt-5.2` | Use OpenAI with specified model |
| `gemini` | any value | **Force Gemini** (validates credentials) |
| `openai` | any value | **Force OpenAI** (validates credentials) |

**Credential Validation:**

When you explicitly specify a provider, the script validates that the required credentials are present:

- `--provider gemini` checks for:
  - AI Studio mode: `GOOGLE_API_KEY`
  - Vertex AI mode: `GOOGLE_CLOUD_PROJECT`
  
- `--provider openai` checks for:
  - `OPENAI_API_KEY`

If credentials are missing, you'll get a clear error message with hints on how to fix it.

**Use Cases:**

```bash
# Scenario: Both GOOGLE_API_KEY and OPENAI_API_KEY are set
# Want to use OpenAI instead of default Gemini
python $DOC_AUDIT_SKILL_PATH/scripts/run_audit.py \
  --document blocks.jsonl \
  --rules rules.json \
  --provider openai

# Scenario: Testing different providers for comparison
python $DOC_AUDIT_SKILL_PATH/scripts/run_audit.py \
  --document blocks.jsonl \
  --rules rules.json \
  --provider gemini \
  --output manifest_gemini.jsonl

python $DOC_AUDIT_SKILL_PATH/scripts/run_audit.py \
  --document blocks.jsonl \
  --rules rules.json \
  --provider openai \
  --output manifest_openai.jsonl
```

### Model Selection (`--model`)

- `auto` (default): Auto-select based on available credentials (Gemini preferred if configured)
- `gemini-2.5-flash`, `gemini-3-flash`: Use Google Gemini (via AI Studio or Vertex AI)
- `gpt-5.2`, `gpt-4o`, `gpt-4o-mini`: Use OpenAI (requires `OPENAI_API_KEY`)
- Model defaults are configured in `.claude-work/doc-audit/env.sh`

**Note:** When `--provider` is specified, it overrides the provider inference from `--model`. For example, `--provider gemini --model gpt-5.2` will use Gemini (the model name is treated as a custom model identifier).

### Google Gemini Configuration

The audit tool supports two modes for accessing Google Gemini:

#### AI Studio Mode (Default)

Uses Google AI Studio API with an API key. This is the simplest setup for development and testing.

```bash
# Required environment variable
export GOOGLE_API_KEY="your-api-key"
```

#### Vertex AI Mode

Uses Google Cloud Vertex AI with Application Default Credentials (ADC). This is recommended for production deployments and enterprise use.

```bash
# Enable Vertex AI mode
export GOOGLE_GENAI_USE_VERTEXAI=true

# Required: GCP project ID
export GOOGLE_CLOUD_PROJECT="your-project-id"

# Optional: GCP region (default: us-central1)
export GOOGLE_CLOUD_LOCATION="us-central1"

# Authentication: One of the following
export GOOGLE_APPLICATION_CREDENTIALS="/path/to/service-account.json"

# Base URL (Optional: for proxy or public models available general)
export GOOGLE_VERTEX_BASE_URL='https://aiplatform.googleapis.com'
```

**Note:** When `GOOGLE_GENAI_USE_VERTEXAI=true` is set, the `GOOGLE_API_KEY` environment variable is ignored. The tool will use ADC for authentication instead.

#### Environment Variable Summary

| Variable | Required | Description |
|----------|----------|-------------|
| **AI Studio Mode** | | |
| `GOOGLE_API_KEY` | Yes | API key from Google AI Studio |
| **Vertex AI Mode** | | |
| `GOOGLE_GENAI_USE_VERTEXAI` | Yes | Set to `true` to enable |
| `GOOGLE_CLOUD_PROJECT` | Yes | GCP project ID |
| `GOOGLE_CLOUD_LOCATION` | No | GCP region (default: `us-central1`) |
| `GOOGLE_VERTEX_BASE_URL` | No | Custom API endpoint (for API gateway proxies) |
| `GOOGLE_APPLICATION_CREDENTIALS` | No* | Path to service account JSON |

\* Not required if using `gcloud auth application-default login` or running on GCP (GCE, GKE, Cloud Run)

#### Custom API Endpoint

For scenarios requiring a custom API gateway proxy (e.g., corporate network policies, custom load balancing), you can specify a custom base URL:

```bash
export GOOGLE_GENAI_USE_VERTEXAI=true
export GOOGLE_CLOUD_PROJECT="your-project-id"
export GOOGLE_VERTEX_BASE_URL="https://custom-api-gateway.example.com"
```

When `GOOGLE_VERTEX_BASE_URL` is set, the SDK will route all requests through the specified endpoint instead of the default Vertex AI endpoint. If not set, the SDK automatically determines the appropriate endpoint based on the project and location.

### OpenAI Configuration

#### Default Mode

Uses the official OpenAI API with an API key.

```bash
export OPENAI_API_KEY="sk-..."
```

#### Custom Endpoint

For scenarios requiring a custom API endpoint (e.g., corporate proxy, Azure OpenAI, local LLM server with OpenAI-compatible API), you can specify a custom base URL:

```bash
export OPENAI_API_KEY="sk-..."
export OPENAI_BASE_URL="https://my-proxy.example.com/v1"
```

**Note:** `OPENAI_BASE_URL` is natively supported by the OpenAI Python SDK.

#### Environment Variable Summary

| Variable | Required | Description |
|----------|----------|-------------|
| `OPENAI_API_KEY` | Yes | OpenAI API key |
| `OPENAI_BASE_URL` | No | Custom API endpoint (for proxies, Azure, etc.) |

### Parallel Processing (`--workers`)

The audit script processes multiple text blocks concurrently using asyncio for improved performance:

- **Default**: 8 parallel workers
- **Implementation**: Uses `asyncio.Semaphore` to limit concurrent API calls
- **Rate limiting**: Applied per worker (default 0.05s between calls per worker)

**Usage Examples:**

```bash
# Use default 8 workers
python $DOC_AUDIT_SKILL_PATH/scripts/run_audit.py -d blocks.jsonl -r rules.json

# Increase parallelism for faster processing
python $DOC_AUDIT_SKILL_PATH/scripts/run_audit.py -d blocks.jsonl -r rules.json --workers 8

# Reduce parallelism to avoid rate limits
python $DOC_AUDIT_SKILL_PATH/scripts/run_audit.py -d blocks.jsonl -r rules.json --workers 2 --rate-limit 0.5
```

**Performance Impact:**
- With 8 workers: ~8x faster than sequential processing
- Adjust `--workers` based on your API rate limits and document size
- Progress output shows block completion in real-time (may appear out of order due to parallel execution)

**SDK Support:**
- Both Google Gemini (`client.aio`) and OpenAI (`AsyncOpenAI`) use native async APIs
- No thread overhead - uses Python's asyncio event loop

### Resume Functionality (Advanced)

The `--resume` flag enables recovery from interrupted audit runs by:

1. **Loading completed UUIDs**: Reads `manifest.jsonl` to get UUIDs of already-processed blocks
2. **Skipping processed blocks**: During iteration, skips blocks whose UUIDs are in the completed set
3. **Appending new results**: New audit results are appended to existing `manifest.jsonl`

#### Resume Use Cases

**Case 1: Simple Resume After Interruption**
```bash
# Initial run (interrupted at block 45/100)
python $DOC_AUDIT_SKILL_PATH/scripts/run_audit.py \
  --document blocks.jsonl \
  --rules rules.json \
  --output manifest.jsonl
# ... interrupted (Ctrl+C, network error, etc.)

# Resume from where it left off
python $DOC_AUDIT_SKILL_PATH/scripts/run_audit.py \
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
python $DOC_AUDIT_SKILL_PATH/scripts/run_audit.py \
  --document blocks.jsonl \
  --rules rules.json \
  --start-block 0 \
  --end-block 99

# Process second chunk (blocks 100-199) - interrupted at block 150
python $DOC_AUDIT_SKILL_PATH/scripts/run_audit.py \
  --document blocks.jsonl \
  --rules rules.json \
  --start-block 100 \
  --end-block 199
# ... interrupted

# Resume second chunk (will skip 100-149, continue from 150)
python $DOC_AUDIT_SKILL_PATH/scripts/run_audit.py \
  --document blocks.jsonl \
  --rules rules.json \
  --start-block 100 \
  --end-block 199 \
  --resume

# Process third chunk (blocks 200-299)
python $DOC_AUDIT_SKILL_PATH/scripts/run_audit.py \
  --document blocks.jsonl \
  --rules rules.json \
  --start-block 200 \
  --end-block 299
```

**Case 3: Re-audit Specific Blocks (Without Resume)**

To re-audit specific blocks (e.g., after changing rules), **do NOT use `--resume`**:

```bash
# Re-audit blocks 10-20 (will overwrite those results in manifest)
python $DOC_AUDIT_SKILL_PATH/scripts/run_audit.py \
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
  "uuid": "12AB34CD",
  "uuid_end": "56EF78AB",
  "p_heading": "2.1 Penalty Clause",
  "p_content": "If Party B delays payment, they shall pay approximately 1%...",
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

**UUID Range Fields:**
- `uuid`: 8-character hex paraId of first paragraph in source block
- `uuid_end`: 8-character hex paraId of last paragraph in source block
- Each violation also includes `uuid` and `uuid_end` (injected by script) for range-restricted text search in `apply_audit_edits.py`

---

## Generate Report

Create HTML audit report from audit manifest with statistics and traceability. This tool is automatically invoked by `workflow.sh`, but can be used independently for:
- **Re-generating reports**: After modifying the template
- **Custom output paths**: Save report to specific location
- **JSON export**: Generate both HTML and JSON output
- **Testing templates**: Preview template changes

### Usage Examples

```bash
# Basic usage with single rule file
python $DOC_AUDIT_SKILL_PATH/scripts/generate_report.py \
  --manifest manifest.jsonl \
  --template $DOC_AUDIT_SKILL_PATH/assets/report_template.html \
  --rules $DOC_AUDIT_SKILL_PATH/assets/default_rules.json \
  --output audit_report.html

# Use multiple rule files (auto-merge, later files override earlier ones)
python $DOC_AUDIT_SKILL_PATH/scripts/generate_report.py \
  -m manifest.jsonl \
  -t $DOC_AUDIT_SKILL_PATH/assets/report_template.html \
  -r $DOC_AUDIT_SKILL_PATH/assets/default_rules.json \
  -r $DOC_AUDIT_SKILL_PATH/assets/bidding_rules.json \
  -o audit_report.html

# Another way to specify multiple files
python $DOC_AUDIT_SKILL_PATH/scripts/generate_report.py \
  -m manifest.jsonl \
  -t $DOC_AUDIT_SKILL_PATH/assets/report_template.html \
  --rules rules1.json rules2.json rules3.json \
  -o audit_report.html

# No rule descriptions in report (not recommended)
python $DOC_AUDIT_SKILL_PATH/scripts/generate_report.py \
  -m manifest.jsonl \
  -t $DOC_AUDIT_SKILL_PATH/assets/report_template.html \
  -o audit_report.html

# Also output JSON data
python $DOC_AUDIT_SKILL_PATH/scripts/generate_report.py \
  -m manifest.jsonl \
  -t $DOC_AUDIT_SKILL_PATH/assets/report_template.html \
  -r rules.json \
  -o report.html \
  --json

# Also output Excel report
python $DOC_AUDIT_SKILL_PATH/scripts/generate_report.py \
  -m manifest.jsonl \
  -t $DOC_AUDIT_SKILL_PATH/assets/report_template.html \
  -r rules.json \
  -o report.html \
  --excel

# For trusted HTML content (disables escaping, not recommended)
python $DOC_AUDIT_SKILL_PATH/scripts/generate_report.py \
  -m manifest.jsonl \
  -t $DOC_AUDIT_SKILL_PATH/assets/report_template.html \
  -o report.html \
  --trusted-html
```

### Key Parameters

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `--manifest` / `-m` | path | Yes | Path to audit manifest JSONL file (from `run_audit.py`) |
| `--output` / `-o` | path | No | Output HTML file path (default: `audit_report.html`) |
| `--template` / `-t` | path | Yes | Path to Jinja2 HTML template |
| `--rules` / `-r` | path | No | Path to audit rules JSON file(s). Can be specified multiple times to merge rules. Later files override earlier ones for duplicate rule IDs. Recommended for displaying full rule details in modal popups. |
| `--trusted-html` | flag | No | Disable HTML escaping (only for trusted inputs) |
| `--json` | flag | No | Also output report data as JSON (same name with `.json` extension) |
| `--excel` | flag | No | Also output report as Excel file (same name with `.xlsx` extension). Requires `openpyxl` package. |

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
- `<output>.xlsx` - Excel report (if `--excel` flag is used, requires `openpyxl`)

---

## Apply Audit Edits

Apply audit results exported from HTML report to Word document with track changes and comments. This is a **post-processing tool** used after manual review of audit results.

**Typical scenario**:
1. Generate audit report using workflow
2. Review issues in HTML report
3. Mark false positives as "blocked"
4. Export non-blocked issues to JSONL
5. Apply edits to Word document using this tool

### Source Document Requirements

⚠️ **No Pre-existing Revisions**: The source Word document should **NOT** contain existing track changes (revisions).

**Why this matters:**
- The script was recently updated to search for `violation_text` in the document's *original text* view (before revisions)
- While the script can now handle documents with existing revisions better than before, **the best practice is still to use clean documents without any track changes**
- Pre-existing `<w:del>` and `<w:ins>` elements can still cause edge cases and increase complexity

**If your document has track changes:**
1. Open the document in Microsoft Word
2. Go to **Review** → **Accept** → **Accept All Changes in Document** (or **Reject All Changes**)
3. Save the document
4. Re-run the audit workflow from the beginning (`parse_document.py` → `run_audit.py` → `generate_report.py`)

**Alternative:** If you must apply edits to a document with existing revisions, use the `--skip-hash` flag (since accepting changes will modify the document hash).

### Usage Examples

```bash
# Basic usage (outputs to <source>_edited.docx)
python $DOC_AUDIT_SKILL_PATH/scripts/apply_audit_edits.py weekly-report_audit_export.jsonl

# Specify output path
python $DOC_AUDIT_SKILL_PATH/scripts/apply_audit_edits.py export.jsonl -o output.docx

# Skip hash verification (if document was modified after export)
python $DOC_AUDIT_SKILL_PATH/scripts/apply_audit_edits.py export.jsonl --skip-hash

# Verbose output (show each edit item processing)
python $DOC_AUDIT_SKILL_PATH/scripts/apply_audit_edits.py export.jsonl -v

# Dry run (validate without saving)
python $DOC_AUDIT_SKILL_PATH/scripts/apply_audit_edits.py export.jsonl --dry-run
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
{"category":"grammar","uuid":"682A7C9F","uuid_end":"682A7CA3","violation_text":"本周的组要工作","violation_reason":"\"组要\"是错别字","fix_action":"replace","revised_text":"本周的主要工作","rule_id":"R001"}
{"category":"logic","uuid":"682A7C9F","uuid_end":"682A7CA3","violation_text":"文件列表如下：","violation_reason":"缺少列表内容","fix_action":"manual","revised_text":"请补充文件列表","rule_id":"R008"}
```

**Field Descriptions:**
- `uuid`: 8-character hex paraId - anchor paragraph where search starts
- `uuid_end`: 8-character hex paraId - end of range for restricted search (ensures text matching stays within the source block)
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
