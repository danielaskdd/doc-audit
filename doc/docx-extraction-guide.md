# Word Document Extraction and Smart Chunking Guide

This guide is for readers who need to submit Word documents to an LLM auditing workflow. It explains the **core extraction challenges**, the **smart chunking optimization strategy**, and the **field definitions with representative examples** in `_blocks.jsonl`.

> This document is based on the current implementation in `skills/doc-audit/scripts/parse_document.py`.

---

## 1. Overview: Why “Extraction + Smart Chunking” Is Required

In document auditing scenarios, Word files introduce three inherent challenges:

- **Complex structure**: heading levels, automatic numbering, merged cells, and repeated table headers.
- **Heterogeneous content**: body text, tables, images, formulas, superscripts, and subscripts coexist.
- **Context limits**: LLMs have input-length limits, so oversized content must be split.

Therefore, the conversion from `.docx` to `_blocks.jsonl` must solve all of the following:

1. **Structure restoration**: recover true heading hierarchy and numbering semantics.
2. **Content fidelity**: preserve tables, images, formulas, superscripts, and subscripts.
3. **Auditable chunking**: retain enough context while staying within model limits.

The resulting `_blocks.jsonl` is the foundational input for the auditing system.

---

## 2. What Extraction Must Solve

### 2.1 Heading Level Detection

- Uses `outlineLvl` to determine heading level (not style-name strings).
- Supports style inheritance (`basedOn`) to complete level metadata.
- Treats `outlineLvl >= 9` as body text.

**Result**: headings at different levels become semantic chunk boundaries.

### 2.2 Automatic Numbering Restoration

- Uses `NumberingResolver` to restore Word auto-numbering (chapter/list/figure indices).
- Makes extracted content align better with what users visually read in Word.

**Result**: strings such as `2.1 Penalty Clause` keep the `2.1` prefix.

### 2.3 Full Table Preservation

- Supports vertical/horizontal merged cells to preserve table semantics for LLM understanding.
- Supports repeated table headers (`w:tblHeader`) so header rows can be retained after table chunking.
- Outputs table content as a JSON 2D array wrapped in `<table>...</table>` for structured LLM input.

**Result**: table semantics are preserved with high fidelity, suitable for rule-based auditing.

### 2.4 Fidelity for Special Content

- **Images**: `<drawing id="..." name="..." path="..." format="..." />`
  - Embedded image (`r:embed`): exported to `<blocks_stem>.image/`, `path` stores relative file path
  - Linked image (`r:link`): not downloaded, `path` stores original external URL
  - `format`: normalized image format derived from content-type/file extension, common values include
    `png`, `jpeg`, `emf`, `wmf`, `gif`, `bmp`, `tiff`, `webp`, `svg`
  - `format` is omitted when format cannot be determined or when drawing has no image blip (e.g., chart shape)
- **Formulas**: OMML → LaTeX, wrapped in `<equation>...</equation>`
- **Super/Subscripts**: `<sup>...</sup>` / `<sub>...</sub>`

**Result**: reduces false judgments caused by missing formulas or inline markers.

### 2.5 Paragraph Anchors for Positioning

Each paragraph depends on Word `w14:paraId`:

- Ensures block boundaries are locatable (`uuid` / `uuid_end`).
- Greatly narrows the search range for downstream annotation and reduces annotation failures.
- Missing `paraId` raises an immediate error (typically from non-standard Word 2013+ files; open and re-save in Word).
- Normal paragraphs: use the first and last paragraph `paraId`.
- Tables: use the first paragraph in the first cell and the last paragraph in the last cell; for vertically merged cells, `paraId` is traced back to the first cell in the merge group to ensure correct edit positioning later.

---

## 3. Smart Chunking Strategy (Core Performance Optimization)

**Why not split only by headings?**

- Body text between headings can still exceed model input limits.
- Tables are often very large and cannot be split arbitrarily.
- If sections are split too aggressively, the audit step may miss cross-paragraph logical consistency issues.

**Goals of smart chunking**

- **Use the LLM context window efficiently**: keep chunk size as close as possible to `IDEAL_BLOCK_CONTENT_TOKENS` without exceeding `MAX_BLOCK_CONTENT_TOKENS`.
- **Preserve hierarchical semantics**: let the model understand each paragraph’s heading and full ancestry; ensure blocks start with a heading and do not include body text from higher-level headings.
- **Avoid table/body context breaks**: keep tables tied to surrounding narrative whenever possible.

The pipeline has four stages:

### Stage A: Heading-Driven Initial Splitting

- Heading paragraphs trigger chunk boundaries.
- Heading text itself is included in block content (as the first paragraph).
- No empty block is emitted when there is no body text between headings.

**Purpose**: create semantically clear block boundaries.

**Heading resolution logic**:

- Primary source: heading paragraphs (`outlineLvl` 0-8); auto-numbering is prefixed via `NumberingResolver`.
- Non-heading start defaults to `Preface/Uncategorized`.
- During long-block splitting, anchor paragraphs become new `heading` values for child blocks.
- Non-first chunks of split tables append: `[表格片段N]`.

> For reliability, the current implementation exits with an error if a heading exceeds 200 characters (it does not continue with truncation flow).

### Stage B: Large Table Row-Based Splitting

Trigger: estimated table token size `> 5000`.

Strategy:

- Split only at **row boundaries**.
- First chunk merges with preceding text, middle chunks stand alone, last chunk merges with following text.
- Middle/last chunks may carry `table_header` (repeated table header rows).

**Purpose**: prevent oversized tables from exhausting context.

> Non-first chunks append ` [表格片段<index>] ` to the original heading. `table_header` comes from Word repeated-header metadata; if unavailable, the field is omitted.

**After table splitting, whether body text and table stay in the same block**:

- Unsplit table: same block as surrounding body text (if present).
- First split chunk: same block as preceding body text (if present).
- Middle split chunks: standalone blocks.
- Last split chunk: same block as following body text (if present).

### Stage C: Anchor-Based Long Block Splitting

Trigger: estimated block token size `> 8000`.

Strategy:

- Select short paragraphs (`<= 100` chars; treated as small-heading candidates) as anchors.
- Anchors become new block `heading` values and remain in block content.
- Tables are not split again at this stage.

**Purpose**: reduce size while preserving semantic continuity. The split granularity does not need to be larger than table-splitting granularity, so tables can remain semantically close to surrounding text.

### Stage D: Small-Block Merging (Hierarchy-Preserving)

Strategy:

- Bottom-up: merge same-level small blocks first, then allow cross-level absorption.
- Never exceed 8000 tokens after merge.
- Skip merging when `table_chunk_role="middle"`.
- Tail Absorption: if a same-level short tail block (`< 1000` tokens) remains, absorb it to avoid unmergeable fragments.

**Purpose**: reduce over-fragmentation and improve audit accuracy.

**Overall outcome**:

- Usually fewer blocks with better per-block context.
- Most blocks end up near 6000 tokens and below 8000 tokens.
- Middle table chunks are not affected by small-block merging.

### Fixed-Level Mode (`--fixlevel=N`)

When `--fixlevel=N` is used:

- Only heading levels `<= N` trigger splitting.
- Table splitting and small-block merging are disabled.
- The final block may still trigger long-block splitting.

This mode is for special text chunking cases and is not used in the default document-audit workflow.

---

## 4. Output Format: `_blocks.jsonl`

The file uses **JSON Lines**:

1. **Line 1**: metadata (`type: "meta"`)
2. **Line 2+**: text blocks (`type: "text"`)

### 4.1 Metadata Line

```json
{
  "type": "meta",
  "source_file": "/absolute/path/to/document.docx",
  "source_hash": "sha256:f59469543f31d54f...",
  "parsed_at": "2026-01-28T18:17:32.175446"
}
```

| Field | Type | Description |
|------|------|------|
| `type` | string | Always `"meta"` |
| `source_file` | string | Absolute path of source document |
| `source_hash` | string | SHA256 of source `.docx`: `sha256:<hex>` |
| `parsed_at` | string | Local timestamp (no timezone) |

### 4.2 Text Block Line

```json
{
  "uuid": "682A7C9F",
  "uuid_end": "5A77B586",
  "heading": "2.1 Penalty Clause",
  "content": "If Party B delays payment...",
  "type": "text",
  "parent_headings": ["Chapter 2 Contract Terms"],
  "level": 2,
  "table_chunk_role": "none",
  "table_header": [["col-1", "col-2", "col-3"]]
}
```

| Field | Type | Description |
|------|------|------|
| `uuid` | string | `paraId` of the first paragraph in the block |
| `uuid_end` | string | `paraId` of the last paragraph in the block |
| `heading` | string | Block heading (including auto-numbering) |
| `content` | string | Concatenated multi-paragraph text |
| `type` | string | Always `"text"` |
| `parent_headings` | list | Ancestor heading chain |
| `level` | int | Heading level (1-9) |
| `table_chunk_role` | string | Table split role: `none/first/middle/last` |
| `table_header` | list | Optional header rows (repeated table header) |

---

## 5. Typical Chunk Examples

### 5.1 Normal Text Block

```json
{
  "uuid": "AA12B3C4",
  "uuid_end": "DD56E7F8",
  "heading": "1.1 Project Scope",
  "content": "This project includes...\n\nThis section describes...",
  "type": "text",
  "parent_headings": ["Chapter 1 General Provisions"],
  "level": 2,
  "table_chunk_role": "none"
}
```

### 5.2 Text Block with Table (Unsplit)

```json
{
  "uuid": "1A2B3C4D",
  "uuid_end": "5E6F7A8B",
  "heading": "2.2 Technical Indicators",
  "content": "Indicators are listed below:\n<table>[['...'], ['...']]</table>",
  "type": "text",
  "parent_headings": ["Chapter 2 Technical Requirements"],
  "level": 2,
  "table_chunk_role": "none"
}
```

### 5.3 Large Table Split (Middle Chunk Standalone)

```json
{
  "uuid": "1234ABCD",
  "uuid_end": "5678EF90",
  "heading": "3.2 Data Scope [表格片段1]",
  "content": "<table>[['...'], ['...']]</table>",
  "type": "text",
  "parent_headings": ["Chapter 3 Data Requirements"],
  "level": 2,
  "table_chunk_role": "middle",
  "table_header": [["Column 1", "Column 2", "Column 3"]]
}
```

### 5.4 Child Block After Anchor-Based Split

```json
{
  "uuid": "A1B2C3D4",
  "uuid_end": "E5F6A7B8",
  "heading": "Background",
  "content": "Background\n...",
  "type": "text",
  "parent_headings": ["3.1 Project Description"],
  "level": 3,
  "table_chunk_role": "none"
}
```

### 5.5 Result After Small-Block Merging

```json
{
  "uuid": "AAAA1111",
  "uuid_end": "BBBB2222",
  "heading": "4.1 Risk Analysis",
  "content": "4.1 Risk Analysis...\n\n4.2 Risk Control Measures...\n\n...",
  "type": "text",
  "parent_headings": ["Chapter 4 Risks"],
  "level": 2,
  "table_chunk_role": "none"
}
```

---

## 6. Threshold Quick Reference

| Constant | Value | Purpose |
|------|----|------|
| `MAX_HEADING_LENGTH` | 200 chars | Max heading length |
| `MAX_ANCHOR_CANDIDATE_LENGTH` | 100 chars | Max anchor length for long-block split |
| `IDEAL_BLOCK_CONTENT_TOKENS` | 6000 | Ideal block size |
| `MAX_BLOCK_CONTENT_TOKENS` | 8000 | Hard block limit |
| `SMALL_TAIL_THRESHOLD` | 1000 | Tail absorption threshold |
| `TABLE_IDEAL_TOKENS` | 3000 | Target table chunk size |
| `TABLE_MAX_TOKENS` | 5000 | Table split trigger |
| `TABLE_MIN_LAST_CHUNK_TOKENS` | 1600 | Min final chunk size |

---

## Appendix: Token Estimation Method

Block size is based on a **heuristic** `estimate_tokens()` method, not exact model tokenization:

- Chinese characters: ~0.75 tokens/char
- JSON structure characters: ~1 token/char
- Other characters: ~0.4 tokens/char
- Plus ~5% buffer and a fixed offset

This keeps chunking behavior relatively stable and controllable across models.
