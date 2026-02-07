# docx_edit Module Structure

`apply_audit_edits.py` has been split into maintainable modules organized by responsibility and editing scenario:

1. `common.py` — Shared constants, data classes, and text normalization utilities.
2. `navigation_mixin.py` — Word DOM traversal, paragraph/table range lookup, and run collection & matching.
3. `table_edit_mixin.py` — Table cell editing: extraction and application for single-cell and multi-cell scenarios.
4. `revision_mixin.py` — Track-changes low-level algorithms: diff computation, `w:del`/`w:ins` XML construction.
5. `workflow_mixin.py` — Delete/replace main workflows (cross-paragraph and single-paragraph) and revision ID management.
6. `comment_workflow_mixin.py` — Comment insertion, manual-action handling, per-item processing, and the overall apply/save flow.

Quick lookup guide:
- **How content is located in Word XML** — `navigation_mixin.py`
- **How table cells are edited** — `table_edit_mixin.py`
- **How revision XML is generated** — `revision_mixin.py`
- **How a single edit item is processed end-to-end** — `comment_workflow_mixin.py`
