#!/usr/bin/env python3
"""
ABOUTME: Applies audit results to Word documents with track changes and comments
ABOUTME: Reads JSONL export from audit report and modifies the source document
"""

import argparse
import json
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import List, Optional

from docx import Document
from docx_edit.common import (
    EditItem,
    EditResult,
    NS,
    extract_longest_segment,
    format_text_preview,
    strip_table_row_numbering,
)
from docx_edit.navigation_mixin import NavigationMixin
from docx_edit.table_edit_mixin import TableEditMixin
from docx_edit.revision_mixin import RevisionMixin
from docx_edit.edit_mixin import EditMixin
from docx_edit.item_search_mixin import ItemSearchMixin
from docx_edit.item_result_mixin import ItemResultMixin
from docx_edit.main_workflow_mixin import MainWorkflowMixin


class AuditEditApplier(
    NavigationMixin,
    TableEditMixin,
    RevisionMixin,
    EditMixin,
    ItemSearchMixin,
    ItemResultMixin,
    MainWorkflowMixin,
):
    """Composite applier assembled from focused mixins for easier maintenance."""

    def apply(self) -> List[EditResult]:
        """Execute all edit operations"""
        # 1. Verify hash
        if not self.skip_hash:
            if not self._verify_hash():
                raise ValueError(
                    f"Document hash mismatch\n"
                    f"Expected: {self.meta.get('source_hash', 'N/A')}\n"
                    f"Document may have been modified. Use --skip-hash to bypass."
                )

        # 2. Load document
        self.doc = Document(str(self.source_path))
        self.body_elem = self.doc._element.body

        # 2.5 Initialize paragraph order cache
        self._init_para_order()

        # 3. Initialize IDs
        self._init_comment_id()
        self._init_change_id()

        # 4. Set unified timestamp for all operations in this run
        self.operation_timestamp = datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ')

        # 5. Process each item
        for i, item in enumerate(self.edit_items):
            if self.verbose:
                print(f"[{i+1}/{len(self.edit_items)}] {item.fix_action}: "
                      f"{item.violation_text[:50]}...")

            result = self._process_item(item)
            self.results.append(result)

            if self.verbose:
                status = "✓" if result.success else "✗"
                if not result.success:
                    print(f"  [{status}]", end="")
                    print(f" {result.error_message}")

        # 5. Debug: count comment markers in document
        if self.verbose:
            try:
                rs = self._xpath(self.body_elem, './/w:commentRangeStart')
                re = self._xpath(self.body_elem, './/w:commentRangeEnd')
                rf = self._xpath(self.body_elem, './/w:commentReference')
                print(
                    f"[Comments] Markers in document.xml: "
                    f"rangeStart={len(rs)} rangeEnd={len(re)} reference={len(rf)}"
                )
            except Exception:
                pass

        # 6. Save comments
        self._save_comments()

        return self.results

    def save(self, dry_run: bool = False):
        """Save modified document"""
        if dry_run:
            print(f"[DRY RUN] Would save to: {self.output_path}")
            return

        self.doc.save(str(self.output_path))
        print(f"Saved to: {self.output_path}")

    def save_failed_items(self) -> Optional[Path]:
        """
        Save failed edit items to JSONL file for retry.

        Returns:
            Path to failed items file if any failures exist, None otherwise
        """
        failed_results = [r for r in self.results if not r.success]

        if not failed_results:
            return None  # No failures, no file created

        # Generate output path: <input>_fail.jsonl
        fail_path = self.jsonl_path.with_stem(self.jsonl_path.stem + '_fail')

        with open(fail_path, 'w', encoding='utf-8') as f:
            # Write enhanced meta line
            meta_line = {
                **self.meta,  # Include all original meta fields
                'type': 'meta',  # Explicitly ensure type field exists (required for retry)
                'original_export': self.jsonl_path.name,
                'generated_at': datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%S%z'),
                'failed_count': len(failed_results),
                'total_count': len(self.edit_items)
            }
            json.dump(meta_line, f, ensure_ascii=False)
            f.write('\n')

            # Write failed items with error information (matching HTML export field order)
            # Note: content field removed - not needed for apply_audit_edits.py processing
            for result in failed_results:
                item = result.item
                data = {
                    'category': item.category,
                    'fix_action': item.fix_action,
                    'violation_reason': item.violation_reason,
                    'violation_text': item.violation_text,
                    'revised_text': item.revised_text,
                    'rule_id': item.rule_id,
                    'uuid': item.uuid,
                    'uuid_end': item.uuid_end,  # Required for retry
                    'heading': item.heading,
                    '_error': result.error_message  # Add error info for debugging
                }
                json.dump(data, f, ensure_ascii=False)
                f.write('\n')

        return fail_path


__all__ = [
    "AuditEditApplier",
    "EditItem",
    "EditResult",
    "NS",
    "extract_longest_segment",
    "strip_table_row_numbering",
]


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Apply audit results to Word document"
    )
    parser.add_argument('jsonl_file', help='Audit export file (JSONL format)')
    parser.add_argument('-o', '--output', help='Output file path')
    parser.add_argument('--author', default='AI',
                       help='Author name for track changes/comments (default: AI)')
    parser.add_argument('--initials',
                       help='Author initials for comments (default: first 2 chars of author)')
    parser.add_argument('--skip-hash', action='store_true',
                       help='Skip hash verification')
    parser.add_argument('--dry-run', action='store_true',
                       help='Validate only, do not save')
    parser.add_argument('-v', '--verbose', action='store_true',
                       help='Verbose output')

    args = parser.parse_args()

    try:
        applier = AuditEditApplier(
            args.jsonl_file,
            output_path=args.output,
            author=args.author,
            initials=args.initials,
            skip_hash=args.skip_hash,
            verbose=args.verbose
        )

        print(f"Source file: {applier.source_path}")
        print(f"Output to: {applier.output_path}")
        print(f"Edit items: {len(applier.edit_items)}")
        if args.verbose:
            print("-" * 50)

        results = applier.apply()

        success_count = sum(1 for r in results if r.success and not r.warning)
        warning_count = sum(1 for r in results if r.success and r.warning)
        fail_count = sum(1 for r in results if not r.success)

        if warning_count > 0:
            print("\nWarning items (fallback to comment):")
            for r in results:
                if r.success and r.warning:
                    text_preview = format_text_preview(r.item.violation_text)
                    print(f"  - [{r.item.rule_id}] {r.error_message}: {text_preview}")

        if fail_count > 0:
            print("\nFailed items:")
            for r in results:
                if not r.success:
                    text_preview = format_text_preview(r.item.violation_text)
                    print(f"  - [{r.item.rule_id}] {r.error_message}: {text_preview}")

        print("-" * 50)
        print(f"Completed: {success_count} succeeded, {warning_count} warnings, {fail_count} failed")

        if not args.dry_run:
            applier.save()
        else:
            applier.save(dry_run=True)

        fail_file = applier.save_failed_items()
        if fail_file:
            print(f"\n{'=' * 50}")
            print(f"Failed items saved to: {fail_file}")
            print("  → You can modify and retry with this file")
            print(f"  → Command: python {sys.argv[0]} {fail_file} --skip-hash")
            print(f"{'=' * 50}")

        return 0

    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        return 1


if __name__ == '__main__':
    sys.exit(main())
