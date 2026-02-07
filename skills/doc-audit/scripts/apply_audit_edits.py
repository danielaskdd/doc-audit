#!/usr/bin/env python3
"""
ABOUTME: Applies audit results to Word documents with track changes and comments
ABOUTME: Reads JSONL export from audit report and modifies the source document
"""

from docx_edit.common import (
    NS,
    DRAWING_PATTERN,
    EditItem,
    EditResult,
    format_text_preview,
    strip_auto_numbering,
    strip_table_row_numbering,
    argparse,
    sys,
)
from docx_edit.navigation_mixin import NavigationMixin
from docx_edit.table_edit_mixin import TableEditMixin
from docx_edit.revision_mixin import RevisionMixin
from docx_edit.workflow_mixin import WorkflowMixin
from docx_edit.comment_workflow_mixin import CommentWorkflowMixin


class AuditEditApplier(
    NavigationMixin,
    TableEditMixin,
    RevisionMixin,
    WorkflowMixin,
    CommentWorkflowMixin,
):
    """Composite applier assembled from focused mixins for easier maintenance."""


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
