"""
This mixin class maps apply-status values to final EditResult output.
"""

from .common import EditItem, EditResult


class ItemResultMixin:
    def _build_result_from_status(
        self,
        item: EditItem,
        success_status: str,
        target_para,
        anchor_para,
        violation_text: str,
        revised_text: str,
        matched_runs_info,
        matched_start: int,
        item_author: str,
        is_cross_paragraph: bool,
        numbering_stripped: bool,
        last_conflict_text: str,
    ) -> EditResult:
        """Build final EditResult for a processed item."""
        if success_status == 'success':
            if numbering_stripped and self.verbose:
                print("  [Success] Matched after stripping auto-numbering")
            return EditResult(True, item)

        if success_status == 'conflict':
            reason = "Multiple changes overlap."
            fallback_para = target_para if target_para is not None else anchor_para
            conflict_item = item
            if last_conflict_text and last_conflict_text != item.violation_text:
                conflict_item = EditItem(
                    uuid=item.uuid,
                    uuid_end=item.uuid_end,
                    violation_text=last_conflict_text,
                    violation_reason=item.violation_reason,
                    fix_action=item.fix_action,
                    revised_text=item.revised_text,
                    category=item.category,
                    rule_id=item.rule_id,
                    heading=item.heading,
                )
            self._apply_fallback_comment(fallback_para, conflict_item, reason)
            if self.verbose:
                print(f"  [Conflict] {reason}")
            return EditResult(
                success=True,
                item=item,
                error_message=reason,
                warning=True,
            )

        if success_status == 'cross_paragraph_fallback':
            return self._build_manual_fallback_result(
                item=item,
                target_para=target_para,
                violation_text=violation_text,
                revised_text=revised_text,
                matched_runs_info=matched_runs_info,
                matched_start=matched_start,
                item_author=item_author,
                is_cross_paragraph=is_cross_paragraph,
                reason="Cross-paragraph delete/replace not supported",
                verbose_label="Cross-paragraph",
            )

        if success_status == 'cross_row_fallback':
            reason = "Cross-row edit not supported"
            self._apply_fallback_comment(target_para, item, reason)
            if self.verbose:
                print("  [Cross-row] Applied fallback comment")
            return EditResult(
                success=True,
                item=item,
                error_message=reason,
                warning=True,
            )

        if success_status == 'cross_cell_fallback':
            return self._build_manual_fallback_result(
                item=item,
                target_para=target_para,
                violation_text=violation_text,
                revised_text=revised_text,
                matched_runs_info=matched_runs_info,
                matched_start=matched_start,
                item_author=item_author,
                is_cross_paragraph=is_cross_paragraph,
                reason="Cross-cell edit not supported",
                verbose_label="Cross-cell",
            )

        if success_status == 'equation_fallback':
            return self._build_manual_fallback_result(
                item=item,
                target_para=target_para,
                violation_text=violation_text,
                revised_text=revised_text,
                matched_runs_info=matched_runs_info,
                matched_start=matched_start,
                item_author=item_author,
                is_cross_paragraph=is_cross_paragraph,
                reason="Equation cannot be edited",
                verbose_label="Equation-only",
            )

        if success_status == 'fallback':
            reason = "No editable runs found"
            if item.fix_action == 'manual':
                self._apply_error_comment(target_para, item)
            else:
                self._apply_fallback_comment(target_para, item, reason)
            if self.verbose:
                print(f"  [Fallback] {reason}")
            return EditResult(
                success=True,
                item=item,
                error_message=reason,
                warning=True,
            )

        self._apply_error_comment(anchor_para, item)
        return EditResult(False, item, "Operation failed")

    def _build_manual_fallback_result(
        self,
        item: EditItem,
        target_para,
        violation_text: str,
        revised_text: str,
        matched_runs_info,
        matched_start: int,
        item_author: str,
        is_cross_paragraph: bool,
        reason: str,
        verbose_label: str,
    ) -> EditResult:
        """Fallback to manual comment and produce warning result."""
        manual_status = self._apply_manual(
            target_para,
            violation_text,
            item.violation_reason,
            revised_text,
            matched_runs_info,
            matched_start,
            item_author,
            is_cross_paragraph,
            fallback_reason=reason,
        )
        if manual_status == 'success':
            if self.verbose:
                print(f"  [{verbose_label}] Applied comment instead")
            return EditResult(
                success=True,
                item=item,
                error_message=reason,
                warning=True,
            )

        self._apply_fallback_comment(target_para, item, reason)
        return EditResult(
            success=True,
            item=item,
            error_message=f"{reason} (comment also failed)",
            warning=True,
        )
