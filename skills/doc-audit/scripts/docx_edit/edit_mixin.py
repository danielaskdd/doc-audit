"""
This mixin class implements delete/replace main workflows (cross-paragraph and single-paragraph) and revision ID management.
"""

import copy
import re
from typing import List, Dict, Tuple, Optional

from lxml import etree

from .common import NS, DRAWING_PATTERN, EQUATION_PATTERN
from utils import sanitize_xml_string


class EditMixin:
    def _init_change_id(self):
        """Initialize next track change ID by scanning existing changes"""
        max_id = -1
        
        for tag in ('ins', 'del'):
            for elem in self.body_elem.iter(f'{{{NS["w"]}}}{tag}'):
                cid = elem.get(f'{{{NS["w"]}}}id')
                if cid:
                    try:
                        max_id = max(max_id, int(cid))
                    except ValueError:
                        pass
        
        self.next_change_id = max_id + 1

    def _get_next_change_id(self) -> str:
        """Get next track change ID and increment counter"""
        cid = str(self.next_change_id)
        self.next_change_id += 1
        return cid

    def _escape_xml(self, text: str) -> str:
        """Escape XML special characters and remove illegal control characters"""
        text = sanitize_xml_string(text)
        return (text
            .replace('&', '&amp;')
            .replace('<', '&lt;')
            .replace('>', '&gt;')
            .replace('"', '&quot;'))

    def _parse_formatted_text(self, text: str) -> List[Tuple[str, Optional[str]]]:
        """
        Parse text with <sup>/<sub> markup into segments with format info.
        
        This function splits text containing superscript/subscript markup into
        segments, each with its associated vertical alignment type.
        
        Args:
            text: Text possibly containing <sup>...</sup> or <sub>...</sub> markup
        
        Returns:
            List of (text_content, vert_align) tuples where:
            - text_content: The actual text without markup tags
            - vert_align: 'superscript' | 'subscript' | None
        
        Examples:
            "x<sup>2</sup>" -> [("x", None), ("2", "superscript")]
            "H<sub>2</sub>O" -> [("H", None), ("2", "subscript"), ("O", None)]
            "normal text" -> [("normal text", None)]
        """
        if '<sup>' not in text and '<sub>' not in text:
            # Fast path: no markup
            return [(text, None)]
        
        segments = []
        pos = 0
        
        # Pattern to match <sup>...</sup> or <sub>...</sub>
        # Non-greedy match to handle multiple tags correctly
        pattern = re.compile(r'<(sup|sub)>(.*?)</\1>', re.DOTALL)
        
        for match in pattern.finditer(text):
            # Add text before this tag (if any)
            if match.start() > pos:
                segments.append((text[pos:match.start()], None))
            
            # Add the tagged content
            tag_type = match.group(1)  # 'sup' or 'sub'
            tag_content = match.group(2)
            vert_align = 'superscript' if tag_type == 'sup' else 'subscript'
            segments.append((tag_content, vert_align))
            
            pos = match.end()
        
        # Add remaining text after last tag (if any)
        if pos < len(text):
            segments.append((text[pos:], None))
        
        return segments

    def _create_run(self, text: str, rPr_xml: str = '') -> etree.Element:
        """
        Create new w:r element(s) with text, supporting <sup>/<sub> markup.
        
        If text contains superscript/subscript markup, returns a document fragment
        with multiple runs (each with appropriate w:vertAlign in w:rPr).
        
        Args:
            text: Text to insert, may contain <sup>...</sup> or <sub>...</sub>
            rPr_xml: Base run properties XML (will be modified to add w:vertAlign)
        
        Returns:
            Single w:r element or document fragment containing multiple w:r elements
        """
        segments = self._parse_formatted_text(text)
        
        if len(segments) == 1 and segments[0][1] is None:
            # Simple case: no formatting, return single run
            run_xml = f'<w:r xmlns:w="{NS["w"]}">{rPr_xml}<w:t>{self._escape_xml(text)}</w:t></w:r>'
            run = etree.fromstring(run_xml)
            t_elem = run.find(f'{{{NS["w"]}}}t')
            if t_elem is not None:
                # Preserve leading/trailing spaces in inserted/replaced text.
                t_elem.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
            return run
        
        # Complex case: need multiple runs with different formatting
        # Parse base rPr to potentially modify it
        if rPr_xml:
            # Check if rPr_xml is already a complete <w:rPr> element
            # _get_rPr_xml() returns the full element, so parse it directly
            if rPr_xml.strip().startswith('<w:rPr') or rPr_xml.strip().startswith('<rPr'):
                # Already a complete element, parse directly
                base_rPr = etree.fromstring(rPr_xml)
            else:
                # Just inner content, wrap in <w:rPr>
                base_rPr = etree.fromstring(f'<w:rPr xmlns:w="{NS["w"]}">{rPr_xml}</w:rPr>')
        else:
            base_rPr = None
        
        # Create a container for multiple runs
        container = etree.Element('container')
        
        for segment_text, vert_align in segments:
            if not segment_text:
                continue  # Skip empty segments
            
            # Start with a copy of base rPr
            if base_rPr is not None:
                run_rPr = etree.fromstring(etree.tostring(base_rPr, encoding='unicode'))
            else:
                run_rPr = etree.Element(f'{{{NS["w"]}}}rPr')
            
            # Add or update w:vertAlign if needed
            if vert_align:
                # Remove existing w:vertAlign if present
                for existing in run_rPr.findall('w:vertAlign', NS):
                    run_rPr.remove(existing)
                
                # Add new w:vertAlign
                vert_align_elem = etree.SubElement(run_rPr, f'{{{NS["w"]}}}vertAlign')
                vert_align_elem.set(f'{{{NS["w"]}}}val', vert_align)
            else:
                # For normal segments (vert_align=None), strip any inherited vertAlign
                # to prevent normal text from being incorrectly rendered as super/subscript
                for existing in run_rPr.findall('w:vertAlign', NS):
                    run_rPr.remove(existing)
            
            # Create run with modified rPr
            run = etree.Element(f'{{{NS["w"]}}}r')
            run.append(run_rPr)
            
            t_elem = etree.SubElement(run, f'{{{NS["w"]}}}t')
            # Preserve leading/trailing spaces in inserted/replaced text.
            t_elem.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
            t_elem.text = segment_text
            
            container.append(run)
        
        # If only one run was created, return it directly
        if len(container) == 1:
            return container[0]
        
        # Otherwise return the container (caller will need to extract children)
        return container

    def _replace_runs(self, _para_elem, affected_runs: List[Dict],
                     new_elements: List[etree.Element]) -> bool:
        """
        Replace affected runs with new elements in the paragraph.
        
        Returns:
            True if replacement succeeded, False if DOM operation failed
        """
        if not affected_runs or not new_elements:
            return False
        
        first_run = affected_runs[0]['elem']
        parent = first_run.getparent()
        
        # Safety check: if element is no longer in DOM, skip
        if parent is None:
            if self.verbose:
                print("  [Warning] _replace_runs failed: first_run has no parent")
            return False
        
        try:
            insert_idx = list(parent).index(first_run)
        except ValueError:
            # Element no longer in parent
            if self.verbose:
                print("  [Warning] _replace_runs failed: first_run not in parent")
            return False
        
        # Remove old runs
        for info in affected_runs:
            try:
                if info['elem'].getparent() is parent:
                    parent.remove(info['elem'])
            except ValueError:
                pass  # Already removed
        
        # Insert new elements
        for i, elem in enumerate(new_elements):
            parent.insert(insert_idx + i, elem)
        
        return True

    def _check_overlap_with_revisions(self, affected_runs: List[Dict]) -> bool:
        """
        Check if any affected run is inside revision markup (w:del or w:ins).
        
        This indicates the text has been modified by a previous rule,
        and we should fallback to comment annotation.
        
        Returns:
            True if overlap detected (should fallback), False otherwise
        """
        for info in affected_runs:
            elem = info.get('elem')
            if elem is None:
                continue
            parent = elem.getparent()
            if parent is not None and parent.tag in (
                f'{{{NS["w"]}}}del',
                f'{{{NS["w"]}}}ins'
            ):
                return True
        return False

    def _find_revision_ancestor(self, elem, para_elem) -> Optional[etree.Element]:
        """
        Find the outermost revision container (w:del or w:ins) for an element.
        
        Walks up the tree from elem to para_elem, looking for revision markup.
        Returns the outermost revision container if found, None otherwise.
        
        Args:
            elem: The element to check (usually a w:r run element)
            para_elem: The paragraph element (stopping point for traversal)
        
        Returns:
            The outermost w:del or w:ins element, or None if not inside revision
        """
        revision_tags = (f'{{{NS["w"]}}}del', f'{{{NS["w"]}}}ins')
        outermost_revision = None
        
        current = elem
        while current is not None and current is not para_elem:
            parent = current.getparent()
            if parent is not None and parent.tag in revision_tags:
                outermost_revision = parent
            current = parent
        
        return outermost_revision

    def _apply_delete_cross_paragraph(self, orig_runs_info: List[Dict],
                                      orig_match_start: int,
                                      violation_text: str,
                                      violation_reason: str,
                                      author: str) -> str:
        """
        Apply delete across multiple paragraphs (body text only).

        Strategy:
        - Delete ranges per paragraph with track changes (w:del)
        - Preserve paragraph structure (do not remove w:p)
        - Add a comment for each paragraph segment
        """
        match_start = orig_match_start
        match_end = match_start + len(violation_text)

        affected = self._find_affected_runs(orig_runs_info, match_start, match_end)
        if not affected:
            if self.verbose:
                print("  [Cross-paragraph delete] No affected runs found")
            return self._set_status_reason(
                'cross_paragraph_fallback',
                'CP_DEL_NO_HIT',
                'delete hit not found',
            )

        real_runs = self._filter_real_runs(affected)
        if not real_runs:
            if self.verbose:
                print("  [Cross-paragraph delete] No real runs after filtering")
            return self._set_status_reason(
                'cross_paragraph_fallback',
                'CP_DEL_NO_RUN',
                'no editable runs',
            )

        # Do not apply cross-paragraph delete in table mode
        if self._is_table_mode(real_runs) or any(r.get('cell_elem') is not None for r in real_runs):
            if self.verbose:
                print("  [Cross-paragraph delete] Table mode detected, fallback")
            return self._set_status_reason(
                'cross_paragraph_fallback',
                'CP_DEL_TBL_MODE',
                'table mode unsupported',
            )

        # Check overlap with existing revisions
        if self._check_overlap_with_revisions(real_runs):
            if self.verbose:
                print("  [Cross-paragraph delete] Overlap with existing revisions")
            return self._set_status_reason(
                'conflict',
                'CF_OVERLAP',
                'overlaps existing revision',
            )

        para_groups = self._group_runs_by_paragraph(real_runs)
        if not para_groups:
            if self.verbose:
                print("  [Cross-paragraph delete] No paragraph groups")
            return self._set_status_reason(
                'cross_paragraph_fallback',
                'CP_DEL_NO_GRP',
                'run grouping failed',
            )

        if self.verbose:
            print(f"  [Cross-paragraph delete] Processing {len(para_groups)} paragraph(s)")

        # Use shared helper to build prepared deletion items
        prepared = self._prepare_deletion_items(para_groups, match_start, match_end)

        if not prepared:
            if self.verbose:
                print("  [Cross-paragraph delete] No paragraphs prepared for deletion")
            return self._set_status_reason(
                'cross_paragraph_fallback',
                'CP_DEL_NO_SEG',
                'no segments prepared',
            )

        # Pre-generate shared change_id and comment_id for unified deletion
        shared_change_id = self._get_next_change_id()
        comment_id = self.next_comment_id
        self.next_comment_id += 1

        # Use shared helper to apply deletions with paragraph mark merging
        success_count = self._delete_paragraphs_in_unit(
            prepared, shared_change_id, author, comment_id=comment_id
        )

        if success_count == 0:
            if self.verbose:
                print(f"  [Cross-paragraph delete] All paragraphs failed, fallback to comment")
            return self._set_status_reason(
                'cross_paragraph_fallback',
                'CP_DEL_ALL_FAIL',
                'all segments failed',
            )
        
        # Record single comment for the unified deletion
        self.comments.append({
            'id': comment_id,
            'text': violation_reason,
            'author': author
        })
        
        if success_count < len(prepared):
            if self.verbose:
                print(f"  [Cross-paragraph delete] Partial success: {success_count}/{len(prepared)} paragraphs")
        else:
            if self.verbose:
                print(f"  [Cross-paragraph delete] All {success_count} paragraphs succeeded with unified deletion")

        return 'success'

    def _apply_replace_cross_paragraph(self, orig_runs_info: List[Dict],
                                       orig_match_start: int,
                                       violation_text: str,
                                       revised_text: str,
                                       violation_reason: str,
                                       author: str) -> str:
        """
        Apply replace across multiple paragraphs with diff-level edits.

        Performs pre-checks (affected runs, table mode, overlap) then delegates
        to _apply_diff_per_paragraph for diff computation and per-paragraph apply.
        """
        match_start = orig_match_start
        match_end = match_start + len(violation_text)

        affected = self._find_affected_runs(orig_runs_info, match_start, match_end)
        if not affected:
            return self._set_status_reason(
                'cross_paragraph_fallback',
                'CP_REP_NO_HIT',
                'replace hit not found',
            )

        real_runs = self._filter_real_runs(affected)
        if not real_runs:
            return self._set_status_reason(
                'cross_paragraph_fallback',
                'CP_REP_NO_RUN',
                'no editable runs',
            )

        # Do not apply cross-paragraph replace in table mode
        if self._is_table_mode(real_runs) or any(r.get('cell_elem') is not None for r in real_runs):
            return self._set_status_reason(
                'cross_paragraph_fallback',
                'CP_REP_TBL_MODE',
                'table mode unsupported',
            )

        # Check overlap with existing revisions
        if self._check_overlap_with_revisions(real_runs):
            return self._set_status_reason(
                'conflict',
                'CF_OVERLAP',
                'overlaps existing revision',
            )

        # Build paragraph groups in order
        para_groups = self._group_runs_by_paragraph(real_runs)
        if not para_groups:
            return self._set_status_reason(
                'cross_paragraph_fallback',
                'CP_REP_NO_GRP',
                'run grouping failed',
            )

        # Build paragraph ranges within combined_text
        para_segments = self._build_para_segments_from_groups(
            para_groups,
            match_start,
            match_end
        )

        if not para_segments:
            return self._set_status_reason(
                'cross_paragraph_fallback',
                'CP_REP_SEG_FAIL',
                'segment build failed',
            )

        return self._apply_diff_per_paragraph(
            para_segments, violation_text, revised_text,
            violation_reason, author,
            fallback_status='cross_paragraph_fallback',
            strip_runs=False,
        )

    def _apply_delete(self, para_elem, violation_text: str,
                     violation_reason: str,
                     orig_runs_info: List[Dict],
                     orig_match_start: int,
                     author: str) -> str:
        """
        Apply delete operation with track changes and comment annotation.

        Args:
            para_elem: Paragraph element
            violation_text: Text to delete
            violation_reason: Reason for violation (used as comment text)
            orig_runs_info: Pre-computed original runs info from _process_item
            orig_match_start: Pre-computed match position in original text
            author: Track change author (base author + category suffix)

        Returns:
            'success': Deletion applied successfully
            'fallback': Should fallback to comment annotation
            'equation_fallback': Equation-only content cannot be edited
            'conflict': Text overlaps with previous rule modification
        """
        # Use original text position directly
        match_start = orig_match_start
        match_end = match_start + len(violation_text)
        affected = self._find_affected_runs(orig_runs_info, match_start, match_end)

        if not affected:
            return self._set_status_reason(
                'fallback',
                'FB_DEL_NO_HIT',
                'delete hit not found',
            )

        # Filter out synthetic boundary runs (JSON boundaries, para boundaries)
        real_runs = self._filter_real_runs(affected)
        if not real_runs:
            if any(r.get('is_equation', False) for r in affected):
                return self._set_status_reason(
                    'equation_fallback',
                    'EQ_DEL_ONLY',
                    'can not modify equation',
                )
            return self._set_status_reason(
                'fallback',
                'FB_DEL_NO_RUN',
                'no editable runs',
            )

        # Check if text overlaps with previous modifications
        if self._check_overlap_with_revisions(real_runs):
            return self._set_status_reason(
                'conflict',
                'CF_OVERLAP',
                'overlaps existing revision',
            )

        rPr_xml = self._get_rPr_xml(real_runs[0].get('rPr'))
        change_id = self._get_next_change_id()

        # Allocate comment ID for wrapping the deleted text
        comment_id = self.next_comment_id
        self.next_comment_id += 1

        # Calculate split points using real runs only
        # Use original text for mutations (not JSON-escaped text)
        first_run = real_runs[0]
        last_run = real_runs[-1]
        first_orig_text = self._get_run_original_text(first_run)
        last_orig_text = self._get_run_original_text(last_run)

        # Translate offsets from escaped space to original space
        before_offset = self._translate_escaped_offset(first_run, max(0, match_start - first_run['start']))
        after_offset = self._translate_escaped_offset(last_run, max(0, match_end - last_run['start']))

        before_text = first_orig_text[:before_offset]
        after_text = last_orig_text[after_offset:]

        new_elements = []

        # Before text (unchanged)
        if before_text:
            run_or_container = self._create_run(before_text, rPr_xml)
            if run_or_container.tag == 'container':
                new_elements.extend(list(run_or_container))
            else:
                new_elements.append(run_or_container)

        # Comment range start (before deleted text)
        comment_start_xml = f'<w:commentRangeStart xmlns:w="{NS["w"]}" w:id="{comment_id}"/>'
        new_elements.append(etree.fromstring(comment_start_xml))

        # Decode violation_text if in table mode (JSON-escaped)
        if self._is_table_mode(real_runs):
            del_text = self._decode_json_escaped(violation_text)
        else:
            del_text = violation_text

        # Deleted text - use _create_run to handle <sup>/<sub> markup
        change_id = self._get_next_change_id()
        
        # Create runs with proper vertAlign formatting
        run_or_container = self._create_run(del_text, rPr_xml)
        
        if run_or_container.tag == 'container':
            # Multiple runs (has markup) - wrap each in w:del
            for del_run in run_or_container:
                # Change w:t to w:delText
                t_elem = del_run.find(f'{{{NS["w"]}}}t')
                if t_elem is not None:
                    t_elem.tag = f'{{{NS["w"]}}}delText'
                
                # Wrap in w:del
                del_elem = etree.Element(f'{{{NS["w"]}}}del')
                del_elem.set(f'{{{NS["w"]}}}id', change_id)
                del_elem.set(f'{{{NS["w"]}}}author', author)
                del_elem.set(f'{{{NS["w"]}}}date', self.operation_timestamp)
                del_elem.append(del_run)
                new_elements.append(del_elem)
        else:
            # Single run - wrap in w:del
            t_elem = run_or_container.find(f'{{{NS["w"]}}}t')
            if t_elem is not None:
                t_elem.tag = f'{{{NS["w"]}}}delText'
            
            del_elem = etree.Element(f'{{{NS["w"]}}}del')
            del_elem.set(f'{{{NS["w"]}}}id', change_id)
            del_elem.set(f'{{{NS["w"]}}}author', author)
            del_elem.set(f'{{{NS["w"]}}}date', self.operation_timestamp)
            del_elem.append(run_or_container)
            new_elements.append(del_elem)

        # Comment range end and reference (after deleted text)
        comment_end_xml = f'<w:commentRangeEnd xmlns:w="{NS["w"]}" w:id="{comment_id}"/>'
        new_elements.append(etree.fromstring(comment_end_xml))

        comment_ref_xml = f'''<w:r xmlns:w="{NS['w']}">
            <w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>
            <w:commentReference w:id="{comment_id}"/>
        </w:r>'''
        new_elements.append(etree.fromstring(comment_ref_xml))

        # After text (unchanged)
        if after_text:
            run_or_container = self._create_run(after_text, rPr_xml)
            if run_or_container.tag == 'container':
                new_elements.extend(list(run_or_container))
            else:
                new_elements.append(run_or_container)

        self._replace_runs(para_elem, real_runs, new_elements)

        # Record comment with violation_reason as content
        self.comments.append({
            'id': comment_id,
            'text': violation_reason,
            'author': author
        })

        return 'success'

    def _apply_replace(self, para_elem, violation_text: str,
                      revised_text: str,
                      violation_reason: str,
                      orig_runs_info: List[Dict],
                      orig_match_start: int,
                      author: str,
                      skip_comment: bool = False) -> str:
        """
        Apply replace operation with diff-based track changes and comment annotation.

        Strategy: Build all elements in a single pass, preserving original elements
        for 'equal' portions (to keep formatting, images, etc.) and creating
        track changes for delete/insert portions.

        Args:
            para_elem: Paragraph element
            violation_text: Text to replace
            revised_text: New text
            violation_reason: Reason for violation (used as comment text)
            orig_runs_info: Pre-computed original runs info from _process_item
            orig_match_start: Pre-computed match position in original text
            author: Track change author (base author + category suffix)
            skip_comment: If True, skip adding comment (used for multi-cell operations)

        Returns:
            'success': Replace applied (may be partial)
            'fallback': Should fallback to comment annotation
            'equation_fallback': Equation-only content cannot be edited
            'conflict': Modifying text overlaps with previous rule modification
        """
        # Use original text position directly
        match_start = orig_match_start
        match_end = match_start + len(violation_text)
        affected = self._find_affected_runs(orig_runs_info, match_start, match_end)

        if not affected:
            return self._set_status_reason(
                'fallback',
                'FB_REP_NO_HIT',
                'replace hit not found',
            )

        # Filter out synthetic boundary runs (JSON boundaries, para boundaries)
        real_runs = self._filter_real_runs(affected)
        if not real_runs:
            if any(r.get('is_equation', False) for r in affected):
                return self._set_status_reason(
                    'equation_fallback',
                    'EQ_REP_ONLY',
                    'equation-only target',
                )
            return self._set_status_reason(
                'fallback',
                'FB_REP_NO_RUN',
                'no editable runs',
            )

        # Check if we need markup-aware diff (detect <sup>/<sub> tags)
        has_markup = ('<sup>' in violation_text or '<sub>' in violation_text or 
                      '<sup>' in revised_text or '<sub>' in revised_text)
        
        if has_markup:
            # Use markup-aware diff that preserves formatting
            diff_ops = self._calculate_markup_aware_diff(violation_text, revised_text)
        else:
            # Use standard character-level diff for plain text
            # Convert to markup-aware format for consistent handling
            plain_diff = self._calculate_diff(violation_text, revised_text)
            diff_ops = [(op, text, None) for op, text in plain_diff]

        # Check for special element modification (drawing/equation)
        should_reject, reject_reason = self._check_special_element_modification(violation_text, diff_ops, has_markup)
        if should_reject:
            if self.verbose:
                print(f"  [Fallback] {reject_reason}")
            return self._set_status_reason(
                'fallback',
                'FB_REP_SPECIAL',
                reject_reason,
            )

        # Check for conflicts only on delete operations
        current_pos = match_start
        for op_tuple in diff_ops:
            # Extract operation and text (ignore vert_align for conflict checking)
            if len(op_tuple) == 3:
                op, text, _ = op_tuple
            else:
                op, text = op_tuple
            
            if op == 'delete':
                del_end = current_pos + len(text)
                del_affected = self._find_affected_runs(orig_runs_info, current_pos, del_end)
                del_real = self._filter_real_runs(del_affected)
                if self._check_overlap_with_revisions(del_real):
                    return self._set_status_reason(
                        'conflict',
                        'CF_OVERLAP',
                        'overlaps existing revision',
                    )
                current_pos = del_end
            elif op == 'equal':
                current_pos += len(text)
            # insert doesn't consume original text position

        rPr_xml = self._get_rPr_xml(real_runs[0].get('rPr'))

        # Allocate comment ID for wrapping the replaced text (unless skipped)
        comment_id = None
        if not skip_comment:
            comment_id = self.next_comment_id
            self.next_comment_id += 1

        # Calculate split points for before/after text using real runs
        # Use original text for mutations (not JSON-escaped text)
        first_run = real_runs[0]
        last_run = real_runs[-1]
        first_orig_text = self._get_run_original_text(first_run)
        last_orig_text = self._get_run_original_text(last_run)

        # Translate offsets from escaped space to original space
        before_offset = self._translate_escaped_offset(first_run, max(0, match_start - first_run['start']))
        after_offset = self._translate_escaped_offset(last_run, max(0, match_end - last_run['start']))

        before_text = first_orig_text[:before_offset]
        after_text = last_orig_text[after_offset:]

        # Build new elements in a single pass
        new_elements = []

        # Before text (unchanged part before the match)
        if before_text:
            run_or_container = self._create_run(before_text, rPr_xml)
            if run_or_container.tag == 'container':
                new_elements.extend(list(run_or_container))
            else:
                new_elements.append(run_or_container)

        # Comment range start (before replaced text) - only if not skipped
        if not skip_comment:
            comment_start_xml = f'<w:commentRangeStart xmlns:w="{NS["w"]}" w:id="{comment_id}"/>'
            new_elements.append(etree.fromstring(comment_start_xml))

        # Check if we're in table mode (need to decode JSON-escaped text)
        is_table_mode = self._is_table_mode(real_runs)

        # Process diff operations
        violation_pos = 0  # Position within violation_text (plain text, no markup)
        # Track position in original violation_text (with markup tags) for has_markup mode.
        # This allows correct mapping back to orig_runs_info positions when <sup>/<sub> tags
        # cause plain-text positions to diverge from combined_text positions.
        violation_pos_orig = 0
        # Track consumed standalone equation elements for cleanup after replacement
        consumed_equation_elems = []

        for op_tuple in diff_ops:
            # Extract components - handle both 2-tuple and 3-tuple formats
            if len(op_tuple) == 3:
                op, text, vert_align = op_tuple
            else:
                op, text = op_tuple
                vert_align = None
            if op == 'equal':
                # Calculate the original-coordinate length of this segment
                if vert_align == 'superscript':
                    orig_text_len = 5 + len(text) + 6  # <sup>text</sup>
                elif vert_align == 'subscript':
                    orig_text_len = 5 + len(text) + 6  # <sub>text</sub>
                else:
                    orig_text_len = len(text)

                # Use original-coordinate position for run lookup (works for both paths)
                equal_start = match_start + violation_pos_orig
                equal_end = equal_start + orig_text_len

                # Check if this equal segment contains special element placeholders
                has_special_elems = (EQUATION_PATTERN.search(text) is not None or
                                     DRAWING_PATTERN.search(text) is not None)

                if has_markup and not has_special_elems:
                    # No equations: recreate text with proper formatting (original behavior)
                    # Decode if in table mode
                    equal_text = self._decode_json_escaped(text) if is_table_mode else text

                    # Wrap with markup if vert_align is specified
                    if vert_align == 'superscript':
                        equal_text = f'<sup>{equal_text}</sup>'
                    elif vert_align == 'subscript':
                        equal_text = f'<sub>{equal_text}</sub>'

                    # Create run with proper vertAlign formatting
                    run_or_container = self._create_run(equal_text, rPr_xml)
                    if run_or_container.tag == 'container':
                        new_elements.extend(list(run_or_container))
                    else:
                        new_elements.append(run_or_container)
                else:
                    # Preserve original elements (images, equations)
                    equal_runs = self._find_affected_runs(orig_runs_info, equal_start, equal_end)

                    if equal_runs:
                        # Check if this is a single special element run that matches exactly
                        if (len(equal_runs) == 1 and
                            equal_runs[0].get('is_drawing') and
                            equal_runs[0]['start'] == equal_start and
                            equal_runs[0]['end'] == equal_end):
                            # Copy original image element
                            new_elements.append(copy.deepcopy(equal_runs[0]['elem']))
                        elif (len(equal_runs) == 1 and
                              equal_runs[0].get('is_equation') and
                              equal_runs[0]['start'] == equal_start and
                              equal_runs[0]['end'] == equal_end):
                            # Copy original equation element
                            omath = equal_runs[0].get('omath_elem')
                            if omath is not None:
                                new_elements.append(copy.deepcopy(omath))
                                consumed_equation_elems.append(equal_runs[0])
                        else:
                            # Multiple runs or partial match - iterate
                            for eq_run in equal_runs:
                                escaped_text = eq_run['text']
                                orig_text = self._get_run_original_text(eq_run)

                                # Calculate offsets in escaped space, then translate
                                escaped_start = max(0, equal_start - eq_run['start'])
                                escaped_end = min(len(escaped_text), equal_end - eq_run['start'])

                                if eq_run.get('is_drawing'):
                                    # Image run - copy entire element if fully contained
                                    if escaped_start == 0 and escaped_end == len(escaped_text):
                                        new_elements.append(copy.deepcopy(eq_run['elem']))
                                elif eq_run.get('is_equation'):
                                    # Equation run - copy entire m:oMath if fully contained
                                    if escaped_start == 0 and escaped_end == len(escaped_text):
                                        omath = eq_run.get('omath_elem')
                                        if omath is not None:
                                            new_elements.append(copy.deepcopy(omath))
                                            consumed_equation_elems.append(eq_run)
                                else:
                                    # Text run
                                    # Translate to original space
                                    orig_start = self._translate_escaped_offset(eq_run, escaped_start)
                                    orig_end = self._translate_escaped_offset(eq_run, escaped_end)
                                    portion = orig_text[orig_start:orig_end]

                                    if portion:
                                        # For has_markup text runs, wrap with vert_align
                                        if has_markup and vert_align == 'superscript':
                                            portion = f'<sup>{portion}</sup>'
                                        elif has_markup and vert_align == 'subscript':
                                            portion = f'<sub>{portion}</sub>'
                                        run_or_container = self._create_run(portion, rPr_xml)
                                        if run_or_container.tag == 'container':
                                            new_elements.extend(list(run_or_container))
                                        else:
                                            new_elements.append(run_or_container)
                    else:
                        # No runs found, create text directly
                        # Decode if in table mode and use _create_run to handle markup
                        equal_text = self._decode_json_escaped(text) if is_table_mode else text
                        if has_markup:
                            if vert_align == 'superscript':
                                equal_text = f'<sup>{equal_text}</sup>'
                            elif vert_align == 'subscript':
                                equal_text = f'<sub>{equal_text}</sub>'
                        run_or_container = self._create_run(equal_text, rPr_xml)
                        if run_or_container.tag == 'container':
                            new_elements.extend(list(run_or_container))
                        else:
                            new_elements.append(run_or_container)

                violation_pos += len(text)
                violation_pos_orig += orig_text_len

            elif op == 'delete':
                # Decode if in table mode
                del_text = self._decode_json_escaped(text) if is_table_mode else text
                
                # Wrap with markup if vert_align is specified
                if vert_align == 'superscript':
                    del_text = f'<sup>{del_text}</sup>'
                elif vert_align == 'subscript':
                    del_text = f'<sub>{del_text}</sub>'
                
                change_id = self._get_next_change_id()
                
                # Create runs with proper vertAlign formatting
                run_or_container = self._create_run(del_text, rPr_xml)
                
                if run_or_container.tag == 'container':
                    # Multiple runs (has markup) - wrap each in w:del
                    for del_run in run_or_container:
                        # Change w:t to w:delText
                        t_elem = del_run.find(f'{{{NS["w"]}}}t')
                        if t_elem is not None:
                            t_elem.tag = f'{{{NS["w"]}}}delText'
                        
                        # Wrap in w:del
                        del_elem = etree.Element(f'{{{NS["w"]}}}del')
                        del_elem.set(f'{{{NS["w"]}}}id', change_id)
                        del_elem.set(f'{{{NS["w"]}}}author', author)
                        del_elem.set(f'{{{NS["w"]}}}date', self.operation_timestamp)
                        del_elem.append(del_run)
                        new_elements.append(del_elem)
                else:
                    # Single run - wrap in w:del
                    t_elem = run_or_container.find(f'{{{NS["w"]}}}t')
                    if t_elem is not None:
                        t_elem.tag = f'{{{NS["w"]}}}delText'
                    
                    del_elem = etree.Element(f'{{{NS["w"]}}}del')
                    del_elem.set(f'{{{NS["w"]}}}id', change_id)
                    del_elem.set(f'{{{NS["w"]}}}author', author)
                    del_elem.set(f'{{{NS["w"]}}}date', self.operation_timestamp)
                    del_elem.append(run_or_container)
                    new_elements.append(del_elem)
                
                violation_pos += len(text)
                # Track original position for delete ops (same logic as equal)
                if vert_align == 'superscript':
                    violation_pos_orig += 5 + len(text) + 6
                elif vert_align == 'subscript':
                    violation_pos_orig += 5 + len(text) + 6
                else:
                    violation_pos_orig += len(text)

            elif op == 'insert':
                # Decode if in table mode
                ins_text = self._decode_json_escaped(text) if is_table_mode else text
                
                # Wrap with markup if vert_align is specified
                if vert_align == 'superscript':
                    ins_text = f'<sup>{ins_text}</sup>'
                elif vert_align == 'subscript':
                    ins_text = f'<sub>{ins_text}</sub>'

                # Handle soft line breaks for inserted text
                if '\n' in ins_text:
                    ins_elements = self._build_ins_elements_with_breaks(ins_text, rPr_xml, author)
                    new_elements.extend(ins_elements)
                    # insert doesn't consume violation_pos
                    continue
                
                change_id = self._get_next_change_id()
                
                # Create runs with proper vertAlign formatting
                run_or_container = self._create_run(ins_text, rPr_xml)
                
                if run_or_container.tag == 'container':
                    # Multiple runs (has markup) - wrap each in w:ins
                    for ins_run in run_or_container:
                        # Wrap in w:ins
                        ins_elem = etree.Element(f'{{{NS["w"]}}}ins')
                        ins_elem.set(f'{{{NS["w"]}}}id', change_id)
                        ins_elem.set(f'{{{NS["w"]}}}author', author)
                        ins_elem.set(f'{{{NS["w"]}}}date', self.operation_timestamp)
                        ins_elem.append(ins_run)
                        new_elements.append(ins_elem)
                else:
                    # Single run - wrap in w:ins
                    ins_elem = etree.Element(f'{{{NS["w"]}}}ins')
                    ins_elem.set(f'{{{NS["w"]}}}id', change_id)
                    ins_elem.set(f'{{{NS["w"]}}}author', author)
                    ins_elem.set(f'{{{NS["w"]}}}date', self.operation_timestamp)
                    ins_elem.append(run_or_container)
                    new_elements.append(ins_elem)
                # insert doesn't consume violation_pos

        # Comment range end and reference (after replaced text) - only if not skipped
        if not skip_comment:
            comment_end_xml = f'<w:commentRangeEnd xmlns:w="{NS["w"]}" w:id="{comment_id}"/>'
            new_elements.append(etree.fromstring(comment_end_xml))

            comment_ref_xml = f'''<w:r xmlns:w="{NS['w']}">
                <w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>
                <w:commentReference w:id="{comment_id}"/>
            </w:r>'''
            new_elements.append(etree.fromstring(comment_ref_xml))

        # After text (unchanged part after the match)
        if after_text:
            run_or_container = self._create_run(after_text, rPr_xml)
            if run_or_container.tag == 'container':
                new_elements.extend(list(run_or_container))
            else:
                new_elements.append(run_or_container)

        # Single DOM operation to replace all real runs (not boundary markers)
        self._replace_runs(para_elem, real_runs, new_elements)

        # Remove original standalone equation elements that were deep-copied into new_elements.
        # Without this, standalone m:oMath elements (not inside any w:r) would remain in the
        # paragraph and get pushed to the end, since _filter_real_runs excludes them.
        # For inline equations (inside a w:r), the host w:r is already removed by _replace_runs.
        for eq_info in consumed_equation_elems:
            omath = eq_info.get('omath_elem')
            if omath is not None and omath.getparent() is not None:
                # Only remove if the element is a direct child of the paragraph
                # (standalone equation). If it was inside a w:r, the w:r removal
                # already handled it.
                if eq_info.get('elem') is None:
                    # Standalone equation (no host run) - remove from paragraph
                    omath.getparent().remove(omath)

        # Record comment with violation_reason as content (only if not skipped)
        if not skip_comment:
            self.comments.append({
                'id': comment_id,
                'text': violation_reason,
                'author': author
            })

        return 'success'
