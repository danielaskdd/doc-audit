"""High-level editing actions including revisions and comments."""

import copy
import re
from datetime import datetime, timezone
from typing import Dict, Generator, Iterator, List, Optional, Tuple

from lxml import etree

from .common import (
    COMMENTS_CONTENT_TYPE,
    NS,
    EditItem,
    PackURI,
    Part,
    sanitize_xml_string,
)
from .edit_primitives import DocxEditPrimitivesMixin


class DocxEditActionsMixin(DocxEditPrimitivesMixin):
            def _init_comment_id(self):
                """Initialize next comment ID by scanning existing comments"""
                max_id = -1
        
                # Check comments.xml via OPC API
                try:
                    from docx.opc.constants import RELATIONSHIP_TYPE as RT
                    comments_part = self.doc.part.part_related_by(RT.COMMENTS)
                    comments_xml = etree.fromstring(comments_part.blob)
            
                    for comment in comments_xml.findall(f'{{{NS["w"]}}}comment'):
                        cid = comment.get(f'{{{NS["w"]}}}id')
                        if cid:
                            try:
                                max_id = max(max_id, int(cid))
                            except ValueError:
                                pass
                except (KeyError, AttributeError):
                    pass
        
                # Also check document.xml for comment references
                for tag in ('commentRangeStart', 'commentRangeEnd', 'commentReference'):
                    for elem in self.body_elem.iter(f'{{{NS["w"]}}}{tag}'):
                        cid = elem.get(f'{{{NS["w"]}}}id')
                        if cid:
                            try:
                                max_id = max(max_id, int(cid))
                            except ValueError:
                                pass
        
                self.next_comment_id = max_id + 1

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
                    return etree.fromstring(run_xml)
        
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
                    return 'cross_paragraph_fallback'

                real_runs = self._filter_real_runs(affected)
                if not real_runs:
                    if self.verbose:
                        print("  [Cross-paragraph delete] No real runs after filtering")
                    return 'cross_paragraph_fallback'

                # Do not apply cross-paragraph delete in table mode
                if self._is_table_mode(real_runs) or any(r.get('cell_elem') is not None for r in real_runs):
                    if self.verbose:
                        print("  [Cross-paragraph delete] Table mode detected, fallback")
                    return 'cross_paragraph_fallback'

                # Check overlap with existing revisions
                if self._check_overlap_with_revisions(real_runs):
                    if self.verbose:
                        print("  [Cross-paragraph delete] Overlap with existing revisions")
                    return 'conflict'

                para_groups = self._group_runs_by_paragraph(real_runs)
                if not para_groups:
                    if self.verbose:
                        print("  [Cross-paragraph delete] No paragraph groups")
                    return 'cross_paragraph_fallback'

                if self.verbose:
                    print(f"  [Cross-paragraph delete] Processing {len(para_groups)} paragraph(s)")

                # Use shared helper to build prepared deletion items
                prepared = self._prepare_deletion_items(para_groups, match_start, match_end)

                if not prepared:
                    if self.verbose:
                        print("  [Cross-paragraph delete] No paragraphs prepared for deletion")
                    return 'cross_paragraph_fallback'

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
                    return 'cross_paragraph_fallback'
        
                # Record single comment for the unified deletion
                self.comments.append({
                    'id': comment_id,
                    'text': violation_reason,
                    'author': f"{author}-R"
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

                Strategy:
                - Compute diff between violation_text and revised_text
                - Apply per-paragraph replace only on changed segments
                - If diff deletes paragraph boundary ('\\n'), merge paragraphs by removing
                  the boundary and appending following paragraph content to the previous one
                """
                match_start = orig_match_start
                match_end = match_start + len(violation_text)

                affected = self._find_affected_runs(orig_runs_info, match_start, match_end)
                if not affected:
                    return 'cross_paragraph_fallback'

                real_runs = self._filter_real_runs(affected)
                if not real_runs:
                    return 'cross_paragraph_fallback'

                # Do not apply cross-paragraph replace in table mode
                if self._is_table_mode(real_runs) or any(r.get('cell_elem') is not None for r in real_runs):
                    return 'cross_paragraph_fallback'

                # Check overlap with existing revisions
                if self._check_overlap_with_revisions(real_runs):
                    return 'conflict'

                # Compute diff ops
                has_markup = ('<sup>' in violation_text or '<sub>' in violation_text or
                              '<sup>' in revised_text or '<sub>' in revised_text)
                if has_markup:
                    diff_ops = self._calculate_markup_aware_diff(violation_text, revised_text)
                else:
                    plain_diff = self._calculate_diff(violation_text, revised_text)
                    diff_ops = [(op, text, None) for op, text in plain_diff]

                # Check for special element modification (drawing/equation)
                should_reject, reject_reason = self._check_special_element_modification(violation_text, diff_ops, has_markup)
                if should_reject:
                    if self.verbose:
                        print(f"  [Fallback] {reject_reason}")
                    return 'cross_paragraph_fallback'

                # Detect deleted paragraph boundaries
                boundary_positions = [idx for idx, ch in enumerate(violation_text) if ch == '\n']
                boundary_pos_to_idx = {pos: i for i, pos in enumerate(boundary_positions)}
                deleted_boundary_indices = set()
                orig_pos = 0
                for op_tuple in diff_ops:
                    op, text, _ = op_tuple if len(op_tuple) == 3 else (*op_tuple, None)
                    if op == 'delete' and '\n' in text:
                        for i, ch in enumerate(text):
                            if ch == '\n':
                                bpos = orig_pos + i
                                if bpos in boundary_pos_to_idx:
                                    deleted_boundary_indices.add(boundary_pos_to_idx[bpos])
                    if op in ('equal', 'delete'):
                        orig_pos += len(text)

                # Build combined text for the match range
                combined_text = ''.join(r.get('text', '') for r in orig_runs_info)
                match_text = combined_text[match_start:match_end]

                # Build paragraph groups in order
                para_groups = self._group_runs_by_paragraph(real_runs)
                if not para_groups:
                    return 'cross_paragraph_fallback'

                # Build paragraph ranges within combined_text
                para_ranges = []
                for group in para_groups:
                    runs = group['runs']
                    para_start = min(r['start'] for r in runs)
                    para_end = max(r['end'] for r in runs)
                    overlap_start = max(match_start, para_start)
                    overlap_end = min(match_end, para_end)
                    if overlap_start < overlap_end:
                        para_ranges.append({
                            'para_elem': group['para_elem'],
                            'overlap_start': overlap_start,
                            'overlap_end': overlap_end,
                            'para_start': para_start
                        })

                if not para_ranges:
                    return 'cross_paragraph_fallback'

                # Helper: extract revised segment for an original range
                def extract_revised_segment(seg_start: int, seg_end: int) -> str:
                    orig_pos = 0
                    parts = []
                    for op_tuple in diff_ops:
                        op, text, _ = op_tuple if len(op_tuple) == 3 else (*op_tuple, None)
                        if op == 'equal':
                            op_start = orig_pos
                            op_end = orig_pos + len(text)
                            # overlap with [seg_start, seg_end)
                            if op_end > seg_start and op_start < seg_end:
                                take_start = max(seg_start, op_start) - op_start
                                take_end = min(seg_end, op_end) - op_start
                                parts.append(text[take_start:take_end])
                            orig_pos += len(text)
                        elif op == 'delete':
                            orig_pos += len(text)
                        elif op == 'insert':
                            # insert occurs at current orig_pos
                            if seg_start <= orig_pos <= seg_end:
                                parts.append(text)
                    return ''.join(parts)

                # Apply per-paragraph replace
                any_applied = False
                for pr in para_ranges:
                    para_elem = pr['para_elem']
                    seg_start = pr['overlap_start'] - match_start
                    seg_end = pr['overlap_end'] - match_start
                    orig_segment = match_text[seg_start:seg_end]
                    revised_segment = extract_revised_segment(seg_start, seg_end)

                    if orig_segment == revised_segment:
                        continue

                    para_runs_info, _ = self._collect_runs_info_original(para_elem)
                    match_start_in_para = pr['overlap_start'] - pr['para_start']
                    status = self._apply_replace(
                        para_elem,
                        orig_segment,
                        revised_segment,
                        violation_reason,
                        para_runs_info,
                        match_start_in_para,
                        author
                    )

                    if status == 'conflict':
                        return 'conflict'
                    if status not in ('success',):
                        return 'cross_paragraph_fallback'
                    any_applied = True

                # Merge paragraphs if boundary deleted
                if deleted_boundary_indices:
                    # Collect paragraph elements in order
                    para_elems = [pr['para_elem'] for pr in para_ranges]
                    # Merge from right to left to avoid index shifts
                    for b_idx in sorted(deleted_boundary_indices, reverse=True):
                        if b_idx < 0 or b_idx + 1 >= len(para_elems):
                            continue
                        prev_para = para_elems[b_idx]
                        next_para = para_elems[b_idx + 1]
                        # Move children (except pPr) from next to prev
                        for child in list(next_para):
                            if child.tag == f'{{{NS["w"]}}}pPr':
                                continue
                            next_para.remove(child)
                            prev_para.append(child)
                        parent = next_para.getparent()
                        if parent is not None:
                            try:
                                parent.remove(next_para)
                            except ValueError:
                                pass
                        # Remove from list to keep indices consistent
                        para_elems.pop(b_idx + 1)

                return 'success' if any_applied else 'success'

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
                    return 'fallback'

                # Filter out synthetic boundary runs (JSON boundaries, para boundaries)
                real_runs = self._filter_real_runs(affected)
                if not real_runs:
                    if any(r.get('is_equation', False) for r in affected):
                        return 'equation_fallback'
                    return 'fallback'

                # Check if text overlaps with previous modifications
                if self._check_overlap_with_revisions(real_runs):
                    return 'conflict'

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
                # Use "-R" suffix to distinguish comment author from track change author
                self.comments.append({
                    'id': comment_id,
                    'text': violation_reason,
                    'author': f"{author}-R"
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
                    return 'fallback'

                # Filter out synthetic boundary runs (JSON boundaries, para boundaries)
                real_runs = self._filter_real_runs(affected)
                if not real_runs:
                    if any(r.get('is_equation', False) for r in affected):
                        return 'equation_fallback'
                    return 'fallback'

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
                    return 'fallback'

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
                            return 'conflict'
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

                for op_tuple in diff_ops:
                    # Extract components - handle both 2-tuple and 3-tuple formats
                    if len(op_tuple) == 3:
                        op, text, vert_align = op_tuple
                    else:
                        op, text = op_tuple
                        vert_align = None
                    if op == 'equal':
                        # When has_markup=True, position mapping between plain text (violation_pos)
                        # and combined_text (orig_runs_info positions) is incorrect due to <sup>/<sub> tags.
                        # Skip run preservation and just recreate the text.
                        if has_markup:
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
                            # No markup: preserve original elements (especially images)
                            equal_start = match_start + violation_pos
                            equal_end = equal_start + len(text)
                            equal_runs = self._find_affected_runs(orig_runs_info, equal_start, equal_end)

                            if equal_runs:
                                # Check if this is a single image run that matches exactly
                                if (len(equal_runs) == 1 and
                                    equal_runs[0].get('is_drawing') and
                                    equal_runs[0]['start'] == equal_start and
                                    equal_runs[0]['end'] == equal_end):
                                    # Copy original image element
                                    new_elements.append(copy.deepcopy(equal_runs[0]['elem']))
                                else:
                                    # Extract text from the equal portion
                                    # Handle partial runs at boundaries
                                    # Use original text for document mutations
                                    for eq_run in equal_runs:
                                        escaped_text = eq_run['text']
                                        orig_text = self._get_run_original_text(eq_run)

                                        # Calculate offsets in escaped space, then translate
                                        escaped_start = max(0, equal_start - eq_run['start'])
                                        escaped_end = min(len(escaped_text), equal_end - eq_run['start'])

                                        # Translate to original space
                                        orig_start = self._translate_escaped_offset(eq_run, escaped_start)
                                        orig_end = self._translate_escaped_offset(eq_run, escaped_end)
                                        portion = orig_text[orig_start:orig_end]

                                        if eq_run.get('is_drawing'):
                                            # Image run - copy entire element if fully contained
                                            if escaped_start == 0 and escaped_end == len(escaped_text):
                                                new_elements.append(copy.deepcopy(eq_run['elem']))
                                        elif portion:
                                            run_or_container = self._create_run(portion, rPr_xml)
                                            if run_or_container.tag == 'container':
                                                new_elements.extend(list(run_or_container))
                                            else:
                                                new_elements.append(run_or_container)
                            else:
                                # No runs found, create text directly
                                # Decode if in table mode and use _create_run to handle markup
                                equal_text = self._decode_json_escaped(text) if is_table_mode else text
                                run_or_container = self._create_run(equal_text, rPr_xml)
                                if run_or_container.tag == 'container':
                                    new_elements.extend(list(run_or_container))
                                else:
                                    new_elements.append(run_or_container)

                        violation_pos += len(text)

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

                # Record comment with violation_reason as content (only if not skipped)
                # Use "-R" suffix to distinguish comment author from track change author
                if not skip_comment:
                    self.comments.append({
                        'id': comment_id,
                        'text': violation_reason,
                        'author': f"{author}-R"
                    })

                return 'success'

            def _apply_error_comment(self, para_elem, item: EditItem, author_override: str = None) -> bool:
                """
                Insert an unselected comment at the end of paragraph for failed items.
        
                Comment format: {WHY}<violation_reason>  {WHERE}<violation_text>{SUGGEST}<revised_text>
                Author: author_override if provided, otherwise {self.author}-{category}
        
                Args:
                    para_elem: Paragraph element to attach comment
                    item: EditItem with violation details
                    author_override: Optional author name to use instead of default
                """
                comment_id = self.next_comment_id
                self.next_comment_id += 1
        
                # Insert only commentReference at the end of paragraph (no range selection)
                ref_xml = f'''<w:r xmlns:w="{NS['w']}">
                    <w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>
                    <w:commentReference w:id="{comment_id}"/>
                </w:r>'''
                para_elem.append(etree.fromstring(ref_xml))
        
                # Record comment content with custom format and author
                comment_text = f"{{WHY}}{item.violation_reason}  {{WHERE}}{item.violation_text}{{SUGGEST}}{item.revised_text}"
                comment_author = author_override if author_override else self._author_for_item(item)
                self.comments.append({
                    'id': comment_id,
                    'text': comment_text,
                    'author': comment_author
                })
        
                return True

            def _apply_fallback_comment(self, para_elem, item: EditItem, reason: str = "") -> bool:
                """
                Insert a fallback comment when delete/replace/manual operation cannot be applied.
        
                This is different from _apply_error_comment:
                - Error comment: Unexpected failure
                - Fallback comment: Expected fallback (text was modified by previous rule)
        
                Args:
                    para_elem: Paragraph element
                    item: Edit item
                    reason: Reason for fallback (e.g., "Text modified by previous rule")
        
                Returns:
                    True (always succeeds)
                """
                comment_id = self.next_comment_id
                self.next_comment_id += 1
        
                # Insert commentReference at end of paragraph
                ref_xml = f'''<w:r xmlns:w="{NS['w']}">
                    <w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>
                    <w:commentReference w:id="{comment_id}"/>
                </w:r>'''
                para_elem.append(etree.fromstring(ref_xml))
        
                # Format: [FALLBACK] reason | {WHY} ... {WHERE} ... {SUGGEST} ...
                comment_text = f"[FALLBACK]{reason} {{WHY}}{item.violation_reason}  {{WHERE}}{item.violation_text} {{SUGGEST}}{item.revised_text}"
        
                self.comments.append({
                    'id': comment_id,
                    'text': comment_text,
                    'author': self._author_for_item(item)
                })
        
                return True

            def _apply_cell_fallback_comment(self, para_elem, cell_violation: str, 
                                              cell_revised: str, reason: str, 
                                              item: EditItem) -> bool:
                """
                Insert a fallback comment for a single failed cell in multi-cell operation.
        
                Args:
                    para_elem: Paragraph element in the failed cell
                    cell_violation: Cell-specific violation text
                    cell_revised: Cell-specific revised text
                    reason: Reason for failure (e.g., "Text not found")
                    item: Original edit item (for rule_id and violation_reason)
        
                Returns:
                    True (always succeeds)
                """
                comment_id = self.next_comment_id
                self.next_comment_id += 1
        
                # Insert commentReference at end of paragraph
                ref_xml = f'''<w:r xmlns:w="{NS['w']}">
                    <w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>
                    <w:commentReference w:id="{comment_id}"/>
                </w:r>'''
                para_elem.append(etree.fromstring(ref_xml))
        
                # Format: [CELL FAILED] reason | {WHY} ... {WHERE} ... {SUGGEST} ...
                comment_text = f"[CELL FAILED] {reason}\n{{WHY}}{item.violation_reason}  {{WHERE}}{cell_violation}{{SUGGEST}}{cell_revised}"
        
                self.comments.append({
                    'id': comment_id,
                    'text': comment_text,
                    'author': self._author_for_item(item)
                })
        
                return True

            def _apply_manual(self, para_elem, violation_text: str,
                             violation_reason: str, revised_text: str,
                             orig_runs_info: List[Dict],
                             orig_match_start: int,
                             author: str,
                             is_cross_paragraph: bool = False,
                             fallback_reason: Optional[str] = None) -> str:
                """
                Apply manual operation by adding a Word comment.
        
                This method preserves internal run structure (including images, revision 
                markup, etc.) by inserting comment markers at element boundaries rather
                than replacing the content.
        
                Strategy:
                1. Find the precise start/end positions in the original text
                2. For each boundary (start/end), check if the run is inside revision markup
                3. If inside revision: insert marker outside the revision container
                4. If not inside revision: split run if needed, insert marker at precise position
        
                Cross-paragraph support:
                - When is_cross_paragraph=True, commentRangeStart/End can span multiple paragraphs
                - Uses 'para_elem' field from runs to determine which paragraph to insert markers into
        
                Args:
                    para_elem: Paragraph element (anchor paragraph, may not be used in cross-para mode)
                    violation_text: Text to mark with comment
                    violation_reason: Reason to show in comment
                    revised_text: Suggestion to show in comment
                    orig_runs_info: Pre-computed original runs info from _process_item
                    orig_match_start: Pre-computed match position in original text
                    author: Comment author (base author + category suffix)
                    is_cross_paragraph: If True, handle cross-paragraph comment range
        
                Returns:
                    'success': Comment added successfully
                    'fallback': Should fallback to error comment
                """
                # Use original text position directly
                match_start = orig_match_start
                match_end = match_start + len(violation_text)
                affected = self._find_affected_runs(orig_runs_info, match_start, match_end)

                if not affected:
                    return 'fallback'

                # Filter out all synthetic boundary markers (JSON and paragraph boundaries)
                real_runs = self._filter_real_runs(affected, include_equations=True)

                if not real_runs:
                    return 'fallback'

                comment_id = self.next_comment_id
                self.next_comment_id += 1

                first_run_info = real_runs[0]
                last_run_info = real_runs[-1]

                def _doc_key_for_run(run_info: Dict) -> Optional[Tuple[int, int]]:
                    """Compute document-order key using host para if available, else para_elem."""
                    host_para = run_info.get('host_para_elem')
                    para = run_info.get('para_elem', para_elem)
                    key = None
                    if host_para is not None:
                        key = self._get_run_doc_key(run_info.get('elem'), host_para)
                    if key is None and para is not None and para is not host_para:
                        key = self._get_run_doc_key(run_info.get('elem'), para)
                    return key

                start_key = _doc_key_for_run(first_run_info)
                end_key = _doc_key_for_run(last_run_info)

                if start_key is not None and end_key is None:
                    # No valid end run to anchor range - fallback to reference-only comment
                    fallback_para = first_run_info.get('host_para_elem')
                    if fallback_para is None:
                        fallback_para = first_run_info.get('para_elem')
                    if fallback_para is None:
                        fallback_para = para_elem
                    if fallback_para is None:
                        print("  [Warning] Reference-only fallback failed: no anchor paragraph")
                        return 'fallback'
                    ref_xml = f'''<w:r xmlns:w="{NS['w']}">
                        <w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>
                        <w:commentReference w:id="{comment_id}"/>
                    </w:r>'''
                    fallback_para.append(etree.fromstring(ref_xml))
                    comment_text = f"[FALLBACK]Reference-only comment: {violation_reason}"
                    if revised_text:
                        comment_text += f"\nSuggestion: {revised_text}"
                    self.comments.append({
                        'id': comment_id,
                        'text': comment_text,
                        'author': author
                    })
                    return 'success'

                if start_key is not None and end_key is not None and end_key < start_key:
                    candidate = None
                    for run_info in reversed(real_runs):
                        cand_key = _doc_key_for_run(run_info)
                        if cand_key is not None and cand_key >= start_key:
                            candidate = run_info
                            break
                    if candidate is None:
                        fallback_para = first_run_info.get('host_para_elem')
                        if fallback_para is None:
                            fallback_para = first_run_info.get('para_elem')
                        if fallback_para is None:
                            fallback_para = para_elem
                        if fallback_para is None:
                            print("  [Warning] Reference-only fallback failed: no anchor paragraph")
                            return 'fallback'
                        ref_xml = f'''<w:r xmlns:w="{NS['w']}">
                            <w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>
                            <w:commentReference w:id="{comment_id}"/>
                        </w:r>'''
                        fallback_para.append(etree.fromstring(ref_xml))
                        comment_text = f"[FALLBACK]Reference-only comment: {violation_reason}"
                        if revised_text:
                            comment_text += f"\nSuggestion: {revised_text}"
                        self.comments.append({
                            'id': comment_id,
                            'text': comment_text,
                            'author': author
                        })
                        if self.verbose:
                            print("  [Debug] ParaId order inverted; fallback to reference-only comment")
                        return 'success'
                    if self.verbose:
                        print("  [Debug] ParaId order inverted; adjusted end run")
                    last_run_info = candidate
                    # Clamp match_end to avoid splitting beyond the adjusted end run
                    match_end = min(match_end, last_run_info.get('end', match_end))

                first_run = first_run_info.get('elem')
                last_run = last_run_info.get('elem')
                rPr_xml = self._get_rPr_xml(first_run_info.get('rPr'))
                no_split_start = first_run_info.get('is_equation', False)
                no_split_end = last_run_info.get('is_equation', False)

                # Host paragraph anchors (for vMerge continue cases where runs are reused)
                start_host_para = first_run_info.get('host_para_elem')
                end_host_para = last_run_info.get('host_para_elem')
                start_use_host = start_host_para is not None and start_host_para is not first_run_info.get('para_elem')
                end_use_host = end_host_para is not None and end_host_para is not last_run_info.get('para_elem')

                # If any run is a host mismatch (vMerge continue reuse), prefer host-para anchors
                has_host_mismatch = any(
                    r.get('host_para_elem') is not None and r.get('host_para_elem') is not r.get('para_elem')
                    for r in real_runs
                )
                if has_host_mismatch:
                    start_use_host = True
                    end_use_host = True

                if first_run is None:
                    start_use_host = True
                    no_split_start = True
                if last_run is None:
                    end_use_host = True
                    no_split_end = True

                if start_use_host and start_host_para is None:
                    start_host_para = first_run_info.get('para_elem', para_elem)
                if end_use_host and end_host_para is None:
                    end_host_para = last_run_info.get('para_elem', para_elem)

                # Choose a reference paragraph that actually has visible content.
                reference_para = None
                for run_info in reversed(real_runs):
                    para_candidate = run_info.get('host_para_elem')
                    if para_candidate is None:
                        para_candidate = run_info.get('para_elem')
                    if para_candidate is None:
                        continue
                    try:
                        _, para_text = self._collect_runs_info_original(para_candidate)
                        if para_text.strip():
                            reference_para = para_candidate
                            break
                    except Exception:
                        continue
                if reference_para is None:
                    reference_para = end_host_para if end_host_para is not None else start_host_para
                    if reference_para is None:
                        reference_para = para_elem

                # If host anchors are inverted, fallback to reference-only comment
                if start_use_host and end_use_host and start_host_para is not None and reference_para is not None:
                    self._init_para_order()
                    start_idx = self._para_order.get(id(start_host_para))
                    end_idx = self._para_order.get(id(reference_para))
                    if start_idx is not None and end_idx is not None and end_idx < start_idx:
                        fallback_para = start_host_para if start_host_para is not None else para_elem
                        if fallback_para is None:
                            print("  [Warning] Reference-only fallback failed: no anchor paragraph")
                            return 'fallback'
                        ref_xml = f'''<w:r xmlns:w="{NS['w']}">
                            <w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>
                            <w:commentReference w:id="{comment_id}"/>
                        </w:r>'''
                        fallback_para.append(etree.fromstring(ref_xml))
                        comment_text = f"[FALLBACK]Reference-only comment: {violation_reason}"
                        if revised_text:
                            comment_text += f"\nSuggestion: {revised_text}"
                        self.comments.append({
                            'id': comment_id,
                            'text': comment_text,
                            'author': author
                        })
                        if self.verbose:
                            print("  [Debug] ParaId order inverted; fallback to reference-only comment")
                        return 'success'

                # Get parent paragraphs (may be different in cross-paragraph mode)
                if is_cross_paragraph:
                    first_para = first_run_info.get('para_elem', para_elem)
                    last_para = last_run_info.get('para_elem', para_elem)
                else:
                    first_para = para_elem
                    last_para = para_elem

                # Check if start/end runs are inside revision markup
                start_revision = self._find_revision_ancestor(first_run, first_para)
                end_revision = self._find_revision_ancestor(last_run, last_para)

                # Calculate text split points using real runs
                # Use original text for mutations (not JSON-escaped text)
                first_orig_text = self._get_run_original_text(first_run_info)
                last_orig_text = self._get_run_original_text(last_run_info)

                # Translate offsets from escaped space to original space
                before_offset = self._translate_escaped_offset(first_run_info, max(0, match_start - first_run_info['start']))
                after_offset = self._translate_escaped_offset(last_run_info, max(0, match_end - last_run_info['start']))

                before_text = first_orig_text[:before_offset]
                after_text = last_orig_text[after_offset:]
        
                # Create comment markers
                range_start = etree.fromstring(
                    f'<w:commentRangeStart xmlns:w="{NS["w"]}" w:id="{comment_id}"/>'
                )
                range_end = etree.fromstring(
                    f'<w:commentRangeEnd xmlns:w="{NS["w"]}" w:id="{comment_id}"/>'
                )
                comment_ref = etree.fromstring(f'''<w:r xmlns:w="{NS['w']}">
                    <w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>
                    <w:commentReference w:id="{comment_id}"/>
                </w:r>''')
        
                # === Handle START position ===
                if start_use_host:
                    # Insert commentRangeStart at start of host paragraph (no run split)
                    target_para = start_host_para
                    if target_para is None:
                        print("  [Warning] Reference-only fallback failed: no anchor paragraph")
                        return 'fallback'
                    insert_idx = 0
                    for i, child in enumerate(list(target_para)):
                        if child.tag == f'{{{NS["w"]}}}pPr':
                            continue
                        insert_idx = i
                        break
                    target_para.insert(insert_idx, range_start)
                elif start_revision is not None:
                    # Start is inside revision: insert commentRangeStart before revision container
                    parent = start_revision.getparent()
                    if parent is not None:
                        idx = list(parent).index(start_revision)
                        parent.insert(idx, range_start)
                elif no_split_start:
                    # Start is in a non-text run (equation): insert range start before the run
                    parent = first_run.getparent() if first_run is not None else None
                    if parent is None:
                        return 'fallback'
                    try:
                        idx = list(parent).index(first_run)
                    except ValueError:
                        return 'fallback'
                    parent.insert(idx, range_start)
                else:
                    # Start is in normal run: may need to split
                    parent = first_run.getparent()
                    if parent is None:
                        return 'fallback'
            
                    try:
                        idx = list(parent).index(first_run)
                    except ValueError:
                        return 'fallback'
            
                    if before_text:
                        # Need to split: create run for before_text, insert before first_run
                        run_or_container = self._create_run(before_text, rPr_xml)
                        if run_or_container.tag == 'container':
                            # Unwrap container: insert each child run
                            for child_run in run_or_container:
                                parent.insert(idx, child_run)
                                idx += 1
                        else:
                            parent.insert(idx, run_or_container)
                            idx += 1
                
                        # Update first_run's text content (remove before_text portion)
                        t_elem = first_run.find('w:t', NS)
                        if t_elem is not None and t_elem.text:
                            t_elem.text = t_elem.text[len(before_text):]
            
                    # Insert commentRangeStart before the (possibly modified) first_run
                    parent.insert(idx, range_start)
        
                # === Handle END position ===
                if end_use_host:
                    # Insert commentRangeEnd and reference at end of host paragraph (no run split)
                    target_para = reference_para if reference_para is not None else end_host_para
                    if target_para is None:
                        print("  [Warning] Reference-only fallback failed: no anchor paragraph")
                        return 'fallback'
                    if self.verbose and reference_para is not None and end_host_para is not None and reference_para is not end_host_para:
                        print("  [Debug] End anchor moved to previous non-empty paragraph")
                    target_para.append(range_end)
                    target_para.append(comment_ref)
                elif end_revision is not None:
                    # End is inside revision: insert commentRangeEnd after revision container
                    parent = end_revision.getparent()
                    if parent is not None:
                        idx = list(parent).index(end_revision)
                        parent.insert(idx + 1, range_end)
                        parent.insert(idx + 2, comment_ref)
                elif no_split_end:
                    # End is in a non-text run (equation): insert range end after the run
                    parent = last_run.getparent() if last_run is not None else None
                    if parent is None:
                        return 'fallback'
                    try:
                        idx = list(parent).index(last_run)
                    except ValueError:
                        return 'fallback'
                    parent.insert(idx + 1, range_end)
                    parent.insert(idx + 2, comment_ref)
                else:
                    # End is in normal run: may need to split
                    parent = last_run.getparent()
                    if parent is None:
                        return 'fallback'
            
                    try:
                        idx = list(parent).index(last_run)
                    except ValueError:
                        return 'fallback'
            
                    if after_text:
                        # Need to split: update last_run to remove after_text, create new run for after_text
                        t_elem = last_run.find('w:t', NS)
                        if t_elem is not None and t_elem.text:
                            original_text = t_elem.text
                            # Keep only the portion up to match_end
                            keep_len = len(original_text) - len(after_text)
                            t_elem.text = original_text[:keep_len]
                
                        # Insert commentRangeEnd after last_run
                        parent.insert(idx + 1, range_end)
                        # Insert commentReference after range_end
                        parent.insert(idx + 2, comment_ref)
                        # Create after_run and insert after comment_ref
                        run_or_container = self._create_run(after_text, rPr_xml)
                        if run_or_container.tag == 'container':
                            # Unwrap container: insert each child run
                            insert_pos = idx + 3
                            for child_run in run_or_container:
                                parent.insert(insert_pos, child_run)
                                insert_pos += 1
                        else:
                            parent.insert(idx + 3, run_or_container)
                    else:
                        # No split needed: insert commentRangeEnd and reference after last_run
                        parent.insert(idx + 1, range_end)
                        parent.insert(idx + 2, comment_ref)

                # Record comment content
                if fallback_reason:
                    comment_text = f"[FALLBACK]{fallback_reason} {violation_reason}"
                else:
                    comment_text = violation_reason
                if revised_text:
                    comment_text += f"\nSuggestion: {revised_text}"
        
                self.comments.append({
                    'id': comment_id,
                    'text': comment_text,
                    'author': author
                })
        
                return 'success'

            def _save_comments(self):
                """Save comments to comments.xml using OPC API"""
                if not self.comments:
                    return

                from docx.opc.constants import RELATIONSHIP_TYPE as RT
        
                # Try to get existing comments.xml
                try:
                    comments_part = self.doc.part.part_related_by(RT.COMMENTS)
                    comments_xml = etree.fromstring(comments_part.blob)
                except (KeyError, AttributeError):
                    # Create new comments.xml
                    comments_xml = etree.fromstring(
                        f'<w:comments xmlns:w="{NS["w"]}" xmlns:w14="{NS["w14"]}"/>'
                    )
        
                # Add comments
                for comment in self.comments:
                    comment_elem = etree.SubElement(
                        comments_xml, f'{{{NS["w"]}}}comment'
                    )
                    comment_elem.set(f'{{{NS["w"]}}}id', str(comment['id']))
                    # Support independent author for each comment (author-category suffix)
                    comment_author = comment.get('author', self.author)
                    comment_elem.set(f'{{{NS["w"]}}}author', comment_author)
                    comment_elem.set(f'{{{NS["w"]}}}date', self.operation_timestamp)
                    # Use self.initials for all comments with author prefix matching self.author
                    # This includes: AI-<category> (all share same initials)
                    if comment_author.startswith(self.author):
                        comment_initials = self.initials
                    else:
                        comment_initials = comment_author[:2] if len(comment_author) >= 2 else comment_author
                    comment_elem.set(f'{{{NS["w"]}}}initials', comment_initials)

                    # Add paragraph with formatted text (handle <sup>/<sub> markup)
                    p = etree.SubElement(comment_elem, f'{{{NS["w"]}}}p')
            
                    # Parse text for superscript/subscript markup
                    # sanitize_xml_string first to remove illegal control characters
                    sanitized_text = sanitize_xml_string(comment['text'])
                    segments = self._parse_formatted_text(sanitized_text)
            
                    for segment_text, vert_align in segments:
                        if not segment_text:
                            continue
                
                        r = etree.SubElement(p, f'{{{NS["w"]}}}r')
                
                        # Add run properties with vertAlign if needed
                        if vert_align:
                            rPr = etree.SubElement(r, f'{{{NS["w"]}}}rPr')
                            vert_elem = etree.SubElement(rPr, f'{{{NS["w"]}}}vertAlign')
                            vert_elem.set(f'{{{NS["w"]}}}val', vert_align)
                
                        t = etree.SubElement(r, f'{{{NS["w"]}}}t')
                        # Preserve whitespace (Word drops leading/trailing spaces without this)
                        t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
                        t.text = segment_text

                if self.verbose:
                    try:
                        total = len(comments_xml.findall(f'{{{NS["w"]}}}comment'))
                        print(f"[Comments] comments.xml count after merge: {total}")
                    except Exception:
                        pass
        
                # Save via OPC
                blob = etree.tostring(
                    comments_xml, xml_declaration=True, encoding='UTF-8'
                )

                try:
                    comments_part = self.doc.part.part_related_by(RT.COMMENTS)
                    comments_part._blob = blob
                except (KeyError, AttributeError):
                    # Create new Part
                    comments_part = Part(
                        PackURI('/word/comments.xml'),
                        COMMENTS_CONTENT_TYPE,
                        blob,
                        self.doc.part.package
                    )
                    self.doc.part.relate_to(comments_part, RT.COMMENTS)

                if self.verbose:
                    try:
                        print(f"[Comments] Saved {len(self.comments)} comment(s) to comments.xml")
                    except Exception:
                        pass
