"""
This mixin class implements the overall apply/save flow and omment insertion, manual-action handling, per-item processing.
Handling how a single edit item is processed end-to-end.
"""

from .common import (
    NS,
    EditItem,
    EditResult,
    COMMENTS_CONTENT_TYPE,
    extract_longest_segment,
    format_text_preview,
    build_numbering_variants,
    strip_numbering_by_mode,
    strip_table_row_number_only,
    strip_table_row_numbering,
    normalize_table_json,
    DEBUG_MARKER
)
from typing import List, Dict, Optional, Tuple

from docx.opc.packuri import PackURI
from docx.opc.part import Part
from lxml import etree

from utils import sanitize_xml_string


class CommentWorkflowMixin:
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

    def _process_item(self, item: EditItem) -> EditResult:
        """Process a single edit item"""
        anchor_para = None
        try:
            # Strip leading/trailing whitespace from search and replacement text
            # to prevent matching failures caused by whitespace in JSONL data
            violation_text = item.violation_text.strip()
            revised_text = item.revised_text.strip()
            
            # 1. Find anchor paragraph by ID
            anchor_para = self._find_para_node_by_id(item.uuid)
            if anchor_para is None:
                return EditResult(False, item,
                    f"Paragraph ID {item.uuid} not found (may be in header/footer or ID changed)")
            
            item_author = self._author_for_item(item)

            # Handle mixed body/table content (violation_text contains <table> tags)
            has_table_tag = '<table>' in violation_text or '</table>' in violation_text
            if has_table_tag:
                if item.fix_action in ('delete', 'replace'):
                    reason = "Mixed body/table content is invalid"
                    self._apply_fallback_comment(anchor_para, item, reason)
                    if self.verbose:
                        print(f"  [Mixed content] {reason}")
                    return EditResult(success=True, item=item, error_message=reason, warning=True)
                else:  # manual
                    longest = extract_longest_segment(violation_text)
                    if longest:
                        if self.verbose:
                            print(f"  [Mixed content] Extracted longest segment for comment: '{format_text_preview(longest)}'")
                        violation_text = longest

            # 2. Search for text from anchor paragraph using ORIGINAL text (before revisions)
            # Store match results to pass to apply methods (avoid double matching)
            # IMPORTANT: Search is restricted to uuid -> uuid_end range to prevent
            # accidental modifications to content in other text blocks
            target_para = None
            matched_runs_info = None
            matched_start = -1
            numbering_stripped = False
            
            # Strategy: Try single-paragraph search first, then cross-paragraph if needed
            is_cross_paragraph = False
            numbering_variants = build_numbering_variants(violation_text)
            
            # Try original text first (using revision-free view) - single paragraph
            for para in self._iter_paragraphs_in_range(anchor_para, item.uuid_end):
                runs_info_orig, _ = self._collect_runs_info_original(para)

                pos, matched_override = self._find_in_runs_with_normalization(
                    runs_info_orig, violation_text
                )
                if pos != -1:
                    target_para = para
                    matched_runs_info = runs_info_orig
                    matched_start = pos
                    if matched_override is not None:
                        violation_text = matched_override
                    break
            
            # Fallback 1: Try stripping auto-numbering if original match failed
            if target_para is None:
                for stripped_violation, strip_mode in numbering_variants:
                    for para in self._iter_paragraphs_in_range(anchor_para, item.uuid_end):
                        runs_info_orig, _ = self._collect_runs_info_original(para)

                        pos, matched_override = self._find_in_runs_with_normalization(
                            runs_info_orig, stripped_violation
                        )
                        if pos != -1:
                            target_para = para
                            matched_runs_info = runs_info_orig
                            matched_start = pos
                            numbering_stripped = True
                            violation_text = matched_override or stripped_violation

                            # Handle revised_text for replace operation
                            if item.fix_action == 'replace':
                                stripped_revised, revised_has_numbering = strip_numbering_by_mode(revised_text, strip_mode)
                                if revised_has_numbering:
                                    # Both have numbering: strip both
                                    revised_text = stripped_revised
                                # else: Only violation_text had numbering, keep revised_text as-is

                            break
                    if target_para is not None:
                        break
            
            # Fallback 2: Try cross-paragraph search (within uuid → uuid_end range)
            # This also handles table content with JSON format matching
            boundary_error = None
            if target_para is None:
                # Collect runs across all paragraphs in range
                cross_runs, cross_text, is_multi_para, boundary_error = self._collect_runs_info_across_paragraphs(
                    anchor_para, item.uuid_end
                )

                if boundary_error:
                    # Boundary error detected - will handle below
                    if self.verbose:
                        print(f"  [Boundary] {boundary_error}")
                elif is_multi_para:
                    # Only use cross-paragraph mode if there are actually multiple paragraphs
                    # Build search attempts: original, numbering-stripped variants, and newline-unescaped
                    search_attempts: List[Tuple[str, Optional[str]]] = [(violation_text, None)]
                    for stripped_violation, strip_mode in numbering_variants:
                        search_attempts.append((stripped_violation, strip_mode))
                    if '\\n' in violation_text:
                        search_attempts.append((violation_text.replace('\\n', '\n'), None))
                    if self._is_table_mode(cross_runs) and violation_text.startswith('["'):
                        stripped_row_only, was_row_stripped = strip_table_row_number_only(violation_text)
                        if was_row_stripped:
                            search_attempts.append((stripped_row_only, "table_row_number"))
                        stripped_table_text, was_table_stripped = strip_table_row_numbering(violation_text)
                        if was_table_stripped and stripped_table_text != stripped_row_only:
                            search_attempts.append((stripped_table_text, "table_row"))

                    for search_text, strip_mode in search_attempts:
                        if self._is_table_mode(cross_runs):
                            pos = cross_text.find(search_text)
                            matched_override = None
                            if pos == -1 and search_text:
                                normalized_table_text, norm_to_orig = self._normalize_table_text_for_search(cross_runs)
                                pos_norm = normalized_table_text.find(search_text)
                                if pos_norm != -1:
                                    norm_end = pos_norm + len(search_text) - 1
                                    if 0 <= pos_norm < len(norm_to_orig) and 0 <= norm_end < len(norm_to_orig):
                                        orig_start = norm_to_orig[pos_norm]
                                        orig_end = norm_to_orig[norm_end] + 1
                                        pos = orig_start
                                        matched_override = cross_text[orig_start:orig_end]
                        else:
                            pos, matched_override = self._find_in_runs_with_normalization(
                                cross_runs, search_text
                            )

                        if pos != -1:
                            # Found match across paragraphs
                            target_para = anchor_para  # Use anchor as reference
                            matched_runs_info = cross_runs
                            matched_start = pos
                            is_cross_paragraph = True
                            if matched_override is not None:
                                violation_text = matched_override
                            else:
                                violation_text = search_text

                            if strip_mode:
                                numbering_stripped = True
                                if item.fix_action == 'replace':
                                    if strip_mode == "table_row_number":
                                        stripped_revised, revised_has_numbering = strip_table_row_number_only(revised_text)
                                    elif strip_mode == "table_row":
                                        stripped_revised, revised_has_numbering = strip_table_row_numbering(revised_text)
                                    else:
                                        stripped_revised, revised_has_numbering = strip_numbering_by_mode(revised_text, strip_mode)
                                    if revised_has_numbering:
                                        revised_text = stripped_revised

                            if self.verbose:
                                print(f"  [Success] Found in cross-paragraph mode")
                            break
            
            # Fallback 3: Try table search if violation_text looks like JSON array
            if target_para is None and (violation_text.startswith('["') or violation_text.startswith('[["')):
                # Normalize to remove duplicate brackets at boundaries (LLM artifacts)
                violation_text = normalize_table_json(violation_text)
                if item.fix_action == 'replace':
                    revised_text = normalize_table_json(revised_text)
                
                # Find all tables in range
                tables_in_range = self._find_tables_in_range(anchor_para, item.uuid_end)
                
                # Try with original violation_text first, then with row numbering stripped
                search_attempts = [
                    (violation_text, None),  # Original text, not stripped
                ]

                stripped_row_only, was_row_stripped = strip_table_row_number_only(violation_text)
                if was_row_stripped:
                    search_attempts.append((stripped_row_only, "table_row_number"))

                stripped_table_text, was_table_stripped = strip_table_row_numbering(violation_text)
                if was_table_stripped and stripped_table_text != stripped_row_only:
                    search_attempts.append((stripped_table_text, "table_row"))
                
                for search_text, strip_mode in search_attempts:
                    for table_idx, table_elem in enumerate(tables_in_range):
                        # Get the first and last paragraph in this table
                        table_paras = list(table_elem.iter(f'{{{NS["w"]}}}p'))
                        if not table_paras:
                            continue
                        
                        first_table_para = table_paras[0]
                        last_table_para = table_paras[-1]
                        
                        # Get their paraIds
                        first_para_id = first_table_para.get('{http://schemas.microsoft.com/office/word/2010/wordml}paraId')
                        last_para_id = last_table_para.get('{http://schemas.microsoft.com/office/word/2010/wordml}paraId')
                        
                        if not first_para_id or not last_para_id:
                            continue
                        
                        # Collect table content in JSON format
                        try:
                            table_runs, table_text, _, _ = self._collect_runs_info_in_table(
                                first_table_para, last_para_id, table_elem
                            )
                            
                            # Search for search_text in table content
                            pos = table_text.find(search_text)
                            matched_text_override = None
                            if pos == -1 and search_text:
                                # Fallback: normalize table text by stripping per-paragraph whitespace
                                # to match parse_document.py behavior, and map back to original indices.
                                normalized_table_text, norm_to_orig = self._normalize_table_text_for_search(table_runs)
                                pos_norm = normalized_table_text.find(search_text)
                                if pos_norm != -1:
                                    norm_end = pos_norm + len(search_text) - 1
                                    if 0 <= pos_norm < len(norm_to_orig) and 0 <= norm_end < len(norm_to_orig):
                                        orig_start = norm_to_orig[pos_norm]
                                        orig_end = norm_to_orig[norm_end] + 1
                                        pos = orig_start
                                        matched_text_override = table_text[orig_start:orig_end]
                            
                            # Debug logging for specific content
                            if pos == -1 and DEBUG_MARKER and (DEBUG_MARKER in table_text or DEBUG_MARKER in search_text):
                                print(f"\n  [DEBUG] Table matching failed for row containing '{DEBUG_MARKER}':")
                                
                                # Extract only the row containing DEBUG_MARKER from table_text
                                def extract_matching_row(text, marker):
                                    """Extract the row containing the marker from table JSON format."""
                                    # Table format: ["cell1", "cell2"], ["cell3", "cell4"]
                                    # Split by row boundaries '], ['
                                    if not text.startswith('["'):
                                        return None
                                    
                                    # Remove outer brackets and split by row separator
                                    rows_text = text[2:-2] if text.endswith('"]') else text[2:]
                                    rows = rows_text.split('"], ["')
                                    
                                    for row in rows:
                                        if marker in row:
                                            return '["' + row + '"]'
                                    return None
                                
                                table_row = extract_matching_row(table_text, DEBUG_MARKER)
                                search_row = extract_matching_row(search_text, DEBUG_MARKER)
                                
                                if table_row:
                                    print(f"  [DEBUG] Table row content:")
                                    print(f"    {table_row}")
                                else:
                                    print(f"  [DEBUG] Table content (marker '{DEBUG_MARKER}' not found in individual row):")
                                    print(f"    {table_text[:300]}...")
                                
                                if search_row:
                                    print(f"  [DEBUG] Searching for:")
                                    print(f"    {search_row}")
                                else:
                                    print(f"  [DEBUG] Search content (marker '{DEBUG_MARKER}' not found in individual row):")
                                    print(f"    {search_text[:300]}...")
                                
                                # Show character-level diff if both rows found
                                if table_row and search_row:
                                    min_len = min(len(table_row), len(search_row))
                                    for i in range(min_len):
                                        if table_row[i] != search_row[i]:
                                            print(f"  [DEBUG] First difference at position {i}:")
                                            print(f"    Table: ...{repr(table_row[max(0,i-10):i+30])}...")
                                            print(f"    Search: ...{repr(search_row[max(0,i-10):i+30])}...")
                                            break
                                    else:
                                        # No difference found in common length, check length difference
                                        if len(table_row) != len(search_row):
                                            print(f"  [DEBUG] Length mismatch: table={len(table_row)}, search={len(search_row)}")
                            
                            if pos != -1:
                                # Found match in this table!
                                target_para = first_table_para  # Use first para as anchor
                                matched_runs_info = table_runs
                                matched_start = pos
                                is_cross_paragraph = True  # Table mode is always cross-paragraph
                                
                                # Update violation_text to the matched version(stripped or not)
                                violation_text = matched_text_override or search_text
                                
                                # For replace operations, also strip row numbering from revised_text
                                if strip_mode and item.fix_action == 'replace':
                                    if strip_mode == "table_row_number":
                                        stripped_revised, revised_was_stripped = strip_table_row_number_only(revised_text)
                                    else:
                                        stripped_revised, revised_was_stripped = strip_table_row_numbering(revised_text)
                                    if revised_was_stripped:
                                        revised_text = stripped_revised
                                
                                if self.verbose:
                                    if strip_mode:
                                        print(f"  [Success] Found in table after stripping row numbering")
                                    else:
                                        print(f"  [Success] Found in table (JSON format)")
                                break

                        except (ValueError, KeyError, IndexError, AttributeError) as e:
                            # If table processing fails, continue to next table
                            if self.verbose:
                                print(f"  [Warning] Skipping table: {e}")
                            continue
                    
                    # If found, break outer loop
                    if target_para is not None:
                        break
            
            # Fallback 2.5: Try non-JSON table search (raw text mode)
            # For plain text violation_text that may be in table cells with multiple paragraphs
            if target_para is None and not violation_text.startswith('["') and not violation_text.startswith('[["'):
                # Find all tables in range
                tables_in_range = self._find_tables_in_range(anchor_para, item.uuid_end)
                
                for table_elem in tables_in_range:
                    # Search for raw text in each cell independently
                    result = self._search_in_table_cell_raw(
                        table_elem, violation_text, anchor_para, item.uuid_end
                    )
                    
                    if result:
                        target_para, matched_runs_info, matched_start, matched_text, strip_mode = result
                        # Update violation_text with the actual matched text (handles fallback normalization)
                        violation_text = matched_text
                        # Cell content is always treated as single-paragraph for now
                        is_cross_paragraph = False
                        if strip_mode:
                            numbering_stripped = True
                            if item.fix_action == 'replace':
                                stripped_revised, revised_has_numbering = strip_numbering_by_mode(revised_text, strip_mode)
                                if revised_has_numbering:
                                    revised_text = stripped_revised
                        
                        if self.verbose:
                            print(f"  [Success] Found in table cell (plain text mode)")
                        break
            
            if target_para is None:
                # Check if we have a boundary error from table/row crossing
                if boundary_error:
                    # Special handling for boundary_crossed: try searching in tables first
                    if boundary_error == 'boundary_crossed':
                        # Find all tables in range
                        tables_in_range = self._find_tables_in_range(anchor_para, item.uuid_end)
                        
                        # Try searching raw text in each table cell
                        for table_elem in tables_in_range:
                            # Iterate all cells in this table
                            for tc in table_elem.iter(f'{{{NS["w"]}}}tc'):
                                # Collect all paragraphs in this cell
                                cell_paras = tc.findall(f'{{{NS["w"]}}}p')
                                if not cell_paras:
                                    continue
                                
                                # Build cell's raw text content (with \n between paragraphs)
                                cell_text_parts = []
                                cell_para_runs_map = {}  # Map from para to runs_info
                                
                                for cell_para in cell_paras:
                                    para_runs, para_text = self._collect_runs_info_original(cell_para)
                                    cell_text_parts.append(para_text)
                                    cell_para_runs_map[id(cell_para)] = (para_runs, para_text)
                                
                                # Join with \n (actual newline, not JSON-escaped)
                                cell_combined_text = '\n'.join(cell_text_parts)
                                
                                # Normalize to match parse_document.py behavior (removes trailing whitespace)
                                cell_normalized = self._normalize_text_for_search(cell_combined_text)
                                
                                # Build search attempts for raw text in cell
                                search_attempts: List[Tuple[str, Optional[str]]] = [(violation_text, None)]
                                search_attempts.extend(build_numbering_variants(violation_text))
                                if '\\n' in violation_text:
                                    newline_text = violation_text.replace('\\n', '\n')
                                    search_attempts.append((newline_text, None))
                                    search_attempts.extend(build_numbering_variants(newline_text))

                                # De-duplicate while preserving order
                                seen = set()
                                deduped_attempts: List[Tuple[str, Optional[str]]] = []
                                for text, mode in search_attempts:
                                    key = (text, mode)
                                    if key in seen:
                                        continue
                                    seen.add(key)
                                    deduped_attempts.append((text, mode))

                                match_pos = -1
                                matched_search_text = violation_text
                                matched_strip_mode: Optional[str] = None
                                for search_text, strip_mode in deduped_attempts:
                                    match_pos = cell_normalized.find(search_text)
                                    if match_pos != -1:
                                        matched_search_text = search_text
                                        matched_strip_mode = strip_mode
                                        break

                                if match_pos != -1:
                                    # Found match in this cell!
                                    if self.verbose:
                                        print(f"  [Success] Found in table cell (raw text match)")
                                    if matched_strip_mode:
                                        numbering_stripped = True
                                        violation_text = matched_search_text
                                        if item.fix_action == 'replace':
                                            stripped_revised, revised_has_numbering = strip_numbering_by_mode(revised_text, matched_strip_mode)
                                            if revised_has_numbering:
                                                revised_text = stripped_revised
                                    else:
                                        violation_text = matched_search_text
                                    
                                    # Determine which paragraph(s) contain the match
                                    # For simplicity, if match is within first paragraph, use it
                                    # Otherwise, this is a cross-paragraph match within the cell
                                    current_offset = 0
                                    matched_para = None
                                    matched_para_runs = None
                                    matched_para_start = -1
                                    
                                    for cell_para in cell_paras:
                                        para_id_obj = id(cell_para)
                                        if para_id_obj not in cell_para_runs_map:
                                            continue
                                        
                                        para_runs, para_text = cell_para_runs_map[para_id_obj]
                                        para_len = len(para_text)
                                        
                                        # Check if match starts in this paragraph
                                        if current_offset <= match_pos < current_offset + para_len:
                                            matched_para = cell_para
                                            matched_para_runs = para_runs
                                            matched_para_start = match_pos - current_offset
                                            break
                                        
                                        current_offset += para_len + 1  # +1 for \n separator
                                    
                                    if matched_para is not None:
                                        # Use the matched paragraph
                                        target_para = matched_para
                                        matched_runs_info = matched_para_runs
                                        matched_start = matched_para_start
                                        is_cross_paragraph = False  # Single para in cell
                                        break
                                else:
                                    # Debug: log snippet from marker on non-JSON match failure
                                    marker_idx = cell_combined_text.find(DEBUG_MARKER)
                                    if DEBUG_MARKER and marker_idx != -1:
                                        snippet_len = len(DEBUG_MARKER) + 60
                                        snippet = cell_combined_text[marker_idx:marker_idx + snippet_len]
                                        print(f"  [DEBUG] Non-JSON cell content from marker: {repr(snippet)}")

                            # If found, break table loop
                            if target_para is not None:
                                break
                    
                    # If still not found after table search, try body text search
                    if target_para is None:
                        if self.verbose:
                            print(f"  [Boundary] Table search failed, trying body text...")
                        
                        # Collect ALL body paragraphs in range, grouped by continuity
                        # This handles: table→body, body→table, and interleaved scenarios
                        body_segments = []  # List of segments, each segment = [(para, runs, text), ...]
                        current_segment = []
                        
                        for para in self._iter_paragraphs_in_range(anchor_para, item.uuid_end):
                            if self._is_paragraph_in_table(para):
                                # Encountered table paragraph - end current body segment
                                if current_segment:
                                    body_segments.append(current_segment)
                                    current_segment = []
                                continue
                            
                            # Body paragraph
                            para_runs, para_text = self._collect_runs_info_original(para)
                            # Skip empty paragraphs to match parse_document.py behavior
                            if not para_text.strip():
                                continue
                            
                            current_segment.append((para, para_runs, para_text))
                        
                        # Don't forget the last segment
                        if current_segment:
                            body_segments.append(current_segment)
                        
                        # Search in each body segment (try both original and stripped numbering)
                        for segment_idx, body_paras_data in enumerate(body_segments):
                            # Build combined text with \n separator (like _collect_runs_info_in_body)
                            all_runs = []
                            pos = 0
                            
                            for i, (para, para_runs, para_text) in enumerate(body_paras_data):
                                # Add paragraph runs with updated positions
                                for run in para_runs:
                                    run_copy = dict(run)
                                    run_copy['para_elem'] = para
                                    run_copy['start'] = run['start'] + pos
                                    run_copy['end'] = run['end'] + pos
                                    all_runs.append(run_copy)
                                
                                pos += len(para_text)
                                
                                # Add paragraph boundary (except after last)
                                if i < len(body_paras_data) - 1:
                                    all_runs.append({
                                        'text': '\n',
                                        'start': pos,
                                        'end': pos + 1,
                                        'para_elem': para,
                                        'is_para_boundary': True
                                    })
                                    pos += 1
                            
                            # Try multiple search patterns (original and stripped numbering)
                            # Use build_numbering_variants() to support multi-line numbering removal
                            search_attempts: List[Tuple[str, Optional[str]]] = [(violation_text, None)]
                            numbering_variants = build_numbering_variants(violation_text)
                            search_attempts.extend(numbering_variants)
                            
                            match_pos = -1
                            matched_text = violation_text
                            matched_override = None
                            matched_is_stripped = False
                            
                            for search_text, is_stripped in search_attempts:
                                match_pos, matched_override = self._find_in_runs_with_normalization(
                                    all_runs, search_text
                                )
                                if match_pos != -1:
                                    matched_text = matched_override or search_text
                                    matched_is_stripped = is_stripped
                                    break
                            
                            if match_pos != -1:
                                # Found match in this segment!
                                target_para = body_paras_data[0][0]  # Use first para as anchor
                                matched_runs_info = all_runs
                                matched_start = match_pos
                                is_cross_paragraph = len(body_paras_data) > 1
                                violation_text = matched_text  # Update violation_text to matched version
                                numbering_stripped = matched_is_stripped
                                
                                if self.verbose:
                                    if is_cross_paragraph:
                                        print(f"  [Success] Found in body segment {segment_idx + 1} (cross-paragraph)")
                                    else:
                                        print(f"  [Success] Found in body segment {segment_idx + 1}")
                                break  # Stop searching other segments
                    
                    # If still not found after segmented body text search, apply fallback
                    if target_para is None:
                        reason = ""
                        if boundary_error == 'boundary_crossed':
                            reason = "Violation text not found(C)"  # Table/body boundary crossed
                        else:
                            reason = f"Boundary error: {boundary_error}"

                        # Debug: log body paragraph content from marker on boundary failure
                        if DEBUG_MARKER:
                            try:
                                for para in self._iter_paragraphs_in_range(anchor_para, item.uuid_end):
                                    if self._is_paragraph_in_table(para):
                                        continue
                                    _, para_text = self._collect_runs_info_original(para)
                                    marker_idx = para_text.find(DEBUG_MARKER)
                                    if marker_idx != -1:
                                        snippet_len = len(DEBUG_MARKER) + 60
                                        snippet = para_text[marker_idx:marker_idx + snippet_len]
                                        print(f"  [DEBUG] Body search target: {repr(violation_text)}")
                                        break
                            except Exception:
                                # Debug logging should never break apply flow
                                pass

                        self._apply_fallback_comment(anchor_para, item, reason)
                        if self.verbose:
                            print(f"  [Boundary] {reason}")
                        return EditResult(
                            success=True,
                            item=item,
                            error_message=reason,
                            warning=True
                        )
                
                # Only proceed with fallback if target_para is still None
                # (boundary_error handling may have found a match)
                if target_para is None:
                    # Debug: log body paragraph content from marker on overall match failure
                    if DEBUG_MARKER:
                        try:
                            for para in self._iter_paragraphs_in_range(anchor_para, item.uuid_end):
                                if self._is_paragraph_in_table(para):
                                    continue
                                _, para_text = self._collect_runs_info_original(para)
                                marker_idx = para_text.find(DEBUG_MARKER)
                                if marker_idx != -1:
                                    snippet_len = len(DEBUG_MARKER) + 60
                                    snippet = para_text[marker_idx:marker_idx + snippet_len]
                                    print(f"  [DEBUG] Body search target: {repr(violation_text)}")
                                    break
                        except Exception:
                            # Debug logging should never break apply flow
                            pass

                    # Calculate total segments for error message
                    # Variables are initialized in their respective code paths above
                    # If not set, default to empty lists
                    tables_count = len(tables_in_range) if 'tables_in_range' in dir() else 0
                    body_count = len(body_segments) if 'body_segments' in dir() else 0
                    total_segments = tables_count + body_count
                    
                    if total_segments == 1:
                        reason = "Violation text not found(S)"  # Single segment
                    else:
                        reason = f"Violation text not found(M)"  # Multiple segments
                    
                    # For manual fix_action, text not found is expected (not an error)
                    if item.fix_action == 'manual':
                        self._apply_error_comment(anchor_para, item)
                        return EditResult(
                            success=True,
                            item=item,
                            error_message=reason,
                            warning=True
                        )
                    else:
                        # For delete/replace, text not found is an error
                        self._apply_error_comment(anchor_para, item)
                        return EditResult(False, item, reason)
            
            # 3. Apply operation based on fix_action
            # Pass matched_runs_info and matched_start to avoid double matching
            success_status = None
            
            if item.fix_action == 'delete':
                # Check for cross-paragraph: only fallback if ACTUAL match spans multiple paragraphs
                if is_cross_paragraph:
                    # Get affected runs in the matched range
                    match_end = matched_start + len(violation_text)
                    affected = self._find_affected_runs(matched_runs_info, matched_start, match_end)

                    # Filter out boundary markers (paragraph and JSON boundaries)
                    real_runs = [r for r in affected
                                 if not r.get('is_para_boundary', False)
                                 and not r.get('is_json_boundary', False)
                                 and not r.get('is_json_escape', False)]

                    # Check if actual match spans multiple paragraphs
                    para_elems = set(r.get('para_elem') for r in real_runs if r.get('para_elem') is not None)

                    # Check if match spans multiple table rows
                    if self._check_cross_row_boundary(real_runs):
                        # Multi-row delete: use cell-by-cell deletion
                        if self.verbose:
                            print(f"  [Multi-row] Applying cell-by-cell deletion")
                        success_status = self._apply_delete_multi_cell(
                            real_runs, violation_text,
                            item.violation_reason, item_author, matched_start,
                            item
                        )
                    # Check if match spans multiple table cells (within same row)
                    elif self._check_cross_cell_boundary(real_runs):
                        # Multi-cell delete: use cell-by-cell deletion
                        if self.verbose:
                            print(f"  [Multi-cell] Applying cell-by-cell deletion")
                        success_status = self._apply_delete_multi_cell(
                            real_runs, violation_text,
                            item.violation_reason, item_author, matched_start,
                            item
                        )
                    elif len(para_elems) > 1:
                        # Actually spans multiple paragraphs
                        if self._is_table_mode(real_runs) or any(r.get('cell_elem') is not None for r in real_runs):
                            if self.verbose:
                                print(f"  [Cross-paragraph] delete spans {len(para_elems)} paragraphs, fallback to comment")
                            success_status = 'cross_paragraph_fallback'
                        else:
                            if self.verbose:
                                print(f"  [Cross-paragraph] delete spans {len(para_elems)} paragraphs, applying cross-paragraph delete")
                            success_status = self._apply_delete_cross_paragraph(
                                matched_runs_info, matched_start, violation_text,
                                item.violation_reason, item_author
                            )
                    else:
                        # All content is in single paragraph - safe to delete
                        if real_runs:
                            target_para = real_runs[0].get('para_elem', target_para)
                        success_status = self._apply_delete(
                            target_para, violation_text,
                            item.violation_reason,
                            matched_runs_info, matched_start,
                            item_author
                        )
                else:
                    success_status = self._apply_delete(
                        target_para, violation_text,
                        item.violation_reason,
                        matched_runs_info, matched_start,
                        item_author
                    )
            elif item.fix_action == 'replace':
                # Check for cross-paragraph: only fallback if ACTUAL match spans multiple paragraphs
                if is_cross_paragraph:
                    # Get affected runs in the matched range
                    match_end = matched_start + len(violation_text)
                    affected = self._find_affected_runs(matched_runs_info, matched_start, match_end)

                    # Filter out boundary markers (paragraph and JSON boundaries)
                    real_runs = [r for r in affected
                                 if not r.get('is_para_boundary', False)
                                 and not r.get('is_json_boundary', False)
                                 and not r.get('is_json_escape', False)]

                    # Check if actual match spans multiple paragraphs
                    para_elems = set(r.get('para_elem') for r in real_runs if r.get('para_elem') is not None)

                    # Check if match spans multiple table rows
                    is_cross_row = self._check_cross_row_boundary(real_runs)

                    # Detect JSON cell boundaries even if only one cell has text runs.
                    # This handles cases where empty cells (e.g., stripped row numbers)
                    # produce boundary markers but no runs with cell_elem.
                    has_cell_boundary = any(r.get('is_cell_boundary') for r in affected)
                    is_cross_cell = self._check_cross_cell_boundary(real_runs) or has_cell_boundary or is_cross_row

                    if is_cross_row and self.verbose:
                        print(f"  [Cross-row] replace spans multiple rows, trying cell-by-cell extraction...")
                    if self.verbose and has_cell_boundary and not self._check_cross_cell_boundary(real_runs):
                        print(f"  [Cross-cell] Boundary markers detected (empty cells), forcing cell extraction")

                    if success_status is None and is_cross_cell:
                        # Try to extract single-cell edit first
                        if self.verbose:
                            print(f"  [Cross-cell] Detected cross-cell match, trying cell-by-cell extraction...")
                        
                        single_cell = self._try_extract_single_cell_edit(
                            violation_text, revised_text, affected, matched_start
                        )
                        
                        if single_cell:
                            # Successfully extracted single-cell edit - all changes in one cell
                            success_status = self._apply_replace_in_cell_paragraphs(
                                single_cell['cell_violation'],
                                single_cell['cell_revised'],
                                single_cell['cell_runs'],
                                item.violation_reason,
                                item_author,
                                skip_comment=False
                            )
                            if self.verbose and success_status == 'success':
                                print(f"  [Single-cell] All changes in one cell, applied track change")
                            elif success_status not in ('success', 'conflict'):
                                success_status = 'cross_cell_fallback'
                        else:
                            # Try multi-cell extraction - changes distributed across cells
                            multi_cells = self._try_extract_multi_cell_edits(
                                violation_text, revised_text, affected, matched_start
                            )
                            
                            if multi_cells:
                                # Successfully extracted multi-cell edits - each change within its own cell
                                if self.verbose:
                                    print(f"  [Multi-cell] Found {len(multi_cells)} cells with changes, applying track changes...")
                                
                                success_count = 0
                                failed_cells = []  # List of (cell_edit, cell_para, error_reason)
                                first_success_para = None
                                
                                for cell_edit in multi_cells:
                                    cell_status = self._apply_replace_in_cell_paragraphs(
                                        cell_edit['cell_violation'],
                                        cell_edit['cell_revised'],
                                        cell_edit['cell_runs'],
                                        item.violation_reason,
                                        item_author,
                                        skip_comment=True
                                    )

                                    if cell_status != 'success':
                                        # Get first paragraph for error tracking
                                        first_para = None
                                        for run in cell_edit['cell_runs']:
                                            para = run.get('para_elem')
                                            if para is not None:
                                                first_para = para
                                                break
                                        failed_cells.append((cell_edit, first_para, f"Apply failed: {cell_status}"))
                                        continue  # Continue processing next cell

                                    # Success - track for overall comment
                                    success_count += 1
                                    if first_success_para is None:
                                        # Get first paragraph for comment placement
                                        for run in cell_edit['cell_runs']:
                                            para = run.get('para_elem')
                                            if para is not None:
                                                first_success_para = para
                                                break
                                
                                # Handle results based on success/failure counts
                                if success_count > 0:
                                    # At least one cell succeeded - add overall comment
                                    if first_success_para is not None:
                                        comment_id = self.next_comment_id
                                        self.next_comment_id += 1
                                        
                                        # Insert commentReference at end of first successful paragraph
                                        ref_xml = f'''<w:r xmlns:w="{NS['w']}">
                                            <w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>
                                            <w:commentReference w:id="{comment_id}"/>
                                        </w:r>'''
                                        first_success_para.append(etree.fromstring(ref_xml))
                                        
                                        # Record comment with -R suffix author
                                        self.comments.append({
                                            'id': comment_id,
                                            'text': item.violation_reason,
                                            'author': f"{item_author}-R"
                                        })
                                    
                                    # Add individual comments for failed cells
                                    for cell_edit, cell_para, error_reason in failed_cells:
                                        if cell_para is not None:
                                            self._apply_cell_fallback_comment(
                                                cell_para,
                                                cell_edit['cell_violation'],
                                                cell_edit['cell_revised'],
                                                error_reason,
                                                item
                                            )
                                    
                                    success_status = 'success'
                                    if self.verbose:
                                        if failed_cells:
                                            print(f"  [Multi-cell] Partial success: {success_count}/{len(multi_cells)} cells processed, {len(failed_cells)} failed")
                                        else:
                                            print(f"  [Multi-cell] Successfully applied track changes to all {len(multi_cells)} cells")
                                else:
                                    # All cells failed - fallback to overall comment
                                    success_status = 'cross_cell_fallback'
                            else:
                                # Changes cross cell boundaries - fallback to comment
                                if self.verbose:
                                    print(f"  [Cross-cell] Changes cross cell boundaries, fallback to comment")
                                success_status = 'cross_cell_fallback'
                        if success_status is None and is_cross_row:
                            # Cell extraction failed for cross-row content
                            if self.verbose:
                                print(f"  [Cross-row] Cell extraction failed, fallback to comment")
                            success_status = 'cross_row_fallback'
                    elif len(para_elems) > 1:
                        # Actually spans multiple paragraphs
                        if self._is_table_mode(real_runs) or any(r.get('cell_elem') is not None for r in real_runs):
                            if self.verbose:
                                print(f"  [Cross-paragraph] replace spans {len(para_elems)} paragraphs, fallback to comment")
                            # Debug: log snippet from marker on cross-paragraph fallback
                            try:
                                if DEBUG_MARKER:
                                    combined_text = ''.join(r.get('text', '') for r in matched_runs_info)
                                    marker_idx = combined_text.find(DEBUG_MARKER)
                                    if marker_idx != -1:
                                        snippet_len = len(DEBUG_MARKER) + 60
                                        snippet = combined_text[marker_idx:marker_idx + snippet_len]
                                        print(f"  [DEBUG] Cross-paragraph content from marker: {repr(snippet)}")
                            except Exception:
                                # Debug logging should never break apply flow
                                pass
                            success_status = 'cross_paragraph_fallback'
                        else:
                            if self.verbose:
                                print(f"  [Cross-paragraph] replace spans {len(para_elems)} paragraphs, applying cross-paragraph replace")
                            success_status = self._apply_replace_cross_paragraph(
                                matched_runs_info, matched_start, violation_text,
                                revised_text, item.violation_reason, item_author
                            )
                    else:
                        # All content is in single paragraph - safe to replace
                        if real_runs:
                            target_para = real_runs[0].get('para_elem', target_para)
                        success_status = self._apply_replace(
                            target_para, violation_text, revised_text,
                            item.violation_reason,
                            matched_runs_info, matched_start,
                            item_author
                        )
                else:
                    success_status = self._apply_replace(
                        target_para, violation_text, revised_text,
                        item.violation_reason,
                        matched_runs_info, matched_start,
                        item_author
                    )
            elif item.fix_action == 'manual':
                success_status = self._apply_manual(
                    target_para, violation_text,
                    item.violation_reason, revised_text,
                    matched_runs_info, matched_start,
                    item_author,
                    is_cross_paragraph  # Pass cross-paragraph flag
                )
            else:
                # Insert error comment for unknown action type
                self._apply_error_comment(anchor_para, item)
                return EditResult(False, item, f"Unknown action type: {item.fix_action}")
            
            # Handle results
            if success_status == 'success':
                if numbering_stripped and self.verbose:
                    print(f"  [Success] Matched after stripping auto-numbering")
                return EditResult(True, item)
            elif success_status == 'conflict':
                # Text overlaps with previous rule modification - mark as warning
                reason = "Multiple changes overlap."
                self._apply_fallback_comment(target_para, item, reason)
                if self.verbose:
                    print(f"  [Conflict] {reason}")
                return EditResult(
                    success=True,
                    item=item,
                    error_message=reason,
                    warning=True
                )
            elif success_status == 'cross_paragraph_fallback':
                # Cross-paragraph delete/replace not supported - fallback to manual comment
                reason = (
                    "Cross-paragraph delete/replace not supported"
                )
                # Apply manual comment instead
                manual_status = self._apply_manual(
                    target_para, violation_text,
                    item.violation_reason, revised_text,
                    matched_runs_info, matched_start,
                    item_author,
                    is_cross_paragraph,
                    fallback_reason=reason
                )
                if manual_status == 'success':
                    if self.verbose:
                        print(f"  [Cross-paragraph] Applied comment instead")
                    return EditResult(
                        success=True,
                        item=item,
                        error_message=reason,
                        warning=True
                    )
                else:
                    # Manual comment also failed - use fallback comment
                    self._apply_fallback_comment(target_para, item, reason)
                    return EditResult(
                        success=True,
                        item=item,
                        error_message=f"{reason} (comment also failed)",
                        warning=True
                    )
            elif success_status == 'cross_row_fallback':
                # Cross-row delete/replace not supported - Word doesn't support cross-row comments
                reason = "Cross-row edit not supported"
                # Use fallback comment (non-selected) since cross-row comments are not supported
                self._apply_fallback_comment(target_para, item, reason)
                if self.verbose:
                    print(f"  [Cross-row] Applied fallback comment")
                return EditResult(
                    success=True,
                    item=item,
                    error_message=reason,
                    warning=True
                )
            elif success_status == 'cross_cell_fallback':
                # Cross-cell delete/replace not supported - fallback to manual comment
                reason = "Cross-cell edit not supported"
                # Apply manual comment instead (same row comment is supported)
                manual_status = self._apply_manual(
                    target_para, violation_text,
                    item.violation_reason, revised_text,
                    matched_runs_info, matched_start,
                    item_author,
                    is_cross_paragraph,
                    fallback_reason=reason
                )
                if manual_status == 'success':
                    if self.verbose:
                        print(f"  [Cross-cell] Applied comment instead")
                    return EditResult(
                        success=True,
                        item=item,
                        error_message=reason,
                        warning=True
                    )
                else:
                    # Manual comment also failed - use fallback comment
                    self._apply_fallback_comment(target_para, item, reason)
                    return EditResult(
                        success=True,
                        item=item,
                        error_message=f"{reason} (comment also failed)",
                        warning=True
                    )
            elif success_status == 'equation_fallback':
                # Equation-only content cannot be edited - fallback to manual comment
                reason = "Equation cannot be edited"
                manual_status = self._apply_manual(
                    target_para, violation_text,
                    item.violation_reason, revised_text,
                    matched_runs_info, matched_start,
                    item_author,
                    is_cross_paragraph,
                    fallback_reason=reason
                )
                if manual_status == 'success':
                    if self.verbose:
                        print(f"  [Equation-only] Applied comment instead")
                    return EditResult(
                        success=True,
                        item=item,
                        error_message=reason,
                        warning=True
                    )
                else:
                    # Manual comment also failed - use fallback comment
                    self._apply_fallback_comment(target_para, item, reason)
                    return EditResult(
                        success=True,
                        item=item,
                        error_message=f"{reason} (comment also failed)",
                        warning=True
                    )
            elif success_status == 'fallback':
                # Fallback to comment annotation - mark as warning for all fix_actions
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
                    warning=True
                )
            else:
                # Unexpected return value or old boolean False
                self._apply_error_comment(anchor_para, item)
                return EditResult(False, item, "Operation failed")
                
        except Exception as e:
            # Insert error comment on exception if anchor paragraph exists
            if anchor_para is not None:
                self._apply_error_comment(anchor_para, item)
            return EditResult(False, item, str(e))
