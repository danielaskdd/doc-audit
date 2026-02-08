"""
This mixin class implements table-specific edit handling,.
Handling how table cells are edited, including extraction and application for single-cell and multi-cell scenarios.
"""

from typing import List, Dict, Tuple, Optional
from .common import (
    NS,
    EditItem,
    format_text_preview,
)
from lxml import etree


class TableEditMixin:
    def _check_cross_row_boundary(self, affected_runs: List[Dict]) -> bool:
        """
        Check if affected runs span multiple table rows.

        Args:
            affected_runs: List of run info dicts from _find_affected_runs

        Returns:
            True if runs span multiple rows, False otherwise
        """
        rows = set()
        for info in affected_runs:
            row = info.get('row_elem')
            if row is not None:
                rows.add(id(row))
        return len(rows) > 1

    def _group_runs_by_cell(self, runs: List[Dict]) -> Dict[Tuple, List[Dict]]:
        """
        Group runs by (para_elem, cell_elem) pairs.
        
        This allows processing runs cell-by-cell for multi-cell operations
        like delete/replace across table cells.
        
        IMPORTANT: This method assumes runs are already in document order
        (as provided by _collect_runs_info_* methods). The grouping preserves
        this order within each cell group.
        
        Args:
            runs: List of run info dicts (must be in document order)
        
        Returns:
            Dict mapping (para_elem, cell_elem) -> list of runs in that cell
        """
        groups = {}
        for run in runs:
            para = run.get('para_elem')
            cell = run.get('cell_elem')
            key = (id(para) if para is not None else None, 
                   id(cell) if cell is not None else None)
            if key not in groups:
                groups[key] = []
            groups[key].append(run)
        return groups

    def _group_runs_by_paragraph(self, runs: List[Dict]) -> List[Dict]:
        """
        Group runs by paragraph in document order.

        Args:
            runs: List of run info dicts (must be in document order)

        Returns:
            List of dicts: [{'para_elem': para, 'runs': [...]}]
        """
        groups = {}
        order = []
        for run in runs:
            para = run.get('para_elem')
            if para is None:
                continue
            key = id(para)
            if key not in groups:
                groups[key] = {'para_elem': para, 'runs': []}
                order.append(key)
            groups[key]['runs'].append(run)
        return [groups[k] for k in order]

    def _build_para_segments_from_groups(
        self,
        para_groups: List[Dict],
        match_start: int,
        match_end: int
    ) -> List[Dict]:
        """
        Build paragraph segments from grouped runs.

        Args:
            para_groups: Output of _group_runs_by_paragraph
            match_start: Absolute start offset of the match in combined_text
            match_end: Absolute end offset of the match in combined_text

        Returns:
            List of segment dicts for _apply_diff_per_paragraph
        """
        para_segments = []
        for group in para_groups:
            runs = group['runs']
            if not runs:
                continue
            para_start = min(r['start'] for r in runs)
            para_end = max(r['end'] for r in runs)
            overlap_start = max(match_start, para_start)
            overlap_end = min(match_end, para_end)
            if overlap_start < overlap_end:
                para_segments.append({
                    'para_elem': group['para_elem'],
                    'seg_start': overlap_start - match_start,
                    'seg_end': overlap_end - match_start,
                    'match_pos_in_para': overlap_start - para_start,
                })
        return para_segments

    def _init_para_order(self):
        """Initialize paragraph order cache for document-order comparisons."""
        if self.body_elem is None:
            return
        self._para_list = list(self._xpath(self.body_elem, './/w:p'))
        self._para_order = {id(p): i for i, p in enumerate(self._para_list)}
        self._para_id_list = [
            p.get('{http://schemas.microsoft.com/office/word/2010/wordml}paraId')
            for p in self._para_list
        ]

    def _get_run_doc_key(self, run_elem, para_elem,
                         para_order: Optional[Dict[int, int]] = None) -> Optional[Tuple[int, int]]:
        """
        Get a stable document-order key for a run element.

        Returns (para_index, run_index_in_para) or None if unavailable.
        This is used to detect inverted ordering when runs are reused
        across rows (e.g., vMerge continue cells).
        """
        if run_elem is None or para_elem is None:
            return None
        if para_order is None:
            para_order = self._para_order
        if not para_order:
            self._init_para_order()
            para_order = self._para_order
        para_idx = para_order.get(id(para_elem))
        if para_idx is None:
            # Refresh cache in case body_elem was replaced in tests
            self._init_para_order()
            para_order = self._para_order
            para_idx = para_order.get(id(para_elem))
        if para_idx is None:
            return None
        try:
            run_list = list(para_elem.iter(f'{{{NS["w"]}}}r'))
            run_idx = run_list.index(run_elem)
        except ValueError:
            run_idx = 0
        return (para_idx, run_idx)

    def _is_multi_row_json(self, text: str) -> bool:
        """
        Check if text is multi-row JSON format.
        
        Multi-row JSON format: '["cell1", "cell2"], ["cell3", "cell4"]'
        Single-row JSON format: '["cell1", "cell2"]'
        
        Args:
            text: Text to check
        
        Returns:
            True if text contains row boundary marker '"], ["'
        """
        return text.startswith('["') and '"], ["' in text

    def _apply_delete_multi_cell(self, real_runs: List[Dict], violation_text: str,
                                  violation_reason: str, author: str, match_start: int,
                                  item: EditItem) -> str:
        """
        Apply delete operation across multiple cells/rows.
        
        Strategy: Delete content in each cell independently by wrapping runs
        with <w:del> tags, properly handling partial run deletion when the
        match doesn't align to run boundaries.
        
        Args:
            real_runs: List of real runs (boundaries filtered out)
            violation_text: Text being deleted (may be JSON format)
            violation_reason: Reason for deletion
            author: Track change author
            match_start: Absolute position where violation_text starts
            item: Original EditItem (for fallback comment)
        
        Returns:
            'success' if at least one cell deleted, 'fallback' if all failed
        """
        if not real_runs:
            return 'fallback'

        match_end = match_start + len(violation_text)

        # Group runs by cell
        cell_groups = self._group_runs_by_cell(real_runs)

        # Track success/failure for each cell
        success_count = 0
        failed_cells = []  # List of (para_elem, cell_violation, error_reason)
        first_success_para = None

        # Process each cell independently
        for (para_id, cell_id), cell_runs in cell_groups.items():
            if not cell_runs:
                continue

            # De-duplicate runs by their underlying elem reference
            # This prevents duplicates when vMerge='continue' cells copy runs from restart cells
            seen_elems = set()
            unique_cell_runs = []
            for run in cell_runs:
                elem_id = id(run.get('elem'))
                if elem_id not in seen_elems:
                    seen_elems.add(elem_id)
                    unique_cell_runs.append(run)
            cell_runs = unique_cell_runs

            if not cell_runs:
                continue

            # Get the paragraph element from first run
            para_elem = cell_runs[0].get('para_elem')
            if para_elem is None:
                failed_cells.append((None, "", "No paragraph found"))
                continue

            # Find cell boundaries in the match range
            cell_start = min(r['start'] for r in cell_runs)
            cell_end = max(r['end'] for r in cell_runs)

            # Calculate intersection with match range
            del_start_in_cell = max(cell_start, match_start)
            del_end_in_cell = min(cell_end, match_end)

            if del_start_in_cell >= del_end_in_cell:
                # No overlap with this cell
                continue

            # Identify affected runs in this cell
            affected_cell_runs = [r for r in cell_runs
                                  if r['end'] > del_start_in_cell and r['start'] < del_end_in_cell]

            if not affected_cell_runs:
                failed_cells.append((para_elem, "", "No affected runs found"))
                continue

            # Extract cell violation text for error reporting
            cell_violation_parts = []
            first_run = affected_cell_runs[0]
            last_run = affected_cell_runs[-1]
            before_offset = self._translate_escaped_offset(
                first_run, max(0, del_start_in_cell - first_run['start']))
            after_offset = self._translate_escaped_offset(
                last_run, max(0, del_end_in_cell - last_run['start']))

            for run in affected_cell_runs:
                if run.get('is_drawing'):
                    cell_violation_parts.append(run['text'])
                    continue
                orig_text = self._get_run_original_text(run)
                if run is first_run and run is last_run:
                    cell_violation_parts.append(orig_text[before_offset:after_offset])
                elif run is first_run:
                    cell_violation_parts.append(orig_text[before_offset:])
                elif run is last_run:
                    cell_violation_parts.append(orig_text[:after_offset])
                else:
                    cell_violation_parts.append(orig_text)
            cell_violation = ''.join(cell_violation_parts)

            # Check if any affected run overlaps with existing revisions
            if self._check_overlap_with_revisions(affected_cell_runs):
                failed_cells.append((para_elem, cell_violation, "Overlaps with existing revision"))
                continue

            if not cell_violation:
                failed_cells.append((para_elem, "", "No text to delete"))
                continue

            para_groups = self._group_runs_by_paragraph(affected_cell_runs)
            if not para_groups:
                failed_cells.append((para_elem, cell_violation, "No paragraph groups found"))
                continue

            prepared = self._prepare_deletion_items(
                para_groups, del_start_in_cell, del_end_in_cell
            )
            if not prepared:
                failed_cells.append((para_elem, cell_violation, "No paragraphs prepared for deletion"))
                continue

            shared_change_id = self._get_next_change_id()
            success_paragraphs = self._delete_paragraphs_in_unit(
                prepared, shared_change_id, author, comment_id=None
            )

            if success_paragraphs == 0:
                failed_cells.append((para_elem, cell_violation, "Failed to replace runs"))
                continue

            success_count += 1
            if first_success_para is None:
                first_success_para = para_elem
        
        # Handle results based on success/failure counts
        if success_count > 0:
            # At least one cell deleted - add overall comment
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
                    'text': violation_reason,
                    'author': f"{author}-R"
                })
            
            # Add individual comments for failed cells
            for para_elem, cell_violation, error_reason in failed_cells:
                if para_elem is not None:
                    self._apply_cell_fallback_comment(
                        para_elem,
                        cell_violation,
                        "",  # No revised text for delete operation
                        error_reason,
                        item
                    )
            
            if self.verbose:
                if failed_cells:
                    print(f"  [Multi-cell delete] Partial success: {success_count}/{len(cell_groups)} cells processed, {len(failed_cells)} failed")
                else:
                    print(f"  [Multi-cell delete] Successfully deleted from all {len(cell_groups)} cells")
            elif failed_cells:
                # Summary logging for partial failures in non-verbose mode
                print(f"  [Warning] {len(failed_cells)} cell(s) failed during delete operation")
            
            return 'success'
        else:
            # All cells failed - fallback to overall comment
            return 'fallback'

    def _try_extract_multi_cell_edits(self, violation_text: str, revised_text: str,
                                        affected_runs: List[Dict], match_start: int) -> Optional[List[Dict]]:
        """
        Try to extract multi-cell edits when text spans multiple cells.
        
        This method analyzes the diff to determine if changes are distributed across
        multiple cells, with each change staying within its own cell boundary.
        
        Args:
            violation_text: Original violation text (may span multiple cells)
            revised_text: Revised text (may span multiple cells)
            affected_runs: List of run info dicts that would be affected
            match_start: Absolute position where violation_text starts in document
        
        Returns:
            List of dicts, each with keys: 'cell_violation', 'cell_revised', 'cell_elem', 'cell_runs'
            Or None if any change crosses cell boundaries
        """
        # Calculate diff to find changed positions
        diff_ops = self._calculate_diff(violation_text, revised_text)
        
        # Track positions of all changes (delete/insert operations).
        # Keep order aligned with diff_ops so we can map insertions reliably
        # even when they land on a cell's right boundary.
        change_positions = []
        current_pos = 0
        
        for op, text in diff_ops:
            if op == 'delete':
                # Record the position range of deleted text
                change_positions.append((current_pos, current_pos + len(text)))
                current_pos += len(text)
            elif op == 'insert':
                # Record the insertion point (zero-length range)
                change_positions.append((current_pos, current_pos))
                # insert doesn't consume original text position
            elif op == 'equal':
                current_pos += len(text)
        
        if not change_positions:
            # No changes detected
            return None
        
        # Find which cell each change affects
        # Map: cell_id -> list of change indices
        cell_to_changes = {}
        
        base_offset = match_start
        
        for change_idx, (change_start_rel, change_end_rel) in enumerate(change_positions):
            # Translate relative positions to absolute positions
            change_start = base_offset + change_start_rel
            change_end = base_offset + change_end_rel
            
            is_insert = (change_start == change_end)
            
            # Find which cell contains this change
            change_cell = None
            for run in affected_runs:
                if (run.get('is_json_boundary') or 
                    run.get('is_json_escape') or 
                    run.get('is_para_boundary')):
                    continue
                
                # Check if this run is affected by the change
                is_affected = False
                if is_insert:
                    if run['start'] <= change_start <= run['end']:
                        is_affected = True
                else:
                    if run['end'] > change_start and run['start'] < change_end:
                        is_affected = True
                
                if is_affected:
                    cell = run.get('cell_elem')
                    if cell is not None:
                        if change_cell is None:
                            change_cell = cell
                        elif id(change_cell) != id(cell):
                            # Change spans multiple cells
                            return None
            
            if change_cell is not None:
                cell_id = id(change_cell)
                if cell_id not in cell_to_changes:
                    cell_to_changes[cell_id] = []
                cell_to_changes[cell_id].append(change_idx)
            else:
                # Change landed on JSON boundary marker (e.g., ", " or "], [")
                # Cannot be correctly mapped to a cell - this indicates a cross-cell
                # operation like merge/split that requires fallback to comment
                return None
        
        if not cell_to_changes:
            return None
        
        # Group affected runs by cell
        cell_runs_map = {}
        for run in affected_runs:
            if (run.get('is_json_boundary') or 
                run.get('is_json_escape') or 
                run.get('is_para_boundary')):
                continue
            
            cell = run.get('cell_elem')
            if cell is not None:
                cell_id = id(cell)
                if cell_id not in cell_runs_map:
                    cell_runs_map[cell_id] = {'cell': cell, 'runs': []}
                cell_runs_map[cell_id]['runs'].append(run)
        
        # Extract cell-specific edits
        result_edits = []
        
        for cell_id, change_indices in cell_to_changes.items():
            if cell_id not in cell_runs_map:
                continue
            
            cell_info = cell_runs_map[cell_id]
            cell_runs = cell_info['runs']
            
            # Find cell boundaries in violation_text
            cell_start = min(r['start'] for r in cell_runs)
            cell_end = max(r['end'] for r in cell_runs)
            
            relative_start = cell_start - match_start
            relative_end = cell_end - match_start
            
            # Extract cell_violation
            if self._is_table_mode(affected_runs):
                cell_violation_escaped = violation_text[relative_start:relative_end]
                cell_violation = self._decode_json_escaped(cell_violation_escaped)
            else:
                cell_violation = violation_text[relative_start:relative_end]

            # Build cell_revised by applying diff operations within cell range.
            # Use change_indices to decide insert ownership. This prevents
            # dropping inserts at exact cell right boundary (e.g., "A" -> "A!").
            violation_pos = 0
            revised_accumulator = ''
            has_changes_in_cell = False  # Track whether any changes affect this cell
            change_cursor = 0
            cell_change_indices = set(change_indices)
            
            for op, text in diff_ops:
                if op == 'equal':
                    chunk_start = violation_pos
                    chunk_end = violation_pos + len(text)
                    
                    # Check if this chunk overlaps with cell range
                    if chunk_end > relative_start and chunk_start < relative_end:
                        overlap_start = max(0, relative_start - chunk_start)
                        overlap_end = min(len(text), relative_end - chunk_start)
                        revised_accumulator += text[overlap_start:overlap_end]
                    
                    violation_pos += len(text)
                
                elif op == 'delete':
                    chunk_start = violation_pos
                    chunk_end = violation_pos + len(text)
                    
                    # Check if this deletion affects the current cell
                    if chunk_end > relative_start and chunk_start < relative_end:
                        has_changes_in_cell = True
                    
                    violation_pos += len(text)
                    change_cursor += 1
                
                elif op == 'insert':
                    # Change-to-cell ownership is already resolved above.
                    # Respect that mapping to avoid boundary insertion loss.
                    if change_cursor in cell_change_indices:
                        revised_accumulator += text
                        has_changes_in_cell = True
                    change_cursor += 1
            
            # Use accumulator if there were changes (even if empty), otherwise keep original
            cell_revised = revised_accumulator if has_changes_in_cell else cell_violation
            
            # Fix: Decode cell_revised in table mode (same as cell_violation)
            # cell_revised is accumulated from escaped diff chunks and needs decoding
            # to match cell_violation which is already decoded
            if self._is_table_mode(affected_runs) and cell_revised != cell_violation:
                _cell_revised = cell_revised
                cell_revised = self._decode_json_escaped(cell_revised)
                if _cell_revised != cell_revised:
                    print(f"  [Multi-cell] Decoded revised_text json: {format_text_preview(_cell_revised, 60)}")

            result_edits.append({
                'cell_violation': cell_violation,
                'cell_revised': cell_revised,
                'cell_elem': cell_info['cell'],
                'cell_runs': cell_runs
            })
        
        return result_edits if result_edits else None

    def _apply_replace_in_cell_paragraphs(
        self,
        cell_violation: str,
        cell_revised: str,
        cell_runs: List[Dict],
        violation_reason: str,
        author: str,
        skip_comment: bool = False
    ) -> str:
        """Apply replace within a table cell that may contain multiple paragraphs.

        Follows the body cross-paragraph pattern: segments per paragraph,
        calls _apply_replace separately for each changed paragraph.

        Returns: 'success', 'cross_cell_fallback', 'conflict'
        """
        # 1. Collect unique paragraphs, sort by document order
        cell_paras = set()
        for run in cell_runs:
            para = run.get('para_elem')
            if para is not None:
                cell_paras.add(para)

        if not cell_paras:
            return 'cross_cell_fallback'

        if len(cell_paras) == 1:
            # Single paragraph — direct apply (existing simple path)
            para = next(iter(cell_paras))
            runs_info, text = self._collect_runs_info_original(para)
            pos = text.find(cell_violation)
            if pos == -1:
                return 'cross_cell_fallback'
            return self._apply_replace(
                para, cell_violation, cell_revised,
                violation_reason, runs_info, pos, author,
                skip_comment=skip_comment
            )

        # 2. Multiple paragraphs — sort, filter empty
        para_list = []
        for para in cell_paras:
            first_run_pos = None
            for run in cell_runs:
                if run.get('para_elem') is para:
                    first_run_pos = run.get('start', 0)
                    break
            para_list.append((first_run_pos or 0, para))
        para_list.sort(key=lambda x: x[0])

        # Build non-empty paragraph list (matching _collect_runs_info_in_table behavior)
        non_empty_paras = []
        for _, para in para_list:
            _, p_text = self._collect_runs_info_original(para)
            if p_text.strip():
                non_empty_paras.append(para)

        # 3. Build paragraph texts using real paragraph content.
        # This preserves in-paragraph line breaks (w:br -> "\n") and
        # avoids treating them as paragraph delimiters.
        para_texts = []
        for para in non_empty_paras:
            para_runs, _ = self._collect_runs_info_original(para)
            para_runs = self._strip_runs_whitespace(para_runs)
            para_texts.append(''.join(r.get('text', '') for r in para_runs))

        # Validate that the cell violation matches the reconstructed paragraphs.
        # Paragraph delimiters are represented by a single "\n" between paragraphs.
        expected_cell_text = '\n'.join(para_texts)
        if cell_violation != expected_cell_text:
            return 'cross_cell_fallback'

        # 4. Build para_segments from actual paragraph lengths
        para_segments = []
        offset = 0
        for i, para in enumerate(non_empty_paras):
            part_len = len(para_texts[i])
            para_segments.append({
                'para_elem': para,
                'seg_start': offset,
                'seg_end': offset + part_len,
                'match_pos_in_para': None,  # use find()
            })
            offset += part_len + 1  # +1 for \n

        return self._apply_diff_per_paragraph(
            para_segments, cell_violation, cell_revised,
            violation_reason, author,
            skip_comment=skip_comment,
            fallback_status='cross_cell_fallback',
            strip_runs=True,
        )

    @staticmethod
    def _extract_revised_segment(diff_ops: List, seg_start: int, seg_end: int) -> str:
        """Extract revised text corresponding to original range [seg_start, seg_end) from diff ops.

        Each diff op is a tuple of (op, text, markup_or_none). 'equal' and 'delete'
        consume positions in the original text; 'insert' adds text at the current
        original position without consuming it.
        """
        orig_pos = 0
        parts = []
        for op_tuple in diff_ops:
            op, text, _ = op_tuple if len(op_tuple) == 3 else (*op_tuple, None)
            if op == 'equal':
                op_start = orig_pos
                op_end = orig_pos + len(text)
                if op_end > seg_start and op_start < seg_end:
                    take_start = max(seg_start, op_start) - op_start
                    take_end = min(seg_end, op_end) - op_start
                    parts.append(text[take_start:take_end])
                orig_pos += len(text)
            elif op == 'delete':
                orig_pos += len(text)
            elif op == 'insert':
                if seg_start <= orig_pos <= seg_end:
                    parts.append(text)
        return ''.join(parts)
