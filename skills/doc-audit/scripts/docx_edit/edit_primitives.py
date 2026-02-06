"""Low-level edit primitives for run diff and table extraction."""

import copy
import difflib
import json
import re
from typing import Dict, Generator, Iterator, List, Optional, Tuple

from lxml import etree
from .common import DRAWING_PATTERN, EQUATION_PATTERN, NS, EditItem


class DocxEditPrimitivesMixin:
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
        
                # Track positions of all changes (delete/insert operations)
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

                    # Build cell_revised by applying diff operations within cell range
                    violation_pos = 0
                    revised_accumulator = ''
                    has_changes_in_cell = False  # Track whether any changes affect this cell
            
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
                
                        elif op == 'insert':
                            # Check if insertion point is within cell range (exclude right boundary)
                            # Insertions at exact cell boundary (violation_pos == relative_end) are 
                            # ambiguous and should trigger cross-cell fallback
                            if relative_start <= violation_pos < relative_end:
                                revised_accumulator += text
                                has_changes_in_cell = True
            
                    # Use accumulator if there were changes (even if empty), otherwise keep original
                    cell_revised = revised_accumulator if has_changes_in_cell else cell_violation
            
                    # Fix: Decode cell_revised in table mode (same as cell_violation)
                    # cell_revised is accumulated from escaped diff chunks and needs decoding
                    # to match cell_violation which is already decoded
                    if self._is_table_mode(affected_runs) and cell_revised != cell_violation:
                        _cell_revised = cell_revised
                        cell_revised = self._decode_json_escaped(cell_revised)
                        if _cell_revised != cell_revised:
                            print(f"  [Multi-cell] Decoded cell_revised to avoid JSON artifacts: {format_text_preview(_cell_revised, 60)}")

                    result_edits.append({
                        'cell_violation': cell_violation,
                        'cell_revised': cell_revised,
                        'cell_elem': cell_info['cell'],
                        'cell_runs': cell_runs
                    })
        
                return result_edits if result_edits else None

            def _try_extract_single_cell_edit(self, violation_text: str, revised_text: str,
                                               affected_runs: List[Dict], match_start: int) -> Optional[Dict]:
                """
                Try to extract single-cell edit when text spans multiple cells.

                This method analyzes the diff between violation_text and revised_text to
                determine if all changes are confined to a single table cell. If so, it
                extracts the cell-specific violation and revised text.

                Args:
                    violation_text: Original violation text (may span multiple cells)
                    revised_text: Revised text (may span multiple cells)
                    affected_runs: List of run info dicts that would be affected
                    match_start: Absolute position where violation_text starts in document

                Returns:
                    Dict with keys: 'cell_violation', 'cell_revised', 'cell_elem', 'cell_runs'
                    Or None if changes span multiple cells
                """
                # Calculate diff to find changed positions
                diff_ops = self._calculate_diff(violation_text, revised_text)
        
                # Track positions of all changes (delete/insert operations)
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
        
                # Find the cell(s) containing all changes
                cells_with_changes = set()
        
                # Use match_start as base offset for coordinate translation
                # change_positions are relative to violation_text (0-based)
                # match_start is the absolute position where violation_text starts in document
                base_offset = match_start
        
                for change_idx, (change_start_rel, change_end_rel) in enumerate(change_positions):
                    # Translate relative positions to absolute positions
                    change_start = base_offset + change_start_rel
                    change_end = base_offset + change_end_rel
            
                    # Determine if this is an insert (zero-length) or delete operation
                    is_insert = (change_start == change_end)
                    # Find runs affected by this change
                    for run_idx, run in enumerate(affected_runs):
                        # Skip synthetic boundary markers
                        if (run.get('is_json_boundary') or 
                            run.get('is_json_escape') or 
                            run.get('is_para_boundary')):
                            continue
                
                        # Check if this run is affected by the change (using absolute positions)
                        is_affected = False
                        if is_insert:
                            # Insert operation: check if insertion point is within run boundaries
                            # A run contains the insertion point if start <= pos <= end
                            if run['start'] <= change_start <= run['end']:
                                is_affected = True
                        else:
                            # Delete operation: check if run overlaps with the deletion range
                            if run['end'] > change_start and run['start'] < change_end:
                                is_affected = True
                
                        if is_affected:
                            cell = run.get('cell_elem')
                            if cell is not None:
                                cells_with_changes.add(id(cell))
                if len(cells_with_changes) != 1:
                    # Changes span multiple cells or no cell found
                    return None
        
                # All changes are in a single cell - extract cell-specific content
                # Find the target cell element
                target_cell_id = next(iter(cells_with_changes))
                target_cell = None
                cell_runs = []
        
                for run in affected_runs:
                    if run.get('is_json_boundary') or run.get('is_json_escape') or run.get('is_para_boundary'):
                        continue
                    cell = run.get('cell_elem')
                    if cell is not None and id(cell) == target_cell_id:
                        target_cell = cell
                        cell_runs.append(run)
        
                if target_cell is None or not cell_runs:
                    return None
        
                # Extract the portion of violation_text and revised_text within this cell
                cell_start = min(r['start'] for r in cell_runs)
                cell_end = max(r['end'] for r in cell_runs)
        
                # Convert absolute positions to relative offsets within violation_text
                # cell_start/cell_end are absolute positions in combined_text
                # violation_text starts at match_start, so we need relative offsets
                relative_start = cell_start - match_start
                relative_end = cell_end - match_start
        
                # Extract cell_violation using relative offsets
                # Account for JSON escaping if in table mode
                if self._is_table_mode(affected_runs):
                    cell_violation_escaped = violation_text[relative_start:relative_end]
                    cell_violation = self._decode_json_escaped(cell_violation_escaped)
                else:
                    cell_violation = violation_text[relative_start:relative_end]

                # Build cell_revised by applying diff operations to cell_violation
                # This handles cases where insertion/deletion changes text length
                cell_revised = cell_violation

                # Apply diff operations that fall within the cell range
                violation_pos = 0
                revised_accumulator = ''
                has_changes_in_cell = False
        
                for op, text in diff_ops:
                    if op == 'equal':
                        chunk_start = violation_pos
                        chunk_end = violation_pos + len(text)
                
                        # Check if this chunk overlaps with cell range [relative_start, relative_end)
                        if chunk_end > relative_start and chunk_start < relative_end:
                            # Calculate overlap
                            overlap_start = max(0, relative_start - chunk_start)
                            overlap_end = min(len(text), relative_end - chunk_start)
                            revised_accumulator += text[overlap_start:overlap_end]
                
                        violation_pos += len(text)
            
                    elif op == 'delete':
                        chunk_start = violation_pos
                        chunk_end = violation_pos + len(text)
                
                        # Skip deleted text if it falls within cell range
                        # (already handled by not including it in revised_accumulator)
                        if chunk_end > relative_start and chunk_start < relative_end:
                            has_changes_in_cell = True
                        violation_pos += len(text)
            
                    elif op == 'insert':
                        # Insert operations don't have a position in violation_text
                        # Check if insertion point is within cell range
                        if relative_start <= violation_pos <= relative_end:
                            revised_accumulator += text
                            has_changes_in_cell = True
        
                cell_revised = revised_accumulator if has_changes_in_cell else cell_violation

                # Fix: Decode cell_revised in table mode (same as cell_violation)
                # cell_revised is accumulated from escaped diff chunks and needs decoding
                # to avoid injecting JSON boundary characters into the cell text.
                if self._is_table_mode(affected_runs) and cell_revised != cell_violation:
                    _cell_revised = cell_revised
                    cell_revised = self._decode_json_escaped(cell_revised)
                    if _cell_revised != cell_revised:
                        print("  [Single-cell] Decoded cell_revised to avoid JSON artifacts: {format_text_preview(_cell_revised, 60)}")

                return {
                    'cell_violation': cell_violation,
                    'cell_revised': cell_revised,
                    'cell_elem': target_cell,
                    'cell_runs': cell_runs
                }

            def _extract_text_from_run_element(self, run_elem, rPr) -> str:
                """
                Extract text from a single run element with superscript/subscript markup.
        
                This matches parse_document.py's extract_text_from_run() behavior to ensure
                violation_text from LLM matches the text extracted during apply phase.
        
                Superscript/subscript detection:
                - Check w:rPr/w:vertAlign[@w:val="superscript|subscript"]
                - Wrap text with <sup>...</sup> or <sub>...</sub> markup
        
                Args:
                    run_elem: The w:r run element to extract text from
                    rPr: Pre-fetched w:rPr element (may be None)
        
                Returns:
                    Extracted text with superscript/subscript markup if applicable
                """
                # Check for vertical alignment (superscript/subscript)
                vert_align = None
                if rPr is not None:
                    vert_align_elem = rPr.find('w:vertAlign', NS)
                    if vert_align_elem is not None:
                        vert_align = vert_align_elem.get(f'{{{NS["w"]}}}val')
        
                # Extract text content from run
                text_parts = []
                for elem in run_elem:
                    if elem.tag == f'{{{NS["w"]}}}t':
                        text = elem.text or ''
                        if text:
                            text_parts.append(text)
                    elif elem.tag == f'{{{NS["w"]}}}delText':
                        # For deleted text in revision markup
                        text = elem.text or ''
                        if text:
                            text_parts.append(text)
                    elif elem.tag == f'{{{NS["w"]}}}tab':
                        text_parts.append('\t')
                    elif elem.tag == f'{{{NS["w"]}}}br':
                        # Handle line breaks - textWrapping or no type = soft line break
                        br_type = elem.get(f'{{{NS["w"]}}}type')
                        if br_type in (None, 'textWrapping'):
                            text_parts.append('\n')
                        # Skip page and column breaks (layout elements)
        
                combined_text = ''.join(text_parts)
        
                # Apply superscript/subscript markup
                if combined_text and vert_align in ('superscript', 'subscript'):
                    if vert_align == 'superscript':
                        return f'<sup>{combined_text}</sup>'
                    else:  # subscript
                        return f'<sub>{combined_text}</sub>'
        
                return combined_text

            def _collect_runs_info_original(self, para_elem) -> Tuple[List[Dict], str]:
                """
                Collect run info representing ORIGINAL text (before track changes).
                This is used for searching violation text in documents that have been edited.
        
                Logic:
                - Include: <w:delText> in <w:del> elements (deleted text was part of original)
                - Include: Normal <w:t>, <w:tab>, <w:br> NOT inside <w:ins> or <w:del> elements
                - Exclude: <w:t> inside <w:ins> elements (inserted text didn't exist in original)
                - Superscript/subscript: Wrapped with <sup>/<sub> tags for LLM matching
        
                Returns:
                    Tuple of (runs_info, combined_text)
                    runs_info: [{'text': str, 'start': int, 'end': int}, ...]
                    combined_text: Full text string with <sup>/<sub> markup
                """
                runs_info = []
                pos = 0
                w_ns = NS['w']
                wp_ns = NS['wp']
                m_ns = NS['m']
                w_ins_tag = f'{{{w_ns}}}ins'
                w_r_tag = f'{{{w_ns}}}r'
                m_omath_tag = f'{{{m_ns}}}oMath'
                m_omathpara_tag = f'{{{m_ns}}}oMathPara'

                def append_equation(omath_elem, host_run=None, host_rPr=None) -> None:
                    """Append a synthetic run entry for an OMML equation."""
                    nonlocal pos
                    try:
                        from omml import convert_omml_to_latex
                        latex = convert_omml_to_latex(omath_elem)
                    except Exception:
                        return
                    if not latex:
                        return
                    eq_text = f'<equation>{latex}</equation>'
                    runs_info.append({
                        'text': eq_text,
                        'start': pos,
                        'end': pos + len(eq_text),
                        'elem': host_run,
                        'rPr': host_rPr,
                        'is_equation': True
                    })
                    pos += len(eq_text)

                def append_run(run_elem) -> None:
                    """Append run entries in order, splitting drawings and equations."""
                    nonlocal pos
                    rPr = run_elem.find('w:rPr', NS)

                    # Check for vertical alignment (superscript/subscript)
                    vert_align = None
                    if rPr is not None:
                        vert_align_elem = rPr.find('w:vertAlign', NS)
                        if vert_align_elem is not None:
                            vert_align = vert_align_elem.get(f'{{{w_ns}}}val')

                    # Buffer for accumulating text before a drawing/equation
                    text_buffer = []

                    def flush_text_buffer() -> None:
                        """Helper to flush accumulated text as a separate entry."""
                        nonlocal pos
                        if not text_buffer:
                            return

                        combined_text = ''.join(text_buffer)

                        # Apply superscript/subscript markup if needed
                        if vert_align in ('superscript', 'subscript'):
                            if vert_align == 'superscript':
                                combined_text = f'<sup>{combined_text}</sup>'
                            else:
                                combined_text = f'<sub>{combined_text}</sub>'

                        runs_info.append({
                            'text': combined_text,
                            'start': pos,
                            'end': pos + len(combined_text),
                            'elem': run_elem,
                            'rPr': rPr,
                            'is_drawing': False
                        })
                        pos += len(combined_text)
                        text_buffer.clear()

                    # Process children in order to preserve text/image/equation positions
                    for elem in run_elem:
                        if elem.tag == f'{{{w_ns}}}t':
                            text = elem.text or ''
                            if text:
                                text_buffer.append(text)
                        elif elem.tag == f'{{{w_ns}}}delText':
                            text = elem.text or ''
                            if text:
                                text_buffer.append(text)
                        elif elem.tag == f'{{{w_ns}}}tab':
                            text_buffer.append('\t')
                        elif elem.tag == f'{{{w_ns}}}br':
                            br_type = elem.get(f'{{{w_ns}}}type')
                            if br_type in (None, 'textWrapping'):
                                text_buffer.append('\n')
                        elif elem.tag == f'{{{w_ns}}}drawing':
                            # Flush any accumulated text before the drawing
                            flush_text_buffer()

                            # Create separate entry for drawing
                            inline = elem.find(f'{{{wp_ns}}}inline')
                            if inline is not None:
                                doc_pr = inline.find(f'{{{wp_ns}}}docPr')
                                if doc_pr is not None:
                                    img_id = doc_pr.get('id', '')
                                    img_name = doc_pr.get('name', '')
                                    drawing_text = f'<drawing id="{img_id}" name="{img_name}" />'

                                    runs_info.append({
                                        'text': drawing_text,
                                        'start': pos,
                                        'end': pos + len(drawing_text),
                                        'elem': run_elem,
                                        'rPr': rPr,
                                        'is_drawing': True,
                                        'drawing_elem': elem  # Store reference to drawing element
                                    })
                                    pos += len(drawing_text)
                        elif elem.tag == m_omath_tag:
                            # Flush any accumulated text before the equation
                            flush_text_buffer()
                            append_equation(elem, run_elem, rPr)
                        elif elem.tag == m_omathpara_tag:
                            # Flush any accumulated text before the equation block
                            flush_text_buffer()
                            for omath in elem:
                                if omath.tag == m_omath_tag:
                                    append_equation(omath, run_elem, rPr)

                    # Flush any remaining text after all elements
                    flush_text_buffer()

                def walk(node) -> None:
                    tag = node.tag
                    if tag == w_ins_tag:
                        # Inserted text = NOT part of original, skip completely
                        return
                    if tag == w_r_tag:
                        append_run(node)
                        return
                    if tag == m_omath_tag:
                        append_equation(node)
                        return
                    if tag == m_omathpara_tag:
                        for child in node:
                            if child.tag == m_omath_tag:
                                append_equation(child)
                            else:
                                walk(child)
                        return
                    for child in node:
                        walk(child)

                # Walk paragraph content in document order
                for child in para_elem:
                    walk(child)
        
                combined_text = ''.join(r['text'] for r in runs_info)
                return runs_info, combined_text

            def _get_combined_text(self, runs_info: List[Dict]) -> str:
                """Get combined text from runs"""
                return ''.join(r['text'] for r in runs_info)

            def _find_affected_runs(self, runs_info: List[Dict],
                                   match_start: int, match_end: int) -> List[Dict]:
                """Find runs that overlap with the target text range"""
                affected = []
                for info in runs_info:
                    if info['end'] > match_start and info['start'] < match_end:
                        affected.append(info)
                return affected

            def _filter_real_runs(self, runs: List[Dict], include_equations: bool = False) -> List[Dict]:
                """
                Filter out synthetic boundary runs (JSON boundaries, paragraph boundaries).

                Synthetic runs are injected for text matching but don't have actual
                document elements (elem/rPr). This method filters them out before
                applying document modifications.

                Args:
                    runs: List of run info dicts
                    include_equations: If True, keep synthetic equation runs for comment anchoring

                Returns:
                    List of runs that have actual document elements
                """
                return [r for r in runs
                        if not r.get('is_json_boundary', False)
                        and not r.get('is_json_escape', False)
                        and not r.get('is_para_boundary', False)
                        and (include_equations or not r.get('is_equation', False))]

            def _get_run_original_text(self, run: Dict) -> str:
                """
                Get the original (unescaped) text for a run.

                In table mode, runs have 'original_text' with unescaped content
                and 'text' with JSON-escaped content. For document mutations,
                we always want the original text.

                Args:
                    run: Run info dict

                Returns:
                    Original text content (unescaped)
                """
                return run.get('original_text', run['text'])

            def _is_table_mode(self, runs: List[Dict]) -> bool:
                """
                Check if we're operating in table mode (runs have JSON-escaped text).

                Args:
                    runs: List of run info dicts

                Returns:
                    True if any run has 'original_text' field (indicating table mode)
                """
                return any(r.get('original_text') is not None for r in runs)

            def _decode_json_escaped(self, text: str) -> str:
                """
                Decode JSON-escaped text back to original.

                In table mode, violation_text and revised_text from the LLM contain
                JSON escape sequences (like \" for quotes, \\n for newlines).
                This method decodes them back to the original characters.

                Args:
                    text: JSON-escaped text string

                Returns:
                    Decoded original text
                """
                if not text:
                    return text
                try:
                    # json.loads expects a complete JSON string with quotes
                    return json.loads('"' + text + '"')
                except json.JSONDecodeError:
                    # Fallback: if decode fails, return original text
                    return text

            def _translate_escaped_offset(self, run: Dict, escaped_offset: int) -> int:
                """
                Translate an offset from escaped text space to original text space.

                In table mode, run['text'] contains JSON-escaped content where
                characters like " become \\". This method translates offsets
                from the escaped space to the original text space.

                Args:
                    run: Run info dict with 'text' and possibly 'original_text'
                    escaped_offset: Offset within run['text']

                Returns:
                    Corresponding offset in original text
                """
                if 'original_text' not in run:
                    return escaped_offset  # No escaping, offset is the same

                escaped_text = run['text']
                original_text = run['original_text']

                if escaped_offset <= 0:
                    return 0
                if escaped_offset >= len(escaped_text):
                    return len(original_text)

                # Try to decode the escaped prefix as JSON string content
                escaped_prefix = escaped_text[:escaped_offset]
                try:
                    # json.loads expects a complete JSON string
                    original_prefix = json.loads('"' + escaped_prefix + '"')
                    return len(original_prefix)
                except json.JSONDecodeError:
                    # Partial escape sequence - find the last valid boundary
                    for i in range(escaped_offset - 1, -1, -1):
                        try:
                            original_prefix = json.loads('"' + escaped_text[:i] + '"')
                            return len(original_prefix)
                        except json.JSONDecodeError:
                            continue
                    return 0

            def _get_rPr_xml(self, rPr_elem) -> str:
                """Convert rPr element to XML string"""
                if rPr_elem is None:
                    return ''
                return etree.tostring(rPr_elem, encoding='unicode')

            def _prepare_deletion_items(self, para_groups: List[Dict],
                                        match_start: int, match_end: int) -> List[Dict]:
                """
                Build prepared deletion items from paragraph-grouped runs.

                Shared preparation logic for _apply_delete_cross_paragraph and
                _apply_delete_multi_cell. Computes before_text, del_text, after_text
                for each paragraph in the deletion range.

                Args:
                    para_groups: List of {'para_elem': ..., 'runs': [...]} dicts
                                 (from _group_runs_by_paragraph)
                    match_start: Start position of the matched violation text
                    match_end: End position of the matched violation text

                Returns:
                    List of dicts with keys:
                    - para_elem, affected_runs, rPr_xml, before_text, del_text, after_text
                """
                prepared = []
                for group in para_groups:
                    para_elem = group['para_elem']
                    para_runs = group['runs']

                    # Identify affected runs within this paragraph
                    first_run = None
                    last_run = None
                    affected_runs_in_para = []
                    for run in para_runs:
                        if run['end'] > match_start and run['start'] < match_end:
                            if first_run is None:
                                first_run = run
                            last_run = run
                            affected_runs_in_para.append(run)

                    if first_run is None or last_run is None:
                        if self.verbose:
                            print(f"  [Prepare] Skipping paragraph: no affected runs")
                        continue

                    rPr_xml = self._get_rPr_xml(first_run.get('rPr'))

                    before_offset = self._translate_escaped_offset(
                        first_run, max(0, match_start - first_run['start']))
                    after_offset = self._translate_escaped_offset(
                        last_run, max(0, match_end - last_run['start']))

                    first_orig_text = self._get_run_original_text(first_run)
                    last_orig_text = self._get_run_original_text(last_run)

                    before_text = first_orig_text[:before_offset]
                    after_text = last_orig_text[after_offset:]

                    # Build deleted text from affected runs in this paragraph
                    del_parts = []
                    for run in affected_runs_in_para:
                        if run.get('is_drawing'):
                            del_parts.append(run['text'])
                            continue
                        orig_text = self._get_run_original_text(run)
                        if run is first_run and run is last_run:
                            del_parts.append(orig_text[before_offset:after_offset])
                        elif run is first_run:
                            del_parts.append(orig_text[before_offset:])
                        elif run is last_run:
                            del_parts.append(orig_text[:after_offset])
                        else:
                            del_parts.append(orig_text)

                    del_text = ''.join(del_parts)
                    if not del_text:
                        if self.verbose:
                            print(f"  [Prepare] Skipping paragraph: no text to delete")
                        continue

                    prepared.append({
                        'para_elem': para_elem,
                        'affected_runs': affected_runs_in_para,
                        'rPr_xml': rPr_xml,
                        'before_text': before_text,
                        'del_text': del_text,
                        'after_text': after_text,
                    })

                return prepared

            def _delete_paragraphs_in_unit(self, prepared: List[Dict],
                                            shared_change_id: str, author: str,
                                            comment_id: Optional[int] = None) -> int:
                """
                Delete content across multiple paragraphs in a single unit.

                Shared core loop for _apply_delete_cross_paragraph and _apply_delete_multi_cell.
                Handles building deletion elements, replacing runs, and adding paragraph mark
                deletion for non-last paragraphs so Word displays a unified deletion.

                Args:
                    prepared: List from _prepare_deletion_items()
                    shared_change_id: Change ID shared across all paragraphs in this unit
                    author: Track change author
                    comment_id: If set, insert commentRangeStart before first deletion
                                and commentRangeEnd+ref after last deletion

                Returns:
                    Number of paragraphs successfully processed
                """
                if not prepared:
                    return 0

                success_count = 0
                is_first_para = True

                for item_idx, item in enumerate(prepared):
                    para_elem = item['para_elem']
                    affected_runs = item['affected_runs']
                    rPr_xml = item['rPr_xml']
                    before_text = item['before_text']
                    del_text = item['del_text']
                    after_text = item['after_text']

                    is_last_para = (item_idx == len(prepared) - 1)

                    # Build new elements for this paragraph
                    new_elements = []

                    if before_text:
                        run_or_container = self._create_run(before_text, rPr_xml)
                        if run_or_container.tag == 'container':
                            new_elements.extend(list(run_or_container))
                        else:
                            new_elements.append(run_or_container)

                    # Comment range start (only for first paragraph, if comment_id provided)
                    if is_first_para and comment_id is not None:
                        comment_start_xml = (
                            f'<w:commentRangeStart xmlns:w="{NS["w"]}" '
                            f'w:id="{comment_id}"/>'
                        )
                        new_elements.append(etree.fromstring(comment_start_xml))

                    # Deleted text with shared change_id
                    self._append_del_elements(
                        new_elements, del_text, rPr_xml, author,
                        change_id=shared_change_id
                    )

                    # Comment range end and reference (only for last paragraph)
                    if is_last_para and comment_id is not None:
                        comment_end_xml = (
                            f'<w:commentRangeEnd xmlns:w="{NS["w"]}" '
                            f'w:id="{comment_id}"/>'
                        )
                        new_elements.append(etree.fromstring(comment_end_xml))
                        comment_ref_xml = (
                            f'<w:r xmlns:w="{NS["w"]}">'
                            f'<w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>'
                            f'<w:commentReference w:id="{comment_id}"/>'
                            f'</w:r>'
                        )
                        new_elements.append(etree.fromstring(comment_ref_xml))

                    if after_text:
                        run_or_container = self._create_run(after_text, rPr_xml)
                        if run_or_container.tag == 'container':
                            new_elements.extend(list(run_or_container))
                        else:
                            new_elements.append(run_or_container)

                    # Replace runs and check success
                    replace_success = self._replace_runs(
                        para_elem, affected_runs, new_elements
                    )

                    if replace_success:
                        # For non-last paragraphs, mark paragraph mark (¶) as deleted
                        # so Word merges consecutive deleted paragraphs into one unified
                        # deletion. Without this, Word displays each deletion separately.
                        if not is_last_para:
                            pPr = para_elem.find(f'{{{NS["w"]}}}pPr')
                            if pPr is None:
                                pPr = etree.SubElement(para_elem, f'{{{NS["w"]}}}pPr')
                                # pPr must be the first child of w:p
                                para_elem.insert(0, pPr)
                            rPr_in_pPr = pPr.find(f'{{{NS["w"]}}}rPr')
                            if rPr_in_pPr is None:
                                rPr_in_pPr = etree.SubElement(pPr, f'{{{NS["w"]}}}rPr')
                            del_mark = etree.SubElement(
                                rPr_in_pPr, f'{{{NS["w"]}}}del'
                            )
                            del_mark.set(f'{{{NS["w"]}}}id', shared_change_id)
                            del_mark.set(f'{{{NS["w"]}}}author', author)
                            del_mark.set(f'{{{NS["w"]}}}date', self.operation_timestamp)

                        success_count += 1
                    else:
                        if self.verbose:
                            print(f"  [Delete unit] Failed to replace runs in paragraph")

                    is_first_para = False

                return success_count

            def _append_del_elements(self, new_elements: List[etree.Element],
                                     del_text: str, rPr_xml: str, author: str,
                                     change_id: Optional[str] = None):
                """Append deletion elements for del_text into new_elements.
        
                Args:
                    new_elements: List to append deletion elements to
                    del_text: Text to delete
                    rPr_xml: Run properties XML
                    author: Author name
                    change_id: Optional pre-generated change ID. If None, generates new ID.
                """
                if not del_text:
                    return

                change_id = change_id or self._get_next_change_id()
                run_or_container = self._create_run(del_text, rPr_xml)

                if run_or_container.tag == 'container':
                    # Multiple runs (has markup) - wrap each in w:del
                    for del_run in run_or_container:
                        t_elem = del_run.find(f'{{{NS["w"]}}}t')
                        if t_elem is not None:
                            t_elem.tag = f'{{{NS["w"]}}}delText'

                        del_elem = etree.Element(f'{{{NS["w"]}}}del')
                        del_elem.set(f'{{{NS["w"]}}}id', change_id)
                        del_elem.set(f'{{{NS["w"]}}}author', author)
                        del_elem.set(f'{{{NS["w"]}}}date', self.operation_timestamp)
                        del_elem.append(del_run)
                        new_elements.append(del_elem)
                else:
                    t_elem = run_or_container.find(f'{{{NS["w"]}}}t')
                    if t_elem is not None:
                        t_elem.tag = f'{{{NS["w"]}}}delText'

                    del_elem = etree.Element(f'{{{NS["w"]}}}del')
                    del_elem.set(f'{{{NS["w"]}}}id', change_id)
                    del_elem.set(f'{{{NS["w"]}}}author', author)
                    del_elem.set(f'{{{NS["w"]}}}date', self.operation_timestamp)
                    del_elem.append(run_or_container)
                    new_elements.append(del_elem)

            def _build_ins_elements_with_breaks(self, text: str, rPr_xml: str, author: str) -> List[etree.Element]:
                """
                Build insertion elements for text, converting '\\n' to <w:br/>.
                Returns a list of w:ins elements.
                """
                if not text:
                    return []

                normalized = text.replace('\r\n', '\n').replace('\r', '\n')
                lines = normalized.split('\n')
                ins_elements = []
                change_id = self._get_next_change_id()

                for idx, line in enumerate(lines):
                    if line:
                        run_or_container = self._create_run(line, rPr_xml)
                        if run_or_container.tag == 'container':
                            for ins_run in run_or_container:
                                ins_elem = etree.Element(f'{{{NS["w"]}}}ins')
                                ins_elem.set(f'{{{NS["w"]}}}id', change_id)
                                ins_elem.set(f'{{{NS["w"]}}}author', author)
                                ins_elem.set(f'{{{NS["w"]}}}date', self.operation_timestamp)
                                ins_elem.append(ins_run)
                                ins_elements.append(ins_elem)
                        else:
                            ins_elem = etree.Element(f'{{{NS["w"]}}}ins')
                            ins_elem.set(f'{{{NS["w"]}}}id', change_id)
                            ins_elem.set(f'{{{NS["w"]}}}author', author)
                            ins_elem.set(f'{{{NS["w"]}}}date', self.operation_timestamp)
                            ins_elem.append(run_or_container)
                            ins_elements.append(ins_elem)

                    # Insert line break between lines
                    if idx < len(lines) - 1:
                        br_run = etree.Element(f'{{{NS["w"]}}}r')
                        etree.SubElement(br_run, f'{{{NS["w"]}}}br')
                        ins_elem = etree.Element(f'{{{NS["w"]}}}ins')
                        ins_elem.set(f'{{{NS["w"]}}}id', change_id)
                        ins_elem.set(f'{{{NS["w"]}}}author', author)
                        ins_elem.set(f'{{{NS["w"]}}}date', self.operation_timestamp)
                        ins_elem.append(br_run)
                        ins_elements.append(ins_elem)

                return ins_elements

            def _check_special_element_modification(self, violation_text: str, diff_ops: List, has_markup: bool = False) -> Tuple[bool, str]:
                """
                Check if any diff operation modifies content inside special elements (<drawing> or <equation>).
        
                Strategy:
                1. Find all special element position ranges in violation_text
                2. If has_markup=True, map ranges to plain-text coordinate space (diff ops work on plain text)
                3. Track position through diff ops
                4. If any delete/insert operation overlaps with special element ranges, reject
        
                Args:
                    violation_text: Original violation text (may contain special elements and markup)
                    diff_ops: List of diff operations from _calculate_diff or _calculate_markup_aware_diff
                    has_markup: If True, diff_ops work on plain text (markup stripped), need coordinate mapping
        
                Returns:
                    Tuple of (should_reject, reason)
                    - should_reject: True if modification involves special element content
                    - reason: Description of why rejection is needed
                """
                # Find all special element position ranges in violation_text (original coordinates)
                special_ranges_orig = []  # [(start, end, element_type), ...]
        
                for match in DRAWING_PATTERN.finditer(violation_text):
                    special_ranges_orig.append((match.start(), match.end(), 'drawing'))
        
                for match in EQUATION_PATTERN.finditer(violation_text):
                    special_ranges_orig.append((match.start(), match.end(), 'equation'))
        
                # If no special elements, only check if inserting new ones
                if not special_ranges_orig:
                    # Rebuild complete inserted text from consecutive insert operations
                    # This handles the case where markup-aware diff splits insertions containing
                    # <equation> tags into multiple chunks (e.g., '<equation>H', '2', 'O</equation>')
                    # where individual chunks don't match the full pattern
                    full_insert_text = ''.join(
                        op_tuple[1] for op_tuple in diff_ops 
                        if op_tuple[0] == 'insert'
                    )
            
                    if full_insert_text:
                        if DRAWING_PATTERN.search(full_insert_text):
                            return True, "Cannot insert drawing via revision markup"
                        if EQUATION_PATTERN.search(full_insert_text):
                            return True, "Cannot insert equation via revision markup"
                    return False, ""
        
                # Check if diff operations would modify content inside special elements
                # Strategy: Track position through diff ops and check overlap with special element ranges
        
                # Map special_ranges to plain-text coordinates if markup is present
                if has_markup and special_ranges_orig:
                    # Build mapping from original position to plain-text position
                    segments = self._parse_formatted_text(violation_text)
                    orig_to_plain = {}  # Map original char index -> plain char index
                    plain_pos = 0
                    orig_pos = 0
            
                    for text, vert_align in segments:
                        # Original text may have <sup>text</sup> (11 chars for "text")
                        # Plain text has just "text" (4 chars)
                        if vert_align == 'superscript':
                            # Original: <sup>text</sup>
                            markup_before = '<sup>'
                            markup_after = '</sup>'
                        elif vert_align == 'subscript':
                            # Original: <sub>text</sub>
                            markup_before = '<sub>'
                            markup_after = '</sub>'
                        else:
                            # No markup
                            markup_before = ''
                            markup_after = ''
                
                        # Skip markup_before in original, map content
                        orig_pos += len(markup_before)
                        for char in text:
                            orig_to_plain[orig_pos] = plain_pos
                            orig_pos += 1
                            plain_pos += 1
                        orig_pos += len(markup_after)
            
                    # Transform special_ranges to plain-text coordinates
                    special_ranges = []
                    for elem_start_orig, elem_end_orig, elem_type in special_ranges_orig:
                        # Find plain-text positions for element boundaries
                        # Use the first and last content positions within the element
                        elem_start_plain = None
                        elem_end_plain = None
                
                        # Find first content position in element
                        for orig_idx in range(elem_start_orig, elem_end_orig):
                            if orig_idx in orig_to_plain:
                                elem_start_plain = orig_to_plain[orig_idx]
                                break
                
                        # Find last content position in element
                        for orig_idx in range(elem_end_orig - 1, elem_start_orig - 1, -1):
                            if orig_idx in orig_to_plain:
                                elem_end_plain = orig_to_plain[orig_idx] + 1  # +1 for exclusive end
                                break
                
                        # If element is entirely within markup tags (no content mapped), skip it
                        # This shouldn't happen for <drawing> or <equation> but handle gracefully
                        if elem_start_plain is not None and elem_end_plain is not None:
                            special_ranges.append((elem_start_plain, elem_end_plain, elem_type))
                else:
                    # No markup: use original coordinates directly
                    special_ranges = special_ranges_orig
        
                # Pre-check: Rebuild full insert text and check for new special elements
                # This must be done BEFORE the main loop to catch markup-split insertions
                # (e.g., '<equation>H', '2', 'O</equation>' from markup-aware diff)
                full_insert_text = ''.join(
                    op_tuple[1] for op_tuple in diff_ops 
                    if op_tuple[0] == 'insert'
                )
        
                if full_insert_text:
                    if DRAWING_PATTERN.search(full_insert_text):
                        return True, "Cannot insert drawing via revision markup"
                    if EQUATION_PATTERN.search(full_insert_text):
                        return True, "Cannot insert equation via revision markup"
        
                # Track cumulative deletion coverage for each special element
                # This handles cases where markup-aware diff splits a complete deletion into multiple delete ops
                # (e.g., deleting <equation>H<sub>2</sub>O</equation> produces separate deletes for non-markup and subscript segments)
                elem_deleted_ranges = {i: [] for i in range(len(special_ranges))}
        
                # First pass: collect all delete operations and their coverage of each special element
                # Also check if any special element would survive in "equal" segments when has_markup=True
                current_pos = 0
                for op_tuple in diff_ops:
                    # Handle both 2-tuple and 3-tuple formats
                    op = op_tuple[0]
                    text = op_tuple[1]

                    if op == 'equal' and has_markup:
                        # When has_markup=True, equal segments preserve content unchanged
                        # No need to check special elements - they remain intact
                        current_pos += len(text)

                    elif op == 'delete':
                        del_start = current_pos
                        del_end = current_pos + len(text)

                        # Record overlap with each special element
                        for idx, (elem_start, elem_end, elem_type) in enumerate(special_ranges):
                            if del_end > elem_start and del_start < elem_end:
                                # Calculate overlap relative to element start
                                overlap_start = max(del_start, elem_start) - elem_start
                                overlap_end = min(del_end, elem_end) - elem_start
                                elem_deleted_ranges[idx].append((overlap_start, overlap_end))

                        current_pos += len(text)

                    elif op == 'equal':
                        current_pos += len(text)

                    elif op == 'insert':
                        # Check if inserting at a position inside special element
                        for elem_start, elem_end, elem_type in special_ranges:
                            if elem_start < current_pos < elem_end:
                                return True, f"Cannot insert inside {elem_type}"
                
                        # Note: Full insert text is already checked before loop (see pre-check above)
                        # Individual chunk checks are redundant and buggy - removed

                # Second pass: check if any special element was partially deleted
                for idx, (elem_start, elem_end, elem_type) in enumerate(special_ranges):
                    deleted_ranges = elem_deleted_ranges[idx]
            
                    if not deleted_ranges:
                        continue  # Element not affected by deletions
            
                    elem_len = elem_end - elem_start
            
                    # Merge overlapping ranges and calculate total coverage
                    merged_ranges = []
                    for start, end in sorted(deleted_ranges):
                        if merged_ranges and start <= merged_ranges[-1][1]:
                            # Overlapping or adjacent - merge
                            merged_ranges[-1] = (merged_ranges[-1][0], max(merged_ranges[-1][1], end))
                        else:
                            merged_ranges.append((start, end))
            
                    total_deleted = sum(end - start for start, end in merged_ranges)
            
                    if total_deleted == elem_len:
                        # Complete deletion - reject (cannot track change delete special elements)
                        return True, f"Cannot delete {elem_type} via revision markup"
                    else:
                        # Partial deletion - reject
                        return True, f"Cannot partially modify {elem_type} content"

                return False, ""

            def _calculate_diff(self, old_text: str, new_text: str) -> List[Tuple[str, str]]:
                """
                Calculate minimal diff between two texts.
                Returns: [('equal', text), ('delete', text), ('insert', text), ...]
                """
                matcher = difflib.SequenceMatcher(None, old_text, new_text)
                operations = []
        
                for tag, i1, i2, j1, j2 in matcher.get_opcodes():
                    if tag == 'equal':
                        operations.append(('equal', old_text[i1:i2]))
                    elif tag == 'delete':
                        operations.append(('delete', old_text[i1:i2]))
                    elif tag == 'insert':
                        operations.append(('insert', new_text[j1:j2]))
                    elif tag == 'replace':
                        operations.append(('delete', old_text[i1:i2]))
                        operations.append(('insert', new_text[j1:j2]))
        
                return operations

            def _calculate_markup_aware_diff(self, old_text: str, new_text: str) -> List[Tuple[str, str, Optional[str]]]:
                """
                Calculate diff with markup awareness for <sup>/<sub> tags.
        
                This method parses both texts into segments with formatting info,
                then performs diff on the text content only (ignoring markup tags).
        
                Args:
                    old_text: Original text with possible <sup>/<sub> markup
                    new_text: New text with possible <sup>/<sub> markup
        
                Returns:
                    List of (operation, text, vert_align) tuples where:
                    - operation: 'equal' | 'delete' | 'insert'
                    - text: The actual text content (without markup tags)
                    - vert_align: 'superscript' | 'subscript' | None
        
                Examples:
                    old: "x<sup>2</sup>"  new: "x<sup>3</sup>"
                    → [('equal', 'x', None), ('delete', '2', 'superscript'), ('insert', '3', 'superscript')]
                """
                # Parse both texts into segments
                old_segments = self._parse_formatted_text(old_text)
                new_segments = self._parse_formatted_text(new_text)
        
                # Build plain text versions for diffing (text content only, no markup)
                old_plain = ''.join(text for text, _ in old_segments)
                new_plain = ''.join(text for text, _ in new_segments)
        
                # Perform character-level diff on plain text
                matcher = difflib.SequenceMatcher(None, old_plain, new_plain)
                operations = []
        
                for tag, i1, i2, j1, j2 in matcher.get_opcodes():
                    if tag == 'equal':
                        # Find segments in old_text that cover this range
                        ops = self._map_range_to_segments(old_segments, i1, i2, 'equal')
                        operations.extend(ops)
                    elif tag == 'delete':
                        # Find segments in old_text
                        ops = self._map_range_to_segments(old_segments, i1, i2, 'delete')
                        operations.extend(ops)
                    elif tag == 'insert':
                        # Find segments in new_text
                        ops = self._map_range_to_segments(new_segments, j1, j2, 'insert')
                        operations.extend(ops)
                    elif tag == 'replace':
                        # Delete from old, insert from new
                        ops_del = self._map_range_to_segments(old_segments, i1, i2, 'delete')
                        ops_ins = self._map_range_to_segments(new_segments, j1, j2, 'insert')
                        operations.extend(ops_del)
                        operations.extend(ops_ins)
        
                return operations

            def _map_range_to_segments(self, segments: List[Tuple[str, Optional[str]]], 
                                        start: int, end: int, operation: str) -> List[Tuple[str, str, Optional[str]]]:
                """
                Map a character range to segments with formatting info.
        
                Args:
                    segments: List of (text, vert_align) tuples
                    start: Start position in plain text
                    end: End position in plain text
                    operation: 'equal' | 'delete' | 'insert'
        
                Returns:
                    List of (operation, text, vert_align) tuples
                """
                result = []
                pos = 0
        
                for text, vert_align in segments:
                    segment_start = pos
                    segment_end = pos + len(text)
            
                    # Check if this segment overlaps with [start, end)
                    if segment_end <= start:
                        # Before the range
                        pos = segment_end
                        continue
                    if segment_start >= end:
                        # After the range
                        break
            
                    # Calculate overlap
                    overlap_start = max(0, start - segment_start)
                    overlap_end = min(len(text), end - segment_start)
            
                    if overlap_start < overlap_end:
                        chunk = text[overlap_start:overlap_end]
                        result.append((operation, chunk, vert_align))
            
                    pos = segment_end
        
                return result
