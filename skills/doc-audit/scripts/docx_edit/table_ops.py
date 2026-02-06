"""Table-oriented run collection and boundary detection."""

from typing import Dict, Generator, Iterator, List, Optional, Tuple

from lxml import etree

from .common import NS, json_escape


class DocxTableMixin:
        def _get_cell_merge_properties(self, tcPr):
            """
            Get cell merge properties (gridSpan and vMerge).
        
            Args:
                tcPr: Cell properties element (w:tcPr)
        
            Returns:
                Tuple of (grid_span, vmerge_type)
                - grid_span: Number of columns this cell spans (default 1)
                - vmerge_type: 'restart' | 'continue' | None
            """
            grid_span = 1
            vmerge_type = None
        
            if tcPr is not None:
                # Check gridSpan (horizontal merge)
                gs = tcPr.find(f'{{{NS["w"]}}}gridSpan')
                if gs is not None:
                    try:
                        grid_span = int(gs.get(f'{{{NS["w"]}}}val'))
                    except (ValueError, TypeError):
                        grid_span = 1
            
                # Check vMerge (vertical merge)
                vmerge_elem = tcPr.find(f'{{{NS["w"]}}}vMerge')
                if vmerge_elem is not None:
                    vmerge_val = vmerge_elem.get(f'{{{NS["w"]}}}val')
                    if vmerge_val == 'restart':
                        vmerge_type = 'restart'
                    else:
                        # None or 'continue' both mean continue
                        vmerge_type = 'continue'
        
            return grid_span, vmerge_type

        def _collect_runs_info_across_paragraphs(self, start_para, uuid_end: str) -> Tuple[List[Dict], str, bool, Optional[str]]:
            """
            Collect run info across multiple paragraphs (uuid → uuid_end range).

            Behavior:
            - Body text: Paragraph boundaries are represented as '\\n'
            - Table content (same row): JSON format ["cell1", "cell2"] with '", "' between cells
            - Cross body/table boundary: Returns boundary_error
            - Cross table row boundary: Returns row_boundary_error (Word doesn't support cross-row comments)

            Args:
                start_para: Starting paragraph element
                uuid_end: End boundary paraId (inclusive)

            Returns:
                Tuple of (runs_info, combined_text, is_cross_paragraph, boundary_error)
                - runs_info: List of run info dicts with 'text', 'start', 'end', 'para_elem', etc.
                - combined_text: Full text string
                - is_cross_paragraph: True if text spans multiple paragraphs
                - boundary_error: None if OK, or error type string:
                  - 'boundary_crossed': Crossed body/table or different tables
            """
            # 1. Detect if start paragraph is in a table
            start_in_table = self._is_paragraph_in_table(start_para)
            start_table = self._find_ancestor_table(start_para) if start_in_table else None

            # 2. Find end paragraph and check its context
            end_para = self._resolve_end_para(start_para, uuid_end)
            end_in_table = self._is_paragraph_in_table(end_para) if end_para is not None else False
            end_table = self._find_ancestor_table(end_para) if end_in_table else None

            # 3. Check boundary consistency
            if start_in_table != end_in_table:
                # Crossed body/table boundary
                return [], '', False, 'boundary_crossed'

            if start_in_table and start_table is not end_table:
                # Crossed different tables
                return [], '', False, 'boundary_crossed'

            # Note: Row boundary check is NOT done here.
            # For multi-row table blocks, we collect content row by row and let
            # the caller check if the actual match spans multiple rows.

            # 4. Dispatch to appropriate collector
            if start_in_table:
                return self._collect_runs_info_in_table(start_para, uuid_end, start_table, end_para=end_para)
            else:
                return self._collect_runs_info_in_body(start_para, uuid_end, end_para=end_para)

        def _collect_runs_info_in_body(self, start_para, uuid_end: str, end_para=None) -> Tuple[List[Dict], str, bool, Optional[str]]:
            """
            Collect run info across paragraphs in document body (not in table).
            Paragraph boundaries are represented as '\\n'.

            Args:
                start_para: Starting paragraph element
                uuid_end: End boundary paraId (inclusive)

            Returns:
                Tuple of (runs_info, combined_text, is_cross_paragraph, boundary_error)
            """
            if end_para is None:
                end_para = self._resolve_end_para(start_para, uuid_end)
            runs_info = []
            pos = 0
            para_count = 0

            for para in self._iter_paragraphs_in_range(start_para, uuid_end):
                para_count += 1

                # Collect runs for this paragraph
                para_runs, para_text = self._collect_runs_info_original(para)
                para_id = para.get('{http://schemas.microsoft.com/office/word/2010/wordml}paraId')

                # Skip empty paragraphs to match parse_document.py behavior
                if not para_text.strip():
                    continue

                # Add paragraph element reference to each run
                for run in para_runs:
                    run['para_elem'] = para
                    run['host_para_elem'] = para
                    run['host_para_id'] = para_id
                    run['start'] += pos
                    run['end'] += pos
                    runs_info.append(run)

                pos += len(para_text)

                # Add paragraph boundary (except after last paragraph)
                is_last_para = False
                if end_para is not None:
                    is_last_para = para is end_para
                else:
                    para_id = para.get('{http://schemas.microsoft.com/office/word/2010/wordml}paraId')
                    is_last_para = (para_id == uuid_end)
                if not is_last_para:
                    # Not the last paragraph - add boundary marker
                    runs_info.append({
                        'text': '\n',
                        'start': pos,
                        'end': pos + 1,
                        'para_elem': para,
                        'is_para_boundary': True  # Mark as paragraph boundary
                    })
                    pos += 1

            combined_text = ''.join(r['text'] for r in runs_info)
            is_cross_paragraph = para_count > 1

            return runs_info, combined_text, is_cross_paragraph, None

        def _collect_runs_info_in_table(self, start_para, uuid_end: str, table_elem, end_para=None) -> Tuple[List[Dict], str, bool, Optional[str]]:
            """
            Collect run info within a table using JSON format, handling multiple rows and merged cells.

            For each row, format matches parse_document.py: ["cell1", "cell2", "cell3"]
            - Cell boundaries within row: '", "'
            - Paragraph boundaries within cell: '\\n' (JSON-escaped as '\\n')
            - Row boundaries: '"], ["' (close previous row, open new row)
            - Content is JSON-escaped to match LLM output
        
            Merged cell handling (consistent with TableExtractor):
            - Horizontal merge (gridSpan): Content in first cell only, spans multiple columns
            - Vertical merge (vMerge):
              - restart: Start of merge region, content recorded in vmerge_content
              - continue: Copy content from vmerge_content if available (within uuid range)
              - If restart is outside uuid range, continue cells remain empty
        
            Note: This method collects content across multiple rows. The caller should
            check if the actual match spans multiple rows using _check_cross_row_boundary().

            Args:
                start_para: Starting paragraph element
                uuid_end: End boundary paraId (inclusive)
                table_elem: The table element (w:tbl)

            Returns:
                Tuple of (runs_info, combined_text, is_cross_paragraph, boundary_error)
                - runs_info includes 'cell_elem' and 'row_elem' fields
                - runs_info includes 'original_text' for actual content (before JSON escaping)
            """
            if end_para is None:
                end_para = self._resolve_end_para(start_para, uuid_end)

            def is_end_para(para) -> bool:
                if end_para is not None:
                    return para is end_para
                para_id = para.get('{http://schemas.microsoft.com/office/word/2010/wordml}paraId')
                return para_id == uuid_end

            # Get number of columns from tblGrid
            tbl_grid = table_elem.find(f'{{{NS["w"]}}}tblGrid')
            num_cols = 0
            if tbl_grid is not None:
                num_cols = len(tbl_grid.findall(f'{{{NS["w"]}}}gridCol'))
        
            if num_cols == 0:
                # No columns defined, fallback to paragraph iteration
                return self._collect_runs_info_in_table_legacy(start_para, uuid_end, table_elem)
        
            runs_info = []
            pos = 0
            in_range = False
            first_row_in_range = True
            vmerge_content = {}  # {grid_col: {'runs': [...], 'text': str}}
            last_para = start_para
            reached_end = False  # Flag to stop when uuid_end is found
            cols_in_range = set()  # Track which columns are within uuid range
        
            # Iterate rows (w:tr elements)
            for tr in table_elem.findall(f'{{{NS["w"]}}}tr'):
                if reached_end:
                    break
            
                grid_col = 0
                row_data = [None] * num_cols  # Pre-fill with None to match TableExtractor behavior
            
                # Iterate cells (w:tc elements) in this row
                for tc in tr.findall(f'{{{NS["w"]}}}tc'):
                    if reached_end:
                        break
                    # Get cell properties
                    tcPr = tc.find(f'{{{NS["w"]}}}tcPr')
                    grid_span, vmerge_type = self._get_cell_merge_properties(tcPr)
                
                    # Collect all paragraphs in this cell
                    cell_paras = tc.findall(f'{{{NS["w"]}}}p')
                
                    # Check if any paragraph in this cell is in range
                    cell_in_range = False
                    for para in cell_paras:
                        para_id = para.get('{http://schemas.microsoft.com/office/word/2010/wordml}paraId')
                        if para is start_para:
                            in_range = True
                            cell_in_range = True
                        if in_range and not reached_end and para_id:
                            cell_in_range = True
                            last_para = para
                            if is_end_para(para):
                                reached_end = True
                                break
                
                    # Collect cell content if in range
                    cell_runs = []
                    cell_text = ''
                
                    if cell_in_range:
                        # Mark all columns spanned by this cell as in range
                        cols_in_range.update(range(grid_col, grid_col + grid_span))
                        # Determine cell content based on vMerge type
                        if vmerge_type == 'restart':
                            # Merge restart: collect content and store in vmerge_content
                            for para_idx, para in enumerate(cell_paras):
                                para_id = para.get('{http://schemas.microsoft.com/office/word/2010/wordml}paraId')
                                if not in_range and para is not start_para:
                                    continue
                                if para is start_para:
                                    in_range = True
                            
                                if para_idx > 0:
                                    # Add paragraph boundary marker
                                    cell_runs.append({
                                        'text': '\\n',
                                        'is_json_escape': True,
                                        'is_para_boundary': True,
                                        'para_elem': para,
                                        'cell_elem': tc,
                                        'row_elem': tr
                                    })
                                    cell_text += '\\n'
                            
                                para_runs, para_orig_text = self._collect_runs_info_original(para)
                                # Strip paragraph text to match table_extractor.py behavior (line 144/163)
                                # table_extractor.py calls para_text.strip() before checking if empty
                                para_orig_text = para_orig_text.strip()
                            
                                # Skip empty paragraphs (match table_extractor.py behavior at line 144/163)
                                # table_extractor.py strips each paragraph and skips if empty
                                if not para_orig_text:
                                    continue
                            
                                para_runs = self._strip_runs_whitespace(para_runs)
                                for run in para_runs:
                                    original_text = run['text']
                                    escaped_text = json_escape(original_text)
                                    cell_runs.append({
                                        **run,
                                        'original_text': original_text,
                                        'text': escaped_text,
                                        'para_elem': para,
                                        'host_para_elem': para,
                                        'host_para_id': para_id,
                                        'cell_elem': tc,
                                        'row_elem': tr
                                    })
                                    cell_text += escaped_text
                            
                                if is_end_para(para):
                                    break
                        
                            # Store in vmerge_content for future continue cells
                            vmerge_content[grid_col] = {
                                'runs': cell_runs,
                                'text': cell_text
                            }
                    
                        elif vmerge_type == 'continue':
                            # Merge continue: copy from vmerge_content if available
                            if grid_col in vmerge_content:
                                # Deep copy runs to avoid reference issues
                                cell_runs = [dict(r) for r in vmerge_content[grid_col]['runs']]
                                cell_text = vmerge_content[grid_col]['text']
                                # Update row_elem to current row
                                for run in cell_runs:
                                    run['row_elem'] = tr
                                    # Override host para to current row cell para
                                    if cell_paras:
                                        host_para = cell_paras[0]
                                        host_para_id = host_para.get('{http://schemas.microsoft.com/office/word/2010/wordml}paraId')
                                        run['host_para_elem'] = host_para
                                        run['host_para_id'] = host_para_id
                            # else: restart outside range, keep empty
                    
                        else:
                            # Normal cell (no vMerge)
                            for para_idx, para in enumerate(cell_paras):
                                para_id = para.get('{http://schemas.microsoft.com/office/word/2010/wordml}paraId')
                                if not in_range and para is not start_para:
                                    continue
                                if para is start_para:
                                    in_range = True
                            
                                if para_idx > 0:
                                    cell_runs.append({
                                        'text': '\\n',
                                        'is_json_escape': True,
                                        'is_para_boundary': True,
                                        'para_elem': para,
                                        'cell_elem': tc,
                                        'row_elem': tr
                                    })
                                    cell_text += '\\n'
                            
                                para_runs, para_orig_text = self._collect_runs_info_original(para)
                                # Strip paragraph text to match table_extractor.py behavior (line 144/163)
                                # table_extractor.py calls para_text.strip() before checking if empty
                                para_orig_text = para_orig_text.strip()
                            
                                # Skip empty paragraphs (match table_extractor.py behavior at line 144/163)
                                # table_extractor.py strips each paragraph and skips if empty
                                if not para_orig_text:
                                    continue
                            
                                para_runs = self._strip_runs_whitespace(para_runs)
                                for run in para_runs:
                                    original_text = run['text']
                                    escaped_text = json_escape(original_text)
                                    cell_runs.append({
                                        **run,
                                        'original_text': original_text,
                                        'text': escaped_text,
                                        'para_elem': para,
                                        'host_para_elem': para,
                                        'host_para_id': para_id,
                                        'cell_elem': tc,
                                        'row_elem': tr
                                    })
                                    cell_text += escaped_text
                            
                                if is_end_para(para):
                                    break
                        
                            # Check if empty cell inherits from vMerge (TableExtractor logic)
                            if not cell_text and grid_col in vmerge_content:
                                cell_runs = [dict(r) for r in vmerge_content[grid_col]['runs']]
                                cell_text = vmerge_content[grid_col]['text']
                                for run in cell_runs:
                                    run['row_elem'] = tr
                                    # Override host para to current row cell para
                                    if cell_paras:
                                        host_para = cell_paras[0]
                                        host_para_id = host_para.get('{http://schemas.microsoft.com/office/word/2010/wordml}paraId')
                                        run['host_para_elem'] = host_para
                                        run['host_para_id'] = host_para_id
                            elif cell_text:
                                # Non-empty normal cell ends vMerge region
                                vmerge_content.pop(grid_col, None)
                
                    # Store cell data at grid_col position (matching TableExtractor behavior)
                    # gridSpan cells only occupy the starting position, skipped columns remain None
                    if grid_col < num_cols and (cell_in_range or cell_runs):
                        row_data[grid_col] = (cell_runs, cell_text)
                
                    # Advance grid column by gridSpan
                    grid_col += grid_span
                
                    # Check if we've reached the end - stop immediately when uuid_end is found
                    if cell_in_range:
                        for para in cell_paras:
                            if is_end_para(para):
                                reached_end = True
                                break
                
                    # If we've reached the end, stop processing subsequent cells in this row
                    if reached_end:
                        break
            
                # After processing all cells in row, add to runs_info if row is in range
                if in_range and any(cell_tuple is not None and cell_tuple[0] for cell_tuple in row_data):
                    # Add row opening marker (first row gets '["', subsequent get '"], ["')
                    if first_row_in_range:
                        runs_info.append({
                            'text': '["',
                            'start': pos,
                            'end': pos + 2,
                            'para_elem': last_para,
                            'is_json_boundary': True
                        })
                        pos += 2
                        first_row_in_range = False
                    else:
                        runs_info.append({
                            'text': '"], ["',
                            'start': pos,
                            'end': pos + 6,
                            'para_elem': last_para,
                            'is_json_boundary': True,
                            'is_row_boundary': True
                        })
                        pos += 6
                
                    # Output cells based on cols_in_range to preserve gridSpan behavior
                    # For each column in range:
                    #   - If row_data[col] has data: output cell content
                    #   - If row_data[col] is None but col in cols_in_range: output "" (gridSpan placeholder)
                    #   - If col not in cols_in_range: skip (out of range)
                    output_col_count = 0
                    for col_idx in range(num_cols):
                        if col_idx not in cols_in_range:
                            continue  # Skip columns outside uuid range
                    
                        if output_col_count > 0:
                            # Add cell boundary marker
                            runs_info.append({
                                'text': '", "',
                                'start': pos,
                                'end': pos + 4,
                                'para_elem': last_para,
                                'is_json_boundary': True,
                                'is_cell_boundary': True
                            })
                            pos += 4
                    
                        if row_data[col_idx] is not None:
                            # Cell has content - output runs
                            cell_runs, cell_text = row_data[col_idx]
                            for run in cell_runs:
                                run['start'] = pos
                                run['end'] = pos + len(run['text'])
                                runs_info.append(run)
                                pos += len(run['text'])
                        # else: gridSpan placeholder - output nothing (empty string between ", ")
                    
                        output_col_count += 1
            
        
            # Add closing '"]'
            if in_range:
                runs_info.append({
                    'text': '"]',
                    'start': pos,
                    'end': pos + 2,
                    'para_elem': last_para,
                    'is_json_boundary': True
                })
                pos += 2
        
            combined_text = ''.join(r['text'] for r in runs_info)
            is_cross_paragraph = True  # Table mode is always treated as cross-paragraph
        
            return runs_info, combined_text, is_cross_paragraph, None

        def _collect_runs_info_in_table_legacy(self, start_para, uuid_end: str, table_elem) -> Tuple[List[Dict], str, bool, Optional[str]]:
            """
            Legacy table collection method (paragraph iteration without merge handling).
        
            Used as fallback when tblGrid is not available.
            """
            end_para = self._resolve_end_para(start_para, uuid_end)

            def is_end_para(para) -> bool:
                if end_para is not None:
                    return para is end_para
                para_id = para.get('{http://schemas.microsoft.com/office/word/2010/wordml}paraId')
                return para_id == uuid_end

            runs_info = []
            pos = 0
            current_cell = None
            current_row = None
            in_range = False
            last_para = start_para

            # Add opening '["'
            runs_info.append({
                'text': '["',
                'start': pos,
                'end': pos + 2,
                'para_elem': start_para,
                'is_json_boundary': True
            })
            pos += 2

            # Iterate all paragraphs in the table (in document order)
            for para in table_elem.iter(f'{{{NS["w"]}}}p'):
                # Check if we've entered the range
                if para is start_para:
                    in_range = True

                if not in_range:
                    continue

                last_para = para

                # Get cell and row for this paragraph
                cell = self._find_ancestor_cell(para)
                row = self._find_ancestor_row(para)

                # Handle row transition
                if row is not current_row:
                    if current_row is not None:
                        # New row: close previous row and open new row '"], ["'
                        runs_info.append({
                            'text': '"], ["',
                            'start': pos,
                            'end': pos + 6,
                            'para_elem': para,
                            'is_json_boundary': True,
                            'is_row_boundary': True
                        })
                        pos += 6
                        current_cell = None  # Reset cell tracking for new row
                    current_row = row

                # Handle cell transition (within same row)
                if cell is not current_cell:
                    if current_cell is not None:
                        # New cell in same row: add '", "'
                        runs_info.append({
                            'text': '", "',
                            'start': pos,
                            'end': pos + 4,
                            'para_elem': para,
                            'is_json_boundary': True,
                            'is_cell_boundary': True
                        })
                        pos += 4
                    current_cell = cell
                elif current_cell is not None:
                    # Same cell, new paragraph: add '\n' (which is \\n in JSON)
                    runs_info.append({
                        'text': '\\n',
                        'start': pos,
                        'end': pos + 2,
                        'para_elem': para,
                        'is_para_boundary': True,
                        'is_json_escape': True
                    })
                    pos += 2

                # Collect paragraph content (with JSON escaping)
                para_runs, _ = self._collect_runs_info_original(para)
                para_runs = self._strip_runs_whitespace(para_runs)

                for run in para_runs:
                    original_text = run['text']
                    escaped_text = json_escape(original_text)

                    run['original_text'] = original_text
                    run['text'] = escaped_text
                    run['para_elem'] = para
                    run['host_para_elem'] = para
                    run['host_para_id'] = para.get('{http://schemas.microsoft.com/office/word/2010/wordml}paraId')
                    run['cell_elem'] = cell
                    run['row_elem'] = row
                    run['start'] = pos
                    run['end'] = pos + len(escaped_text)
                    runs_info.append(run)

                    pos += len(escaped_text)

                # Check if we've reached the end
                if is_end_para(para):
                    break

            # Add closing '"]'
            runs_info.append({
                'text': '"]',
                'start': pos,
                'end': pos + 2,
                'para_elem': last_para,
                'is_json_boundary': True
            })
            pos += 2

            combined_text = ''.join(r['text'] for r in runs_info)
            is_cross_paragraph = True  # Table mode is always treated as cross-paragraph

            return runs_info, combined_text, is_cross_paragraph, None

        def _extract_text_in_range_from_table(self, table_elem, uuid_start: str, uuid_end: str) -> str:
            """
            Extract table text between uuid_start and uuid_end in JSON row format.

            Returns a string like:
              ["cell1", "cell2"], ["cell3", "cell4"]

            This mirrors parse_document-style row formatting and handles gridSpan/vMerge:
            - gridSpan: repeats cell content across spanned columns
            - vMerge restart/continue: repeats content only when restart is within range
            """
            rows_text = []
            in_range = False
            reached_end = False
            vmerge_content = {}  # {grid_col: text}
            end_para = self._find_last_para_with_id_in_table(table_elem, uuid_end)

            for tr in table_elem.findall(f'{{{NS["w"]}}}tr'):
                if reached_end:
                    break
                row_cells = []
                row_in_range = False
                grid_col = 0

                for tc in tr.findall(f'{{{NS["w"]}}}tc'):
                    if reached_end:
                        break

                    tcPr = tc.find(f'{{{NS["w"]}}}tcPr')
                    grid_span, vmerge_type = self._get_cell_merge_properties(tcPr)
                    cell_paras = tc.findall(f'{{{NS["w"]}}}p')

                    cell_text_parts = []
                    cell_has_range = False

                    for para in cell_paras:
                        para_id = para.get('{http://schemas.microsoft.com/office/word/2010/wordml}paraId')
                        if para_id == uuid_start:
                            in_range = True

                        if in_range:
                            cell_has_range = True
                            _, para_text = self._collect_runs_info_original(para)
                            if para_text:
                                cell_text_parts.append(para_text)

                        if in_range:
                            if end_para is not None:
                                if para is end_para:
                                    reached_end = True
                                    break
                            elif para_id == uuid_end:
                                reached_end = True
                                break

                    cell_text = '\n'.join(cell_text_parts).replace('\x07', '')

                    if vmerge_type == 'restart':
                        if cell_has_range:
                            for col in range(grid_col, grid_col + grid_span):
                                vmerge_content[col] = cell_text
                    elif vmerge_type == 'continue':
                        if grid_col in vmerge_content:
                            cell_text = vmerge_content.get(grid_col, '')
                    else:
                        if not cell_text and grid_col in vmerge_content:
                            cell_text = vmerge_content.get(grid_col, '')
                        elif cell_text:
                            for col in range(grid_col, grid_col + grid_span):
                                vmerge_content.pop(col, None)

                    if in_range or cell_has_range:
                        row_in_range = True
                        escaped = json_escape(cell_text)
                        for _ in range(grid_span):
                            row_cells.append(escaped)

                    grid_col += grid_span

                if row_in_range:
                    rows_text.append('["' + '", "'.join(row_cells) + '"]')

            return ', '.join(rows_text)

        def _check_cross_cell_boundary(self, affected_runs: List[Dict]) -> bool:
            """
            Check if affected runs span multiple table cells.

            Args:
                affected_runs: List of run info dicts from _find_affected_runs

            Returns:
                True if runs span multiple cells, False otherwise
            """
            cells = set()
            for info in affected_runs:
                cell = info.get('cell_elem')
                if cell is not None:
                    cells.add(id(cell))
            return len(cells) > 1

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
