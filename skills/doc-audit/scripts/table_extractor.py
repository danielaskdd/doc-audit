#!/usr/bin/env python3
"""
ABOUTME: Extracts tables from DOCX with proper merged cell handling
ABOUTME: Outputs 2D array with content in first cell of merged region
"""

from docx.table import Table
from docx.oxml.ns import qn
from typing import List

class TableExtractor:
    """
    Extract table content handling merged cells correctly.
    
    Merged cells in DOCX:
    - Horizontal: w:gridSpan specifies how many columns cell spans
    - Vertical: w:vMerge with val="restart" starts merge, subsequent cells continue
    
    Output format:
    - 2D list of strings
    - Merged cell content in top-left position only
    - Other positions in merged region are empty strings
    """
    
    @staticmethod
    def extract(table: Table, numbering_resolver=None) -> List[List[str]]:
        """
        Extract table to 2D string array.
        
        Args:
            table: python-docx Table object
            numbering_resolver: Optional NumberingResolver for extracting numbering
            
        Returns:
            List of rows, each row is list of cell text strings
        """
        tbl = table._tbl
        
        # Get number of columns from tblGrid
        tbl_grid = tbl.find(qn('w:tblGrid'))
        num_cols = 0
        if tbl_grid is not None:
            num_cols = len(tbl_grid.findall(qn('w:gridCol')))
        
        if num_cols == 0:
            return []
        
        # Process each row by directly iterating <w:tr> elements
        grid = []
        
        for tr in tbl.findall(qn('w:tr')):
            row_data = [''] * num_cols  # Pre-fill with empty strings
            grid_col = 0
            
            # Iterate actual <w:tc> elements (each physical cell appears once)
            for tc in tr.findall(qn('w:tc')):
                tcPr = tc.find(qn('w:tcPr'))
                
                # Check gridSpan (horizontal merge)
                grid_span = 1
                if tcPr is not None:
                    gs = tcPr.find(qn('w:gridSpan'))
                    if gs is not None:
                        grid_span = int(gs.get(qn('w:val')))
                
                # Check vMerge (vertical merge)
                vmerge_val = None
                if tcPr is not None:
                    vm = tcPr.find(qn('w:vMerge'))
                    if vm is not None:
                        vmerge_val = vm.get(qn('w:val'))  # 'restart' or None (means 'continue')
                
                # Only extract text if NOT a vMerge continuation
                cell_text = ''
                if vmerge_val != 'continue' and not (vmerge_val is None and tcPr is not None and tcPr.find(qn('w:vMerge')) is not None):
                    # Get cell text with numbering support
                    if numbering_resolver is not None:
                        # Extract text with numbering labels
                        cell_paragraphs = []
                        for para_elem in tc.findall(qn('w:p')):
                            # Get text content
                            para_text = ''
                            for t_elem in para_elem.findall('.//'+qn('w:t')):
                                if t_elem.text:
                                    para_text += t_elem.text
                            
                            # Get numbering label
                            label = numbering_resolver.get_label(para_elem)
                            
                            # Combine label and text
                            if label:
                                full_text = f"{label} {para_text}".strip()
                            else:
                                full_text = para_text.strip()
                            
                            if full_text:
                                cell_paragraphs.append(full_text)
                        
                        cell_text = '\n'.join(cell_paragraphs).replace('\x07', '')
                    else:
                        # Fallback to simple text extraction
                        # Cannot use cell.text here, must extract from XML
                        para_texts = []
                        for para_elem in tc.findall(qn('w:p')):
                            para_text = ''
                            for t_elem in para_elem.findall('.//'+qn('w:t')):
                                if t_elem.text:
                                    para_text += t_elem.text
                            if para_text:
                                para_texts.append(para_text.strip())
                        cell_text = '\n'.join(para_texts).replace('\x07', '')
                
                # Place content at starting grid position only
                if grid_col < num_cols:
                    row_data[grid_col] = cell_text
                
                # Move grid position by gridSpan
                grid_col += grid_span
            
            grid.append(row_data)
        
        return grid
