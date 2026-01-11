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
        result = TableExtractor.extract_with_metadata(table, numbering_resolver)
        return result['rows']
    
    @staticmethod
    def extract_with_metadata(table: Table, numbering_resolver=None) -> dict:
        """
        Extract table to 2D string array with metadata (paraIds, header info).
        
        Args:
            table: python-docx Table object
            numbering_resolver: Optional NumberingResolver for extracting numbering
            
        Returns:
            Dict with:
            - rows: 2D list of cell text strings
            - para_ids: 2D list of paraIds (first paraId in each cell, or None)
            - para_ids_end: 2D list of paraIds (last paraId in each cell, or None)
            - header_indices: List of row indices marked as table headers
        """
        tbl = table._tbl
        
        # Get number of columns from tblGrid
        tbl_grid = tbl.find(qn('w:tblGrid'))
        num_cols = 0
        if tbl_grid is not None:
            num_cols = len(tbl_grid.findall(qn('w:gridCol')))
        
        if num_cols == 0:
            return {'rows': [], 'para_ids': [], 'para_ids_end': [], 'header_indices': []}
        
        # Detect header rows using w:tblHeader attribute
        header_indices = []
        for idx, tr in enumerate(tbl.findall(qn('w:tr'))):
            trPr = tr.find(qn('w:trPr'))
            if trPr is not None:
                tbl_header = trPr.find(qn('w:tblHeader'))
                if tbl_header is not None:
                    header_indices.append(idx)
        
        # Process each row by directly iterating <w:tr> elements
        grid = []
        para_ids_grid = []
        para_ids_end_grid = []  # Track last paraId in each cell
        
        for tr in tbl.findall(qn('w:tr')):
            row_data = [''] * num_cols  # Pre-fill with empty strings
            row_para_ids = [None] * num_cols  # Pre-fill with None
            row_para_ids_end = [None] * num_cols  # Pre-fill with None for last paraId
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
                cell_para_id = None
                cell_para_id_end = None  # Track last paraId in cell
                if vmerge_val != 'continue' and not (vmerge_val is None and tcPr is not None and tcPr.find(qn('w:vMerge')) is not None):
                    # Get cell text with numbering support
                    if numbering_resolver is not None:
                        # Extract text with numbering labels
                        cell_paragraphs = []
                        for para_elem in tc.findall(qn('w:p')):
                            # Capture paraId from each paragraph
                            para_id_attr = para_elem.get('{http://schemas.microsoft.com/office/word/2010/wordml}paraId')
                            if para_id_attr:
                                if cell_para_id is None:
                                    cell_para_id = para_id_attr  # First paraId
                                cell_para_id_end = para_id_attr  # Always update to get last
                            
                            # Get text content (including line breaks)
                            para_text = ''
                            for run in para_elem.findall('.//'+qn('w:r')):
                                for child in run:
                                    tag = child.tag.split('}')[-1]
                                    if tag == 't' and child.text:
                                        para_text += child.text
                                    elif tag == 'br':
                                        # Handle line breaks - textWrapping or no type = soft line break
                                        br_type = child.get(qn('w:type'))
                                        if br_type in (None, 'textWrapping'):
                                            para_text += '\n'
                                        # Skip page and column breaks (layout elements)
                            
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
                        # Fallback to simple text extraction (including line breaks)
                        # Cannot use cell.text here, must extract from XML
                        para_texts = []
                        for para_elem in tc.findall(qn('w:p')):
                            # Capture paraId from each paragraph
                            para_id_attr = para_elem.get('{http://schemas.microsoft.com/office/word/2010/wordml}paraId')
                            if para_id_attr:
                                if cell_para_id is None:
                                    cell_para_id = para_id_attr  # First paraId
                                cell_para_id_end = para_id_attr  # Always update to get last
                            
                            para_text = ''
                            for run in para_elem.findall('.//'+qn('w:r')):
                                for child in run:
                                    tag = child.tag.split('}')[-1]
                                    if tag == 't' and child.text:
                                        para_text += child.text
                                    elif tag == 'br':
                                        # Handle line breaks - textWrapping or no type = soft line break
                                        br_type = child.get(qn('w:type'))
                                        if br_type in (None, 'textWrapping'):
                                            para_text += '\n'
                                        # Skip page and column breaks (layout elements)
                            if para_text:
                                para_texts.append(para_text.strip())
                        cell_text = '\n'.join(para_texts).replace('\x07', '')
                
                # Place content at starting grid position only
                if grid_col < num_cols:
                    row_data[grid_col] = cell_text
                    row_para_ids[grid_col] = cell_para_id
                    row_para_ids_end[grid_col] = cell_para_id_end
                
                # Move grid position by gridSpan
                grid_col += grid_span
            
            grid.append(row_data)
            para_ids_grid.append(row_para_ids)
            para_ids_end_grid.append(row_para_ids_end)
        
        return {
            'rows': grid,
            'para_ids': para_ids_grid,
            'para_ids_end': para_ids_end_grid,
            'header_indices': header_indices
        }
