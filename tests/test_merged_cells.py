#!/usr/bin/env python3
"""
Unit tests for merged cell handling in apply_audit_edits.py

Tests verify that vMerge (vertical merge) and gridSpan (horizontal merge)
are handled consistently with TableExtractor.
"""

import sys
from pathlib import Path
from lxml import etree

# Add parent directory to path for imports
sys.path.insert(0, str(Path(__file__).parent.parent / 'skills' / 'doc-audit' / 'scripts'))

from apply_audit_edits import AuditEditApplier, NS  # type: ignore


def create_table_with_vmerge(para_ids):
    """
    Create a table with vertically merged cells.
    
    Structure:
      Col0    | Col1    | Col2
      --------+---------+--------
      Cell A  | Merged1 | Data1   (restart)
      Cell B  | Merged1 | Data2   (continue)
      Cell C  | Merged1 | Data3   (continue)
    
    Args:
        para_ids: List of paragraph IDs [p1, p2, p3, ...]
    
    Returns:
        Table element (w:tbl)
    """
    table_xml = f'''<w:tbl xmlns:w="{NS['w']}" xmlns:w14="{NS['w14']}">
        <w:tblGrid>
            <w:gridCol/>
            <w:gridCol/>
            <w:gridCol/>
        </w:tblGrid>
        <w:tr>
            <w:tc>
                <w:p w14:paraId="{para_ids[0]}">
                    <w:r><w:t>Cell A</w:t></w:r>
                </w:p>
            </w:tc>
            <w:tc>
                <w:tcPr>
                    <w:vMerge w:val="restart"/>
                </w:tcPr>
                <w:p w14:paraId="{para_ids[1]}">
                    <w:r><w:t>Merged1</w:t></w:r>
                </w:p>
            </w:tc>
            <w:tc>
                <w:p w14:paraId="{para_ids[2]}">
                    <w:r><w:t>Data1</w:t></w:r>
                </w:p>
            </w:tc>
        </w:tr>
        <w:tr>
            <w:tc>
                <w:p w14:paraId="{para_ids[3]}">
                    <w:r><w:t>Cell B</w:t></w:r>
                </w:p>
            </w:tc>
            <w:tc>
                <w:tcPr>
                    <w:vMerge/>
                </w:tcPr>
                <w:p w14:paraId="{para_ids[4]}">
                    <w:r><w:t></w:t></w:r>
                </w:p>
            </w:tc>
            <w:tc>
                <w:p w14:paraId="{para_ids[5]}">
                    <w:r><w:t>Data2</w:t></w:r>
                </w:p>
            </w:tc>
        </w:tr>
        <w:tr>
            <w:tc>
                <w:p w14:paraId="{para_ids[6]}">
                    <w:r><w:t>Cell C</w:t></w:r>
                </w:p>
            </w:tc>
            <w:tc>
                <w:tcPr>
                    <w:vMerge/>
                </w:tcPr>
                <w:p w14:paraId="{para_ids[7]}">
                    <w:r><w:t></w:t></w:r>
                </w:p>
            </w:tc>
            <w:tc>
                <w:p w14:paraId="{para_ids[8]}">
                    <w:r><w:t>Data3</w:t></w:r>
                </w:p>
            </w:tc>
        </w:tr>
    </w:tbl>'''
    
    return etree.fromstring(table_xml)


def create_table_with_gridspan(para_ids):
    """
    Create a table with horizontally merged cells (gridSpan).
    
    Structure:
      Col0         | Col1  | Col2
      -------------+-------+-------
      Wide Cell (spans 2)  | Data1
      Cell A       | Cell B| Data2
    
    Args:
        para_ids: List of paragraph IDs
    
    Returns:
        Table element (w:tbl)
    """
    table_xml = f'''<w:tbl xmlns:w="{NS['w']}" xmlns:w14="{NS['w14']}">
        <w:tblGrid>
            <w:gridCol/>
            <w:gridCol/>
            <w:gridCol/>
        </w:tblGrid>
        <w:tr>
            <w:tc>
                <w:tcPr>
                    <w:gridSpan w:val="2"/>
                </w:tcPr>
                <w:p w14:paraId="{para_ids[0]}">
                    <w:r><w:t>Wide Cell</w:t></w:r>
                </w:p>
            </w:tc>
            <w:tc>
                <w:p w14:paraId="{para_ids[1]}">
                    <w:r><w:t>Data1</w:t></w:r>
                </w:p>
            </w:tc>
        </w:tr>
        <w:tr>
            <w:tc>
                <w:p w14:paraId="{para_ids[2]}">
                    <w:r><w:t>Cell A</w:t></w:r>
                </w:p>
            </w:tc>
            <w:tc>
                <w:p w14:paraId="{para_ids[3]}">
                    <w:r><w:t>Cell B</w:t></w:r>
                </w:p>
            </w:tc>
            <w:tc>
                <w:p w14:paraId="{para_ids[4]}">
                    <w:r><w:t>Data2</w:t></w:r>
                </w:p>
            </w:tc>
        </w:tr>
    </w:tbl>'''
    
    return etree.fromstring(table_xml)


def test_vmerge_restart_in_range():
    """Test vMerge restart within uuid range - content should be repeated"""
    print("Testing vMerge restart in range...")
    
    # Create a mock applier instance
    class MockApplier:
        def __init__(self):
            pass
        
        def _xpath(self, elem, expr):
            return elem.xpath(expr, namespaces=NS)
        
        def _find_ancestor_cell(self, para_elem):
            parent = para_elem.getparent()
            while parent is not None:
                if parent.tag == f'{{{NS["w"]}}}tc':
                    return parent
                parent = parent.getparent()
            return None
        
        def _find_ancestor_row(self, para_elem):
            parent = para_elem.getparent()
            while parent is not None:
                if parent.tag == f'{{{NS["w"]}}}tr':
                    return parent
                parent = parent.getparent()
            return None
        
        def _get_cell_merge_properties(self, tcPr):
            return AuditEditApplier._get_cell_merge_properties(self, tcPr)
        
        def _collect_runs_info_original(self, para_elem):
            # Simple text extraction for testing
            text = ''.join(t.text or '' for t in para_elem.findall('.//w:t', NS))
            if not text:
                return [], ''
            return [{'text': text, 'start': 0, 'end': len(text), 'elem': None, 'rPr': None}], text
        
        def _collect_runs_info_in_table(self, start_para, uuid_end, table_elem):
            return AuditEditApplier._collect_runs_info_in_table(
                self, start_para, uuid_end, table_elem
            )
    
    # Create table with vMerge
    para_ids = [f"ID{i:02d}" for i in range(9)]
    table = create_table_with_vmerge(para_ids)
    
    # Find start paragraph (first row, first cell)
    start_para = table.xpath(f'.//w:p[@w14:paraId="{para_ids[0]}"]', namespaces=NS)[0]
    uuid_end = para_ids[8]  # Last paragraph
    
    applier = MockApplier()
    runs_info, combined_text, is_cross_para, boundary_error = applier._collect_runs_info_in_table(
        start_para, uuid_end, table
    )
    
    # Expected output: ["Cell A", "Merged1", "Data1"], ["Cell B", "Merged1", "Data2"], ["Cell C", "Merged1", "Data3"]
    expected_text = '["Cell A", "Merged1", "Data1"], ["Cell B", "Merged1", "Data2"], ["Cell C", "Merged1", "Data3"]'
    
    assert combined_text == expected_text, f"Expected:\n{expected_text}\nGot:\n{combined_text}"
    assert is_cross_para
    assert boundary_error is None
    
    print("✓ vMerge restart test passed")


def test_vmerge_continue_out_of_range():
    """Test vMerge continue when restart is outside uuid range - should be empty"""
    print("Testing vMerge continue out of range...")
    
    class MockApplier:
        def __init__(self):
            pass
        
        def _xpath(self, elem, expr):
            return elem.xpath(expr, namespaces=NS)
        
        def _find_ancestor_cell(self, para_elem):
            parent = para_elem.getparent()
            while parent is not None:
                if parent.tag == f'{{{NS["w"]}}}tc':
                    return parent
                parent = parent.getparent()
            return None
        
        def _find_ancestor_row(self, para_elem):
            parent = para_elem.getparent()
            while parent is not None:
                if parent.tag == f'{{{NS["w"]}}}tr':
                    return parent
                parent = parent.getparent()
            return None
        
        def _get_cell_merge_properties(self, tcPr):
            return AuditEditApplier._get_cell_merge_properties(self, tcPr)
        
        def _collect_runs_info_original(self, para_elem):
            text = ''.join(t.text or '' for t in para_elem.findall('.//w:t', NS))
            if not text:
                return [], ''
            return [{'text': text, 'start': 0, 'end': len(text), 'elem': None, 'rPr': None}], text
        
        def _collect_runs_info_in_table(self, start_para, uuid_end, table_elem):
            return AuditEditApplier._collect_runs_info_in_table(
                self, start_para, uuid_end, table_elem
            )
    
    # Create table with vMerge
    para_ids = [f"ID{i:02d}" for i in range(9)]
    table = create_table_with_vmerge(para_ids)
    
    # Start from second row (vMerge restart is in first row, outside range)
    start_para = table.xpath(f'.//w:p[@w14:paraId="{para_ids[3]}"]', namespaces=NS)[0]
    uuid_end = para_ids[8]
    
    applier = MockApplier()
    runs_info, combined_text, is_cross_para, boundary_error = applier._collect_runs_info_in_table(
        start_para, uuid_end, table
    )
    
    # Expected: Second column should be empty since restart is outside range
    expected_text = '["Cell B", "", "Data2"], ["Cell C", "", "Data3"]'
    
    assert combined_text == expected_text, f"Expected:\n{expected_text}\nGot:\n{combined_text}"
    
    print("✓ vMerge out of range test passed")


def test_gridspan_handling():
    """Test gridSpan (horizontal merge) handling"""
    print("Testing gridSpan handling...")
    
    class MockApplier:
        def __init__(self):
            pass
        
        def _xpath(self, elem, expr):
            return elem.xpath(expr, namespaces=NS)
        
        def _find_ancestor_cell(self, para_elem):
            parent = para_elem.getparent()
            while parent is not None:
                if parent.tag == f'{{{NS["w"]}}}tc':
                    return parent
                parent = parent.getparent()
            return None
        
        def _find_ancestor_row(self, para_elem):
            parent = para_elem.getparent()
            while parent is not None:
                if parent.tag == f'{{{NS["w"]}}}tr':
                    return parent
                parent = parent.getparent()
            return None
        
        def _get_cell_merge_properties(self, tcPr):
            return AuditEditApplier._get_cell_merge_properties(self, tcPr)
        
        def _collect_runs_info_original(self, para_elem):
            text = ''.join(t.text or '' for t in para_elem.findall('.//w:t', NS))
            if not text:
                return [], ''
            return [{'text': text, 'start': 0, 'end': len(text), 'elem': None, 'rPr': None}], text
        
        def _collect_runs_info_in_table(self, start_para, uuid_end, table_elem):
            return AuditEditApplier._collect_runs_info_in_table(
                self, start_para, uuid_end, table_elem
            )
    
    # Create table with gridSpan
    para_ids = [f"ID{i:02d}" for i in range(5)]
    table = create_table_with_gridspan(para_ids)
    
    start_para = table.xpath(f'.//w:p[@w14:paraId="{para_ids[0]}"]', namespaces=NS)[0]
    uuid_end = para_ids[4]
    
    applier = MockApplier()
    runs_info, combined_text, is_cross_para, boundary_error = applier._collect_runs_info_in_table(
        start_para, uuid_end, table
    )
    
    # Expected: First row has wide cell spanning 2 columns, then Data1
    # Second row has Cell A, Cell B, Data2
    expected_text = '["Wide Cell", "Data1"], ["Cell A", "Cell B", "Data2"]'
    
    assert combined_text == expected_text, f"Expected:\n{expected_text}\nGot:\n{combined_text}"
    
    print("✓ gridSpan test passed")


def test_get_cell_merge_properties():
    """Test _get_cell_merge_properties helper method"""
    print("Testing _get_cell_merge_properties...")
    
    applier = AuditEditApplier.__new__(AuditEditApplier)
    
    # Test case 1: vMerge restart with gridSpan
    tcPr_xml = f'''<w:tcPr xmlns:w="{NS['w']}">
        <w:gridSpan w:val="2"/>
        <w:vMerge w:val="restart"/>
    </w:tcPr>'''
    tcPr = etree.fromstring(tcPr_xml)
    
    grid_span, vmerge_type = applier._get_cell_merge_properties(tcPr)
    assert grid_span == 2, f"Expected gridSpan=2, got {grid_span}"
    assert vmerge_type == 'restart', f"Expected vmerge_type='restart', got {vmerge_type}"
    
    # Test case 2: vMerge continue (no val attribute)
    tcPr_xml = f'''<w:tcPr xmlns:w="{NS['w']}">
        <w:vMerge/>
    </w:tcPr>'''
    tcPr = etree.fromstring(tcPr_xml)
    
    grid_span, vmerge_type = applier._get_cell_merge_properties(tcPr)
    assert grid_span == 1, f"Expected gridSpan=1, got {grid_span}"
    assert vmerge_type == 'continue', f"Expected vmerge_type='continue', got {vmerge_type}"
    
    # Test case 3: Normal cell (no merge)
    tcPr_xml = f'''<w:tcPr xmlns:w="{NS['w']}"/>'''
    tcPr = etree.fromstring(tcPr_xml)
    
    grid_span, vmerge_type = applier._get_cell_merge_properties(tcPr)
    assert grid_span == 1, f"Expected gridSpan=1, got {grid_span}"
    assert vmerge_type is None, f"Expected vmerge_type=None, got {vmerge_type}"
    
    # Test case 4: None tcPr
    grid_span, vmerge_type = applier._get_cell_merge_properties(None)
    assert grid_span == 1, f"Expected gridSpan=1, got {grid_span}"
    assert vmerge_type is None, f"Expected vmerge_type=None, got {vmerge_type}"
    
    print("✓ _get_cell_merge_properties test passed")


def run_all_tests():
    """Run all tests"""
    print("=" * 60)
    print("Running merged cell handling tests")
    print("=" * 60)
    
    try:
        test_get_cell_merge_properties()
        test_vmerge_restart_in_range()
        test_vmerge_continue_out_of_range()
        test_gridspan_handling()
        
        print("=" * 60)
        print("All tests passed! ✓")
        print("=" * 60)
        return 0
    except AssertionError as e:
        print(f"\n✗ Test failed: {e}")
        return 1
    except Exception as e:
        print(f"\n✗ Unexpected error: {e}")
        import traceback
        traceback.print_exc()
        return 1


if __name__ == '__main__':
    sys.exit(run_all_tests())
