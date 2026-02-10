#!/usr/bin/env python3
"""
ABOUTME: Unit tests for apply_audit_edits.py
"""

import sys
from pathlib import Path

TESTS_DIR = Path(__file__).parent
sys.path.insert(0, str(TESTS_DIR))

import pytest  # noqa: E402
from lxml import etree  # noqa: E402

from _apply_audit_edits_helpers import (  # noqa: E402
    apply_module, NS, NSMAP,
    create_paragraph_xml,
    create_mock_applier,
    create_mock_body_with_paragraphs,
    create_table_with_cells, create_multi_row_table,
    create_table_with_cells_with_br,
)


def create_table_with_vmerge_continue(para_ids):
    """Create a 2-row table with a vMerge continue cell in row 2, col 2."""
    table_xml = f'''<w:tbl xmlns:w="{NS['w']}" xmlns:w14="{NS['w14']}">
        <w:tblGrid>
            <w:gridCol/>
            <w:gridCol/>
            <w:gridCol/>
        </w:tblGrid>
        <w:tr>
            <w:tc>
                <w:p w14:paraId="{para_ids[0]}">
                    <w:r><w:t>R1C1</w:t></w:r>
                </w:p>
            </w:tc>
            <w:tc>
                <w:tcPr>
                    <w:vMerge w:val="restart"/>
                </w:tcPr>
                <w:p w14:paraId="{para_ids[1]}">
                    <w:r><w:t>Merged</w:t></w:r>
                </w:p>
            </w:tc>
            <w:tc>
                <w:p w14:paraId="{para_ids[2]}">
                    <w:r><w:t>R1C3</w:t></w:r>
                </w:p>
            </w:tc>
        </w:tr>
        <w:tr>
            <w:tc>
                <w:p w14:paraId="{para_ids[3]}">
                    <w:r><w:t>R2C1</w:t></w:r>
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
                    <w:r><w:t>R2C3</w:t></w:r>
                </w:p>
            </w:tc>
        </w:tr>
    </w:tbl>'''
    return etree.fromstring(table_xml)


class TestTableDetection:
    """Tests for table detection helper methods"""

    def test_is_paragraph_in_table(self):
        """Paragraph inside table should be detected"""
        applier = create_mock_applier()

        # Create table with one cell
        tbl = create_table_with_cells([['Cell content']], ['AAA'])
        body = etree.Element(f'{{{NS["w"]}}}body', nsmap=NSMAP)
        body.append(tbl)
        applier.body_elem = body

        # Find paragraph
        para = body.find('.//w:p', NSMAP)
        assert applier._is_paragraph_in_table(para) is True

    def test_is_paragraph_not_in_table(self):
        """Paragraph outside table should not be detected as in table"""
        applier = create_mock_applier()

        body = create_mock_body_with_paragraphs(['AAA'])
        applier.body_elem = body

        para = body.find('.//w:p', NSMAP)
        assert applier._is_paragraph_in_table(para) is False

    def test_find_ancestor_cell(self):
        """Should find the cell containing a paragraph"""
        applier = create_mock_applier()

        tbl = create_table_with_cells([['Cell1'], ['Cell2']], ['AAA', 'BBB'])
        body = etree.Element(f'{{{NS["w"]}}}body', nsmap=NSMAP)
        body.append(tbl)
        applier.body_elem = body

        para = body.find('.//w:p[@w14:paraId="AAA"]', NSMAP)
        cell = applier._find_ancestor_cell(para)
        assert cell is not None
        assert cell.tag == f'{{{NS["w"]}}}tc'

    def test_find_ancestor_row(self):
        """Should find the row containing a paragraph"""
        applier = create_mock_applier()

        tbl = create_table_with_cells([['Cell1'], ['Cell2']], ['AAA', 'BBB'])
        body = etree.Element(f'{{{NS["w"]}}}body', nsmap=NSMAP)
        body.append(tbl)
        applier.body_elem = body

        para = body.find('.//w:p[@w14:paraId="AAA"]', NSMAP)
        row = applier._find_ancestor_row(para)
        assert row is not None
        assert row.tag == f'{{{NS["w"]}}}tr'



class TestTableBoundaryDetection:
    """Tests for boundary detection in _collect_runs_info_across_paragraphs"""

    def test_body_to_table_boundary_crossed(self):
        """Crossing from body to table should return boundary_crossed error"""
        applier = create_mock_applier()

        # Create body with paragraph followed by table
        body = etree.Element(f'{{{NS["w"]}}}body', nsmap=NSMAP)

        # Add body paragraph
        p = etree.SubElement(body, f'{{{NS["w"]}}}p')
        p.set(f'{{{NS["w14"]}}}paraId', 'AAA')
        r = etree.SubElement(p, f'{{{NS["w"]}}}r')
        t = etree.SubElement(r, f'{{{NS["w"]}}}t')
        t.text = 'Body text'

        # Add table
        tbl = create_table_with_cells([['Table cell']], ['BBB'])
        body.append(tbl)

        applier.body_elem = body

        start_para = body.find('.//w:p[@w14:paraId="AAA"]', NSMAP)

        runs_info, combined_text, is_cross_para, boundary_error = applier._collect_runs_info_across_paragraphs(
            start_para, 'BBB'
        )

        assert boundary_error == 'boundary_crossed'

    def test_table_to_body_boundary_crossed(self):
        """Crossing from table to body should return boundary_crossed error"""
        applier = create_mock_applier()

        # Create body with table followed by body paragraph
        body = etree.Element(f'{{{NS["w"]}}}body', nsmap=NSMAP)

        # Add table first
        tbl = create_table_with_cells([['Table cell']], ['AAA'])
        body.append(tbl)

        # Add body paragraph after table
        p = etree.SubElement(body, f'{{{NS["w"]}}}p')
        p.set(f'{{{NS["w14"]}}}paraId', 'BBB')
        r = etree.SubElement(p, f'{{{NS["w"]}}}r')
        t = etree.SubElement(r, f'{{{NS["w"]}}}t')
        t.text = 'Body text after table'

        applier.body_elem = body

        start_para = body.find('.//w:p[@w14:paraId="AAA"]', NSMAP)

        runs_info, combined_text, is_cross_para, boundary_error = applier._collect_runs_info_across_paragraphs(
            start_para, 'BBB'
        )

        assert boundary_error == 'boundary_crossed'

    def test_table_row_boundary_collected(self):
        """Multi-row table content should be collected (row boundary checked later)"""
        applier = create_mock_applier()

        # Create table with two rows
        tbl = create_multi_row_table(
            [[['Row1 Cell1'], ['Row1 Cell2']], [['Row2 Cell1'], ['Row2 Cell2']]],
            ['AAA', 'BBB', 'CCC', 'DDD']
        )
        body = etree.Element(f'{{{NS["w"]}}}body', nsmap=NSMAP)
        body.append(tbl)
        applier.body_elem = body

        start_para = body.find('.//w:p[@w14:paraId="AAA"]', NSMAP)

        # Content from row 1 to row 2 should be collected (no upfront error)
        runs_info, combined_text, is_cross_para, boundary_error = applier._collect_runs_info_across_paragraphs(
            start_para, 'CCC'
        )

        # No upfront boundary error - content is collected across rows
        assert boundary_error is None
        # Content from both rows should be present
        assert 'Row1 Cell1' in combined_text
        assert 'Row2 Cell1' in combined_text
        # Row boundary marker should be present
        assert '"], ["' in combined_text

        # _check_cross_row_boundary should detect if a match spans rows
        # Get runs that span from row 1 to row 2
        all_affected = applier._find_affected_runs(runs_info, 0, len(combined_text))
        real_runs = [r for r in all_affected
                     if not r.get('is_json_boundary', False)
                     and not r.get('is_json_escape', False)]
        assert applier._check_cross_row_boundary(real_runs) is True

    def test_same_row_no_boundary_error(self):
        """Same row content should not trigger boundary error"""
        applier = create_mock_applier()

        # Create table with one row, multiple cells
        tbl = create_table_with_cells([['Cell1'], ['Cell2'], ['Cell3']], ['AAA', 'BBB', 'CCC'])
        body = etree.Element(f'{{{NS["w"]}}}body', nsmap=NSMAP)
        body.append(tbl)
        applier.body_elem = body

        start_para = body.find('.//w:p[@w14:paraId="AAA"]', NSMAP)

        runs_info, combined_text, is_cross_para, boundary_error = applier._collect_runs_info_across_paragraphs(
            start_para, 'CCC'
        )

        assert boundary_error is None

    def test_iter_range_duplicate_paraid_in_table(self):
        """Duplicate paraId across table rows should extend range to last occurrence"""
        applier = create_mock_applier()

        # Two-row table with duplicate paraId (vertical merge scenario)
        tbl = create_multi_row_table(
            [[['Row1']], [['Row2']]],
            ['AAA', 'AAA']
        )
        body = etree.Element(f'{{{NS["w"]}}}body', nsmap=NSMAP)
        body.append(tbl)

        # Add paragraph after table to ensure range does not spill over
        p_after = etree.SubElement(body, f'{{{NS["w"]}}}p')
        p_after.set(f'{{{NS["w14"]}}}paraId', 'ZZZ')
        r = etree.SubElement(p_after, f'{{{NS["w"]}}}r')
        t = etree.SubElement(r, f'{{{NS["w"]}}}t')
        t.text = 'After table'

        applier.body_elem = body

        start_para = body.find('.//w:tbl//w:p[@w14:paraId="AAA"]', NSMAP)
        table_paras = body.findall('.//w:tbl//w:p[@w14:paraId="AAA"]', NSMAP)
        range_paras = list(applier._iter_paragraphs_in_range(start_para, 'AAA'))

        assert len(range_paras) == len(table_paras)
        assert all(applier._is_paragraph_in_table(p) for p in range_paras)



class TestTableJsonFormat:
    """Tests for JSON format in table row content collection"""

    def test_single_cell_json_format(self):
        """Single cell should produce JSON format ["content"]"""
        applier = create_mock_applier()

        tbl = create_table_with_cells([['Hello']], ['AAA'])
        body = etree.Element(f'{{{NS["w"]}}}body', nsmap=NSMAP)
        body.append(tbl)
        applier.body_elem = body

        start_para = body.find('.//w:p[@w14:paraId="AAA"]', NSMAP)

        runs_info, combined_text, is_cross_para, boundary_error = applier._collect_runs_info_across_paragraphs(
            start_para, 'AAA'
        )

        assert boundary_error is None
        assert combined_text == '["Hello"]'

    def test_multi_cell_same_row_json_format(self):
        """Multiple cells in same row should produce JSON format ["cell1", "cell2"]"""
        applier = create_mock_applier()

        tbl = create_table_with_cells([['Cell1'], ['Cell2'], ['Cell3']], ['AAA', 'BBB', 'CCC'])
        body = etree.Element(f'{{{NS["w"]}}}body', nsmap=NSMAP)
        body.append(tbl)
        applier.body_elem = body

        start_para = body.find('.//w:p[@w14:paraId="AAA"]', NSMAP)

        runs_info, combined_text, is_cross_para, boundary_error = applier._collect_runs_info_across_paragraphs(
            start_para, 'CCC'
        )

        assert boundary_error is None
        assert combined_text == '["Cell1", "Cell2", "Cell3"]'

    def test_multi_para_in_cell_json_format(self):
        """Multiple paragraphs in same cell should use \\n separator"""
        applier = create_mock_applier()

        # Cell with two paragraphs
        tbl = create_table_with_cells([['Para1', 'Para2']], ['AAA', 'BBB'])
        body = etree.Element(f'{{{NS["w"]}}}body', nsmap=NSMAP)
        body.append(tbl)
        applier.body_elem = body

        start_para = body.find('.//w:p[@w14:paraId="AAA"]', NSMAP)

        runs_info, combined_text, is_cross_para, boundary_error = applier._collect_runs_info_across_paragraphs(
            start_para, 'BBB'
        )

        assert boundary_error is None
        # In JSON, newlines are escaped as \n
        assert combined_text == '["Para1\\nPara2"]'

    def test_json_escape_quotes(self):
        """Content with quotes should be JSON-escaped"""
        applier = create_mock_applier()

        tbl = create_table_with_cells([['He said "hello"']], ['AAA'])
        body = etree.Element(f'{{{NS["w"]}}}body', nsmap=NSMAP)
        body.append(tbl)
        applier.body_elem = body

        start_para = body.find('.//w:p[@w14:paraId="AAA"]', NSMAP)

        runs_info, combined_text, is_cross_para, boundary_error = applier._collect_runs_info_across_paragraphs(
            start_para, 'AAA'
        )

        assert boundary_error is None
        # The quotes inside should be escaped
        assert combined_text == '["He said \\"hello\\""]'



class TestCrossCellBoundary:
    """Tests for cross-cell boundary detection"""

    def test_check_cross_cell_boundary_single_cell(self):
        """Runs from single cell should not trigger cross-cell"""
        applier = create_mock_applier()

        tbl = create_table_with_cells([['Cell content']], ['AAA'])
        body = etree.Element(f'{{{NS["w"]}}}body', nsmap=NSMAP)
        body.append(tbl)
        applier.body_elem = body

        start_para = body.find('.//w:p[@w14:paraId="AAA"]', NSMAP)
        cell = applier._find_ancestor_cell(start_para)

        # Simulate run info with cell_elem
        runs_info = [
            {'text': 'Cell content', 'cell_elem': cell, 'start': 2, 'end': 14}
        ]

        assert applier._check_cross_cell_boundary(runs_info) is False

    def test_check_cross_cell_boundary_multiple_cells(self):
        """Runs from multiple cells should trigger cross-cell"""
        applier = create_mock_applier()

        tbl = create_table_with_cells([['Cell1'], ['Cell2']], ['AAA', 'BBB'])
        body = etree.Element(f'{{{NS["w"]}}}body', nsmap=NSMAP)
        body.append(tbl)
        applier.body_elem = body

        para_aaa = body.find('.//w:p[@w14:paraId="AAA"]', NSMAP)
        para_bbb = body.find('.//w:p[@w14:paraId="BBB"]', NSMAP)
        cell_aaa = applier._find_ancestor_cell(para_aaa)
        cell_bbb = applier._find_ancestor_cell(para_bbb)

        # Simulate run info spanning two cells
        runs_info = [
            {'text': 'Cell1', 'cell_elem': cell_aaa, 'start': 2, 'end': 7},
            {'text': 'Cell2', 'cell_elem': cell_bbb, 'start': 12, 'end': 17}
        ]

        assert applier._check_cross_cell_boundary(runs_info) is True


class TestMultiCellExtraction:
    """Tests for multi-cell diff extraction behavior."""

    def test_extract_insert_at_cell_right_boundary(self):
        """
        Insertion at a cell's right boundary should stay in that cell.

        Regression case:
        - violation: ["A", "B"]
        - revised:   ["A!", "B"]
        """
        applier = create_mock_applier()

        tbl = create_table_with_cells([['A'], ['B']], ['AAA', 'BBB'])
        body = etree.Element(f'{{{NS["w"]}}}body', nsmap=NSMAP)
        body.append(tbl)
        applier.body_elem = body

        start_para = body.find('.//w:p[@w14:paraId="AAA"]', NSMAP)
        runs_info, combined_text, is_cross_para, boundary_error = applier._collect_runs_info_across_paragraphs(
            start_para, 'BBB'
        )

        assert boundary_error is None
        assert is_cross_para is True
        assert combined_text == '["A", "B"]'

        violation_text = '["A", "B"]'
        revised_text = '["A!", "B"]'
        match_start = combined_text.find(violation_text)
        assert match_start != -1

        affected = applier._find_affected_runs(
            runs_info,
            match_start,
            match_start + len(violation_text),
        )

        result = applier._try_extract_multi_cell_edits(
            violation_text,
            revised_text,
            affected,
            match_start,
        )

        assert result is not None
        assert len(result) == 1
        assert result[0]['cell_violation'] == 'A'
        assert result[0]['cell_revised'] == 'A!'


class TestCellParagraphBoundaryHandling:
    """Tests for separating paragraph boundaries from in-paragraph line breaks."""

    def test_replace_multi_para_cell_with_br(self):
        """Cell with w:br inside paragraph and multiple paragraphs should replace."""
        applier = create_mock_applier()

        tbl = create_table_with_cells_with_br(
            [["Line1\nLine2", "Para2"]],
            ["AAA", "BBB"]
        )
        body = etree.Element(f'{{{NS["w"]}}}body', nsmap=NSMAP)
        body.append(tbl)
        applier.body_elem = body

        start_para = body.find('.//w:p[@w14:paraId="AAA"]', NSMAP)
        runs_info, combined_text, is_cross_para, boundary_error = applier._collect_runs_info_across_paragraphs(
            start_para, 'BBB'
        )

        assert boundary_error is None
        assert is_cross_para is True
        assert combined_text == '["Line1\\nLine2\\nPara2"]'

        violation_text = '["Line1\\nLine2\\nPara2"]'
        revised_text = '["Line1\\nLine2\\nPara2 (updated)"]'
        match_start = combined_text.find(violation_text)
        assert match_start != -1

        affected = applier._find_affected_runs(
            runs_info,
            match_start,
            match_start + len(violation_text),
        )

        single_cell = applier._try_extract_single_cell_edit(
            violation_text,
            revised_text,
            affected,
            match_start,
        )
        assert single_cell is not None

        status = applier._apply_replace_in_cell_paragraphs(
            single_cell['cell_violation'],
            single_cell['cell_revised'],
            single_cell['cell_runs'],
            "Reason",
            "Test",
            skip_comment=True
        )
        assert status == 'success'

    def test_replace_single_para_cell_with_br(self):
        """Single paragraph cell with w:br should replace successfully."""
        applier = create_mock_applier()

        tbl = create_table_with_cells_with_br(
            [["Alpha\nBeta"]],
            ["AAA"]
        )
        body = etree.Element(f'{{{NS["w"]}}}body', nsmap=NSMAP)
        body.append(tbl)
        applier.body_elem = body

        start_para = body.find('.//w:p[@w14:paraId="AAA"]', NSMAP)
        runs_info, combined_text, is_cross_para, boundary_error = applier._collect_runs_info_across_paragraphs(
            start_para, 'AAA'
        )

        assert boundary_error is None
        assert combined_text == '["Alpha\\nBeta"]'

        violation_text = '["Alpha\\nBeta"]'
        revised_text = '["Alpha\\nBeta!"]'
        match_start = combined_text.find(violation_text)
        assert match_start != -1

        affected = applier._find_affected_runs(
            runs_info,
            match_start,
            match_start + len(violation_text),
        )

        single_cell = applier._try_extract_single_cell_edit(
            violation_text,
            revised_text,
            affected,
            match_start,
        )
        assert single_cell is not None

        status = applier._apply_replace_in_cell_paragraphs(
            single_cell['cell_violation'],
            single_cell['cell_revised'],
            single_cell['cell_runs'],
            "Reason",
            "Test",
            skip_comment=True
        )
        assert status == 'success'

    def test_replace_multi_para_cell_soft_break_deletion(self):
        """Deleting a soft break should not merge paragraphs in a cell."""
        applier = create_mock_applier()

        tbl = create_table_with_cells_with_br(
            [["A\nB", "C"]],
            ["AAA", "BBB"]
        )
        body = etree.Element(f'{{{NS["w"]}}}body', nsmap=NSMAP)
        body.append(tbl)
        applier.body_elem = body

        start_para = body.find('.//w:p[@w14:paraId="AAA"]', NSMAP)
        runs_info, combined_text, is_cross_para, boundary_error = applier._collect_runs_info_across_paragraphs(
            start_para, 'BBB'
        )

        assert boundary_error is None
        assert is_cross_para is True
        assert combined_text == '["A\\nB\\nC"]'

        violation_text = '["A\\nB\\nC"]'
        revised_text = '["AB\\nC"]'
        match_start = combined_text.find(violation_text)
        assert match_start != -1

        affected = applier._find_affected_runs(
            runs_info,
            match_start,
            match_start + len(violation_text),
        )

        single_cell = applier._try_extract_single_cell_edit(
            violation_text,
            revised_text,
            affected,
            match_start,
        )
        assert single_cell is not None

        status = applier._apply_replace_in_cell_paragraphs(
            single_cell['cell_violation'],
            single_cell['cell_revised'],
            single_cell['cell_runs'],
            "Reason",
            "Test",
            skip_comment=True
        )
        assert status == 'success'

        para_elems = body.findall('.//w:tc//w:p', NSMAP)
        assert len(para_elems) == 2

    def test_replace_multi_para_cell_preserves_leading_space(self):
        """Leading spaces at paragraph edges should be preserved in replacements."""
        applier = create_mock_applier()

        tbl = create_table_with_cells(
            [["  Beta", "Gamma"]],
            ["AAA", "BBB"]
        )
        body = etree.Element(f'{{{NS["w"]}}}body', nsmap=NSMAP)
        body.append(tbl)
        applier.body_elem = body

        start_para = body.find('.//w:p[@w14:paraId="AAA"]', NSMAP)
        runs_info, combined_text, is_cross_para, boundary_error = applier._collect_runs_info_across_paragraphs(
            start_para, 'BBB'
        )

        assert boundary_error is None
        assert is_cross_para is True
        assert combined_text == '["Beta\\nGamma"]'

        violation_text = '["Beta\\nGamma"]'
        revised_text = '["Beta!\\nGamma"]'
        match_start = combined_text.find(violation_text)
        assert match_start != -1

        affected = applier._find_affected_runs(
            runs_info,
            match_start,
            match_start + len(violation_text),
        )

        single_cell = applier._try_extract_single_cell_edit(
            violation_text,
            revised_text,
            affected,
            match_start,
        )
        assert single_cell is not None

        status = applier._apply_replace_in_cell_paragraphs(
            single_cell['cell_violation'],
            single_cell['cell_revised'],
            single_cell['cell_runs'],
            "Reason",
            "Test",
            skip_comment=True
        )
        assert status == 'success'

        para_aaa = body.find('.//w:p[@w14:paraId="AAA"]', NSMAP)
        assert para_aaa is not None
        para_text = ''.join(para_aaa.itertext())
        assert para_text.startswith('  ')


# ============================================================
# Tests: Real Document Table (tests/test.docx)
# ============================================================


class TestFilterRealRuns:
    """Test that _filter_real_runs correctly filters out synthetic boundary markers"""

    def test_filter_removes_json_boundary_runs(self):
        """Test that JSON boundary runs (["  and "]) are filtered out"""
        applier = create_mock_applier()

        runs = [
            {'text': '["', 'start': 0, 'end': 2, 'is_json_boundary': True},
            {'text': 'hello', 'start': 2, 'end': 7, 'elem': 'mock_elem', 'rPr': None},
            {'text': '"]', 'start': 7, 'end': 9, 'is_json_boundary': True},
        ]

        filtered = applier._filter_real_runs(runs)

        assert len(filtered) == 1
        assert filtered[0]['text'] == 'hello'
        assert 'elem' in filtered[0]

    def test_filter_removes_cell_boundary_runs(self):
        """Test that cell boundary runs (", ") are filtered out"""
        applier = create_mock_applier()

        runs = [
            {'text': '["', 'start': 0, 'end': 2, 'is_json_boundary': True},
            {'text': 'cell1', 'start': 2, 'end': 7, 'elem': 'mock1', 'rPr': None},
            {'text': '", "', 'start': 7, 'end': 11, 'is_json_boundary': True, 'is_cell_boundary': True},
            {'text': 'cell2', 'start': 11, 'end': 16, 'elem': 'mock2', 'rPr': None},
            {'text': '"]', 'start': 16, 'end': 18, 'is_json_boundary': True},
        ]

        filtered = applier._filter_real_runs(runs)

        assert len(filtered) == 2
        assert filtered[0]['text'] == 'cell1'
        assert filtered[1]['text'] == 'cell2'

    def test_filter_removes_para_boundary_runs(self):
        """Test that paragraph boundary runs are filtered out"""
        applier = create_mock_applier()

        runs = [
            {'text': 'para1', 'start': 0, 'end': 5, 'elem': 'mock1', 'rPr': None},
            {'text': '\n', 'start': 5, 'end': 6, 'is_para_boundary': True},
            {'text': 'para2', 'start': 6, 'end': 11, 'elem': 'mock2', 'rPr': None},
        ]

        filtered = applier._filter_real_runs(runs)

        assert len(filtered) == 2
        assert filtered[0]['text'] == 'para1'
        assert filtered[1]['text'] == 'para2'

    def test_filter_removes_json_escape_runs(self):
        """Test that JSON escape marker runs are filtered out"""
        applier = create_mock_applier()

        runs = [
            {'text': 'before', 'start': 0, 'end': 6, 'elem': 'mock1', 'rPr': None},
            {'text': '\\', 'start': 6, 'end': 7, 'is_json_escape': True},
            {'text': 'after', 'start': 7, 'end': 12, 'elem': 'mock2', 'rPr': None},
        ]

        filtered = applier._filter_real_runs(runs)

        assert len(filtered) == 2
        assert filtered[0]['text'] == 'before'
        assert filtered[1]['text'] == 'after'

    def test_filter_keeps_all_real_runs(self):
        """Test that real runs with elem are preserved"""
        applier = create_mock_applier()

        runs = [
            {'text': 'run1', 'start': 0, 'end': 4, 'elem': 'mock1', 'rPr': None},
            {'text': 'run2', 'start': 4, 'end': 8, 'elem': 'mock2', 'rPr': 'mock_rPr'},
            {'text': 'run3', 'start': 8, 'end': 12, 'elem': 'mock3', 'rPr': None, 'is_drawing': True},
        ]

        filtered = applier._filter_real_runs(runs)

        assert len(filtered) == 3
        assert all('elem' in r for r in filtered)


# ============================================================
# Test Class: Original Text Helpers (Issue 2 - JSON escaped text for mutations)
# ============================================================


class TestOriginalTextHelpers:
    """Test helper methods for handling JSON-escaped vs original text"""

    def test_get_run_original_text_with_original(self):
        """Test that _get_run_original_text returns original_text when present"""
        applier = create_mock_applier()

        run = {
            'text': 'He said \\"hello\\"',  # JSON escaped
            'original_text': 'He said "hello"',  # Original
            'start': 0,
            'end': 17
        }

        result = applier._get_run_original_text(run)
        assert result == 'He said "hello"'

    def test_get_run_original_text_without_original(self):
        """Test that _get_run_original_text returns text when no original_text"""
        applier = create_mock_applier()

        run = {
            'text': 'simple text',
            'start': 0,
            'end': 11
        }

        result = applier._get_run_original_text(run)
        assert result == 'simple text'

    def test_translate_escaped_offset_no_escaping(self):
        """Test offset translation when no escaping present"""
        applier = create_mock_applier()

        run = {
            'text': 'hello world',
            'start': 0,
            'end': 11
        }

        # No original_text means no escaping, offset unchanged
        assert applier._translate_escaped_offset(run, 0) == 0
        assert applier._translate_escaped_offset(run, 5) == 5
        assert applier._translate_escaped_offset(run, 11) == 11

    def test_translate_escaped_offset_with_quote_escaping(self):
        """Test offset translation with escaped quotes"""
        applier = create_mock_applier()

        # Original: He said "hi"
        # Escaped:  He said \"hi\"
        run = {
            'text': 'He said \\"hi\\"',  # 14 chars
            'original_text': 'He said "hi"',  # 12 chars
            'start': 0,
            'end': 14
        }

        # Before the first escape
        assert applier._translate_escaped_offset(run, 0) == 0
        assert applier._translate_escaped_offset(run, 7) == 7  # 'He said'

        # After both escapes (end of string)
        assert applier._translate_escaped_offset(run, 14) == 12

    def test_translate_escaped_offset_with_backslash_escaping(self):
        """Test offset translation with escaped backslashes"""
        applier = create_mock_applier()

        # Original: path\to\file
        # Escaped:  path\\to\\file
        run = {
            'text': 'path\\\\to\\\\file',  # 14 chars
            'original_text': 'path\\to\\file',  # 12 chars
            'start': 0,
            'end': 14
        }

        assert applier._translate_escaped_offset(run, 0) == 0
        assert applier._translate_escaped_offset(run, 4) == 4  # 'path'
        assert applier._translate_escaped_offset(run, 14) == 12

    def test_translate_escaped_offset_with_newline_escaping(self):
        """Test offset translation with escaped newlines"""
        applier = create_mock_applier()

        # Original: line1\nline2 (11 chars with actual newline)
        # Escaped:  line1\\nline2 (12 chars with \n as two chars)
        run = {
            'text': 'line1\\nline2',  # 12 chars
            'original_text': 'line1\nline2',  # 11 chars
            'start': 0,
            'end': 12
        }

        assert applier._translate_escaped_offset(run, 0) == 0
        assert applier._translate_escaped_offset(run, 5) == 5  # 'line1'
        assert applier._translate_escaped_offset(run, 12) == 11

    def test_translate_escaped_offset_boundary_cases(self):
        """Test boundary cases for offset translation"""
        applier = create_mock_applier()

        run = {
            'text': 'test\\"text',  # 10 chars
            'original_text': 'test"text',  # 9 chars
            'start': 0,
            'end': 10
        }

        # Negative offset returns 0
        assert applier._translate_escaped_offset(run, -1) == 0

        # Offset beyond end returns original length
        assert applier._translate_escaped_offset(run, 100) == 9

    def test_is_table_mode_with_original_text(self):
        """Test _is_table_mode returns True when runs have original_text"""
        applier = create_mock_applier()

        runs = [
            {'text': 'escaped', 'original_text': 'original', 'start': 0, 'end': 7},
            {'text': 'normal', 'start': 7, 'end': 13}
        ]

        assert applier._is_table_mode(runs) is True

    def test_is_table_mode_without_original_text(self):
        """Test _is_table_mode returns False when no runs have original_text"""
        applier = create_mock_applier()

        runs = [
            {'text': 'normal1', 'start': 0, 'end': 7},
            {'text': 'normal2', 'start': 7, 'end': 14}
        ]

        assert applier._is_table_mode(runs) is False

    def test_decode_json_escaped_quotes(self):
        """Test _decode_json_escaped decodes escaped quotes"""
        applier = create_mock_applier()

        assert applier._decode_json_escaped('He said \\"hello\\"') == 'He said "hello"'

    def test_decode_json_escaped_backslash(self):
        """Test _decode_json_escaped decodes escaped backslashes"""
        applier = create_mock_applier()

        assert applier._decode_json_escaped('path\\\\to\\\\file') == 'path\\to\\file'

    def test_decode_json_escaped_newline(self):
        """Test _decode_json_escaped decodes escaped newlines"""
        applier = create_mock_applier()

        assert applier._decode_json_escaped('line1\\nline2') == 'line1\nline2'

    def test_decode_json_escaped_empty(self):
        """Test _decode_json_escaped handles empty string"""
        applier = create_mock_applier()

        assert applier._decode_json_escaped('') == ''

    def test_decode_json_escaped_no_escapes(self):
        """Test _decode_json_escaped returns unchanged text with no escapes"""
        applier = create_mock_applier()

        assert applier._decode_json_escaped('normal text') == 'normal text'


# ============================================================
# Test Class: Apply Methods with Escaped Text (Issue 2)
# ============================================================


class TestApplyMethodsWithEscapedText:
    """Test that apply methods correctly use original text for mutations"""

    def test_apply_delete_with_escaped_quote(self):
        """Test _apply_delete correctly decodes JSON-escaped violation_text"""
        # Create paragraph with text containing quotes
        p = create_paragraph_xml('He said "hello" to me')
        applier = create_mock_applier()

        # Simulate table mode where run has both escaped and original text
        runs_info = [{
            'text': 'He said \\"hello\\" to me',  # JSON escaped (23 chars)
            'original_text': 'He said "hello" to me',  # Original (21 chars)
            'start': 0,
            'end': 23,
            'elem': p.find('.//w:r', namespaces={'w': NS['w']}),
            'rPr': None,
            'para_elem': p
        }]

        # In table mode, LLM sees JSON-escaped content, so it returns escaped text
        # violation_text is what the LLM returns (escaped form)
        violation_text = '\\"hello\\"'  # LLM returns escaped version

        # Find position in the escaped combined text
        combined_text = runs_info[0]['text']
        match_start = combined_text.find('\\"hello\\"')
        assert match_start != -1

        result = applier._apply_delete(
            p, violation_text, 'test reason',
            runs_info, match_start, 'test-author'
        )

        # Should succeed without corrupting the document
        assert result == 'success'

        # Verify the delText contains the DECODED text (actual quotes, not escaped)
        del_text_elem = p.find('.//w:delText', namespaces={'w': NS['w']})
        assert del_text_elem is not None
        # The delText should contain decoded text: "hello" not \"hello\"
        assert del_text_elem.text == '"hello"', f"delText should be decoded, got: {del_text_elem.text}"
        assert '\\' not in del_text_elem.text, "delText should not contain backslashes"

    def test_apply_replace_with_escaped_quote(self):
        """Test _apply_replace correctly decodes JSON-escaped text"""
        p = create_paragraph_xml('Value is "old" here')
        applier = create_mock_applier()

        runs_info = [{
            'text': 'Value is \\"old\\" here',  # JSON escaped (21 chars)
            'original_text': 'Value is "old" here',  # Original (19 chars)
            'start': 0,
            'end': 21,
            'elem': p.find('.//w:r', namespaces={'w': NS['w']}),
            'rPr': None,
            'para_elem': p
        }]

        # LLM returns escaped versions - replace entire "old" with "newval"
        # This tests that quotes at different positions get decoded properly
        violation_text = '\\"old\\"'
        revised_text = '\\"newval\\"'

        combined_text = runs_info[0]['text']
        match_start = combined_text.find('\\"old\\"')
        assert match_start != -1

        result = applier._apply_replace(
            p, violation_text, revised_text, 'test reason',
            runs_info, match_start, 'test-author'
        )

        assert result == 'success'

        # The diff between \"old\" and \"newval\" will have:
        # - equal: \" (decoded to ")
        # - delete: old
        # - insert: newval
        # - equal: \" (decoded to ")
        # So we check that the output doesn't contain backslash escapes

        # Get all text content from the paragraph
        all_text = etree.tostring(p, encoding='unicode')

        # Should not contain literal \\" sequences in the output
        assert '\\"' not in all_text or '&quot;' in all_text, \
            f"Output should not contain escaped quotes: {all_text[:200]}"

    def test_apply_replace_decodes_full_delete(self):
        """Test _apply_replace decodes when entire text is replaced"""
        p = create_paragraph_xml('Say "hi"')
        applier = create_mock_applier()

        runs_info = [{
            'text': 'Say \\"hi\\"',  # JSON escaped
            'original_text': 'Say "hi"',  # Original
            'start': 0,
            'end': 10,
            'elem': p.find('.//w:r', namespaces={'w': NS['w']}),
            'rPr': None,
            'para_elem': p
        }]

        # Replace entire "hi" with "bye" - different lengths force full delete/insert
        violation_text = '\\"hi\\"'
        revised_text = '\\"bye\\"'

        combined_text = runs_info[0]['text']
        match_start = combined_text.find('\\"hi\\"')
        assert match_start != -1

        result = applier._apply_replace(
            p, violation_text, revised_text, 'test reason',
            runs_info, match_start, 'test-author'
        )

        assert result == 'success'

        # Verify ins/del content doesn't have escaped backslashes
        ins_elems = p.findall('.//w:ins//w:t', namespaces={'w': NS['w']})
        for ins_elem in ins_elems:
            if ins_elem.text:
                assert '\\' not in ins_elem.text, \
                    f"Inserted text should not have backslashes: {ins_elem.text}"

    def test_apply_replace_preserves_inserted_leading_space(self):
        """Inserted leading space should be preserved with xml:space='preserve'."""
        p = create_paragraph_xml('FPGALASH是否存在')
        applier = create_mock_applier()

        run_elem = p.find('.//w:r', namespaces={'w': NS['w']})
        runs_info = [{
            'text': 'FPGALASH是否存在',
            'start': 0,
            'end': len('FPGALASH是否存在'),
            'elem': run_elem,
            'rPr': None,
            'para_elem': p
        }]

        result = applier._apply_replace(
            p,
            'FPGALASH是否存在',
            'FPGA FLASH是否存在',
            'test reason',
            runs_info,
            0,
            'test-author'
        )

        assert result == 'success'

        ins_t = p.find('.//w:ins//w:t', namespaces={'w': NS['w']})
        assert ins_t is not None
        assert ins_t.text == ' F'
        assert ins_t.get('{http://www.w3.org/XML/1998/namespace}space') == 'preserve'

    def test_apply_manual_with_escaped_quote(self):
        """Test _apply_manual correctly splits runs with quoted content"""
        p = create_paragraph_xml('Text "quoted" more')
        applier = create_mock_applier()

        runs_info = [{
            'text': 'Text \\"quoted\\" more',  # JSON escaped (20 chars)
            'original_text': 'Text "quoted" more',  # Original (18 chars)
            'start': 0,
            'end': 20,
            'elem': p.find('.//w:r', namespaces={'w': NS['w']}),
            'rPr': None,
            'para_elem': p
        }]

        violation_text = '"quoted"'

        combined_text = runs_info[0]['text']
        match_start = combined_text.find('\\"quoted\\"')
        assert match_start != -1

        result = applier._apply_manual(
            p, violation_text, 'test reason', 'suggestion',
            runs_info, match_start, 'test-author'
        )

        assert result == 'success'

        # Verify comment markers were inserted
        comment_start = p.find('.//w:commentRangeStart', namespaces={'w': NS['w']})
        comment_end = p.find('.//w:commentRangeEnd', namespaces={'w': NS['w']})
        assert comment_start is not None
        assert comment_end is not None

    def test_before_after_text_uses_original(self):
        """Test that before/after text splits use original (unescaped) text"""
        p = create_paragraph_xml('prefix "target" suffix')
        applier = create_mock_applier()

        # The run element
        run_elem = p.find('.//w:r', namespaces={'w': NS['w']})

        runs_info = [{
            'text': 'prefix \\"target\\" suffix',  # 24 chars escaped
            'original_text': 'prefix "target" suffix',  # 22 chars original
            'start': 0,
            'end': 24,
            'elem': run_elem,
            'rPr': None,
            'para_elem': p
        }]

        # Delete "target" (with quotes)
        violation_text = '"target"'

        # Position in escaped text
        combined_text = runs_info[0]['text']
        match_start = combined_text.find('\\"target\\"')

        result = applier._apply_delete(
            p, violation_text, 'test reason',
            runs_info, match_start, 'test-author'
        )

        assert result == 'success'

        # Get all text runs to verify splits
        all_runs = p.findall('.//w:r', namespaces={'w': NS['w']})

        # Find the before text (should be 'prefix ' not 'prefix \\')
        texts = []
        for run in all_runs:
            t = run.find('w:t', namespaces={'w': NS['w']})
            if t is not None and t.text:
                texts.append(t.text)

        # Should have correct unescaped prefix and suffix
        # 'prefix ' and ' suffix' should be present as separate runs
        all_text = ''.join(texts)
        assert 'prefix ' in all_text or all_text.startswith('prefix')
        assert ' suffix' in all_text or all_text.endswith('suffix')

        # Should NOT have escaped backslashes in the output
        assert '\\' not in all_text, f"Output should not contain backslashes: {all_text}"


# ============================================================
# Test Class: Apply Methods with JSON Boundary Runs (Issue 1)
# ============================================================


class TestApplyMethodsWithBoundaryRuns:
    """Test that apply methods don't fail when runs contain JSON boundary markers"""

    def test_apply_delete_filters_boundary_runs(self):
        """Test _apply_delete filters out JSON boundary runs before processing"""
        p = create_paragraph_xml('cell content')
        applier = create_mock_applier()

        run_elem = p.find('.//w:r', namespaces={'w': NS['w']})

        # Simulated table mode runs with JSON boundaries
        runs_info = [
            {'text': '["', 'start': 0, 'end': 2, 'is_json_boundary': True, 'para_elem': p},
            {'text': 'cell content', 'start': 2, 'end': 14, 'elem': run_elem, 'rPr': None, 'para_elem': p},
            {'text': '"]', 'start': 14, 'end': 16, 'is_json_boundary': True, 'para_elem': p},
        ]

        # Delete "cell" - note we need to account for the JSON prefix
        violation_text = 'cell'
        match_start = 2  # After '["'

        result = applier._apply_delete(
            p, violation_text, 'test reason',
            runs_info, match_start, 'test-author'
        )

        # Should succeed without KeyError on boundary runs
        assert result == 'success'

    def test_apply_replace_filters_boundary_runs(self):
        """Test _apply_replace filters out JSON boundary runs before processing"""
        p = create_paragraph_xml('old value')
        applier = create_mock_applier()

        run_elem = p.find('.//w:r', namespaces={'w': NS['w']})

        runs_info = [
            {'text': '["', 'start': 0, 'end': 2, 'is_json_boundary': True, 'para_elem': p},
            {'text': 'old value', 'start': 2, 'end': 11, 'elem': run_elem, 'rPr': None, 'para_elem': p},
            {'text': '"]', 'start': 11, 'end': 13, 'is_json_boundary': True, 'para_elem': p},
        ]

        violation_text = 'old'
        revised_text = 'new'
        match_start = 2

        result = applier._apply_replace(
            p, violation_text, revised_text, 'test reason',
            runs_info, match_start, 'test-author'
        )

        assert result == 'success'

    def test_apply_manual_filters_boundary_runs(self):
        """Test _apply_manual filters out JSON boundary runs before processing"""
        p = create_paragraph_xml('target text')
        applier = create_mock_applier()

        run_elem = p.find('.//w:r', namespaces={'w': NS['w']})

        runs_info = [
            {'text': '["', 'start': 0, 'end': 2, 'is_json_boundary': True, 'para_elem': p},
            {'text': 'target text', 'start': 2, 'end': 13, 'elem': run_elem, 'rPr': None, 'para_elem': p},
            {'text': '"]', 'start': 13, 'end': 15, 'is_json_boundary': True, 'para_elem': p},
        ]

        violation_text = 'target'
        match_start = 2

        result = applier._apply_manual(
            p, violation_text, 'test reason', 'suggestion',
            runs_info, match_start, 'test-author'
        )

        assert result == 'success'

    def test_apply_delete_with_only_boundary_runs_returns_fallback(self):
        """Test that _apply_delete returns fallback if all runs are boundaries"""
        p = create_paragraph_xml('x')
        applier = create_mock_applier()

        # All runs are boundary markers (no real runs)
        runs_info = [
            {'text': '["', 'start': 0, 'end': 2, 'is_json_boundary': True, 'para_elem': p},
            {'text': '", "', 'start': 2, 'end': 6, 'is_json_boundary': True, 'is_cell_boundary': True, 'para_elem': p},
            {'text': '"]', 'start': 6, 'end': 8, 'is_json_boundary': True, 'para_elem': p},
        ]

        # Try to delete something that spans only boundaries
        violation_text = '["'
        match_start = 0

        result = applier._apply_delete(
            p, violation_text, 'test reason',
            runs_info, match_start, 'test-author'
        )

        # Should fallback since no real runs to modify
        assert result == 'fallback'


class TestStripTableRowNumbering:
    """Tests for strip_table_row_numbering with multi-row JSON strings"""

    def test_multi_row_json_string(self):
        text = '["9", "C_RXD+", "RS422发送+", "422接口"], ["10", "C_RXD-", "RS422发送-", "422接口"]'
        expected = '["", "C_RXD+", "RS422发送+", "422接口"], ["", "C_RXD-", "RS422发送-", "422接口"]'

        stripped, was_stripped = apply_module.strip_table_row_numbering(text)

        assert was_stripped is True
        assert stripped == expected

    def test_full_table_json_array_string(self):
        text = '[[\"9\", \"A\"], [\"10\", \"B\"]]'
        expected = '["", "A"], ["", "B"]'

        stripped, was_stripped = apply_module.strip_table_row_numbering(text)

        assert was_stripped is True
        assert stripped == expected

    def test_full_table_json_array_with_whitespace(self):
        text = '[\n  ["9", "A"],\n  ["10", "B"]\n]'
        expected = '["", "A"], ["", "B"]'

        stripped, was_stripped = apply_module.strip_table_row_numbering(text)

        assert was_stripped is True
        assert stripped == expected

    def test_full_table_json_array_no_numbering(self):
        text = '[[\"A\", \"B\"], [\"C\", \"D\"]]'
        expected = '["A", "B"], ["C", "D"]'

        stripped, was_stripped = apply_module.strip_table_row_numbering(text)

        assert was_stripped is True
        assert stripped == expected

    def test_first_cell_with_punctuation(self):
        text = '["9)", "A"], ["10.", "B"]'
        expected = '["", "A"], ["", "B"]'

        stripped, was_stripped = apply_module.strip_table_row_numbering(text)

        assert was_stripped is True
        assert stripped == expected


# ============================================================
# Comment Range Ordering (vMerge continue)
# ============================================================

class TestManualCommentOrdering:
    def test_vmerge_continue_end_ordering(self):
        """Comment range should not be inverted when end is in vMerge continue cell."""
        applier = create_mock_applier()

        tbl = create_table_with_vmerge_continue(
            ['P1', 'P2', 'P3', 'P4', 'P5', 'P6']
        )
        body = etree.Element(f'{{{NS["w"]}}}body', nsmap=NSMAP)
        body.append(tbl)
        applier.body_elem = body

        start_para = body.find('.//w:p[@w14:paraId="P1"]', NSMAP)
        runs_info, combined_text, is_cross_para, boundary_error = applier._collect_runs_info_across_paragraphs(
            start_para, 'P6'
        )

        assert boundary_error is None
        assert is_cross_para is True

        violation_text = 'R2C1", "Merged'
        pos = combined_text.find(violation_text)
        assert pos != -1

        status = applier._apply_manual(
            start_para,
            violation_text,
            "Reason",
            "",
            runs_info,
            pos,
            "Test",
            is_cross_paragraph=True
        )
        assert status == 'success'

        all_elems = list(body.iter())
        start_elems = list(body.iter(f'{{{NS["w"]}}}commentRangeStart'))
        end_elems = list(body.iter(f'{{{NS["w"]}}}commentRangeEnd'))

        assert len(start_elems) == 1
        assert len(end_elems) == 1
        assert all_elems.index(start_elems[0]) < all_elems.index(end_elems[0])

    def test_reference_only_fallback_when_no_valid_end(self):
        """When no valid end run is found, fallback to reference-only comment."""
        applier = create_mock_applier()

        # Build paragraph with two runs
        body = etree.Element(f'{{{NS["w"]}}}body', nsmap=NSMAP)
        p = etree.SubElement(body, f'{{{NS["w"]}}}p')
        p.set(f'{{{NS["w14"]}}}paraId', 'AAA')
        r1 = etree.SubElement(p, f'{{{NS["w"]}}}r')
        t1 = etree.SubElement(r1, f'{{{NS["w"]}}}t')
        t1.text = 'AAA'
        r2 = etree.SubElement(p, f'{{{NS["w"]}}}r')
        t2 = etree.SubElement(r2, f'{{{NS["w"]}}}t')
        t2.text = 'BBB'

        applier.body_elem = body

        orig_runs_info, _ = applier._collect_runs_info_original(p)
        violation_text = 'AAABBB'
        match_start = 0

        # Force end_key to be None by stubbing _get_run_doc_key
        first_run_elem = orig_runs_info[0]['elem']

        def fake_get_run_doc_key(run_elem, para_elem, para_order=None):
            if run_elem is first_run_elem:
                return (1, 0)
            return None

        applier._get_run_doc_key = fake_get_run_doc_key

        status = applier._apply_manual(
            p,
            violation_text,
            "Reason",
            "",
            orig_runs_info,
            match_start,
            "Test",
            is_cross_paragraph=False
        )

        assert status == 'success'
        assert list(body.iter(f'{{{NS["w"]}}}commentRangeStart')) == []
        assert list(body.iter(f'{{{NS["w"]}}}commentRangeEnd')) == []
        assert len(list(body.iter(f'{{{NS["w"]}}}commentReference'))) == 1


class TestMultiCellRangeComment:
    def test_range_comment_for_same_row_cells(self):
        """Multi-cell full success in same row should produce a range comment."""
        applier = create_mock_applier()

        tbl = create_table_with_cells([['CellA'], ['CellB']], ['AAA', 'BBB'])
        body = etree.Element(f'{{{NS["w"]}}}body', nsmap=NSMAP)
        body.append(tbl)
        applier.body_elem = body

        cells = body.findall('.//w:tc', NSMAP)
        assert len(cells) == 2

        ok, reason = applier._try_add_multi_cell_range_comment(
            [{'cell_elem': cells[0]}, {'cell_elem': cells[1]}],
            "Range reason",
            "Author"
        )

        assert ok is True
        assert reason == ""
        assert len(applier.comments) == 1
        assert applier.comments[0]['text'] == "Range reason"
        assert applier.comments[0]['author'] == "Author"
        assert len(list(body.iter(f'{{{NS["w"]}}}commentRangeStart'))) == 1
        assert len(list(body.iter(f'{{{NS["w"]}}}commentRangeEnd'))) == 1
        assert len(list(body.iter(f'{{{NS["w"]}}}commentReference'))) == 1

    def test_range_comment_across_rows_supported(self):
        """Manual-style range anchoring should support multi-row table spans."""
        applier = create_mock_applier()

        tbl = create_multi_row_table(
            [[['R1C1'], ['R1C2']], [['R2C1'], ['R2C2']]],
            ['A1', 'A2', 'B1', 'B2']
        )
        body = etree.Element(f'{{{NS["w"]}}}body', nsmap=NSMAP)
        body.append(tbl)
        applier.body_elem = body

        row1_cell = body.find('.//w:tr[1]/w:tc[1]', NSMAP)
        row2_cell = body.find('.//w:tr[2]/w:tc[1]', NSMAP)
        assert row1_cell is not None
        assert row2_cell is not None

        ok, reason = applier._try_add_multi_cell_range_comment(
            [{'cell_elem': row1_cell}, {'cell_elem': row2_cell}],
            "Range reason",
            "Author"
        )
        assert ok is True
        assert reason == ""
        assert len(applier.comments) == 1
        assert applier.comments[0]['text'] == "Range reason"
        assert len(list(body.iter(f'{{{NS["w"]}}}commentRangeStart'))) == 1
        assert len(list(body.iter(f'{{{NS["w"]}}}commentRangeEnd'))) == 1
        assert len(list(body.iter(f'{{{NS["w"]}}}commentReference'))) == 1

    def test_fallback_anchor_comment_prefix_when_range_fails(self):
        """Fallback anchor comment should keep required WHY/WHERE prefix format."""
        applier = create_mock_applier()

        body = etree.Element(f'{{{NS["w"]}}}body', nsmap=NSMAP)
        p = etree.SubElement(body, f'{{{NS["w"]}}}p')
        p.set(f'{{{NS["w14"]}}}paraId', 'AAA')
        r = etree.SubElement(p, f'{{{NS["w"]}}}r')
        t = etree.SubElement(r, f'{{{NS["w"]}}}t')
        t.text = 'Anchor'
        applier.body_elem = body

        ok, reason = applier._try_add_multi_cell_range_comment(
            [{'cell_elem': None}],
            "Range reason",
            "Author"
        )
        assert ok is False
        assert reason == "Missing cell element"

        fallback_reason = f"Range comment failed: {reason}"
        fallback_text = "{WHY}Reason text  {WHERE}Violation text"
        anchor_para = body.find('.//w:tr[1]/w:tc[1]/w:p', NSMAP)
        if anchor_para is None:
            anchor_para = p
        applied = applier._append_reference_only_comment(
            anchor_para,
            fallback_text,
            "Author",
            fallback_reason=fallback_reason
        )
        assert applied is True
        assert len(applier.comments) == 1
        assert applier.comments[0]['text'] == (
            "[FALLBACK]Range comment failed: Missing cell element  "
            "{WHY}Reason text  {WHERE}Violation text"
        )
        assert len(list(body.iter(f'{{{NS["w"]}}}commentRangeStart'))) == 0
        assert len(list(body.iter(f'{{{NS["w"]}}}commentRangeEnd'))) == 0
        assert len(list(body.iter(f'{{{NS["w"]}}}commentReference'))) == 1


# ============================================================
# Main
# ============================================================

if __name__ == '__main__':
    pytest.main([__file__, '-v'])
