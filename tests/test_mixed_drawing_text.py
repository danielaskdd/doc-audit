#!/usr/bin/env python3
"""
Test extraction of mixed text+drawing runs.

This tests the fix for the bug where _collect_runs_info_original would
skip text extraction when w:drawing was found, even though WordprocessingML
allows a single w:r to contain both w:t and w:drawing elements.
"""

import sys
from pathlib import Path
from lxml import etree

# Add parent directory to path for imports
sys.path.insert(0, str(Path(__file__).parent.parent / 'skills' / 'doc-audit' / 'scripts'))

from apply_audit_edits import AuditEditApplier, NS


def test_mixed_text_and_drawing():
    """
    Test that runs containing both text and drawing are split into separate entries.
    
    Before fix: Single entry with is_drawing=True, causing text to be dropped in equal portions
    After fix: Split into 3 entries (text, drawing, text) for proper handling
    """
    # Create a paragraph with a mixed text+drawing run
    para_xml = f'''<w:p xmlns:w="{NS['w']}" xmlns:wp="{NS['wp']}">
        <w:r>
            <w:t>See image </w:t>
            <w:drawing>
                <wp:inline>
                    <wp:docPr id="1" name="Picture 1"/>
                </wp:inline>
            </w:drawing>
            <w:t> for details</w:t>
        </w:r>
    </w:p>'''
    
    para_elem = etree.fromstring(para_xml)
    
    # Create a mock applier instance
    applier = AuditEditApplier.__new__(AuditEditApplier)
    
    # Call _collect_runs_info_original
    runs_info, combined_text = applier._collect_runs_info_original(para_elem)
    
    # Verify results
    print(f"Runs info: {runs_info}")
    print(f"Combined text: '{combined_text}'")
    
    # Should have 3 separate entries (text before, drawing, text after)
    assert len(runs_info) == 3, f"Expected 3 runs, got {len(runs_info)}"
    
    # Should contain both text parts and drawing placeholder
    expected_text = 'See image <drawing id="1" name="Picture 1" /> for details'
    assert combined_text == expected_text, \
        f"Expected '{expected_text}', got '{combined_text}'"
    
    # First entry: text before drawing
    assert runs_info[0]['text'] == 'See image ', \
        f"First entry should be 'See image ', got '{runs_info[0]['text']}'"
    assert runs_info[0]['is_drawing'] == False, \
        "First entry should not be marked as drawing"
    assert runs_info[0]['start'] == 0, "First entry should start at 0"
    assert runs_info[0]['end'] == 10, "First entry should end at 10"
    
    # Second entry: drawing
    assert runs_info[1]['text'] == '<drawing id="1" name="Picture 1" />', \
        f"Second entry should be drawing placeholder, got '{runs_info[1]['text']}'"
    assert runs_info[1]['is_drawing'] == True, \
        "Second entry should be marked as drawing"
    assert 'drawing_elem' in runs_info[1], \
        "Second entry should have drawing_elem reference"
    assert runs_info[1]['start'] == 10, "Second entry should start at 10"
    assert runs_info[1]['end'] == 45, "Second entry should end at 45"
    
    # Third entry: text after drawing
    assert runs_info[2]['text'] == ' for details', \
        f"Third entry should be ' for details', got '{runs_info[2]['text']}'"
    assert runs_info[2]['is_drawing'] == False, \
        "Third entry should not be marked as drawing"
    assert runs_info[2]['start'] == 45, "Third entry should start at 45"
    assert runs_info[2]['end'] == 57, "Third entry should end at 57"
    
    print("✓ Test passed: Mixed text+drawing run is properly split")


def test_drawing_only_run():
    """
    Test that runs containing only drawing (no text) still work.
    """
    para_xml = f'''<w:p xmlns:w="{NS['w']}" xmlns:wp="{NS['wp']}">
        <w:r>
            <w:drawing>
                <wp:inline>
                    <wp:docPr id="2" name="Chart 1"/>
                </wp:inline>
            </w:drawing>
        </w:r>
    </w:p>'''
    
    para_elem = etree.fromstring(para_xml)
    applier = AuditEditApplier.__new__(AuditEditApplier)
    runs_info, combined_text = applier._collect_runs_info_original(para_elem)
    
    print(f"Drawing-only combined text: '{combined_text}'")
    
    assert len(runs_info) == 1, f"Expected 1 run, got {len(runs_info)}"
    expected_text = '<drawing id="2" name="Chart 1" />'
    assert combined_text == expected_text, \
        f"Expected '{expected_text}', got '{combined_text}'"
    
    print("✓ Test passed: Drawing-only run works correctly")


def test_text_only_run():
    """
    Test that runs containing only text (no drawing) still work.
    """
    para_xml = f'''<w:p xmlns:w="{NS['w']}">
        <w:r>
            <w:t>Plain text content</w:t>
        </w:r>
    </w:p>'''
    
    para_elem = etree.fromstring(para_xml)
    applier = AuditEditApplier.__new__(AuditEditApplier)
    runs_info, combined_text = applier._collect_runs_info_original(para_elem)
    
    print(f"Text-only combined text: '{combined_text}'")
    
    assert len(runs_info) == 1, f"Expected 1 run, got {len(runs_info)}"
    assert combined_text == 'Plain text content', \
        f"Expected 'Plain text content', got '{combined_text}'"
    assert runs_info[0].get('is_drawing') != True, \
        "Text-only run should not be marked as drawing"
    
    print("✓ Test passed: Text-only run works correctly")


def test_multiple_mixed_runs():
    """
    Test paragraph with multiple runs, some with drawings, some without.
    
    The second w:r contains both text and drawing, so it gets split into 2 entries.
    Total: 4 entries (text, text, drawing, text)
    """
    para_xml = f'''<w:p xmlns:w="{NS['w']}" xmlns:wp="{NS['wp']}">
        <w:r>
            <w:t>First run</w:t>
        </w:r>
        <w:r>
            <w:t>Second with </w:t>
            <w:drawing>
                <wp:inline>
                    <wp:docPr id="3" name="Icon"/>
                </wp:inline>
            </w:drawing>
        </w:r>
        <w:r>
            <w:t>Third run</w:t>
        </w:r>
    </w:p>'''
    
    para_elem = etree.fromstring(para_xml)
    applier = AuditEditApplier.__new__(AuditEditApplier)
    runs_info, combined_text = applier._collect_runs_info_original(para_elem)
    
    print(f"Multiple runs combined text: '{combined_text}'")
    
    # Should have 4 entries: text (first w:r), text (second w:r before drawing), 
    # drawing (second w:r), text (third w:r)
    assert len(runs_info) == 4, f"Expected 4 runs, got {len(runs_info)}"
    
    expected_text = 'First runSecond with <drawing id="3" name="Icon" />Third run'
    assert combined_text == expected_text, \
        f"Expected '{expected_text}', got '{combined_text}'"
    
    # Check individual runs
    assert runs_info[0].get('is_drawing') != True, "First entry (First run) should not have drawing"
    assert runs_info[1].get('is_drawing') != True, "Second entry (Second with) should not have drawing"
    assert runs_info[2]['is_drawing'] == True, "Third entry (drawing) should be marked as drawing"
    assert runs_info[3].get('is_drawing') != True, "Fourth entry (Third run) should not have drawing"
    
    print("✓ Test passed: Multiple mixed runs extract correctly")


if __name__ == '__main__':
    print("Testing mixed text+drawing extraction in _collect_runs_info_original\n")
    print("=" * 70)
    
    test_mixed_text_and_drawing()
    print()
    test_drawing_only_run()
    print()
    test_text_only_run()
    print()
    test_multiple_mixed_runs()
    
    print("=" * 70)
    print("\n✅ All tests passed!")
