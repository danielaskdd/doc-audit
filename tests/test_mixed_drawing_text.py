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
    Test that runs containing both text and drawing have both extracted.
    
    Before fix: Text would be silently dropped when w:drawing was found
    After fix: Both text and drawing placeholder should be extracted
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
    
    # Should have exactly 1 run record (all content in same w:r element)
    assert len(runs_info) == 1, f"Expected 1 run, got {len(runs_info)}"
    
    # Should contain both text parts and drawing placeholder
    expected_text = 'See image <drawing id="1" name="Picture 1" /> for details'
    assert combined_text == expected_text, \
        f"Expected '{expected_text}', got '{combined_text}'"
    
    # Run should be marked as containing a drawing
    assert runs_info[0]['is_drawing'] == True, \
        "Run should be marked as containing drawing"
    
    print("✓ Test passed: Mixed text+drawing run extracts both components")


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
    
    assert len(runs_info) == 3, f"Expected 3 runs, got {len(runs_info)}"
    
    expected_text = 'First runSecond with <drawing id="3" name="Icon" />Third run'
    assert combined_text == expected_text, \
        f"Expected '{expected_text}', got '{combined_text}'"
    
    # Check individual runs
    assert runs_info[0].get('is_drawing') != True, "First run should not have drawing"
    assert runs_info[1]['is_drawing'] == True, "Second run should have drawing"
    assert runs_info[2].get('is_drawing') != True, "Third run should not have drawing"
    
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
