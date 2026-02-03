#!/usr/bin/env python3
"""
Unit tests for superscript/subscript handling across parse and apply phases.

Tests the complete workflow:
1. Extract: parse_document.py extracts text with <sup>/<sub> markup
2. Match: apply_audit_edits.py matches text with markup 
3. Apply: generate proper Word XML with w:vertAlign
"""

import sys
from pathlib import Path
from lxml import etree

# Add skills directory to path for imports
sys.path.insert(0, str(Path(__file__).parent.parent / 'skills' / 'doc-audit' / 'scripts'))

from parse_document import extract_text_from_run  # noqa: E402  # type: ignore
from apply_audit_edits import AuditEditApplier  # noqa: E402  # type: ignore

# Namespaces for XML construction
NS = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
}

def qn(tag):
    """Qualified name helper for XML construction"""
    if ':' in tag:
        prefix, local = tag.split(':')
        return f'{{{NS[prefix]}}}{local}'
    return tag


# ============================================================
# Test Fixtures: XML Construction Helpers
# ============================================================

def create_run_with_text(text, vert_align=None):
    """
    Create a w:r element with text and optional vertical alignment.
    
    Args:
        text: Text content
        vert_align: 'superscript' | 'subscript' | None
    
    Returns:
        w:r element
    """
    run = etree.Element(qn('w:r'))
    
    if vert_align:
        rPr = etree.SubElement(run, qn('w:rPr'))
        vertAlign = etree.SubElement(rPr, qn('w:vertAlign'))
        vertAlign.set(qn('w:val'), vert_align)
    
    t = etree.SubElement(run, qn('w:t'))
    t.text = text
    
    return run


# ============================================================
# Phase 1 Tests: Extraction
# ============================================================

def test_extract_superscript_basic():
    """Test basic superscript extraction: x² → x<sup>2</sup>"""
    # Create run: "x"
    run_x = create_run_with_text('x')
    text_x = extract_text_from_run(run_x, NS)
    assert text_x == 'x', f"Expected 'x', got '{text_x}'"
    
    # Create run: "2" with superscript
    run_2 = create_run_with_text('2', vert_align='superscript')
    text_2 = extract_text_from_run(run_2, NS)
    assert text_2 == '<sup>2</sup>', f"Expected '<sup>2</sup>', got '{text_2}'"
    
    # Combined: x²
    combined = text_x + text_2
    assert combined == 'x<sup>2</sup>', f"Expected 'x<sup>2</sup>', got '{combined}'"
    
    print("✓ test_extract_superscript_basic passed")


def test_extract_subscript_basic():
    """Test basic subscript extraction: H₂O → H<sub>2</sub>O"""
    run_h = create_run_with_text('H')
    run_2 = create_run_with_text('2', vert_align='subscript')
    run_o = create_run_with_text('O')
    
    text_h = extract_text_from_run(run_h, NS)
    text_2 = extract_text_from_run(run_2, NS)
    text_o = extract_text_from_run(run_o, NS)
    
    combined = text_h + text_2 + text_o
    assert combined == 'H<sub>2</sub>O', f"Expected 'H<sub>2</sub>O', got '{combined}'"
    
    print("✓ test_extract_subscript_basic passed")


def test_extract_mixed_script():
    """Test mixed superscript and subscript: CO₂²⁺ → CO<sub>2</sub><sup>2+</sup>"""
    run_co = create_run_with_text('CO')
    run_2_sub = create_run_with_text('2', vert_align='subscript')
    run_2plus_sup = create_run_with_text('2+', vert_align='superscript')
    
    text_co = extract_text_from_run(run_co, NS)
    text_2_sub = extract_text_from_run(run_2_sub, NS)
    text_2plus_sup = extract_text_from_run(run_2plus_sup, NS)
    
    combined = text_co + text_2_sub + text_2plus_sup
    expected = 'CO<sub>2</sub><sup>2+</sup>'
    assert combined == expected, f"Expected '{expected}', got '{combined}'"
    
    print("✓ test_extract_mixed_script passed")


def test_extract_no_format():
    """Test that normal text is unchanged"""
    run = create_run_with_text('normal text')
    text = extract_text_from_run(run, NS)
    assert text == 'normal text', f"Expected 'normal text', got '{text}'"
    
    print("✓ test_extract_no_format passed")


# ============================================================
# Phase 2 Tests: Matching
# ============================================================

def test_match_superscript_in_collect():
    """Test that _collect_runs_info_original extracts with markup"""
    # Create a paragraph with x²
    para = etree.Element(qn('w:p'))
    para.append(create_run_with_text('x'))
    para.append(create_run_with_text('2', vert_align='superscript'))
    
    # Create applier instance (minimal setup)
    # We'll directly test the method without full initialization
    applier = object.__new__(AuditEditApplier)
    applier.verbose = False
    
    # Call _collect_runs_info_original
    runs_info, combined_text = applier._collect_runs_info_original(para)
    
    assert combined_text == 'x<sup>2</sup>', f"Expected 'x<sup>2</sup>', got '{combined_text}'"
    assert len(runs_info) == 2, f"Expected 2 runs, got {len(runs_info)}"
    
    print("✓ test_match_superscript_in_collect passed")


def test_match_subscript_in_collect():
    """Test subscript extraction in _collect_runs_info_original"""
    # Create paragraph with H₂O
    para = etree.Element(qn('w:p'))
    para.append(create_run_with_text('H'))
    para.append(create_run_with_text('2', vert_align='subscript'))
    para.append(create_run_with_text('O'))
    
    applier = object.__new__(AuditEditApplier)
    applier.verbose = False
    
    runs_info, combined_text = applier._collect_runs_info_original(para)
    
    assert combined_text == 'H<sub>2</sub>O', f"Expected 'H<sub>2</sub>O', got '{combined_text}'"
    assert len(runs_info) == 3, f"Expected 3 runs, got {len(runs_info)}"
    
    print("✓ test_match_subscript_in_collect passed")


# ============================================================
# Phase 3 Tests: Apply (XML Generation)
# ============================================================

def test_parse_formatted_text_superscript():
    """Test _parse_formatted_text with superscript markup"""
    applier = object.__new__(AuditEditApplier)
    
    segments = applier._parse_formatted_text('x<sup>2</sup>')
    assert len(segments) == 2, f"Expected 2 segments, got {len(segments)}"
    assert segments[0] == ('x', None), f"Expected ('x', None), got {segments[0]}"
    assert segments[1] == ('2', 'superscript'), f"Expected ('2', 'superscript'), got {segments[1]}"
    
    print("✓ test_parse_formatted_text_superscript passed")


def test_parse_formatted_text_subscript():
    """Test _parse_formatted_text with subscript markup"""
    applier = object.__new__(AuditEditApplier)
    
    segments = applier._parse_formatted_text('H<sub>2</sub>O')
    assert len(segments) == 3, f"Expected 3 segments, got {len(segments)}"
    assert segments[0] == ('H', None)
    assert segments[1] == ('2', 'subscript')
    assert segments[2] == ('O', None)
    
    print("✓ test_parse_formatted_text_subscript passed")


def test_parse_formatted_text_mixed():
    """Test _parse_formatted_text with mixed formats"""
    applier = object.__new__(AuditEditApplier)
    
    segments = applier._parse_formatted_text('CO<sub>2</sub><sup>2+</sup>')
    assert len(segments) == 3, f"Expected 3 segments, got {len(segments)}"
    assert segments[0] == ('CO', None)
    assert segments[1] == ('2', 'subscript')
    assert segments[2] == ('2+', 'superscript')
    
    print("✓ test_parse_formatted_text_mixed passed")


def test_parse_formatted_text_no_markup():
    """Test _parse_formatted_text with plain text"""
    applier = object.__new__(AuditEditApplier)
    
    segments = applier._parse_formatted_text('normal text')
    assert len(segments) == 1, f"Expected 1 segment, got {len(segments)}"
    assert segments[0] == ('normal text', None)
    
    print("✓ test_parse_formatted_text_no_markup passed")


def test_create_run_simple():
    """Test _create_run with plain text"""
    applier = object.__new__(AuditEditApplier)
    
    run = applier._create_run('test')
    assert run.tag == qn('w:r'), f"Expected w:r tag, got {run.tag}"
    
    t_elem = run.find(qn('w:t'))
    assert t_elem is not None, "Expected w:t element"
    assert t_elem.text == 'test', f"Expected 'test', got '{t_elem.text}'"
    
    print("✓ test_create_run_simple passed")


def test_create_run_superscript():
    """Test _create_run with superscript markup generates proper XML"""
    applier = object.__new__(AuditEditApplier)
    
    result = applier._create_run('x<sup>2</sup>')
    
    # Should return a container with 2 runs
    assert result.tag == 'container', f"Expected container, got {result.tag}"
    runs = list(result)
    assert len(runs) == 2, f"Expected 2 runs, got {len(runs)}"
    
    # First run: "x" (normal)
    run1 = runs[0]
    t1 = run1.find(qn('w:t'))
    assert t1.text == 'x', f"Expected 'x', got '{t1.text}'"
    
    # Second run: "2" (superscript)
    run2 = runs[1]
    t2 = run2.find(qn('w:t'))
    assert t2.text == '2', f"Expected '2', got '{t2.text}'"
    
    # Check w:vertAlign in second run
    rPr2 = run2.find(qn('w:rPr'))
    assert rPr2 is not None, "Expected w:rPr in superscript run"
    
    vertAlign = rPr2.find(qn('w:vertAlign'))
    assert vertAlign is not None, "Expected w:vertAlign element"
    assert vertAlign.get(qn('w:val')) == 'superscript', "Expected superscript value"
    
    print("✓ test_create_run_superscript passed")


def test_create_run_subscript():
    """Test _create_run with subscript markup"""
    applier = object.__new__(AuditEditApplier)
    
    result = applier._create_run('H<sub>2</sub>O')
    
    assert result.tag == 'container'
    runs = list(result)
    assert len(runs) == 3, f"Expected 3 runs, got {len(runs)}"
    
    # Check middle run has subscript
    run2 = runs[1]
    t2 = run2.find(qn('w:t'))
    assert t2.text == '2'
    
    rPr2 = run2.find(qn('w:rPr'))
    vertAlign = rPr2.find(qn('w:vertAlign'))
    assert vertAlign.get(qn('w:val')) == 'subscript', "Expected subscript value"
    
    print("✓ test_create_run_subscript passed")


# ============================================================
# Integration Tests
# ============================================================

def test_roundtrip_superscript():
    """Test complete roundtrip: extract → parse → generate"""
    # 1. Extract phase: simulate Word document
    run_x = create_run_with_text('x')
    run_2 = create_run_with_text('2', vert_align='superscript')
    
    extracted = extract_text_from_run(run_x, NS) + extract_text_from_run(run_2, NS)
    assert extracted == 'x<sup>2</sup>', f"Extract phase failed: {extracted}"
    
    # 2. Parse phase: simulate LLM processing
    applier = object.__new__(AuditEditApplier)
    segments = applier._parse_formatted_text(extracted)
    assert len(segments) == 2, f"Parse phase failed: {len(segments)} segments"
    
    # 3. Generate phase: create Word XML
    result = applier._create_run(extracted)
    assert result.tag == 'container', "Generate phase failed: not a container"
    
    runs = list(result)
    assert len(runs) == 2, f"Generate phase failed: {len(runs)} runs"
    
    # Verify second run has superscript
    rPr = runs[1].find(qn('w:rPr'))
    vertAlign = rPr.find(qn('w:vertAlign'))
    assert vertAlign.get(qn('w:val')) == 'superscript', "Roundtrip failed: no superscript"
    
    print("✓ test_roundtrip_superscript passed")


def test_roundtrip_subscript():
    """Test complete roundtrip for subscript"""
    # Extract
    runs = [
        create_run_with_text('H'),
        create_run_with_text('2', vert_align='subscript'),
        create_run_with_text('O')
    ]
    extracted = ''.join(extract_text_from_run(r, NS) for r in runs)
    assert extracted == 'H<sub>2</sub>O', f"Extract failed: {extracted}"
    
    # Generate
    applier = object.__new__(AuditEditApplier)
    result = applier._create_run(extracted)
    
    assert result.tag == 'container'
    gen_runs = list(result)
    assert len(gen_runs) == 3
    
    # Verify subscript
    rPr = gen_runs[1].find(qn('w:rPr'))
    vertAlign = rPr.find(qn('w:vertAlign'))
    assert vertAlign.get(qn('w:val')) == 'subscript'
    
    print("✓ test_roundtrip_subscript passed")


# ============================================================
# Edge Cases
# ============================================================

def test_empty_superscript():
    """Test empty superscript tag: x<sup></sup>y → skip empty segment"""
    applier = object.__new__(AuditEditApplier)
    
    segments = applier._parse_formatted_text('x<sup></sup>y')
    # Empty segments should be skipped in _create_run
    assert len(segments) == 3  # x, '', y
    assert segments[0] == ('x', None)
    assert segments[1] == ('', 'superscript')
    assert segments[2] == ('y', None)
    
    # But _create_run should skip empty segments
    result = applier._create_run('x<sup></sup>y')
    runs = list(result) if result.tag == 'container' else [result]
    # Should have 2 runs (x and y), empty superscript skipped
    assert len(runs) == 2, f"Expected 2 runs (empty skipped), got {len(runs)}"
    
    print("✓ test_empty_superscript passed")


def test_nested_markup_not_supported():
    """Test that nested markup is not supported (invalid input)"""
    applier = object.__new__(AuditEditApplier)
    
    # This is invalid markup, but parser should handle it
    segments = applier._parse_formatted_text('x<sup>2<sub>3</sub></sup>')
    # The non-greedy regex will match the first complete tag pair it finds
    # So <sub>3</sub> is matched first, leaving <sup>2</sup> around it
    # Actually the regex matches <sup>2<sub>3</sub></sup> as a whole superscript
    assert len(segments) == 2  # 'x' and '2<sub>3</sub>'
    assert segments[0] == ('x', None)
    # The entire content between <sup></sup> is captured, including nested tags
    assert segments[1] == ('2<sub>3</sub>', 'superscript')
    
    print("✓ test_nested_markup_not_supported passed")


# ============================================================
# Test Runner
# ============================================================

def run_all_tests():
    """Run all tests and report results"""
    print("=" * 60)
    print("Running Superscript/Subscript Unit Tests")
    print("=" * 60)
    
    tests = [
        # Phase 1: Extraction
        ("Phase 1: Extract", [
            test_extract_superscript_basic,
            test_extract_subscript_basic,
            test_extract_mixed_script,
            test_extract_no_format,
        ]),
        # Phase 2: Matching
        ("Phase 2: Match", [
            test_match_superscript_in_collect,
            test_match_subscript_in_collect,
        ]),
        # Phase 3: Apply
        ("Phase 3: Apply", [
            test_parse_formatted_text_superscript,
            test_parse_formatted_text_subscript,
            test_parse_formatted_text_mixed,
            test_parse_formatted_text_no_markup,
            test_create_run_simple,
            test_create_run_superscript,
            test_create_run_subscript,
        ]),
        # Integration
        ("Integration Tests", [
            test_roundtrip_superscript,
            test_roundtrip_subscript,
        ]),
        # Edge Cases
        ("Edge Cases", [
            test_empty_superscript,
            test_nested_markup_not_supported,
        ]),
    ]
    
    total_tests = 0
    passed_tests = 0
    failed_tests = []
    
    for phase_name, phase_tests in tests:
        print(f"\n{phase_name}:")
        print("-" * 60)
        
        for test_func in phase_tests:
            total_tests += 1
            try:
                test_func()
                passed_tests += 1
            except AssertionError as e:
                failed_tests.append((test_func.__name__, str(e)))
                print(f"✗ {test_func.__name__} FAILED: {e}")
            except Exception as e:
                failed_tests.append((test_func.__name__, f"Error: {e}"))
                print(f"✗ {test_func.__name__} ERROR: {e}")
    
    # Summary
    print("\n" + "=" * 60)
    print("Test Summary")
    print("=" * 60)
    print(f"Total: {total_tests}")
    print(f"Passed: {passed_tests}")
    print(f"Failed: {len(failed_tests)}")
    
    if failed_tests:
        print("\nFailed Tests:")
        for test_name, error in failed_tests:
            print(f"  - {test_name}: {error}")
        return 1
    else:
        print("\n✓ All tests passed!")
        return 0


if __name__ == '__main__':
    sys.exit(run_all_tests())
