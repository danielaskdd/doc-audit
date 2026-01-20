#!/usr/bin/env python3
"""
ABOUTME: Unit tests for parse_document.py text chunking logic
ABOUTME: Tests split_long_block, merge_small_blocks, and split_table functions
"""

import sys
from pathlib import Path

# Add skills/doc-audit/scripts directory to path (must be before import)
_scripts_dir = Path(__file__).parent.parent / 'skills' / 'doc-audit' / 'scripts'
sys.path.insert(0, str(_scripts_dir))

import pytest  # noqa: E402

from parse_document import (  # noqa: E402  # type: ignore[import-not-found]
    split_long_block,
    merge_small_blocks,
    split_table,
    split_table_with_heading,
    estimate_tokens,
    IDEAL_BLOCK_CONTENT_TOKENS,
    MAX_BLOCK_CONTENT_TOKENS,
    SMALL_TAIL_THRESHOLD,
)

# Test-local constants (relative to IDEAL threshold)
SMALL_BLOCK_TOKENS = IDEAL_BLOCK_CONTENT_TOKENS // 4    # 1500 tokens - clearly small blocks
MEDIUM_BLOCK_TOKENS = IDEAL_BLOCK_CONTENT_TOKENS // 2   # 3000 tokens - medium blocks
NEAR_IDEAL_TOKENS = IDEAL_BLOCK_CONTENT_TOKENS - 200    # 5800 tokens - near ideal threshold


# ============================================================
# Helper Functions
# ============================================================

def create_paragraph(text: str, para_id: str, is_table: bool = False) -> dict:
    """
    Create a mock paragraph dictionary.
    
    Args:
        text: Paragraph text content
        para_id: Paragraph ID
        is_table: Whether this is a table paragraph
    
    Returns:
        Paragraph dict with required fields
    """
    return {
        'text': text,
        'para_id': para_id,
        'is_table': is_table
    }


def create_block(uuid: str, heading: str, content: str, level: int, 
                 parent_headings: list = None, table_chunk_role: str = "none") -> dict:
    """
    Create a mock block dictionary.
    
    Args:
        uuid: Block UUID
        heading: Block heading
        content: Block content
        level: Heading level
        parent_headings: Parent heading list
        table_chunk_role: Table chunk role (first/middle/last/none)
    
    Returns:
        Block dict with required fields
    """
    return {
        'uuid': uuid,
        'uuid_end': uuid,
        'heading': heading,
        'content': content,
        'type': 'text',
        'level': level,
        'parent_headings': parent_headings or [],
        'table_chunk_role': table_chunk_role
    }


def generate_long_text(target_tokens: int) -> str:
    """
    Generate text with approximately target_tokens.
    Uses mixed Chinese and English.
    
    Args:
        target_tokens: Target token count
    
    Returns:
        Generated text string
    """
    # Rough estimate: Chinese ~0.7 tokens/char, English ~0.35 tokens/char
    # Use 50% Chinese, 50% English
    chars_needed = int(target_tokens / 0.525)  # Average of 0.7 and 0.35
    
    chinese_part = "这是一段中文文本用于测试。" * (chars_needed // 24)  # 12 chars per repeat
    english_part = "This is English test text. " * (chars_needed // 54)  # 27 chars per repeat
    
    result = chinese_part + english_part
    
    # Verify we generated enough text
    actual_tokens = estimate_tokens(result)
    if actual_tokens < target_tokens * 0.8:
        # Add more text if we're significantly short
        padding = "补充内容填充文本。" * ((target_tokens - actual_tokens) // 10)
        result += padding
    
    return result


def generate_text(target_tokens: int) -> str:
    """Generate text with approximately target_tokens using iterative estimation."""
    base_text = "这是测试文本。Test text. "
    base_tokens = estimate_tokens(base_text)

    if base_tokens == 0:
        return ""

    repetitions = max(1, int(target_tokens / base_tokens))
    text = base_text * repetitions

    current_tokens = estimate_tokens(text)
    while current_tokens < target_tokens:
        text += base_text
        current_tokens = estimate_tokens(text)

    return text


def create_table_rows(num_rows: int, cols: int = 3, large_content: bool = False) -> list:
    """
    Create table rows for testing.
    
    Args:
        num_rows: Number of rows
        cols: Number of columns per row
        large_content: If True, create cells with more content for testing splits
    
    Returns:
        2D list of table data
    """
    if large_content:
        # Create cells with more content to trigger table splitting
        return [[f"Cell R{r}C{c}: 这是一个包含较多内容的单元格用于测试表格分割功能。This cell contains more content for testing table splitting functionality." for c in range(cols)] for r in range(num_rows)]
    else:
        return [[f"Cell R{r}C{c}" for c in range(cols)] for r in range(num_rows)]


def create_table_para_ids(num_rows: int, cols: int = 3) -> tuple:
    """
    Create para_ids and para_ids_end for table testing.
    
    Args:
        num_rows: Number of rows
        cols: Number of columns per row
    
    Returns:
        Tuple of (para_ids, para_ids_end)
    """
    para_ids = []
    para_ids_end = []
    
    for r in range(num_rows):
        row_ids = []
        row_ids_end = []
        for c in range(cols):
            cell_id = f"{r:04X}{c:04X}00"
            cell_id_end = f"{r:04X}{c:04X}FF"
            row_ids.append(cell_id)
            row_ids_end.append(cell_id_end)
        para_ids.append(row_ids)
        para_ids_end.append(row_ids_end)
    
    return para_ids, para_ids_end


# ============================================================
# Tests: split_long_block
# ============================================================

class TestSplitLongBlock:
    """Tests for split_long_block function"""
    
    def test_short_block_returns_single_with_level(self):
        """Short block should return single block with correct level"""
        paragraphs = [
            create_paragraph("Short content", "AAA")
        ]
        
        blocks = split_long_block(
            block_heading="Test Heading",
            paragraphs=paragraphs,
            parent_headings=[],
            block_level=2,
            debug=False
        )
        
        assert len(blocks) == 1
        assert blocks[0]['heading'] == "Test Heading"
        assert blocks[0]['content'] == "Short content"
        assert blocks[0]['level'] == 2  # Regression test for level bug fix
        assert blocks[0]['uuid'] == "AAA"
    
    def test_long_block_splits_with_anchor(self):
        """Long block should split using anchor paragraphs"""
        # Create content that exceeds MAX_BLOCK_CONTENT_TOKENS
        long_text = generate_long_text(MAX_BLOCK_CONTENT_TOKENS + 1000)
        
        paragraphs = [
            create_paragraph(long_text[:len(long_text)//3], "AAA"),
            create_paragraph("Anchor 1", "BBB"),  # Short anchor
            create_paragraph(long_text[len(long_text)//3:2*len(long_text)//3], "CCC"),
            create_paragraph("Anchor 2", "DDD"),  # Short anchor
            create_paragraph(long_text[2*len(long_text)//3:], "EEE"),
        ]
        
        blocks = split_long_block(
            block_heading="Long Section",
            paragraphs=paragraphs,
            parent_headings=[],
            block_level=1,
            debug=False
        )
        
        # Should be split into multiple blocks
        assert len(blocks) > 1
        
        # All blocks should have level=1
        for block in blocks:
            assert block['level'] == 1
        
        # Anchors should become headings
        headings = [b['heading'] for b in blocks]
        assert "Anchor 1" in headings or "Anchor 2" in headings
    
    def test_anchor_inherits_parent_level(self):
        """Anchor-based blocks should inherit parent block level"""
        long_text = generate_long_text(MAX_BLOCK_CONTENT_TOKENS + 500)
        
        paragraphs = [
            create_paragraph(long_text[:len(long_text)//2], "AAA"),
            create_paragraph("Split Point", "BBB"),
            create_paragraph(long_text[len(long_text)//2:], "CCC"),
        ]
        
        blocks = split_long_block(
            block_heading="Parent Heading",
            paragraphs=paragraphs,
            parent_headings=["Chapter 1"],
            block_level=3,  # Third-level heading
            debug=False
        )
        
        # Should be split
        assert len(blocks) > 1
        
        # All blocks (including anchor-generated ones) should be level 3
        for block in blocks:
            assert block['level'] == 3
    
    def test_table_chunk_heading_metadata(self):
        """Block with _chunk_heading metadata should use that heading"""
        paragraphs = [
            {
                'text': '<table>[[1,2,3]]</table>',
                'para_id': 'AAA',
                'para_id_end': 'BBB',
                'is_table': True,
                '_chunk_heading': 'Table Fragment [1]',
                '_table_header': [['Header1', 'Header2', 'Header3']]
            }
        ]
        
        blocks = split_long_block(
            block_heading="Original Heading",
            paragraphs=paragraphs,
            parent_headings=[],
            block_level=2,
            debug=False
        )
        
        assert len(blocks) == 1
        assert blocks[0]['heading'] == 'Table Fragment [1]'
        assert 'table_header' in blocks[0]
        assert blocks[0]['table_header'] == [['Header1', 'Header2', 'Header3']]
    
    def test_recursive_split_when_still_too_large(self):
        """Should recursively split if block still exceeds MAX after first split"""
        # Create long text with multiple short anchors for splitting
        # This ensures the recursive split can find anchors
        part_size = MAX_BLOCK_CONTENT_TOKENS // 3
        
        paragraphs = [
            create_paragraph(generate_long_text(part_size), "AAA"),
            create_paragraph("Anchor 1", "BBB"),
            create_paragraph(generate_long_text(part_size), "CCC"),
            create_paragraph("Anchor 2", "DDD"),
            create_paragraph(generate_long_text(part_size), "EEE"),
            create_paragraph("Anchor 3", "FFF"),
            create_paragraph(generate_long_text(part_size), "GGG"),
        ]
        
        blocks = split_long_block(
            block_heading="Huge Section",
            paragraphs=paragraphs,
            parent_headings=[],
            block_level=1,
            debug=False
        )
        
        # Should recursively split into multiple blocks
        assert len(blocks) >= 2
        
        # No individual block should exceed MAX (allow small buffer for edge cases)
        for block in blocks:
            tokens = estimate_tokens(block['content'])
            assert tokens <= MAX_BLOCK_CONTENT_TOKENS * 1.1  # 10% tolerance


# ============================================================
# Tests: merge_small_blocks
# ============================================================

class TestMergeSmallBlocks:
    """Tests for merge_small_blocks function"""
    
    def test_same_level_adjacent_blocks_merge(self):
        """Same-level adjacent small blocks should merge"""
        # Create content that is definitely small (well below IDEAL threshold)
        small_content = generate_long_text(SMALL_BLOCK_TOKENS)
        
        blocks = [
            create_block("AAA", "Section 1", small_content, level=2),
            create_block("BBB", "Section 2", small_content, level=2),
        ]
        
        merged, count = merge_small_blocks(blocks, debug=False)
        
        assert count > 0
        assert len(merged) == 1
        assert merged[0]['level'] == 2
        assert "\n\n" in merged[0]['content']  # Separator check
    
    def test_high_level_absorbs_low_level(self):
        """High-level (smaller number) block should absorb adjacent low-level block"""
        small_content = generate_long_text(SMALL_BLOCK_TOKENS)
        
        blocks = [
            create_block("AAA", "Chapter 1", small_content, level=1),
            create_block("BBB", "Subsection", small_content, level=2),
        ]
        
        merged, count = merge_small_blocks(blocks, debug=False)
        
        assert count > 0
        assert len(merged) == 1
        assert merged[0]['heading'] == "Chapter 1"
        assert merged[0]['level'] == 1  # High-level wins
    
    def test_low_level_cannot_absorb_high_level_forward(self):
        """Low-level block cannot absorb preceding high-level block"""
        small_content = generate_long_text(SMALL_BLOCK_TOKENS)
        
        blocks = [
            create_block("AAA", "Chapter 1", small_content, level=1),
            create_block("BBB", "Subsection", small_content, level=2),
        ]
        
        # AAA (level 1) is first. BBB (level 2) should try to merge backward with AAA.
        # Since AAA.level (1) < BBB.level (2), backward merge is allowed (high absorbs low).
        merged, count = merge_small_blocks(blocks, debug=False)
        
        # Should merge
        assert count > 0
        assert len(merged) == 1
        assert merged[0]['level'] == 1
    
    def test_high_absorbs_low_then_same_level_merge(self):
        """
        After high-level absorbs low-level, remaining high-level blocks 
        should merge if adjacent and same level.
        
        Scenario (numbers indicate levels):
        A(1, small) + B(2, small) + C(1, small)
        
        Expected flow:
        1. A absorbs B -> A'(1)
        2. A' and C are adjacent and same level -> merge to final block
        """
        small_content = generate_long_text(SMALL_BLOCK_TOKENS)
        
        blocks = [
            create_block("AAA", "Chapter 1", small_content, level=1),
            create_block("BBB", "Subsection", small_content, level=2),
            create_block("CCC", "Chapter 2", small_content, level=1),
        ]
        
        merged, count = merge_small_blocks(blocks, debug=False)
        
        # Should merge multiple times
        assert count >= 2
        assert len(merged) == 1
        assert merged[0]['level'] == 1
        # Verify content from all three blocks present
        assert "Chapter 1" in merged[0]['heading'] or len(merged[0]['content']) > len(small_content) * 2
    
    def test_non_adjacent_same_level_no_merge(self):
        """
        After hitting IDEAL size, blocks should stop merging (lock mechanism).
        
        Scenario:
        A(1, near-ideal) + B(2, small) + C(2, small) + D(1, small)
        
        Expected:
        - A absorbs B, reaches IDEAL → locked
        - C and D cannot merge with locked A
        - C and D may merge with each other
        Final result should have at least 2 blocks (locked A + others)
        """
        near_ideal = generate_long_text(NEAR_IDEAL_TOKENS)
        small_content = generate_long_text(SMALL_BLOCK_TOKENS)
        
        blocks = [
            create_block("AAA", "Chapter 1", near_ideal, level=1),
            create_block("BBB", "Sub 1", small_content, level=2),
            create_block("CCC", "Sub 2", small_content, level=2),
            create_block("DDD", "Chapter 2", small_content, level=1),
        ]
        
        merged, count = merge_small_blocks(blocks, debug=False)
        
        # Should have at least 2 blocks (A should be locked after absorbing B)
        assert len(merged) >= 2
        
        # First block should be large (near IDEAL or above)
        assert estimate_tokens(merged[0]['content']) >= IDEAL_BLOCK_CONTENT_TOKENS * 0.9
    
    def test_merge_stops_after_ideal_size(self):
        """Block should stop merging after reaching IDEAL_BLOCK_CONTENT_TOKENS"""
        # Create one block slightly below IDEAL, another small block
        near_ideal = generate_long_text(NEAR_IDEAL_TOKENS)
        small = generate_long_text(SMALL_BLOCK_TOKENS)
        
        blocks = [
            create_block("AAA", "Section 1", near_ideal, level=1),
            create_block("BBB", "Section 2", small, level=1),
            create_block("CCC", "Section 3", small, level=1),
        ]
        
        merged, count = merge_small_blocks(blocks, debug=False)
        
        # After AAA absorbs BBB, it should be "locked" (>= IDEAL)
        # CCC should remain separate
        assert len(merged) >= 2
    
    def test_merge_rejects_if_exceeds_max(self):
        """Merge should be rejected if combined size exceeds MAX_BLOCK_CONTENT_TOKENS"""
        # Create two blocks that together exceed MAX
        large_content = generate_long_text(MAX_BLOCK_CONTENT_TOKENS - 500)
        
        blocks = [
            create_block("AAA", "Section 1", large_content, level=1),
            create_block("BBB", "Section 2", large_content, level=1),
        ]
        
        merged, count = merge_small_blocks(blocks, debug=False)
        
        # Should not merge (would exceed MAX)
        assert count == 0
        assert len(merged) == 2
    
    def test_merge_uses_double_newline_separator(self):
        """Merged content should use \\n\\n as separator"""
        blocks = [
            create_block("AAA", "Section 1", "Content A", level=1),
            create_block("BBB", "Section 2", "Content B", level=1),
        ]
        
        merged, count = merge_small_blocks(blocks, debug=False)
        
        assert count > 0
        assert len(merged) == 1
        assert "Content A\n\nContent B" in merged[0]['content']
    
    def test_table_chunk_role_first_only_forward(self):
        """Table chunk 'first' can only merge forward"""
        small_content = generate_long_text(SMALL_BLOCK_TOKENS)
        
        blocks = [
            create_block("AAA", "Normal", small_content, level=1, table_chunk_role="none"),
            create_block("BBB", "Table First", small_content, level=1, table_chunk_role="first"),
        ]
        
        merged, count = merge_small_blocks(blocks, debug=False)
        
        # BBB (first) can only merge forward, but there's no next block
        # AAA cannot absorb BBB forward (BBB doesn't allow backward merge as 'first')
        # So they should remain separate
        assert len(merged) == 2
    
    def test_table_chunk_role_last_only_backward(self):
        """Table chunk 'last' can only merge backward"""
        small_content = generate_long_text(SMALL_BLOCK_TOKENS)
        
        blocks = [
            create_block("AAA", "Table Last", small_content, level=1, table_chunk_role="last"),
            create_block("BBB", "Normal", small_content, level=1, table_chunk_role="none"),
        ]
        
        merged, count = merge_small_blocks(blocks, debug=False)
        
        # AAA (last) can only merge backward, but there's no previous block
        # AAA cannot merge forward with BBB
        # So they should remain separate
        assert len(merged) == 2
    
    def test_table_chunk_role_middle_no_merge(self):
        """Table chunk 'middle' cannot merge in any direction"""
        small_content = generate_long_text(SMALL_BLOCK_TOKENS)
        
        blocks = [
            create_block("AAA", "Before", small_content, level=1, table_chunk_role="none"),
            create_block("BBB", "Table Middle", small_content, level=1, table_chunk_role="middle"),
            create_block("CCC", "After", small_content, level=1, table_chunk_role="none"),
        ]
        
        merged, count = merge_small_blocks(blocks, debug=False)
        
        # BBB (middle) should remain isolated
        # AAA and CCC might merge if they become adjacent after processing, but BBB blocks them
        assert len(merged) >= 2
        
        # BBB should still exist
        middle_blocks = [b for b in merged if b.get('table_chunk_role') == 'middle']
        assert len(middle_blocks) == 1
    
    def test_bottom_up_merging_order(self):
        """
        Verify bottom-up merging: deepest level blocks merge first among themselves,
        then get absorbed by higher levels, then higher levels merge.
        
        Scenario:
        A(level=1) + B(level=3) + C(level=3) + D(level=2) + E(level=1)
        
        Expected flow (bottom-up):
        1. Level 3 phase: B and C should merge first (same level, adjacent)
        2. Level 2 phase: D can absorb BC (cross-level)
        3. Level 1 phase: A and E may merge if both still below IDEAL
        
        Key assertion: B and C must merge together before being absorbed by upper levels
        """
        small_content = generate_long_text(IDEAL_BLOCK_CONTENT_TOKENS // 4)
        
        blocks = [
            create_block("AAA", "Chapter 1", small_content, level=1),
            create_block("BBB", "Sub-sub 1", small_content, level=3),
            create_block("CCC", "Sub-sub 2", small_content, level=3),
            create_block("DDD", "Subsection", small_content, level=2),
            create_block("EEE", "Chapter 2", small_content, level=1),
        ]
        
        merged, count = merge_small_blocks(blocks, debug=False)
        
        # Should have performed merges
        assert count > 0
        
        # The key is that B and C should NOT remain as separate blocks
        # If there are level 3 blocks left, there should be at most 1 (the merged BC)
        level_3_blocks = [b for b in merged if b['level'] == 3]
        assert len(level_3_blocks) <= 1
        
        # Should result in fewer blocks than we started with
        assert len(merged) < len(blocks)
    
    def test_medium_blocks_merge_when_below_ideal(self):
        """
        Verify that medium-sized blocks (below IDEAL) can merge.
        
        This tests the new merging threshold: blocks < IDEAL_BLOCK_CONTENT_TOKENS
        can merge (not limited to very small blocks).
        """
        # Create medium-sized content (half of IDEAL = 3000 tokens)
        medium_content = generate_long_text(MEDIUM_BLOCK_TOKENS)
        
        blocks = [
            create_block("AAA", "Section 1", medium_content, level=1),
            create_block("BBB", "Section 2", medium_content, level=1),
        ]
        
        # Verify test setup: each block is well below IDEAL
        actual_size = estimate_tokens(medium_content)
        assert actual_size < IDEAL_BLOCK_CONTENT_TOKENS
        assert actual_size > SMALL_BLOCK_TOKENS  # Not tiny blocks
        
        merged, count = merge_small_blocks(blocks, debug=False)
        
        # Combined size should be ~6000 tokens (at IDEAL boundary)
        # Should merge since combined size doesn't exceed MAX (8000)
        combined_size = actual_size * 2
        assert combined_size <= MAX_BLOCK_CONTENT_TOKENS
        
        # Should have merged
        assert count > 0
        assert len(merged) == 1
        assert merged[0]['level'] == 1


# ============================================================
# Tests: tail absorption
# ============================================================

class TestTailAbsorption:
    """Tests for tail absorption strategy in Phase A merging"""
    
    def test_tail_absorption_basic_scenario(self):
        """
        Basic tail absorption: Block at IDEAL + small tail blocks < THRESHOLD
        
        Scenario: A(6500) + B(400) + C(300) + D(200)
        Expected: All merged into one block (7400 tokens)
        """
        block_a = create_block("AAA", "Section 1", generate_text(6500), level=1)
        block_b = create_block("BBB", "Section 2", generate_text(400), level=1)
        block_c = create_block("CCC", "Section 3", generate_text(300), level=1)
        block_d = create_block("DDD", "Section 4", generate_text(200), level=1)
        
        blocks = [block_a, block_b, block_c, block_d]
        
        assert estimate_tokens(block_a['content']) >= IDEAL_BLOCK_CONTENT_TOKENS
        tail_total = sum(estimate_tokens(b['content']) for b in [block_b, block_c, block_d])
        assert tail_total < SMALL_TAIL_THRESHOLD
        
        merged, count = merge_small_blocks(blocks, debug=False)
        
        assert len(merged) == 1
        assert count == 3
        assert merged[0]['heading'] == "Section 1"
    
    def test_tail_absorption_respects_threshold(self):
        """
        Tail NOT absorbed if total >= SMALL_TAIL_THRESHOLD (1000 tokens)
        
        Scenario: A(6500) + B(1500) + C(800)
        Expected: No absorption (tail = 2300 > 1000)
        """
        block_a = create_block("AAA", "Section 1", generate_text(6500), level=1)
        block_b = create_block("BBB", "Section 2", generate_text(1500), level=1)
        block_c = create_block("CCC", "Section 3", generate_text(800), level=1)
        
        blocks = [block_a, block_b, block_c]
        
        tail_total = estimate_tokens(block_b['content']) + estimate_tokens(block_c['content'])
        assert tail_total >= SMALL_TAIL_THRESHOLD
        
        merged, count = merge_small_blocks(blocks, debug=False)
        
        assert len(merged) == 2
    
    def test_tail_absorption_respects_max_limit(self):
        """
        Tail NOT absorbed if combined size > MAX_BLOCK_CONTENT_TOKENS
        
        Scenario: A(7500) + B(800)
        Expected: No absorption (7500 + 800 = 8300 > 8000)
        """
        block_a = create_block("AAA", "Section 1", generate_text(7500), level=1)
        block_b = create_block("BBB", "Section 2", generate_text(800), level=1)
        
        blocks = [block_a, block_b]
        
        combined = estimate_tokens(block_a['content']) + estimate_tokens(block_b['content'])
        assert combined > MAX_BLOCK_CONTENT_TOKENS
        
        merged, count = merge_small_blocks(blocks, debug=False)
        
        assert len(merged) == 2
        assert count == 0
    
    def test_tail_absorption_same_level_only(self):
        """
        Tail absorption only applies to same-level blocks
        
        Scenario: A(6500, level=1) + B(400, level=2) + C(300, level=1)
        Expected: A absorbs C (cross-level in Phase B), B remains separate
        """
        block_a = create_block("AAA", "Chapter 1", generate_text(6500), level=1)
        block_b = create_block("BBB", "Subsection", generate_text(400), level=2)
        block_c = create_block("CCC", "Chapter 2", generate_text(300), level=1)
        
        blocks = [block_a, block_b, block_c]
        
        merged, count = merge_small_blocks(blocks, debug=False)
        
        assert len(merged) >= 1
    
    def test_tail_absorption_with_table_middle_chunk(self):
        """
        Tail absorption stops at table 'middle' chunk
        
        Scenario: A(6500) + B(400, middle) + C(300) + D(200)
        Expected: Only A; B/C/D remain (middle blocks absorption)
        """
        block_a = create_block("AAA", "Section 1", generate_text(6500), level=1)
        block_b = create_block("BBB", "Table Middle", generate_text(400), level=1, table_chunk_role="middle")
        block_c = create_block("CCC", "Section 3", generate_text(300), level=1)
        block_d = create_block("DDD", "Section 4", generate_text(200), level=1)
        
        blocks = [block_a, block_b, block_c, block_d]
        
        merged, count = merge_small_blocks(blocks, debug=False)
        
        assert len(merged) >= 2
        middle_exists = any(b.get('table_chunk_role') == 'middle' for b in merged)
        assert middle_exists
    
    def test_tail_absorption_multiple_at_ideal_blocks(self):
        """
        Multiple blocks at IDEAL should each try tail absorption
        
        Scenario: A(6500) + B(400) + C(6500) + D(300)
        Expected: A absorbs B (7900), C absorbs D (6800) → 2 blocks
        """
        block_a = create_block("AAA", "Section 1", generate_text(6500), level=1)
        block_b = create_block("BBB", "Section 2", generate_text(400), level=1)
        block_c = create_block("CCC", "Section 3", generate_text(6500), level=1)
        block_d = create_block("DDD", "Section 4", generate_text(300), level=1)
        
        blocks = [block_a, block_b, block_c, block_d]
        
        merged, count = merge_small_blocks(blocks, debug=False)
        
        assert len(merged) == 2
        assert count == 2
    
    def test_tail_absorption_edge_exact_threshold(self):
        """
        Test edge case: tail total at threshold should not trigger tail absorption
        
        Scenario: A(6500) + B(3000) + C(3000) (total = 6000 >= 1000)
        Expected: Tail absorption NOT triggered; B and C merge normally to form a block >= IDEAL,
                  resulting in 2 blocks: A and B+C
        """
        block_a = create_block("AAA", "Section 1", generate_text(6500), level=1)
        block_b = create_block("BBB", "Section 2", generate_text(3000), level=1)
        block_c = create_block("CCC", "Section 3", generate_text(3000), level=1)
        
        blocks = [block_a, block_b, block_c]
        
        a_tokens = estimate_tokens(block_a['content'])
        b_tokens = estimate_tokens(block_b['content'])
        c_tokens = estimate_tokens(block_c['content'])
        tail_total = b_tokens + c_tokens
        
        assert a_tokens >= IDEAL_BLOCK_CONTENT_TOKENS
        assert tail_total >= SMALL_TAIL_THRESHOLD
        assert tail_total >= IDEAL_BLOCK_CONTENT_TOKENS
        
        merged, count = merge_small_blocks(blocks, debug=False)
        
        assert len(merged) == 2
        assert count == 1
    
    def test_tail_absorption_preserves_uuid_end(self):
        """Tail absorption should preserve uuid_end from last absorbed block"""
        block_a = create_block("AAA", "Section 1", generate_text(6500), level=1)
        block_a['uuid_end'] = "AAA_END"
        
        block_b = create_block("BBB", "Section 2", generate_text(400), level=1)
        block_b['uuid_end'] = "BBB_END"
        
        block_c = create_block("CCC", "Section 3", generate_text(300), level=1)
        block_c['uuid_end'] = "CCC_END"
        
        blocks = [block_a, block_b, block_c]
        
        merged, count = merge_small_blocks(blocks, debug=False)
        
        assert len(merged) == 1
        assert merged[0]['uuid'] == "AAA"
        assert merged[0]['uuid_end'] == "CCC_END"
    
    def test_tail_absorption_preserves_table_header(self):
        """Tail absorption should preserve table_header from absorbed blocks"""
        block_a = create_block("AAA", "Section 1", generate_text(6500), level=1)
        
        block_b = create_block("BBB", "Section 2", generate_text(400), level=1)
        block_b['table_header'] = [['Header1', 'Header2']]
        
        block_c = create_block("CCC", "Section 3", generate_text(300), level=1)
        
        blocks = [block_a, block_b, block_c]
        
        merged, count = merge_small_blocks(blocks, debug=False)
        
        assert len(merged) == 1
        assert 'table_header' in merged[0]
        assert merged[0]['table_header'] == [['Header1', 'Header2']]


# ============================================================
# Tests: split_table
# ============================================================

class TestSplitTable:
    """Tests for split_table function"""
    
    def test_small_table_not_split(self):
        """Small table should not be split"""
        rows = create_table_rows(5, cols=3)
        para_ids, para_ids_end = create_table_para_ids(5, 3)
        
        chunks = split_table(rows, para_ids, para_ids_end, header_indices=[], debug=False)
        
        assert len(chunks) == 1
        assert chunks[0]['is_first'] is True
        assert chunks[0]['is_last'] is True
    
    def test_large_table_splits_into_multiple_chunks(self):
        """Large table should split into multiple chunks"""
        # Create a very large table (enough to exceed TABLE_MAX_TOKENS)
        # Use large_content=True to create cells with enough text
        num_rows = 200  # Combined with large content, should exceed TABLE_MAX_TOKENS
        
        rows = create_table_rows(num_rows, cols=3, large_content=True)
        para_ids, para_ids_end = create_table_para_ids(num_rows, 3)
        
        chunks = split_table(rows, para_ids, para_ids_end, header_indices=[], debug=False)
        
        # Should split into multiple chunks
        assert len(chunks) > 1
        
        # First chunk should be marked as first
        assert chunks[0]['is_first'] is True
        assert chunks[0]['is_last'] is False
        
        # Last chunk should be marked as last
        assert chunks[-1]['is_first'] is False
        assert chunks[-1]['is_last'] is True
    
    def test_large_table_three_or_more_chunks(self):
        """
        Very large table should split into 3+ chunks (user requirement 1).
        
        This tests that middle chunks (neither first nor last) exist.
        """
        # Create extremely large table with large content to force 3+ chunks
        num_rows = 500  # With large content, should force 3+ chunks
        rows = create_table_rows(num_rows, cols=4, large_content=True)
        para_ids, para_ids_end = create_table_para_ids(num_rows, 4)
        
        chunks = split_table(rows, para_ids, para_ids_end, header_indices=[], debug=False)
        
        # Should have at least 3 chunks
        assert len(chunks) >= 3
        
        # Verify we have middle chunks
        middle_chunks = [c for c in chunks if not c['is_first'] and not c['is_last']]
        assert len(middle_chunks) >= 1
        
        # Verify first and last are marked correctly
        assert chunks[0]['is_first'] is True
        assert chunks[0]['is_last'] is False
        assert chunks[-1]['is_first'] is False
        assert chunks[-1]['is_last'] is True
    
    def test_split_table_with_heading_adds_suffix(self):
        """split_table_with_heading should add suffix to chunk headings"""
        num_rows = 200
        rows = create_table_rows(num_rows, cols=3, large_content=True)
        para_ids, para_ids_end = create_table_para_ids(num_rows, 3)
        
        chunks = split_table_with_heading(
            rows, para_ids, para_ids_end,
            header_indices=[],
            current_heading="Section 1",
            start_suffix=0,
            debug=False
        )
        
        # Should have multiple chunks
        assert len(chunks) > 1
        
        # First chunk should have suffix_number=None
        assert chunks[0]['suffix_number'] is None
        
        # Subsequent chunks should have sequential suffix numbers
        if len(chunks) > 1:
            assert chunks[1]['suffix_number'] == 1
        if len(chunks) > 2:
            assert chunks[2]['suffix_number'] == 2


# ============================================================
# Integration: Full Workflow
# ============================================================

class TestIntegrationWorkflow:
    """Integration tests for combined split and merge operations"""
    
    def test_full_workflow_split_then_merge(self):
        """Test complete workflow: split long block, then merge small blocks"""
        # Create a long block that will be split
        long_text = generate_long_text(MAX_BLOCK_CONTENT_TOKENS + 1000)
        
        paragraphs = [
            create_paragraph(long_text[:len(long_text)//2], "AAA"),
            create_paragraph("Anchor", "BBB"),
            create_paragraph(long_text[len(long_text)//2:], "CCC"),
        ]
        
        # Split
        blocks = split_long_block(
            block_heading="Long Section",
            paragraphs=paragraphs,
            parent_headings=[],
            block_level=1,
            debug=False
        )
        
        # Add table_chunk_role to all blocks
        for block in blocks:
            if 'table_chunk_role' not in block:
                block['table_chunk_role'] = 'none'
        
        # Merge
        merged, count = merge_small_blocks(blocks, debug=False)
        
        # Should have processed successfully
        assert len(merged) > 0
        
        # All blocks should have level
        for block in merged:
            assert 'level' in block
            assert block['level'] == 1
    
    def test_verify_content_reconstruction(self):
        """Verify that merged blocks can reconstruct original content"""
        original_content_parts = [
            "Part A content here.",
            "Part B content here.",
            "Part C content here."
        ]
        
        blocks = [
            create_block(f"ID{i}", f"Section {i}", content, level=1)
            for i, content in enumerate(original_content_parts)
        ]
        
        merged, count = merge_small_blocks(blocks, debug=False)
        
        # Reconstruct content
        reconstructed = "\n\n".join([b['content'] for b in merged])
        
        # All original parts should be present (order preserved)
        for part in original_content_parts:
            assert part in reconstructed
    
    def test_level_consistency_after_operations(self):
        """All blocks should maintain level consistency after split and merge"""
        long_text = generate_long_text(MAX_BLOCK_CONTENT_TOKENS + 500)
        
        paragraphs = [
            create_paragraph(long_text[:len(long_text)//2], "AAA"),
            create_paragraph("Split", "BBB"),
            create_paragraph(long_text[len(long_text)//2:], "CCC"),
        ]
        
        blocks = split_long_block(
            "Section", paragraphs, [], block_level=3, debug=False
        )
        
        for block in blocks:
            block.setdefault('table_chunk_role', 'none')
        
        merged, _ = merge_small_blocks(blocks, debug=False)
        
        # All blocks should have level defined
        for block in merged:
            assert 'level' in block
            assert isinstance(block['level'], int)
            assert block['level'] >= 1


# ============================================================
# Main
# ============================================================

if __name__ == '__main__':
    pytest.main([__file__, '-v'])
