#!/usr/bin/env python3
"""
ABOUTME: Parses DOCX documents into text blocks using python-docx
ABOUTME: Extracts automatic numbering, splits by headings, converts tables to JSON
"""

import argparse
import hashlib
import json
import sys
from datetime import datetime
from pathlib import Path

try:
    from docx import Document
except ImportError:
    print("Error: python-docx not installed. Run: pip install python-docx", file=sys.stderr)
    sys.exit(1)

try:
    from numbering_resolver import NumberingResolver
    from table_extractor import TableExtractor
except ImportError:
    print("Error: Required modules not found. Ensure numbering_resolver.py and table_extractor.py are in the same directory.", file=sys.stderr)
    sys.exit(1)


# Constants for content validation
MAX_HEADING_LENGTH = 200      # Maximum heading length in characters
IDEAL_BLOCK_CONTENT_LENGTH = 5000  # Ideal target size for balanced splitting
MAX_BLOCK_CONTENT_LENGTH = 8000  # Maximum block content length in characters (hard limit)
MIN_BLOCK_CONTENT_LENGTH = 500  # Minimum block content length (triggers merging)
MAX_ANCHOR_CANDIDATE_LENGTH = 100  # Maximum length for candidate anchor paragraphs

# Constants for table splitting
TABLE_IDEAL_LENGTH = 3000  # Ideal target size for table chunks
TABLE_MAX_LENGTH = 5000    # Maximum table size before splitting (triggers table splitting)
TABLE_MIN_LAST_CHUNK_LENGTH = 1000  # Minimum last chunk size (merge with previous if smaller)


def print_error(title: str, details: str, solution: str):
    """
    Print a friendly, formatted error message.
    
    Args:
        title: Error title
        details: Detailed error information
        solution: Suggested solution steps
    """
    print("\n" + "=" * 80, file=sys.stderr)
    print(f"ERROR: {title}", file=sys.stderr)
    print("=" * 80, file=sys.stderr)
    print(f"\n{details}", file=sys.stderr)
    print(f"\nSOLUTION:", file=sys.stderr)
    print(solution, file=sys.stderr)
    print("\n" + "=" * 80 + "\n", file=sys.stderr)


def truncate_heading(heading_text: str, para_id: str = None) -> str:
    """
    Truncate heading if it exceeds MAX_HEADING_LENGTH.
    
    Args:
        heading_text: The heading text to check
        para_id: Optional paragraph ID for warning message
        
    Returns:
        str: Original heading if within limit, truncated heading with "..." if too long
    """
    if len(heading_text) > MAX_HEADING_LENGTH:
        truncated = heading_text[:MAX_HEADING_LENGTH - 3] + "..."
        location = f" (para_id: {para_id})" if para_id else ""
        print(
            f"Warning: Heading truncated (length {len(heading_text)} > max {MAX_HEADING_LENGTH}){location}: "
            f"\"{truncated}\"",
            file=sys.stderr
        )
        return truncated
    return heading_text


def validate_heading_length(heading_text: str, para_id: str):
    """
    Validate that heading length does not exceed MAX_HEADING_LENGTH.
    
    Args:
        heading_text: The heading text to validate
        para_id: The paragraph ID for error reporting
        
    Exits:
        sys.exit(1) if heading exceeds maximum length
    """
    if len(heading_text) > MAX_HEADING_LENGTH:
        preview = heading_text[:100] + "..." if len(heading_text) > 100 else heading_text
        print_error(
            f"Heading too long ({len(heading_text)} characters, max {MAX_HEADING_LENGTH})",
            f"The following heading exceeds the maximum allowed length:\n\n  \"{preview}\"\n\n"
            f"Location: Paragraph ID {para_id}\n"
            f"Actual length: {len(heading_text)} characters",
            "  1. Open the document in Microsoft Word\n"
            f"  2. Shorten this heading to {MAX_HEADING_LENGTH} characters or less\n"
            "  3. Re-run the audit workflow"
        )
        sys.exit(1)


def validate_table_length(table_json: str, block_heading: str):
    """
    Validate that table JSON does not exceed MAX_BLOCK_CONTENT_LENGTH.
    
    Args:
        table_json: The JSON representation of the table
        block_heading: The heading of the block containing this table
        
    Exits:
        sys.exit(1) if table exceeds maximum length
    """
    if len(table_json) > MAX_BLOCK_CONTENT_LENGTH:
        print_error(
            f"Table too large ({len(table_json)} characters, max {MAX_BLOCK_CONTENT_LENGTH})",
            f"A table in the document is too large for LLM processing.\n\n"
            f"Location: Under heading \"{block_heading}\"\n"
            f"Table size: {len(table_json)} characters\n\n"
            "Large tables can cause issues with automated auditing.",
            "  1. Open the document in Microsoft Word\n"
            f"  2. Locate the table under heading \"{block_heading}\"\n"
            "  3. Split the table into smaller tables, or\n"
            "  4. Simplify the table content\n"
            "  5. Re-run the audit workflow"
        )
        sys.exit(1)


def find_first_valid_para_id(para_ids: list) -> str:
    """
    Find the first valid paraId in a 2D array of paraIds.
    
    Args:
        para_ids: 2D list of paraIds from table cells
        
    Returns:
        str: First non-None paraId found
        
    Exits:
        sys.exit(1) if no valid paraId found
    """
    for row in para_ids:
        for para_id in row:
            if para_id:
                return para_id
    
    # No valid paraId found
    print_error(
        "Cannot find valid paraId in table",
        "A table was encountered but no cells contain valid paragraph IDs.\n"
        "This may indicate the table was created by incompatible software.",
        "  1. Open the document in Microsoft Word 2013 or later\n"
        "  2. Save the file (Ctrl+S)\n"
        "  3. Re-run the audit workflow"
    )
    sys.exit(1)


def find_last_valid_para_id(para_ids: list) -> str:
    """
    Find the last valid paraId in a 2D array of paraIds.
    
    Args:
        para_ids: 2D list of paraIds from table cells
        
    Returns:
        str: Last non-None paraId found, or first valid if none found in reverse
    """
    # Iterate in reverse order to find last valid paraId
    for row in reversed(para_ids):
        for para_id in reversed(row):
            if para_id:
                return para_id
    
    # Fallback to first valid paraId
    return find_first_valid_para_id(para_ids)


def split_table(table_rows: list, para_ids: list, para_ids_end: list, header_indices: list, debug: bool = False) -> list:
    """
    Split large table into chunks at row boundaries.
    
    Splitting Strategy:
    1. Only split if table JSON exceeds TABLE_MAX_LENGTH (5000 chars)
    2. Calculate target chunks based on TABLE_IDEAL_LENGTH (3000 chars)
    3. Split at row boundaries to achieve balanced chunk sizes
    4. Avoid very small last chunk: if last chunk < 1000 chars, merge with previous
    5. Extract first valid paraId for each chunk as UUID
    
    Output Strategy:
    - First chunk: Merges with preceding content, uses original heading
    - Middle chunks: Standalone blocks with heading suffix [1], [2], etc.
    - Last chunk: Merges with following content, includes table_header if present
    - Non-first chunks include table_header field (extracted from w:tblHeader attribute)
    
    Args:
        table_rows: 2D array of table content
        para_ids: 2D array of paraIds - first paraId in each cell (for uuid)
        para_ids_end: 2D array of paraIds - last paraId in each cell (for uuid_end)
        header_indices: List of row indices that are table headers
        debug: If True, output debug information
        
    Returns:
        List of chunk dicts: [{
            'rows': 2D array subset,
            'para_ids': 2D array subset,
            'para_ids_end': 2D array subset,
            'uuid': first valid paraId in chunk,
            'is_first': True if first chunk,
            'is_last': True if last chunk
        }, ...]
    """
    import math
    
    # Calculate total JSON length
    total_json = json.dumps(table_rows, ensure_ascii=False)
    total_length = len(total_json)
    
    if total_length <= TABLE_MAX_LENGTH:
        # No splitting needed
        uuid = find_first_valid_para_id(para_ids)
        return [{
            'rows': table_rows,
            'para_ids': para_ids,
            'para_ids_end': para_ids_end,
            'uuid': uuid,
            'is_first': True,
            'is_last': True
        }]
    
    # Need to split - calculate target number of chunks
    target_chunks = math.ceil(total_length / TABLE_IDEAL_LENGTH)
    min_chunks_needed = math.ceil(total_length / TABLE_MAX_LENGTH)
    target_chunks = max(target_chunks, min_chunks_needed)
    
    # Split at row boundaries
    chunks = []
    num_rows = len(table_rows)
    target_rows_per_chunk = num_rows / target_chunks
    
    start_row = 0
    for i in range(target_chunks):
        # Calculate end row for this chunk
        if i == target_chunks - 1:
            # Last chunk gets all remaining rows
            end_row = num_rows
        else:
            # Target end row (rounded)
            end_row = min(int((i + 1) * target_rows_per_chunk), num_rows)
            
            # Adjust to avoid very small last chunk
            rows_remaining = num_rows - end_row
            if rows_remaining > 0 and rows_remaining < target_rows_per_chunk * 0.3:
                # Last chunk would be too small, expand this chunk
                end_row = num_rows
        
        # Extract chunk
        chunk_rows = table_rows[start_row:end_row]
        chunk_para_ids = para_ids[start_row:end_row]
        chunk_para_ids_end = para_ids_end[start_row:end_row]
        
        if chunk_rows:
            chunk_uuid = find_first_valid_para_id(chunk_para_ids)
            chunks.append({
                'rows': chunk_rows,
                'para_ids': chunk_para_ids,
                'para_ids_end': chunk_para_ids_end,
                'uuid': chunk_uuid,
                'is_first': (i == 0),
                'is_last': (end_row >= num_rows)
            })
        
        start_row = end_row
        if start_row >= num_rows:
            break
    
    # Post-processing: Merge very small last chunk with previous chunk if possible
    if len(chunks) >= 2:
        last_chunk = chunks[-1]
        last_chunk_json = json.dumps(last_chunk['rows'], ensure_ascii=False)
        
        if len(last_chunk_json) < TABLE_MIN_LAST_CHUNK_LENGTH:
            # Try to merge with previous chunk
            prev_chunk = chunks[-2]
            
            # Calculate combined size
            combined_rows = prev_chunk['rows'] + last_chunk['rows']
            combined_json = json.dumps(combined_rows, ensure_ascii=False)
            
            # Only merge if combined size doesn't exceed max limit
            if len(combined_json) <= TABLE_MAX_LENGTH:
                # Merge the chunks
                merged_para_ids = prev_chunk['para_ids'] + last_chunk['para_ids']
                merged_para_ids_end = prev_chunk['para_ids_end'] + last_chunk['para_ids_end']
                chunks[-2] = {
                    'rows': combined_rows,
                    'para_ids': merged_para_ids,
                    'para_ids_end': merged_para_ids_end,
                    'uuid': prev_chunk['uuid'],  # Keep UUID of first chunk
                    'is_first': prev_chunk['is_first'],
                    'is_last': True  # This becomes the last chunk
                }
                chunks.pop()  # Remove the last chunk
                
                if debug:
                    print(f"[DEBUG] Merged small last chunk ({len(last_chunk_json)} chars) with previous chunk", file=sys.stderr)
                    print(f"  Combined size: {len(combined_json)} chars", file=sys.stderr)
    
    return chunks


def split_table_with_heading(table_rows: list, para_ids: list, para_ids_end: list, header_indices: list, current_heading: str, start_suffix: int = 0, debug: bool = False) -> list:
    """
    Wrapper for split_table that includes heading information in debug output.
    Supports sequential numbering when multiple tables are split in the same block.
    
    Args:
        table_rows: 2D array of table content
        para_ids: 2D array of paraIds - first paraId in each cell (for uuid)
        para_ids_end: 2D array of paraIds - last paraId in each cell (for uuid_end)
        header_indices: List of row indices that are table headers
        current_heading: Current block heading (for generating chunk headings)
        start_suffix: Starting suffix number for non-first chunks (default: 0)
                     When multiple tables in the same block are split, this ensures
                     sequential numbering (e.g., [1], [2] for first table, [3], [4] for second)
        debug: If True, output debug information with headings
        
    Returns:
        Same as split_table(), with each chunk having suffix calculated from start_suffix
    """
    chunks = split_table(table_rows, para_ids, para_ids_end, header_indices, debug=False)
    
    # Add suffix_number to each chunk for later use
    for i, chunk in enumerate(chunks):
        if i == 0:
            chunk['suffix_number'] = None  # First chunk has no suffix
        else:
            chunk['suffix_number'] = start_suffix + i
    
    # Debug output with headings
    if debug and len(chunks) > 1:
        print(f"\n[DEBUG] Table split into {len(chunks)} chunks (final)", file=sys.stderr)
        for i, chunk in enumerate(chunks):
            chunk_json = json.dumps(chunk['rows'], ensure_ascii=False)
            # Generate heading for this chunk
            if chunk['suffix_number'] is None:
                chunk_heading = current_heading
            else:
                chunk_heading = f"{current_heading} [{chunk['suffix_number']}]"
            print(f"  Chunk {i+1}: heading=\"{chunk_heading}\", {len(chunk['rows'])} rows, {len(chunk_json)} chars", file=sys.stderr)
    
    return chunks


def merge_small_blocks(blocks: list, debug: bool = False) -> tuple:
    """
    Merge blocks that are smaller than MIN_BLOCK_CONTENT_LENGTH with adjacent blocks.
    
    Strategy:
    1. Identify blocks smaller than MIN_BLOCK_CONTENT_LENGTH
    2. Try to merge with next block (small block's heading becomes next block's heading)
    3. If merging with next block would exceed MAX_BLOCK_CONTENT_LENGTH, try previous block
    4. Only keep small block separate if both merge directions exceed limit
    
    Args:
        blocks: List of block dictionaries
        debug: If True, return merge count for debug reporting
        
    Returns:
        Tuple of (merged_blocks, merge_count)
    """
    if len(blocks) <= 1:
        return blocks, 0
    
    merged_blocks = []
    i = 0
    merged_count = 0
    
    while i < len(blocks):
        current_block = blocks[i]
        current_length = len(current_block['content'])
        
        # Check if current block is too small and needs merging
        if current_length < MIN_BLOCK_CONTENT_LENGTH and current_length > 0:
            merged = False
            
            # Try merging with next block first
            if i + 1 < len(blocks):
                next_block = blocks[i + 1]
                next_length = len(next_block['content'])
                combined_length = current_length + next_length + 1  # +1 for newline
                
                if combined_length <= MAX_BLOCK_CONTENT_LENGTH:
                    # Merge current into next block
                    # Current block's heading becomes the new heading
                    # UUID range: current's uuid to next's uuid_end
                    merged_content = current_block['content'] + "\n" + next_block['content']
                    merged_block = {
                        "uuid": current_block['uuid'],  # Use current block's UUID
                        "uuid_end": next_block.get('uuid_end', next_block['uuid']),  # Use next block's end UUID
                        "heading": current_block['heading'],  # Use current block's heading
                        "content": merged_content,
                        "type": "text",
                        "parent_headings": current_block['parent_headings']
                    }
                    merged_blocks.append(merged_block)
                    
                    merged = True
                    merged_count += 1
                    i += 2  # Skip both current and next block
                    continue
            
            # If can't merge with next, try merging with previous
            if not merged and len(merged_blocks) > 0:
                prev_block = merged_blocks[-1]
                prev_length = len(prev_block['content'])
                combined_length = prev_length + current_length + 1  # +1 for newline
                
                if combined_length <= MAX_BLOCK_CONTENT_LENGTH:
                    # Merge current into previous block
                    # Previous block remains the heading
                    # UUID range: prev's uuid to current's uuid_end
                    merged_content = prev_block['content'] + "\n" + current_block['content']
                    merged_blocks[-1] = {
                        "uuid": prev_block['uuid'],  # Keep previous UUID
                        "uuid_end": current_block.get('uuid_end', current_block['uuid']),  # Use current block's end UUID
                        "heading": prev_block['heading'],  # Keep previous heading
                        "content": merged_content,
                        "type": "text",
                        "parent_headings": prev_block['parent_headings']
                    }
                    
                    merged = True
                    merged_count += 1
                    i += 1
                    continue
            
            # If neither merge direction works, keep the small block as-is
            if not merged:
                merged_blocks.append(current_block)
                i += 1
        else:
            # Block is within acceptable size range, keep as-is
            merged_blocks.append(current_block)
            i += 1
    
    return merged_blocks, merged_count


def split_long_block(block_heading: str, paragraphs: list, parent_headings: list, debug: bool = False) -> list:
    """
    Split a long text block into smaller blocks using anchor paragraphs.
    
    Strategy (improved for balanced splitting):
    1. Calculate target number of blocks based on IDEAL_BLOCK_CONTENT_LENGTH
    2. Ensure minimum blocks needed to stay under MAX_BLOCK_CONTENT_LENGTH
    3. Find all candidate anchor paragraphs (<= MAX_ANCHOR_CANDIDATE_LENGTH)
    4. Select anchors closest to ideal split positions for balanced distribution
    5. Create blocks using selected anchors as new headings
    
    Args:
        block_heading: Original heading text
        paragraphs: List of dicts with 'text', 'para_id', and 'is_table' keys
        parent_headings: Parent heading stack
        debug: If True, output debug information when splitting occurs
        
    Returns:
        List of block dictionaries (may be split into multiple blocks)
        
    Exits:
        sys.exit(1) if no suitable anchor found and content exceeds limit
    """
    import math
    
    # Check if this block starts with a split table chunk (has _chunk_heading metadata)
    # If so, use that heading instead of block_heading
    effective_heading = block_heading
    table_header = None
    
    if paragraphs and paragraphs[0].get('_chunk_heading'):
        effective_heading = paragraphs[0]['_chunk_heading']
        table_header = paragraphs[0].get('_table_header')
    
    # Calculate total content length
    total_content = "\n".join(p['text'] for p in paragraphs)
    total_length = len(total_content)
    
    if total_length <= MAX_BLOCK_CONTENT_LENGTH:
        # Within limit, return as single block
        # Use first paragraph's para_id as UUID
        # For uuid_end: use para_id_end if last element is a table, otherwise para_id
        last_para = paragraphs[-1] if paragraphs else {}
        uuid_end = last_para.get('para_id_end') or last_para.get('para_id')
        
        block = {
            "uuid": paragraphs[0]['para_id'] if paragraphs else None,
            "uuid_end": uuid_end,
            "heading": effective_heading,
            "content": total_content,
            "type": "text",
            "parent_headings": parent_headings
        }
        
        # Add table_header if present
        if table_header:
            block["table_header"] = table_header
        
        return [block]
    
    # Content exceeds limit, need to split
    # Calculate target number of blocks based on IDEAL_BLOCK_CONTENT_LENGTH
    target_blocks = math.ceil(total_length / IDEAL_BLOCK_CONTENT_LENGTH)
    
    # Ensure we have enough blocks to stay under MAX_BLOCK_CONTENT_LENGTH
    min_blocks_needed = math.ceil(total_length / MAX_BLOCK_CONTENT_LENGTH)
    target_blocks = max(target_blocks, min_blocks_needed)
    
    # Calculate ideal size per block
    target_size = total_length / target_blocks
    
    # Find candidate anchors (short paragraphs, excluding tables and empty placeholders)
    candidates = []
    cumulative_length = 0
    for idx, para in enumerate(paragraphs):
        if not para.get('is_table', False) and 0 < len(para['text']) <= MAX_ANCHOR_CANDIDATE_LENGTH:
            candidates.append({
                'index': idx,
                'text': para['text'],
                'para_id': para['para_id'],
                'position': cumulative_length
            })
        cumulative_length += len(para['text']) + 1  # +1 for newline
    
    if not candidates:
        # No suitable anchor found
        preview = block_heading[:80] + "..." if len(block_heading) > 80 else block_heading
        print_error(
            f"Cannot split long block (no suitable anchor paragraphs found)",
            f"A text block is too long ({total_length} characters, max {MAX_BLOCK_CONTENT_LENGTH})\n"
            f"but no paragraphs <= {MAX_ANCHOR_CANDIDATE_LENGTH} characters were found to use as split points.\n\n"
            f"Location: Under heading \"{preview}\"\n"
            f"Block size: {total_length} characters\n"
            f"Number of paragraphs: {len(paragraphs)}\n"
            f"Calculated target blocks: {target_blocks}",
            "  1. Open the document in Microsoft Word\n"
            f"  2. Locate the section under heading \"{preview}\"\n"
            f"  3. Add short headings or paragraph breaks (≤{MAX_ANCHOR_CANDIDATE_LENGTH} chars) to divide the content\n"
            "  4. Re-run the audit workflow\n\n"
            f"Tip: Short headings like '概述', '背景', '详细说明' can serve as natural split points."
        )
        sys.exit(1)
    
    # Select anchors for splitting (target_blocks - 1 split points needed)
    selected_anchors = []
    remaining_candidates = candidates.copy()
    
    for i in range(1, target_blocks):
        if not remaining_candidates:
            break
        
        # Calculate ideal position for this split
        ideal_position = i * target_size
        
        # Find candidate closest to ideal position
        best_candidate = min(remaining_candidates, key=lambda c: abs(c['position'] - ideal_position))
        selected_anchors.append(best_candidate)
        remaining_candidates.remove(best_candidate)
    
    # Sort selected anchors by index to maintain document order
    selected_anchors.sort(key=lambda a: a['index'])
    
    # Create blocks using selected split points
    result_blocks = []
    prev_idx = 0
    current_parent_headings = parent_headings
    current_block_heading = block_heading
    
    for anchor in selected_anchors:
        split_idx = anchor['index']
        
        # Create block from prev_idx to split_idx (exclusive)
        block_paragraphs = paragraphs[prev_idx:split_idx]
        if block_paragraphs:
            block_content = "\n".join(p['text'] for p in block_paragraphs)
            # For uuid_end: use para_id_end if last element is a table, otherwise para_id
            last_para = block_paragraphs[-1]
            block_uuid_end = last_para.get('para_id_end') or last_para.get('para_id')
            result_blocks.append({
                "uuid": block_paragraphs[0]['para_id'],  # UUID from first paragraph in content
                "uuid_end": block_uuid_end,  # UUID_end from last paragraph (or table's last cell)
                "heading": current_block_heading,
                "content": block_content,
                "type": "text",
                "parent_headings": current_parent_headings,
                "_paragraphs": block_paragraphs  # Keep original paragraphs for potential re-splitting
            })
        
        # Validate anchor as new heading
        validate_heading_length(anchor['text'], anchor['para_id'])
        
        # Update for next block
        current_block_heading = anchor['text']
        # Update parent headings: add previous heading only if not "Preface/Uncategorized"
        if block_heading != "Preface/Uncategorized":
            current_parent_headings = parent_headings + [block_heading]
        
        prev_idx = split_idx  # Don't skip anchor - it becomes first paragraph of next block
    
    # Create final block with remaining paragraphs
    final_paragraphs = paragraphs[prev_idx:]
    if final_paragraphs:
        final_content = "\n".join(p['text'] for p in final_paragraphs)
        # For uuid_end: use para_id_end if last element is a table, otherwise para_id
        last_final_para = final_paragraphs[-1]
        final_uuid_end = last_final_para.get('para_id_end') or last_final_para.get('para_id')
        result_blocks.append({
            "uuid": final_paragraphs[0]['para_id'],  # UUID from first paragraph in content
            "uuid_end": final_uuid_end,  # UUID_end from last paragraph (or table's last cell)
            "heading": current_block_heading,
            "content": final_content,
            "type": "text",
            "parent_headings": current_parent_headings,
            "_paragraphs": final_paragraphs  # Keep original paragraphs for potential re-splitting
        })
    
    # Post-split validation: Check if any block still exceeds MAX_BLOCK_CONTENT_LENGTH
    # If so, recursively split that block (handles sparse anchor scenarios)
    validated_blocks = []
    for block in result_blocks:
        if len(block['content']) > MAX_BLOCK_CONTENT_LENGTH:
            # This block is still too large - need to recursively split it
            # Use the preserved paragraph structure
            block_paragraphs = block.get('_paragraphs', [])
            
            if not block_paragraphs:
                # Fallback: shouldn't happen, but handle gracefully
                preview = block['heading'][:80] + "..." if len(block['heading']) > 80 else block['heading']
                print_error(
                    f"Cannot re-split oversized block (internal error)",
                    f"A block exceeded MAX_BLOCK_CONTENT_LENGTH but paragraph metadata was lost.\n\n"
                    f"Location: Under heading \"{preview}\"\n"
                    f"Block size: {len(block['content'])} characters",
                    "This is an internal error. Please report this issue."
                )
                sys.exit(1)
            
            # Recursively split this oversized block
            # The recursive call will either find more anchors or raise an error
            sub_blocks = split_long_block(
                block['heading'],
                block_paragraphs,
                block['parent_headings'],
                debug
            )
            validated_blocks.extend(sub_blocks)
        else:
            # Remove internal _paragraphs field before adding to final output
            block.pop('_paragraphs', None)
            validated_blocks.append(block)
    
    # Merge small blocks with adjacent blocks to avoid fragmentation
    final_blocks, merge_count = merge_small_blocks(validated_blocks, debug)
    
    # Output debug information if enabled and split occurred (after merging)
    if debug and len(final_blocks) > 1:
        print(f"\n[DEBUG] Block split: \"{block_heading}\"", file=sys.stderr)
        print(f"  Original length: {total_length} characters", file=sys.stderr)
        final_block_lengths = [len(block['content']) for block in final_blocks]
        print(f"  Final result: {len(final_blocks)} blocks: {final_block_lengths} characters", file=sys.stderr)
        if merge_count > 0:
            print(f"  ({merge_count} small block(s) merged)", file=sys.stderr)
    
    return final_blocks


def extract_para_id(para_element) -> str:
    """
    Extract w14:paraId attribute from paragraph element.

    Args:
        para_element: lxml paragraph element

    Returns:
        str: 8-character hex paraId

    Exits:
        sys.exit(1) if paraId attribute is missing (indicates old Word version)
    """
    # Check for w14:paraId attribute
    para_id = para_element.get('{http://schemas.microsoft.com/office/word/2010/wordml}paraId')
    
    if not para_id:
        print("\n" + "=" * 60, file=sys.stderr)
        print("ERROR: Document missing paraId attributes", file=sys.stderr)
        print("=" * 60, file=sys.stderr)
        print("\nThe paragraphs in this document are missing w14:paraId attributes.", file=sys.stderr)
        print("This may be caused by:", file=sys.stderr)
        print("  - Document generated by python-docx or similar tools", file=sys.stderr)
        print("  - Document created by LibreOffice or Google Docs", file=sys.stderr)
        print("  - Document never saved in Microsoft Word 2013+", file=sys.stderr)
        print("\nSOLUTION:", file=sys.stderr)
        print("  1. Open the document in Microsoft Word 2013 or later", file=sys.stderr)
        print("  2. Save the file (Ctrl+S)", file=sys.stderr)
        print("  3. Re-run the audit workflow", file=sys.stderr)
        print("\n" + "=" * 60 + "\n", file=sys.stderr)
        sys.exit(1)
    
    return para_id


def parse_styles_outline_levels(docx_path: str) -> dict:
    """
    Parse styles.xml to extract outlineLvl definitions for each style,
    following style inheritance chain (basedOn).

    Args:
        docx_path: Path to DOCX file

    Returns:
        dict: styleId -> outlineLvl (0-8 for headings, 9 for body text)
    """
    import zipfile
    try:
        from defusedxml import ElementTree as ET
    except ImportError:
        from xml.etree import ElementTree as ET

    styles_outline = {}  # styleId -> outlineLvl (directly defined)
    style_based_on = {}  # styleId -> parent styleId

    try:
        with zipfile.ZipFile(docx_path, 'r') as zf:
            if 'word/styles.xml' not in zf.namelist():
                return styles_outline

            tree = ET.parse(zf.open('word/styles.xml'))
            root = tree.getroot()

            ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

            # First pass: collect outlineLvl and basedOn for all styles
            for style in root.findall(f'.//{{{ns}}}style'):
                style_id = style.get(f'{{{ns}}}styleId')
                if not style_id:
                    continue

                # Check for basedOn (style inheritance)
                based_on = style.find(f'{{{ns}}}basedOn')
                if based_on is not None:
                    parent_id = based_on.get(f'{{{ns}}}val')
                    if parent_id:
                        style_based_on[style_id] = parent_id

                # Check for outlineLvl in style's pPr
                pPr = style.find(f'{{{ns}}}pPr')
                if pPr is not None:
                    outline_lvl_elem = pPr.find(f'{{{ns}}}outlineLvl')
                    if outline_lvl_elem is not None:
                        level = int(outline_lvl_elem.get(f'{{{ns}}}val'))
                        styles_outline[style_id] = level

            # Second pass: resolve inheritance chain for styles without direct outlineLvl
            def get_outline_level(style_id: str, visited: set = None) -> int:
                if visited is None:
                    visited = set()
                if style_id in visited:
                    return None  # Prevent circular references
                visited.add(style_id)

                # If this style directly defines outlineLvl, return it
                if style_id in styles_outline:
                    return styles_outline[style_id]

                # Otherwise check parent style
                if style_id in style_based_on:
                    parent_id = style_based_on[style_id]
                    return get_outline_level(parent_id, visited)

                return None

            # Fill in missing outlineLvl from inheritance chain
            all_style_ids = set(styles_outline.keys()) | set(style_based_on.keys())
            for style_id in all_style_ids:
                if style_id not in styles_outline:
                    level = get_outline_level(style_id)
                    if level is not None:
                        styles_outline[style_id] = level
    except Exception:
        # Silently ignore parsing errors
        pass

    return styles_outline


def get_heading_level(para_element, styles_outline_map: dict) -> int:
    """
    Get heading level from paragraph, checking both direct format and style.
    
    Priority: paragraph outlineLvl > style outlineLvl
    
    Args:
        para_element: lxml paragraph element
        styles_outline_map: dict of styleId -> outlineLvl from styles.xml
        
    Returns:
        int: 0-8 for heading levels (0=level 1, 1=level 2, etc.), None for non-heading
    """
    # 1. Check paragraph direct format
    pPr = para_element.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPr')
    if pPr is not None:
        outline_elem = pPr.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}outlineLvl')
        if outline_elem is not None:
            level = int(outline_elem.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val'))
            # Only 0-8 are true heading levels (9 is body text)
            if level < 9:
                return level
            else:
                return None  # Level 9 is body text
    
    # 2. Check style definition's outlineLvl
    if pPr is not None:
        pStyle_elem = pPr.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pStyle')
        if pStyle_elem is not None:
            style_id = pStyle_elem.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')
            if style_id and style_id in styles_outline_map:
                level = styles_outline_map[style_id]
                if level < 9:
                    return level
                else:
                    return None
    
    return None


def extract_audit_blocks(file_path: str, debug: bool = False) -> list:
    """
    Extract text blocks (chunks) from a DOCX file for auditing.
    
    Uses python-docx with custom numbering resolver to:
    1. Capture automatic numbering (list labels)
    2. Split document by headings
    3. Convert tables to JSON (2D array)
    4. Validate heading lengths and table sizes
    5. Split long blocks using anchor paragraphs
    
    Args:
        file_path: Path to the DOCX file
        debug: If True, output debug information when splitting blocks
        
    Returns:
        List of block dictionaries with heading, content, type, and metadata
    """
    doc = Document(file_path)
    resolver = NumberingResolver(file_path)
    styles_outline = parse_styles_outline_levels(file_path)
    
    blocks = []
    current_heading = "Preface/Uncategorized"
    current_heading_stack = []
    current_parent_headings = []  # Parent headings for current block
    current_paragraphs = []  # Track paragraphs with metadata for splitting
    has_body_content = False  # Track if current block has body content (non-heading paragraphs/tables)
    table_split_counter = 0  # Track cumulative table split suffix numbers within current block
    
    # Iterate through document body elements (paragraphs and tables)
    body = doc._element.body
    
    for element in body:
        tag = element.tag.split('}')[-1]  # Remove namespace
        
        if tag == 'sectPr':  # Document-level section break
            resolver.reset_tracking_state()
            continue
        
        if tag == 'p':  # Paragraph
            # Get paragraph text
            para_text = ''
            ns = {
                'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
                'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
            }
            for run in element.findall('.//w:r', ns):
                for child in run:
                    tag = child.tag.split('}')[-1]  # Remove namespace
                    if tag == 't' and child.text:
                        para_text += child.text
                    elif tag == 'br':
                        # Handle line breaks - textWrapping or no type = soft line break
                        br_type = child.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}type')
                        if br_type in (None, 'textWrapping'):
                            para_text += '\n'
                        # Skip page and column breaks (layout elements)
                    elif tag == 'drawing':
                        # Extract inline images (ignore floating images wp:anchor)
                        inline = child.find('wp:inline', ns)
                        if inline is not None:
                            doc_pr = inline.find('wp:docPr', ns)
                            if doc_pr is not None:
                                img_id = doc_pr.get('id', '')
                                img_name = doc_pr.get('name', '')
                                para_text += f'<drawing id="{img_id}" name="{img_name}" />'
            
            para_text = para_text.strip()
            if not para_text:
                continue
            
            # Get numbering label using our resolver
            label = resolver.get_label(element)
            full_text = f"{label} {para_text}".strip() if label else para_text
            
            # Check if this is a heading using the new function
            outline_level = get_heading_level(element, styles_outline)
            
            if outline_level is not None:
                # This is a heading (outline level 0-8)
                # Extract paraId for this heading
                heading_para_id = extract_para_id(element)
                
                # Validate heading length
                validate_heading_length(full_text, heading_para_id)
                
                # Only save previous block if it has body content
                if has_body_content and current_paragraphs:
                    # Split long blocks if needed
                    split_blocks = split_long_block(current_heading, current_paragraphs, current_parent_headings, debug)
                    blocks.extend(split_blocks)
                    
                    # Reset for new block
                    current_paragraphs = []
                    has_body_content = False
                    table_split_counter = 0  # Reset table split counter for new heading
                
                # Convert 0-based to 1-based level
                level = outline_level + 1
                
                # Truncate heading if needed before storing
                truncated_text = truncate_heading(full_text, heading_para_id)
                
                # Add heading to current_paragraphs
                current_paragraphs.append({
                    'text': truncated_text,
                    'para_id': heading_para_id,
                    'is_table': False
                })
                
                # Update current_heading and parent_headings for the FIRST heading in a block
                # (when current_paragraphs just had this heading added as its first element)
                if len(current_paragraphs) == 1:
                    current_heading = truncated_text
                    # Parent headings = all headings in stack before this heading (at higher levels)
                    current_parent_headings = current_heading_stack[:max(level - 1, 0)]
                
                # Update heading stack
                current_heading_stack = current_heading_stack[:max(level - 1, 0)]
                current_heading_stack.append(truncated_text)
            else:
                # Regular paragraph content
                para_id = extract_para_id(element)
                
                # Store paragraph with metadata for potential splitting
                current_paragraphs.append({
                    'text': full_text,
                    'para_id': para_id,
                    'is_table': False
                })
                
                # Mark that we have body content
                has_body_content = True
            
            # Check for paragraph-level section break (after processing paragraph)
            # sectPr in pPr means this paragraph ends a section
            pPr = element.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPr')
            if pPr is not None:
                sectPr = pPr.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sectPr')
                if sectPr is not None:
                    # Section break after this paragraph - reset tracking
                    resolver.reset_tracking_state()
        
        elif tag == 'tbl':  # Table
            # Reset numbering tracking before table (table start boundary)
            resolver.reset_tracking_state()
            
            # Directly create Table object from XML element to avoid index mismatch
            # (doc.tables may have different order due to nested tables)
            from docx.table import Table
            table = Table(element, doc)
            table_metadata = TableExtractor.extract_with_metadata(table, numbering_resolver=resolver)
            
            table_rows = table_metadata['rows']
            para_ids = table_metadata['para_ids']
            para_ids_end = table_metadata['para_ids_end']  # Last paraId in each cell
            header_indices = table_metadata['header_indices']
            
            # Convert table to JSON to check length
            table_json = json.dumps(table_rows, ensure_ascii=False)
            
            # Check if table needs splitting
            if len(table_json) > TABLE_MAX_LENGTH:
                # Table exceeds limit - split it
                # Pass table_split_counter to ensure sequential numbering across multiple tables
                table_chunks = split_table_with_heading(table_rows, para_ids, para_ids_end, header_indices, current_heading, table_split_counter, debug)
                
                # Extract header rows if any
                header_rows = []
                if header_indices:
                    header_rows = [table_rows[idx] for idx in header_indices if idx < len(table_rows)]
                
                for chunk_idx, chunk in enumerate(table_chunks):
                    chunk_json = json.dumps(chunk['rows'], ensure_ascii=False)
                    # Get uuid_end from last valid paraId in chunk (use para_ids_end for last cell's last paragraph)
                    chunk_para_id_end = find_last_valid_para_id(chunk['para_ids_end'])
                    
                    if chunk['is_first']:
                        # First chunk: add to current_paragraphs (will merge with preceding content)
                        current_paragraphs.append({
                            'text': f"<table>{chunk_json}</table>",
                            'para_id': chunk['uuid'],
                            'para_id_end': chunk_para_id_end,  # Store end paraId for uuid_end calculation
                            'is_table': True
                        })
                        has_body_content = True
                    else:
                        # Middle or last chunk: save current block first
                        if current_paragraphs:
                            split_blocks = split_long_block(current_heading, current_paragraphs, current_parent_headings, debug)
                            blocks.extend(split_blocks)
                            current_paragraphs = []
                            has_body_content = False
                        
                        # Generate heading using suffix_number from chunk
                        if chunk['suffix_number'] is not None:
                            chunk_heading = f"{current_heading} [{chunk['suffix_number']}]"
                        else:
                            chunk_heading = current_heading
                        
                        # Build block for this table chunk
                        # Get uuid_end from last valid paraId in chunk (use para_ids_end for last cell's last paragraph)
                        chunk_uuid_end = find_last_valid_para_id(chunk['para_ids_end'])
                        chunk_block = {
                            "uuid": chunk['uuid'],
                            "uuid_end": chunk_uuid_end,
                            "heading": chunk_heading,
                            "content": f"<table>{chunk_json}</table>",
                            "type": "text",
                            "parent_headings": current_parent_headings
                        }
                        
                        # Add table_header field if headers exist and this isn't the first chunk
                        if header_rows:
                            chunk_block["table_header"] = header_rows
                        
                        if chunk['is_last']:
                            # Last chunk: add to current_paragraphs for merging with following content
                            current_paragraphs.append({
                                'text': f"<table>{chunk_json}</table>",
                                'para_id': chunk['uuid'],
                                'para_id_end': chunk_para_id_end,  # Store end paraId for uuid_end calculation
                                'is_table': True,
                                '_chunk_heading': chunk_heading,
                                '_table_header': header_rows if header_rows else None
                            })
                            has_body_content = True
                        else:
                            # Middle chunk: output immediately as standalone block
                            blocks.append(chunk_block)
                
                # Update table_split_counter: add number of non-first chunks
                # (first chunk doesn't get a suffix, so we count from second chunk onwards)
                table_split_counter += len(table_chunks) - 1
            else:
                # Table is within size limit - no splitting needed
                # Store table as a paragraph with special marker
                # Use first valid paraId from table, and last valid paraId (from para_ids_end) for uuid_end
                table_para_id = find_first_valid_para_id(para_ids)
                table_para_id_end = find_last_valid_para_id(para_ids_end)
                current_paragraphs.append({
                    'text': f"<table>{table_json}</table>",
                    'para_id': table_para_id,
                    'para_id_end': table_para_id_end,  # Store end paraId for uuid_end calculation
                    'is_table': True
                })
                
                # Mark that we have body content
                has_body_content = True
            
            # Reset numbering tracking after table (table end boundary)
            resolver.reset_tracking_state()
    
    # Save final block with splitting if needed
    if current_paragraphs:
        # Split long blocks if needed
        split_blocks = split_long_block(current_heading, current_paragraphs, current_parent_headings, debug)
        blocks.extend(split_blocks)
    
    return blocks


def calculate_file_hash(file_path: str) -> str:
    """
    Calculate SHA256 hash of a file.

    Args:
        file_path: Path to file

    Returns:
        Hash string in format "sha256:hexdigest"
    """
    sha256_hash = hashlib.sha256()
    with open(file_path, "rb") as f:
        for byte_block in iter(lambda: f.read(4096), b""):
            sha256_hash.update(byte_block)
    return f"sha256:{sha256_hash.hexdigest()}"


def create_metadata(file_path: str) -> dict:
    """
    Create metadata object for source document.

    Args:
        file_path: Path to source document

    Returns:
        Metadata dictionary with type, source file info, hash, and timestamp
    """
    doc_path = Path(file_path).resolve()
    return {
        "type": "meta",
        "source_file": str(doc_path),
        "source_hash": calculate_file_hash(file_path),
        "parsed_at": datetime.now().isoformat()
    }


def format_table_for_display(table_data: list) -> str:
    """
    Format table data as readable text for display.

    Args:
        table_data: 2D list of cell values

    Returns:
        Formatted string representation
    """
    if not table_data:
        return "(empty table)"

    # Calculate column widths
    col_widths = []
    for col_idx in range(len(table_data[0]) if table_data else 0):
        max_width = 0
        for row in table_data:
            if col_idx < len(row):
                max_width = max(max_width, len(str(row[col_idx])))
        col_widths.append(min(max_width, 40))  # Cap at 40 chars

    lines = []
    for row in table_data:
        cells = []
        for i, cell in enumerate(row):
            width = col_widths[i] if i < len(col_widths) else 20
            cells.append(str(cell)[:width].ljust(width))
        lines.append(" | ".join(cells))

    return "\n".join(lines)


def save_blocks_jsonl(blocks: list, output_path: str, metadata: dict = None):
    """
    Save blocks to JSONL format (one JSON object per line).
    First line contains metadata if provided.
    Also removes existing manifest.jsonl to ensure clean resume state.

    Args:
        blocks: List of block dictionaries
        output_path: Path to output file
        metadata: Optional metadata dictionary to write as first line
    """
    with open(output_path, 'w', encoding='utf-8') as f:
        # Write metadata as first line if provided
        if metadata:
            f.write(json.dumps(metadata, ensure_ascii=False) + '\n')
        # Write all blocks
        for block in blocks:
            f.write(json.dumps(block, ensure_ascii=False) + '\n')
    
    # Clean up old manifest.jsonl to prevent UUID mismatch in resume mode
    manifest_path = Path(output_path).parent / "manifest.jsonl"
    if manifest_path.exists():
        manifest_path.unlink()
        print(f"Removed existing manifest: {manifest_path}")


def save_blocks_json(blocks: list, output_path: str, metadata: dict = None):
    """
    Save blocks to regular JSON format.
    Also removes existing manifest.jsonl to ensure clean resume state.

    Args:
        blocks: List of block dictionaries
        output_path: Path to output file
        metadata: Optional metadata dictionary to include in output
    """
    output_data = {
        "total_blocks": len(blocks),
        "blocks": blocks
    }
    
    # Add metadata if provided
    if metadata:
        output_data["meta"] = metadata
    
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(output_data, f, indent=2, ensure_ascii=False)
    
    # Clean up old manifest.jsonl to prevent UUID mismatch in resume mode
    manifest_path = Path(output_path).parent / "manifest.jsonl"
    if manifest_path.exists():
        manifest_path.unlink()
        print(f"Removed existing manifest: {manifest_path}")


def main():
    parser = argparse.ArgumentParser(
        description="Parse DOCX documents into text blocks for auditing"
    )
    parser.add_argument(
        "document",
        type=str,
        help="Path to the DOCX file to parse"
    )
    parser.add_argument(
        "--output", "-o",
        type=str,
        help="Output file path (default: {document}_blocks.jsonl)"
    )
    parser.add_argument(
        "--format",
        type=str,
        choices=["jsonl", "json"],
        default="jsonl",
        help="Output format (default: jsonl)"
    )
    parser.add_argument(
        "--preview",
        action="store_true",
        help="Print preview of extracted blocks"
    )
    parser.add_argument(
        "--stats",
        action="store_true",
        help="Print statistics about the document"
    )
    parser.add_argument(
        "--debug",
        action="store_true",
        help="Enable debug output for block splitting operations"
    )

    args = parser.parse_args()

    # Validate input file
    doc_path = Path(args.document)
    if not doc_path.exists():
        print(f"Error: File not found: {args.document}", file=sys.stderr)
        sys.exit(1)

    if doc_path.suffix.lower() != '.docx':
        print(f"Warning: File does not have .docx extension: {args.document}", file=sys.stderr)

    # Extract blocks
    print(f"Parsing document: {args.document}")
    blocks = extract_audit_blocks(args.document, debug=args.debug)
    print(f"Extracted {len(blocks)} text blocks")

    # Print statistics
    if args.stats:
        print("\n--- Document Statistics ---")
        headings = set()
        total_chars = 0
        for block in blocks:
            headings.add(block['heading'])
            total_chars += len(block['content'])

        print(f"Unique headings: {len(headings)}")
        print(f"Total characters: {total_chars:,}")
        print(f"Average block size: {total_chars // len(blocks) if blocks else 0:,} chars")

    # Print preview
    if args.preview:
        print("\n--- Block Preview (first 5) ---")
        for i, block in enumerate(blocks[:5]):
            print(f"\n[Block {i+1}] {block['heading']}")
            print(f"Type: {block['type']}")
            content = block['content'][:300]
            if len(block['content']) > 300:
                content += "..."
            print(f"Content: {content}")

    # Determine output path
    if args.output:
        output_path = args.output
    else:
        output_path = doc_path.stem + "_blocks." + args.format

    # Create metadata
    metadata = create_metadata(args.document)
    print(f"Calculated file hash: {metadata['source_hash'][:20]}...")

    # Save output with metadata
    if args.format == "jsonl":
        save_blocks_jsonl(blocks, output_path, metadata)
    else:
        save_blocks_json(blocks, output_path, metadata)

    print(f"\nSaved to: {output_path}")
    print(f"Source file: {metadata['source_file']}")


if __name__ == "__main__":
    main()
