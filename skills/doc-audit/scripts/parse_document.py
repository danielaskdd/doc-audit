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
MAX_ANCHOR_CANDIDATE_LENGTH = 100  # Maximum length for candidate anchor paragraphs


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


def split_long_block(block_heading: str, paragraphs: list, parent_headings: list) -> list:
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
        
    Returns:
        List of block dictionaries (may be split into multiple blocks)
        
    Exits:
        sys.exit(1) if no suitable anchor found and content exceeds limit
    """
    import math
    
    # Calculate total content length
    total_content = "\n".join(p['text'] for p in paragraphs)
    total_length = len(total_content)
    
    if total_length <= MAX_BLOCK_CONTENT_LENGTH:
        # Within limit, return as single block
        # Use first paragraph's para_id as UUID (assuming it's the heading's para_id)
        return [{
            "uuid": paragraphs[0]['para_id'] if paragraphs else None,
            "heading": block_heading,
            "content": total_content,
            "type": "text",
            "parent_headings": parent_headings
        }]
    
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
    current_uuid = paragraphs[0]['para_id'] if paragraphs else None
    
    for anchor in selected_anchors:
        split_idx = anchor['index']
        
        # Create block from prev_idx to split_idx (exclusive)
        block_paragraphs = paragraphs[prev_idx:split_idx]
        if block_paragraphs:
            block_content = "\n".join(p['text'] for p in block_paragraphs)
            result_blocks.append({
                "uuid": current_uuid,
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
        current_uuid = anchor['para_id']
        # Update parent headings: add previous heading only if not "Preface/Uncategorized"
        if block_heading != "Preface/Uncategorized":
            current_parent_headings = parent_headings + [block_heading]
        
        prev_idx = split_idx + 1  # Skip the anchor itself as it becomes the heading
    
    # Create final block with remaining paragraphs
    final_paragraphs = paragraphs[prev_idx:]
    if final_paragraphs:
        final_content = "\n".join(p['text'] for p in final_paragraphs)
        result_blocks.append({
            "uuid": current_uuid,
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
            
            # Add UUID placeholder at the beginning
            block_paragraphs_with_uuid = [{
                'text': '',
                'para_id': block['uuid'],
                'is_table': False
            }] + block_paragraphs
            
            # Recursively split this oversized block
            # The recursive call will either find more anchors or raise an error
            sub_blocks = split_long_block(
                block['heading'],
                block_paragraphs_with_uuid,
                block['parent_headings']
            )
            validated_blocks.extend(sub_blocks)
        else:
            # Remove internal _paragraphs field before adding to final output
            block.pop('_paragraphs', None)
            validated_blocks.append(block)
    
    return validated_blocks


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
    Parse styles.xml to extract outlineLvl definitions for each style.
    
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
    
    styles_outline = {}
    
    try:
        with zipfile.ZipFile(docx_path, 'r') as zf:
            if 'word/styles.xml' not in zf.namelist():
                return styles_outline
            
            tree = ET.parse(zf.open('word/styles.xml'))
            root = tree.getroot()
            
            # Parse style definitions
            for style in root.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}style'):
                style_id = style.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}styleId')
                if not style_id:
                    continue
                
                # Check for outlineLvl in style's pPr
                pPr = style.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPr')
                if pPr is not None:
                    outline_lvl_elem = pPr.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}outlineLvl')
                    if outline_lvl_elem is not None:
                        level = int(outline_lvl_elem.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val'))
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


def extract_audit_blocks(file_path: str) -> list:
    """
    Extract text blocks from a DOCX file for auditing.
    
    Uses python-docx with custom numbering resolver to:
    1. Capture automatic numbering (list labels)
    2. Split document by headings
    3. Convert tables to JSON (2D array)
    4. Validate heading lengths and table sizes
    5. Split long blocks using anchor paragraphs
    
    Args:
        file_path: Path to the DOCX file
        
    Returns:
        List of block dictionaries with heading, content, type, and metadata
    """
    doc = Document(file_path)
    resolver = NumberingResolver(file_path)
    styles_outline = parse_styles_outline_levels(file_path)
    
    blocks = []
    current_heading = "Preface/Uncategorized"
    current_heading_para_id = None  # paraId of current heading paragraph
    current_heading_stack = []
    current_paragraphs = []  # Track paragraphs with metadata for splitting
    current_first_content_para_id = None  # For Preface blocks without heading
    
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
            for run in element.findall('.//w:r', {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}):
                for t in run.findall('w:t', {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}):
                    if t.text:
                        para_text += t.text
            
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
                
                # Save previous block with splitting if needed
                if current_paragraphs:
                    parent_headings_for_block = current_heading_stack[:-1] if current_heading_stack else []
                    
                    # Add heading's para_id at the beginning for UUID tracking
                    if current_heading_para_id:
                        current_paragraphs.insert(0, {
                            'text': '',  # Empty text for heading UUID placeholder
                            'para_id': current_heading_para_id,
                            'is_table': False
                        })
                    elif current_first_content_para_id:
                        # For Preface blocks, use first content para_id
                        current_paragraphs[0]['para_id'] = current_first_content_para_id
                    
                    # Split long blocks if needed
                    split_blocks = split_long_block(current_heading, current_paragraphs, parent_headings_for_block)
                    blocks.extend(split_blocks)
                    
                    current_paragraphs = []
                    current_first_content_para_id = None  # Reset for next block
                
                # Convert 0-based to 1-based level
                level = outline_level + 1
                
                # Update heading stack and current heading paraId
                current_heading_stack = current_heading_stack[:max(level - 1, 0)]
                current_heading_stack.append(full_text)
                current_heading = full_text
                current_heading_para_id = heading_para_id
            else:
                # Regular paragraph content
                # Extract paraId and track for Preface blocks
                para_id = extract_para_id(element)
                if not current_first_content_para_id and not current_heading_para_id:
                    # This is the first content paragraph under Preface
                    current_first_content_para_id = para_id
                
                # Store paragraph with metadata for potential splitting
                current_paragraphs.append({
                    'text': full_text,
                    'para_id': para_id,
                    'is_table': False
                })
            
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
            table_data = TableExtractor.extract(table, numbering_resolver=resolver)
            
            # Convert table to JSON
            table_json = json.dumps(table_data, ensure_ascii=False)
            
            # Validate table length
            validate_table_length(table_json, current_heading)
            
            # Store table as a paragraph with special marker
            # Generate a pseudo para_id for the table (use a hash of table content)
            table_para_id = hashlib.md5(table_json.encode('utf-8')).hexdigest()[:8]
            current_paragraphs.append({
                'text': f"<table>{table_json}</table>",
                'para_id': table_para_id,
                'is_table': True
            })
            
            # Reset numbering tracking after table (table end boundary)
            resolver.reset_tracking_state()
    
    # Save final block with splitting if needed
    if current_paragraphs:
        parent_headings_for_block = current_heading_stack[:-1] if current_heading_stack else []
        
        # Add heading's para_id at the beginning for UUID tracking
        if current_heading_para_id:
            current_paragraphs.insert(0, {
                'text': '',  # Empty text for heading UUID placeholder
                'para_id': current_heading_para_id,
                'is_table': False
            })
        elif current_first_content_para_id:
            # For Preface blocks, use first content para_id
            current_paragraphs[0]['para_id'] = current_first_content_para_id
        
        # Split long blocks if needed
        split_blocks = split_long_block(current_heading, current_paragraphs, parent_headings_for_block)
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
    blocks = extract_audit_blocks(args.document)
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
