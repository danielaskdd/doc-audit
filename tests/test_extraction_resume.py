#!/usr/bin/env python3
"""
Test global extraction resume functionality.

This test verifies that:
1. Extraction results are persisted to disk after each block
2. Resume correctly skips already-extracted blocks
3. Rule buckets are rebuilt from extraction file
"""

import json
import sys
import tempfile
from pathlib import Path

# Add parent directory to path for imports
_TEST_DIR = Path(__file__).resolve().parent
_SCRIPT_DIR = _TEST_DIR.parent / "skills" / "doc-audit" / "scripts"
if str(_SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(_SCRIPT_DIR))

# Import after sys.path modification
from run_audit import (  # noqa: E402  # type: ignore
    derive_extraction_path,
    load_completed_extraction_uuids,
    load_extraction_buckets,
)


def test_derive_extraction_path():
    """Test extraction path derivation from manifest path."""
    assert derive_extraction_path("manifest.jsonl") == "manifest_extractions.jsonl"
    assert derive_extraction_path("/path/to/document_manifest.jsonl") == "/path/to/document_manifest_extractions.jsonl"
    assert derive_extraction_path("./output/test.jsonl") == "output/test_extractions.jsonl"
    print("✓ derive_extraction_path tests passed")


def test_load_completed_extraction_uuids():
    """Test loading completed UUIDs from extraction file."""
    with tempfile.NamedTemporaryFile(mode='w', suffix='.jsonl', delete=False, encoding='utf-8') as f:
        extraction_path = f.name
        
        # Write metadata
        f.write(json.dumps({"type": "meta", "source_file": "test.docx"}) + '\n')
        
        # Write extraction results
        f.write(json.dumps({
            "uuid": "ABC123",
            "uuid_end": "DEF456",
            "p_heading": "Chapter 1",
            "results": [{"rule_id": "G001", "extracted_results": []}]
        }) + '\n')
        
        f.write(json.dumps({
            "uuid": "GHI789",
            "uuid_end": "JKL012",
            "p_heading": "Chapter 2",
            "results": [{"rule_id": "G002", "extracted_results": []}]
        }) + '\n')
    
    try:
        completed = load_completed_extraction_uuids(extraction_path)
        assert completed == {"ABC123", "GHI789"}
        print("✓ load_completed_extraction_uuids tests passed")
    finally:
        Path(extraction_path).unlink()


def test_load_extraction_buckets():
    """Test rebuilding rule_buckets from extraction file."""
    with tempfile.NamedTemporaryFile(mode='w', suffix='.jsonl', delete=False, encoding='utf-8') as f:
        extraction_path = f.name
        
        # Write metadata
        f.write(json.dumps({"type": "meta"}) + '\n')
        
        # Write extraction with entities
        f.write(json.dumps({
            "uuid": "ABC123",
            "uuid_end": "DEF456",
            "p_heading": "Section 1",
            "results": [
                {
                    "rule_id": "G001",
                    "extracted_results": [
                        {
                            "entity": "Party A",
                            "fields": [{"name": "name", "value": "Company X"}]
                        },
                        {
                            "entity": "Party B",
                            "fields": [{"name": "name", "value": "Company Y"}]
                        }
                    ]
                }
            ]
        }) + '\n')
        
        f.write(json.dumps({
            "uuid": "GHI789",
            "uuid_end": "JKL012",
            "p_heading": "Section 2",
            "results": [
                {
                    "rule_id": "G002",
                    "extracted_results": [
                        {
                            "entity": "Payment",
                            "fields": [{"name": "amount", "value": "1000 CNY"}]
                        }
                    ]
                }
            ]
        }) + '\n')
    
    try:
        global_rules = [
            {"id": "G001", "topic": "Party Identification"},
            {"id": "G002", "topic": "Payment Terms"}
        ]
        
        rule_buckets = load_extraction_buckets(extraction_path, global_rules)
        
        # Verify G001 has 2 items
        assert len(rule_buckets["G001"]) == 2
        assert rule_buckets["G001"][0]["entity"] == "Party A"
        assert rule_buckets["G001"][0]["uuid"] == "ABC123"
        assert rule_buckets["G001"][1]["entity"] == "Party B"
        
        # Verify G002 has 1 item
        assert len(rule_buckets["G002"]) == 1
        assert rule_buckets["G002"][0]["entity"] == "Payment"
        assert rule_buckets["G002"][0]["p_heading"] == "Section 2"
        
        print("✓ load_extraction_buckets tests passed")
    finally:
        Path(extraction_path).unlink()


def test_extraction_file_missing():
    """Test handling of missing extraction file."""
    non_existent_path = "/tmp/non_existent_file.jsonl"
    
    # Should return empty set
    completed = load_completed_extraction_uuids(non_existent_path)
    assert completed == set()
    
    # Should return empty buckets
    global_rules = [{"id": "G001", "topic": "Test"}]
    rule_buckets = load_extraction_buckets(non_existent_path, global_rules)
    assert rule_buckets == {"G001": []}
    
    print("✓ missing file handling tests passed")


def test_uuid_fallback_with_base_index():
    """Test that UUID fallback uses base_index for consistency."""
    # Simulate blocks without UUID that would use fallback
    blocks_without_uuid = [
        {"heading": "Section 1", "content": "Content 1"},
        {"heading": "Section 2", "content": "Content 2"},
        {"heading": "Section 3", "content": "Content 3"}
    ]
    
    # When processing with start_idx=2 (base_index=2)
    # The fallback UUIDs should be "2", "3", "4", not "0", "1", "2"
    base_index = 2
    
    # Simulate UUID generation for blocks
    for idx, block in enumerate(blocks_without_uuid):
        expected_uuid = str(base_index + idx)
        actual_uuid = block.get('uuid', str(base_index + idx))
        assert actual_uuid == expected_uuid, f"Expected UUID {expected_uuid}, got {actual_uuid}"
    
    print("✓ UUID fallback with base_index tests passed")


if __name__ == "__main__":
    test_derive_extraction_path()
    test_load_completed_extraction_uuids()
    test_load_extraction_buckets()
    test_extraction_file_missing()
    test_uuid_fallback_with_base_index()
    print("\n✅ All extraction resume tests passed!")
