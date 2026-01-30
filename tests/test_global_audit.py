#!/usr/bin/env python3
"""
ABOUTME: Unit tests for global audit functions in run_audit.py
ABOUTME: Tests chunk_items_by_token_limit, normalize_extracted_fields, merge_global_violations
"""

import json
import sys
from pathlib import Path

# Add skills/doc-audit/scripts directory to path (must be before import)
_scripts_dir = Path(__file__).parent.parent / 'skills' / 'doc-audit' / 'scripts'
sys.path.insert(0, str(_scripts_dir))

from run_audit import (  # noqa: E402  # type: ignore
    chunk_items_by_token_limit,
    merge_global_violations,
    strip_global_violations,
)
from prompt import (  # noqa: E402  # type: ignore
    normalize_extracted_fields,
    build_global_extract_system_prompt,
    build_global_verify_system_prompt,
)
from utils import estimate_tokens  # noqa: E402  # type: ignore


# ============================================================
# Helper Functions
# ============================================================

def create_sample_rule(rule_id: str = "G001", topic: str = "Test Topic") -> dict:
    """Create a sample global rule for testing."""
    return {
        "id": rule_id,
        "category": "consistency",
        "topic": topic,
        "extraction": "Extract test data",
        "verification": "Check consistency of test data",
        "extracted_entity": "Test Entity",
        "extracted_fields": [
            {"name": "field1", "desc": "First field", "evidence": "Evidence for field1"},
            {"name": "field2", "desc": "Second field", "evidence": "Evidence for field2"}
        ]
    }


def create_sample_item(uuid: str, entity: str = "TestEntity") -> dict:
    """Create a sample extracted item for testing."""
    return {
        "uuid": uuid,
        "uuid_end": uuid,
        "p_heading": f"Heading {uuid}",
        "entity": entity,
        "fields": [
            {"name": "field1", "value": "value1", "evidence": "evidence text 1"},
            {"name": "field2", "value": "value2", "evidence": "evidence text 2"}
        ]
    }


def create_sample_violation(rule_id: str, uuid: str) -> dict:
    """Create a sample violation for testing."""
    return {
        "rule_id": rule_id,
        "uuid": uuid,
        "uuid_end": uuid,
        "violation_text": "sample violation text",
        "violation_reason": "sample reason",
        "fix_action": "manual",
        "revised_text": "sample guidance"
    }


def create_sample_block(uuid: str) -> dict:
    """Create a sample block for testing."""
    return {
        "uuid": uuid,
        "uuid_end": uuid,
        "heading": f"Block Heading {uuid}",
        "content": f"Block content for {uuid}"
    }


# ============================================================
# Tests: normalize_extracted_fields
# ============================================================

class TestNormalizeExtractedFields:
    """Tests for normalize_extracted_fields function"""

    def test_new_format_with_name(self):
        """Test normalization of new format with 'name' key"""
        rule = {
            "extracted_fields": [
                {"name": "delivery_schedule", "desc": "Delivery time", "evidence": "Quote from text"}
            ]
        }
        result = normalize_extracted_fields(rule)
        assert len(result) == 1
        assert result[0]["name"] == "delivery_schedule"
        assert result[0]["desc"] == "Delivery time"
        assert result[0]["evidence_desc"] == "Quote from text"

    def test_legacy_format(self):
        """Test normalization of legacy format with field_name as key"""
        rule = {
            "extracted_fields": [
                {"delivery_schedule": "Delivery time info", "evidence": "Quote from text"},
                {"acceptance_criteria": "Criteria info", "evidence": "Another quote"}
            ]
        }
        result = normalize_extracted_fields(rule)
        assert len(result) == 2
        assert result[0]["name"] == "delivery_schedule"
        assert result[0]["desc"] == "Delivery time info"
        assert result[0]["evidence_desc"] == "Quote from text"
        assert result[1]["name"] == "acceptance_criteria"
        assert result[1]["desc"] == "Criteria info"

    def test_empty_fields(self):
        """Test with empty extracted_fields"""
        rule = {"extracted_fields": []}
        result = normalize_extracted_fields(rule)
        assert result == []

    def test_missing_fields(self):
        """Test with missing extracted_fields key"""
        rule = {}
        result = normalize_extracted_fields(rule)
        assert result == []

    def test_non_dict_items_skipped(self):
        """Test that non-dict items are skipped"""
        rule = {
            "extracted_fields": [
                {"name": "valid", "desc": "Valid field", "evidence": "Quote"},
                "invalid_string",
                123
            ]
        }
        result = normalize_extracted_fields(rule)
        assert len(result) == 1
        assert result[0]["name"] == "valid"


# ============================================================
# Tests: chunk_items_by_token_limit
# ============================================================

class TestChunkItemsByTokenLimit:
    """Tests for chunk_items_by_token_limit function"""

    def test_single_chunk_within_limit(self):
        """Test items that fit in a single chunk"""
        rule = create_sample_rule()
        items = [create_sample_item(f"UUID{i}") for i in range(3)]

        # Use a high limit so everything fits
        chunks = chunk_items_by_token_limit(rule, items, max_tokens=50000)

        assert len(chunks) == 1
        assert len(chunks[0]) == 3

    def test_multiple_chunks(self):
        """Test items that need to be split into multiple chunks"""
        rule = create_sample_rule()
        # Create items with larger content
        items = []
        for i in range(10):
            item = create_sample_item(f"UUID{i}")
            item["fields"] = [
                {"name": "field1", "value": "A" * 500, "evidence": "B" * 500},
                {"name": "field2", "value": "C" * 500, "evidence": "D" * 500}
            ]
            items.append(item)

        # Use a lower limit to force splitting
        chunks = chunk_items_by_token_limit(rule, items, max_tokens=2000)

        assert len(chunks) > 1
        # All items should be accounted for
        total_items = sum(len(chunk) for chunk in chunks)
        assert total_items == 10

    def test_empty_items(self):
        """Test with empty items list"""
        rule = create_sample_rule()
        chunks = chunk_items_by_token_limit(rule, [], max_tokens=50000)
        assert chunks == []

    def test_single_large_item(self):
        """Test with a single item that exceeds limit (should still be included)"""
        rule = create_sample_rule()
        item = create_sample_item("UUID1")
        item["fields"] = [
            {"name": "field1", "value": "X" * 10000, "evidence": "Y" * 10000}
        ]

        chunks = chunk_items_by_token_limit(rule, [item], max_tokens=100)

        # Should still have one chunk with the item
        assert len(chunks) == 1
        assert len(chunks[0]) == 1


# ============================================================
# Tests: merge_global_violations
# ============================================================

class TestMergeGlobalViolations:
    """Tests for merge_global_violations function"""

    def test_merge_into_existing_entry(self):
        """Test merging violations into an existing manifest entry"""
        blocks = [create_sample_block("UUID1"), create_sample_block("UUID2")]
        uuid_to_block_idx = {"UUID1": 0, "UUID2": 1}
        rule_category_map = {"G001": "consistency"}

        existing_entries = [
            (0, {
                "uuid": "UUID1",
                "uuid_end": "UUID1",
                "p_heading": "Heading 1",
                "p_content": "Content 1",
                "is_violation": True,
                "violations": [
                    {"rule_id": "R001", "violation_text": "existing violation", "uuid_end": "UUID1"}
                ]
            })
        ]

        global_violations = [
            create_sample_violation("G001", "UUID1")
        ]

        result = merge_global_violations(
            existing_entries, global_violations, blocks, uuid_to_block_idx, rule_category_map
        )

        # Should have one entry with two violations
        assert len(result) == 1
        _, entry = result[0]
        assert len(entry["violations"]) == 2
        assert entry["is_violation"] is True

    def test_create_new_entry_for_new_uuid(self):
        """Test creating a new entry when UUID doesn't exist in manifest"""
        blocks = [create_sample_block("UUID1"), create_sample_block("UUID2")]
        uuid_to_block_idx = {"UUID1": 0, "UUID2": 1}
        rule_category_map = {"G001": "consistency"}

        existing_entries = []

        global_violations = [
            create_sample_violation("G001", "UUID2")
        ]

        result = merge_global_violations(
            existing_entries, global_violations, blocks, uuid_to_block_idx, rule_category_map
        )

        assert len(result) == 1
        _, entry = result[0]
        assert entry["uuid"] == "UUID2"
        assert len(entry["violations"]) == 1
        assert entry["is_violation"] is True

    def test_deduplication(self):
        """Test that duplicate violations are not added"""
        blocks = [create_sample_block("UUID1")]
        uuid_to_block_idx = {"UUID1": 0}
        rule_category_map = {"G001": "consistency"}

        existing_entries = [
            (0, {
                "uuid": "UUID1",
                "uuid_end": "UUID1",
                "p_heading": "Heading 1",
                "p_content": "Content 1",
                "is_violation": True,
                "violations": [
                    {"rule_id": "G001", "violation_text": "sample violation text", "uuid_end": "UUID1"}
                ]
            })
        ]

        # Add the same violation again
        global_violations = [
            create_sample_violation("G001", "UUID1")
        ]

        result = merge_global_violations(
            existing_entries, global_violations, blocks, uuid_to_block_idx, rule_category_map
        )

        _, entry = result[0]
        # Should still have only one violation (duplicate skipped)
        assert len(entry["violations"]) == 1

    def test_missing_uuid_warning(self, capsys):
        """Test that violations with missing UUID log a warning"""
        blocks = [create_sample_block("UUID1")]
        uuid_to_block_idx = {"UUID1": 0}
        rule_category_map = {"G001": "consistency"}

        global_violations = [
            create_sample_violation("G001", "NONEXISTENT")
        ]

        merge_global_violations(
            [], global_violations, blocks, uuid_to_block_idx, rule_category_map
        )

        captured = capsys.readouterr()
        assert "not found in blocks" in captured.err

    def test_empty_violations(self):
        """Test with no violations to merge"""
        blocks = [create_sample_block("UUID1")]
        uuid_to_block_idx = {"UUID1": 0}
        rule_category_map = {"G001": "consistency"}

        existing_entries = [
            (0, {
                "uuid": "UUID1",
                "uuid_end": "UUID1",
                "p_heading": "Heading 1",
                "p_content": "Content 1",
                "is_violation": False,
                "violations": []
            })
        ]

        result = merge_global_violations(
            existing_entries, [], blocks, uuid_to_block_idx, rule_category_map
        )

        assert len(result) == 1
        _, entry = result[0]
        assert entry["is_violation"] is False


# ============================================================
# Tests: build_global_extract_system_prompt
# ============================================================

class TestBuildGlobalExtractSystemPrompt:
    """Tests for build_global_extract_system_prompt function"""

    def test_prompt_contains_rule_info(self):
        """Test that the prompt contains rule information"""
        rules = [create_sample_rule("G001", "货物交收信息")]
        prompt = build_global_extract_system_prompt(rules)

        assert "G001" in prompt
        assert "货物交收信息" in prompt
        assert "Extract test data" in prompt
        assert "field1" in prompt
        assert "field2" in prompt

    def test_prompt_structure(self):
        """Test that the prompt has expected structure"""
        rules = [create_sample_rule()]
        prompt = build_global_extract_system_prompt(rules)

        assert "information extractor" in prompt.lower()
        assert "results" in prompt
        assert "extracted_results" in prompt
        assert "entity" in prompt


# ============================================================
# Tests: build_global_verify_system_prompt
# ============================================================

class TestBuildGlobalVerifySystemPrompt:
    """Tests for build_global_verify_system_prompt function"""

    def test_prompt_contains_rule_info(self):
        """Test that the prompt contains verification rule information"""
        rule = create_sample_rule("G001", "一致性检查")
        prompt = build_global_verify_system_prompt(rule)

        assert "G001" in prompt
        assert "一致性检查" in prompt
        assert "Check consistency" in prompt

    def test_prompt_structure(self):
        """Test that the prompt has expected structure"""
        rule = create_sample_rule()
        prompt = build_global_verify_system_prompt(rule)

        assert "cross-reference auditor" in prompt.lower()
        assert "violations" in prompt
        assert "violation_text" in prompt
        assert "manual" in prompt


# ============================================================
# Tests: utils.estimate_tokens
# ============================================================

class TestLoadRulesBackwardCompatibility:
    """Tests for backward compatibility of load_rules with legacy type values"""

    def test_block_level_normalized_to_block(self, tmp_path):
        """Test that legacy 'block_level' type is normalized to 'block'"""
        from run_audit import load_rules  # type: ignore
        
        # Create a legacy rules file with block_level type
        legacy_rules = {
            "version": "1.0",
            "type": "block_level",
            "rules": [
                {"id": "R001", "description": "Test rule", "severity": "high", "category": "test"}
            ]
        }
        
        rules_file = tmp_path / "legacy_rules.json"
        rules_file.write_text(json.dumps(legacy_rules, ensure_ascii=False))
        
        # Load rules and verify normalization
        loaded_rules = load_rules(str(rules_file))
        
        # Verify normalization occurred
        assert len(loaded_rules) == 1
        assert loaded_rules[0]["type"] == "block"  # Should be normalized
        assert loaded_rules[0]["id"] == "R001"

    def test_individual_rule_block_level_normalized(self, tmp_path):
        """Test that individual rules with block_level are normalized"""
        from run_audit import load_rules  # type: ignore
        
        rules_data = {
            "rules": [
                {"id": "R001", "type": "block_level", "description": "Test", "severity": "high", "category": "test"},
                {"id": "R002", "type": "block", "description": "Test2", "severity": "medium", "category": "test"},
                {"id": "R003", "description": "Test3", "severity": "low", "category": "test"}  # No type, should get default
            ]
        }
        
        rules_file = tmp_path / "mixed_rules.json"
        rules_file.write_text(json.dumps(rules_data, ensure_ascii=False))
        
        loaded_rules = load_rules(str(rules_file))
        
        assert len(loaded_rules) == 3
        assert loaded_rules[0]["type"] == "block"  # Normalized from block_level
        assert loaded_rules[1]["type"] == "block"  # Already block
        assert loaded_rules[2]["type"] == "block"  # Default

    def test_block_and_global_rules_preserved(self, tmp_path):
        """Test that block and global types are preserved correctly"""
        from run_audit import load_rules  # type: ignore
        
        rules_data = {
            "rules": [
                {"id": "R001", "type": "block", "description": "Block rule", "severity": "high", "category": "test"},
                {"id": "G001", "type": "global", "description": "Global rule", "severity": "medium", "category": "test"}
            ]
        }
        
        rules_file = tmp_path / "mixed_types.json"
        rules_file.write_text(json.dumps(rules_data, ensure_ascii=False))
        
        loaded_rules = load_rules(str(rules_file))
        
        assert len(loaded_rules) == 2
        assert loaded_rules[0]["type"] == "block"
        assert loaded_rules[1]["type"] == "global"


class TestChunkItemsSystemPromptOverhead:
    """
    Test that chunk_items_by_token_limit accounts for system prompt token overhead.
    """
    
    def test_system_prompt_overhead_reduces_available_tokens(self):
        """Verify that system prompt tokens are subtracted from max_tokens budget."""
        
        rule = {
            "id": "G001",
            "topic": "Test consistency rule with long topic description",
            "verification": "Verify that extracted items are consistent " * 50  # Long text
        }
        
        items = [
            {"uuid": "uuid1", "fields": [{"name": "field1", "value": "value1", "evidence": "evidence1"}]},
            {"uuid": "uuid2", "fields": [{"name": "field2", "value": "value2", "evidence": "evidence2"}]},
        ]
        
        max_tokens = 1000
        chunks = chunk_items_by_token_limit(rule, items, max_tokens)
        
        # Should successfully chunk without errors
        assert len(chunks) > 0
        # All items should be included
        total_items = sum(len(chunk) for chunk in chunks)
        assert total_items == len(items)
    
    def test_large_system_prompt_fallback(self, capsys):
        """Test fallback behavior when system prompt exceeds max_tokens."""
        rule = {
            "id": "G002",
            "topic": "X" * 10000,  # Extremely long topic
            "verification": "Y" * 10000  # Extremely long verification
        }
        
        items = [{"uuid": "uuid1", "fields": []}]
        max_tokens = 100  # Small limit
        
        chunks = chunk_items_by_token_limit(rule, items, max_tokens)
        
        # Should use fallback and emit warning
        captured = capsys.readouterr()
        assert "exceeds GLOBAL_AUDIT_MAX_TOKENS" in captured.err
        assert "Using fallback budget" in captured.err
        
        # Should still produce chunks
        assert len(chunks) > 0
    
    def test_single_item_exceeds_budget_warning(self, capsys):
        """Test warning when a single item exceeds available token budget."""
        rule = {
            "id": "G003",
            "topic": "Simple topic",
            "verification": "Simple verification"
        }
        
        # Create a very large single item
        large_item = {
            "uuid": "uuid1",
            "fields": [{"name": f"field{i}", "value": "X" * 1000, "evidence": "E" * 1000} for i in range(100)]
        }
        
        max_tokens = 5000
        chunks = chunk_items_by_token_limit(rule, [large_item], max_tokens)
        
        # Should emit warning with breakdown
        captured = capsys.readouterr()
        assert "exceeds token budget" in captured.err
        assert "(user)" in captured.err
        assert "(system)" in captured.err
        
        # Should still include the item
        assert len(chunks) == 1
        assert len(chunks[0]) == 1


class TestEstimateTokens:
    """Tests for estimate_tokens function"""

    def test_empty_string(self):
        """Test with empty string"""
        assert estimate_tokens("") == 0

    def test_english_text(self):
        """Test with English text"""
        text = "Hello, this is a test sentence."
        tokens = estimate_tokens(text)
        # English text should be roughly 0.4 tokens per char + buffer
        assert tokens > 0
        assert tokens < len(text)  # Should be less than char count

    def test_chinese_text(self):
        """Test with Chinese text"""
        text = "这是一段中文测试文本"
        tokens = estimate_tokens(text)
        # Chinese text should be roughly 0.75 tokens per char
        assert tokens > 0
        # Chinese characters take more tokens per character
        assert tokens > len(text) * 0.5

    def test_json_structure(self):
        """Test with JSON-like structure"""
        text = '{"key": "value", "array": [1, 2, 3]}'
        tokens = estimate_tokens(text)
        # JSON structure chars should contribute significantly
        assert tokens > 0

    def test_mixed_content(self):
        """Test with mixed Chinese/English/JSON"""
        text = '{"name": "测试名称", "value": 123}'
        tokens = estimate_tokens(text)
        assert tokens > 0


# ============================================================
# Tests: strip_global_violations
# ============================================================

class TestStripGlobalViolations:
    """Tests for strip_global_violations function"""

    def test_empty_global_rule_ids_returns_unchanged(self):
        """Test that empty global_rule_ids returns entries unchanged"""
        entries = [
            (0, {
                "uuid": "UUID1",
                "uuid_end": "UUID1",
                "is_violation": True,
                "violations": [
                    {"rule_id": "G001", "violation_text": "text1"},
                    {"rule_id": "R001", "violation_text": "text2"}
                ]
            })
        ]
        
        cleaned, removed = strip_global_violations(entries, set())
        
        assert cleaned == entries
        assert removed == 0

    def test_strips_global_violations_correctly(self):
        """Test that violations with global rule IDs are removed"""
        entries = [
            (0, {
                "uuid": "UUID1",
                "uuid_end": "UUID1",
                "is_violation": True,
                "violations": [
                    {"rule_id": "G001", "violation_text": "global violation 1"},
                    {"rule_id": "G002", "violation_text": "global violation 2"}
                ]
            })
        ]
        
        global_rule_ids = {"G001", "G002"}
        cleaned, removed = strip_global_violations(entries, global_rule_ids)
        
        assert len(cleaned) == 1
        _, entry = cleaned[0]
        assert len(entry["violations"]) == 0
        assert entry["is_violation"] is False
        assert removed == 2

    def test_preserves_block_violations(self):
        """Test that block-level violations are preserved"""
        entries = [
            (0, {
                "uuid": "UUID1",
                "uuid_end": "UUID1",
                "is_violation": True,
                "violations": [
                    {"rule_id": "R001", "violation_text": "block violation 1"},
                    {"rule_id": "R002", "violation_text": "block violation 2"}
                ]
            })
        ]
        
        global_rule_ids = {"G001", "G002"}
        cleaned, removed = strip_global_violations(entries, global_rule_ids)
        
        assert len(cleaned) == 1
        _, entry = cleaned[0]
        assert len(entry["violations"]) == 2
        assert entry["is_violation"] is True
        assert removed == 0

    def test_updates_is_violation_flag(self):
        """Test that is_violation is updated correctly after stripping"""
        # Entry with only global violations
        entries = [
            (0, {
                "uuid": "UUID1",
                "uuid_end": "UUID1",
                "is_violation": True,
                "violations": [
                    {"rule_id": "G001", "violation_text": "global only"}
                ]
            })
        ]
        
        global_rule_ids = {"G001"}
        cleaned, removed = strip_global_violations(entries, global_rule_ids)
        
        _, entry = cleaned[0]
        assert entry["is_violation"] is False
        assert removed == 1

    def test_returns_correct_removed_count(self):
        """Test that removed count is accurate"""
        entries = [
            (0, {
                "uuid": "UUID1",
                "uuid_end": "UUID1",
                "is_violation": True,
                "violations": [
                    {"rule_id": "G001", "violation_text": "v1"},
                    {"rule_id": "G002", "violation_text": "v2"}
                ]
            }),
            (1, {
                "uuid": "UUID2",
                "uuid_end": "UUID2",
                "is_violation": True,
                "violations": [
                    {"rule_id": "G001", "violation_text": "v3"},
                    {"rule_id": "R001", "violation_text": "v4"}
                ]
            })
        ]
        
        global_rule_ids = {"G001", "G002"}
        cleaned, removed = strip_global_violations(entries, global_rule_ids)
        
        assert removed == 3  # 2 from UUID1, 1 from UUID2

    def test_mixed_violations_partial_removal(self):
        """Test partial removal when entry has both block and global violations"""
        entries = [
            (0, {
                "uuid": "UUID1",
                "uuid_end": "UUID1",
                "is_violation": True,
                "violations": [
                    {"rule_id": "R001", "violation_text": "block violation"},
                    {"rule_id": "G001", "violation_text": "global violation 1"},
                    {"rule_id": "R002", "violation_text": "another block"},
                    {"rule_id": "G002", "violation_text": "global violation 2"}
                ]
            })
        ]
        
        global_rule_ids = {"G001", "G002"}
        cleaned, removed = strip_global_violations(entries, global_rule_ids)
        
        _, entry = cleaned[0]
        assert len(entry["violations"]) == 2
        assert entry["violations"][0]["rule_id"] == "R001"
        assert entry["violations"][1]["rule_id"] == "R002"
        assert entry["is_violation"] is True
        assert removed == 2

    def test_entry_without_violations_unchanged(self):
        """Test that entries without violations field are preserved"""
        entries = [
            (0, {
                "uuid": "UUID1",
                "uuid_end": "UUID1",
                "is_violation": False
            })
        ]
        
        global_rule_ids = {"G001"}
        cleaned, removed = strip_global_violations(entries, global_rule_ids)
        
        assert len(cleaned) == 1
        assert cleaned[0] == entries[0]
        assert removed == 0
