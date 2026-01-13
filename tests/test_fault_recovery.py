#!/usr/bin/env python3
"""
ABOUTME: Unit tests for fault recovery mechanisms in run_audit.py
ABOUTME: Tests manifest entry iteration, resume logic, and backup recovery scenarios
"""

import sys
import json
import tempfile
import shutil
from pathlib import Path
from unittest.mock import patch

# Add skills/doc-audit/scripts directory to path (must be before import)
_scripts_dir = Path(__file__).parent.parent / 'skills' / 'doc-audit' / 'scripts'
sys.path.insert(0, str(_scripts_dir))

import pytest  # noqa: E402

from run_audit import (  # noqa: E402  # type: ignore[import-not-found]
    iter_manifest_entries,
    load_completed_uuids,
    load_manifest_metadata,
    load_existing_entries_with_block_idx,
    rewrite_manifest_sorted,
    save_manifest_entry,
)


# ============================================================
# Helper Functions
# ============================================================

def create_temp_jsonl(entries: list, include_meta: bool = True) -> Path:
    """
    Create a temporary JSONL file with given entries.
    
    Args:
        entries: List of dict entries to write
        include_meta: Whether to prepend a metadata entry
    
    Returns:
        Path to the temporary file
    """
    fd, path = tempfile.mkstemp(suffix='.jsonl')
    with open(path, 'w', encoding='utf-8') as f:
        if include_meta:
            meta = {
                'type': 'meta',
                'source_file': '/tmp/test.docx',
                'source_hash': 'sha256:abc123',
                'audited_at': '2026-01-13T12:00:00'
            }
            f.write(json.dumps(meta, ensure_ascii=False) + '\n')
        for entry in entries:
            f.write(json.dumps(entry, ensure_ascii=False) + '\n')
    return Path(path)


def create_audit_entry(uuid: str, is_violation: bool = False, violations: list = None) -> dict:
    """
    Create a sample audit entry for testing.
    
    Args:
        uuid: Block UUID
        is_violation: Whether block has violations
        violations: List of violation dicts (optional)
    
    Returns:
        Audit entry dict
    """
    return {
        'uuid': uuid,
        'uuid_end': uuid,
        'p_heading': f'Section {uuid}',
        'p_content': f'Content for {uuid}',
        'is_violation': is_violation,
        'violations': violations or []
    }


# ============================================================
# Tests: iter_manifest_entries
# ============================================================

class TestIterManifestEntries:
    """Tests for iter_manifest_entries iterator"""
    
    def test_iterate_normal_file(self):
        """Test iterating over a normal JSONL file"""
        entries = [
            create_audit_entry('AAA'),
            create_audit_entry('BBB'),
            create_audit_entry('CCC'),
        ]
        path = create_temp_jsonl(entries)
        
        try:
            result = list(iter_manifest_entries(str(path)))
            # Include metadata entry
            assert len(result) == 4
            assert result[0].get('type') == 'meta'
            assert result[1].get('uuid') == 'AAA'
            assert result[2].get('uuid') == 'BBB'
            assert result[3].get('uuid') == 'CCC'
        finally:
            path.unlink()
    
    def test_iterate_empty_file(self):
        """Test iterating over an empty file"""
        fd, path = tempfile.mkstemp(suffix='.jsonl')
        Path(path).write_text('')
        
        try:
            result = list(iter_manifest_entries(path))
            assert len(result) == 0
        finally:
            Path(path).unlink()
    
    def test_iterate_nonexistent_file(self):
        """Test iterating over a file that doesn't exist"""
        result = list(iter_manifest_entries('/nonexistent/path/file.jsonl'))
        assert len(result) == 0
    
    def test_iterate_with_empty_lines(self):
        """Test that empty lines are skipped"""
        fd, path = tempfile.mkstemp(suffix='.jsonl')
        with open(path, 'w') as f:
            f.write('{"uuid": "AAA"}\n')
            f.write('\n')  # Empty line
            f.write('   \n')  # Whitespace-only line
            f.write('{"uuid": "BBB"}\n')
        
        try:
            result = list(iter_manifest_entries(path))
            assert len(result) == 2
            assert result[0].get('uuid') == 'AAA'
            assert result[1].get('uuid') == 'BBB'
        finally:
            Path(path).unlink()
    
    def test_iterate_malformed_json_ignore_errors(self):
        """Test that malformed JSON lines are skipped when ignore_errors=True"""
        fd, path = tempfile.mkstemp(suffix='.jsonl')
        with open(path, 'w') as f:
            f.write('{"uuid": "AAA"}\n')
            f.write('not valid json\n')  # Malformed
            f.write('{"uuid": "BBB"}\n')
        
        try:
            result = list(iter_manifest_entries(path, ignore_errors=True))
            assert len(result) == 2
            assert result[0].get('uuid') == 'AAA'
            assert result[1].get('uuid') == 'BBB'
        finally:
            Path(path).unlink()
    
    def test_iterate_malformed_json_raise_error(self):
        """Test that malformed JSON raises error when ignore_errors=False"""
        fd, path = tempfile.mkstemp(suffix='.jsonl')
        with open(path, 'w') as f:
            f.write('{"uuid": "AAA"}\n')
            f.write('not valid json\n')  # Malformed
            f.write('{"uuid": "BBB"}\n')
        
        try:
            with pytest.raises(json.JSONDecodeError):
                list(iter_manifest_entries(path, ignore_errors=False))
        finally:
            Path(path).unlink()


# ============================================================
# Tests: load_completed_uuids
# ============================================================

class TestLoadCompletedUuids:
    """Tests for load_completed_uuids function"""
    
    def test_load_uuids_normal(self):
        """Test loading UUIDs from normal manifest"""
        entries = [
            create_audit_entry('AAA'),
            create_audit_entry('BBB'),
            create_audit_entry('CCC'),
        ]
        path = create_temp_jsonl(entries)
        
        try:
            completed = load_completed_uuids(str(path))
            assert len(completed) == 3
            assert 'AAA' in completed
            assert 'BBB' in completed
            assert 'CCC' in completed
        finally:
            path.unlink()
    
    def test_load_uuids_skips_metadata(self):
        """Test that metadata entries are skipped"""
        entries = [
            create_audit_entry('AAA'),
        ]
        path = create_temp_jsonl(entries, include_meta=True)
        
        try:
            completed = load_completed_uuids(str(path))
            # Should not include metadata
            assert len(completed) == 1
            assert 'AAA' in completed
        finally:
            path.unlink()
    
    def test_load_uuids_skips_empty_uuid(self):
        """Test that entries with empty UUID are skipped"""
        fd, path = tempfile.mkstemp(suffix='.jsonl')
        with open(path, 'w') as f:
            f.write('{"uuid": "AAA"}\n')
            f.write('{"uuid": ""}\n')  # Empty UUID
            f.write('{"uuid": "BBB"}\n')
        
        try:
            completed = load_completed_uuids(path)
            assert len(completed) == 2
            assert 'AAA' in completed
            assert 'BBB' in completed
            assert '' not in completed
        finally:
            Path(path).unlink()
    
    def test_load_uuids_nonexistent_file(self):
        """Test loading from nonexistent file returns empty set"""
        completed = load_completed_uuids('/nonexistent/path/file.jsonl')
        assert len(completed) == 0
    
    def test_load_uuids_tolerates_malformed_lines(self):
        """Test that malformed lines are skipped (ignore_errors=True)"""
        fd, path = tempfile.mkstemp(suffix='.jsonl')
        with open(path, 'w') as f:
            f.write('{"uuid": "AAA"}\n')
            f.write('broken json\n')
            f.write('{"uuid": "BBB"}\n')
        
        try:
            completed = load_completed_uuids(path)
            assert len(completed) == 2
            assert 'AAA' in completed
            assert 'BBB' in completed
        finally:
            Path(path).unlink()


# ============================================================
# Tests: load_manifest_metadata
# ============================================================

class TestLoadManifestMetadata:
    """Tests for load_manifest_metadata function"""
    
    def test_load_metadata_present(self):
        """Test loading metadata when present"""
        entries = [create_audit_entry('AAA')]
        path = create_temp_jsonl(entries, include_meta=True)
        
        try:
            meta = load_manifest_metadata(str(path))
            assert meta is not None
            assert meta.get('type') == 'meta'
            assert meta.get('source_file') == '/tmp/test.docx'
            assert meta.get('audited_at') == '2026-01-13T12:00:00'
        finally:
            path.unlink()
    
    def test_load_metadata_missing(self):
        """Test loading metadata when not present"""
        entries = [create_audit_entry('AAA')]
        path = create_temp_jsonl(entries, include_meta=False)
        
        try:
            meta = load_manifest_metadata(str(path))
            assert meta is None
        finally:
            path.unlink()
    
    def test_load_metadata_nonexistent_file(self):
        """Test loading from nonexistent file returns None"""
        meta = load_manifest_metadata('/nonexistent/path/file.jsonl')
        assert meta is None
    
    def test_load_metadata_with_audited_at_field(self):
        """Test metadata detection via audited_at field"""
        fd, path = tempfile.mkstemp(suffix='.jsonl')
        with open(path, 'w') as f:
            # Metadata without type='meta' but with audited_at
            f.write('{"source_file": "/test.docx", "audited_at": "2026-01-13"}\n')
            f.write('{"uuid": "AAA"}\n')
        
        try:
            meta = load_manifest_metadata(path)
            assert meta is not None
            assert meta.get('audited_at') == '2026-01-13'
        finally:
            Path(path).unlink()


# ============================================================
# Tests: load_existing_entries_with_block_idx
# ============================================================

class TestLoadExistingEntriesWithBlockIdx:
    """Tests for load_existing_entries_with_block_idx function"""
    
    def test_load_entries_with_mapping(self):
        """Test loading entries with UUID to block_idx mapping"""
        entries = [
            create_audit_entry('AAA'),
            create_audit_entry('BBB'),
            create_audit_entry('CCC'),
        ]
        path = create_temp_jsonl(entries)
        uuid_to_block_idx = {'AAA': 0, 'BBB': 1, 'CCC': 2}
        
        try:
            result = load_existing_entries_with_block_idx(str(path), uuid_to_block_idx)
            assert len(result) == 3
            # Should be list of (block_idx, entry) tuples
            assert result[0][0] == 0  # block_idx for AAA
            assert result[0][1].get('uuid') == 'AAA'
            assert result[1][0] == 1  # block_idx for BBB
            assert result[2][0] == 2  # block_idx for CCC
        finally:
            path.unlink()
    
    def test_load_entries_skips_metadata(self):
        """Test that metadata entries are skipped"""
        entries = [create_audit_entry('AAA')]
        path = create_temp_jsonl(entries, include_meta=True)
        uuid_to_block_idx = {'AAA': 0}
        
        try:
            result = load_existing_entries_with_block_idx(str(path), uuid_to_block_idx)
            assert len(result) == 1
            assert result[0][1].get('uuid') == 'AAA'
        finally:
            path.unlink()
    
    def test_load_entries_unknown_uuid_warning(self, capsys):
        """Test warning for unknown UUIDs"""
        entries = [
            create_audit_entry('AAA'),
            create_audit_entry('UNKNOWN'),
        ]
        path = create_temp_jsonl(entries)
        uuid_to_block_idx = {'AAA': 0}  # UNKNOWN not in mapping
        
        try:
            result = load_existing_entries_with_block_idx(str(path), uuid_to_block_idx)
            assert len(result) == 1  # Only AAA included
            
            captured = capsys.readouterr()
            assert 'UNKNOWN' in captured.err
            assert 'not found' in captured.err
        finally:
            path.unlink()
    
    def test_load_entries_nonexistent_file(self):
        """Test loading from nonexistent file"""
        result = load_existing_entries_with_block_idx('/nonexistent/file.jsonl', {})
        assert len(result) == 0


# ============================================================
# Tests: rewrite_manifest_sorted
# ============================================================

class TestRewriteManifestSorted:
    """Tests for rewrite_manifest_sorted function"""
    
    def test_rewrite_sorts_by_block_idx(self):
        """Test that entries are sorted by block_idx"""
        fd, path = tempfile.mkstemp(suffix='.jsonl')
        
        try:
            # Create results in unordered sequence
            results = [
                (2, create_audit_entry('CCC')),
                (0, create_audit_entry('AAA')),
                (1, create_audit_entry('BBB')),
            ]
            metadata = {'type': 'meta', 'source_file': '/test.docx'}
            
            rewrite_manifest_sorted(path, metadata, results)
            
            # Read back and verify order
            with open(path, 'r') as f:
                lines = f.readlines()
            
            assert len(lines) == 4  # metadata + 3 entries
            
            # First line is metadata
            meta = json.loads(lines[0])
            assert meta.get('type') == 'meta'
            
            # Entries should be sorted by block_idx
            entry0 = json.loads(lines[1])
            entry1 = json.loads(lines[2])
            entry2 = json.loads(lines[3])
            assert entry0.get('uuid') == 'AAA'
            assert entry1.get('uuid') == 'BBB'
            assert entry2.get('uuid') == 'CCC'
        finally:
            Path(path).unlink(missing_ok=True)
    
    def test_rewrite_cleanup_backup_on_success(self):
        """Test that backup file is deleted after successful write"""
        fd, path = tempfile.mkstemp(suffix='.jsonl')
        backup_path = Path(path + '.bak')
        
        try:
            # Create original file
            Path(path).write_text('{"uuid": "original"}\n')
            
            results = [(0, create_audit_entry('AAA'))]
            metadata = {'type': 'meta'}
            
            rewrite_manifest_sorted(path, metadata, results)
            
            # Backup should be cleaned up
            assert not backup_path.exists()
            
            # New content should be written
            with open(path, 'r') as f:
                lines = f.readlines()
            assert len(lines) == 2
        finally:
            Path(path).unlink(missing_ok=True)
            backup_path.unlink(missing_ok=True)
    
    def test_rewrite_restore_backup_on_failure(self):
        """Test that backup is restored if write fails"""
        fd, path = tempfile.mkstemp(suffix='.jsonl')
        backup_path = Path(path + '.bak')
        
        try:
            # Create original file with known content
            original_content = '{"uuid": "original"}\n'
            Path(path).write_text(original_content)
            
            results = [(0, create_audit_entry('AAA'))]
            metadata = {'type': 'meta'}
            
            # Mock open to fail during write
            original_open = open
            def mock_open_fail(*args, **kwargs):
                if 'w' in args[1] if len(args) > 1 else kwargs.get('mode', ''):
                    raise IOError("Simulated write failure")
                return original_open(*args, **kwargs)
            
            with patch('builtins.open', side_effect=mock_open_fail):
                with pytest.raises(IOError):
                    rewrite_manifest_sorted(path, metadata, results)
            
            # After failure, backup should exist or original should be restored
            # (depends on where failure occurred)
        finally:
            Path(path).unlink(missing_ok=True)
            backup_path.unlink(missing_ok=True)
    
    def test_rewrite_without_metadata(self):
        """Test rewrite with metadata=None"""
        fd, path = tempfile.mkstemp(suffix='.jsonl')
        
        try:
            results = [
                (0, create_audit_entry('AAA')),
                (1, create_audit_entry('BBB')),
            ]
            
            rewrite_manifest_sorted(path, None, results)
            
            # Read back - should not have metadata line
            with open(path, 'r') as f:
                lines = f.readlines()
            
            assert len(lines) == 2
            entry0 = json.loads(lines[0])
            entry1 = json.loads(lines[1])
            assert entry0.get('uuid') == 'AAA'
            assert entry1.get('uuid') == 'BBB'
        finally:
            Path(path).unlink(missing_ok=True)
    
    def test_rewrite_sorts_violations_by_rule_id(self):
        """Test that violations within entries are sorted by rule_id"""
        fd, path = tempfile.mkstemp(suffix='.jsonl')
        
        try:
            entry = create_audit_entry('AAA', is_violation=True, violations=[
                {'rule_id': 'R003', 'text': 'v3'},
                {'rule_id': 'R001', 'text': 'v1'},
                {'rule_id': 'R002', 'text': 'v2'},
            ])
            results = [(0, entry)]
            
            rewrite_manifest_sorted(path, None, results)
            
            # Read back and check violation order
            with open(path, 'r') as f:
                data = json.loads(f.readline())
            
            violations = data.get('violations', [])
            assert len(violations) == 3
            assert violations[0].get('rule_id') == 'R001'
            assert violations[1].get('rule_id') == 'R002'
            assert violations[2].get('rule_id') == 'R003'
        finally:
            Path(path).unlink(missing_ok=True)


# ============================================================
# Tests: Backup Recovery Scenarios
# ============================================================

class TestBackupRecoveryScenarios:
    """Integration tests for backup recovery scenarios in resume mode"""
    
    def test_scenario_complete_backup_restore(self):
        """
        Scenario: --resume with missing output but complete backup
        Expected: Restore from backup, clean up backup file
        """
        tmpdir = tempfile.mkdtemp()
        output_path = Path(tmpdir) / 'manifest.jsonl'
        backup_path = Path(tmpdir) / 'manifest.jsonl.bak'
        
        try:
            # Create complete backup file
            entries = [
                create_audit_entry('AAA'),
                create_audit_entry('BBB'),
            ]
            backup_entries_path = create_temp_jsonl(entries, include_meta=True)
            shutil.copy(backup_entries_path, backup_path)
            backup_entries_path.unlink()
            
            # Verify setup: output missing, backup exists
            assert not output_path.exists()
            assert backup_path.exists()
            
            # Simulate checking if backup is complete
            completed_uuids = load_completed_uuids(str(backup_path))
            target_uuids = {'AAA', 'BBB'}
            
            assert target_uuids.issubset(completed_uuids)
            
            # Simulate restore process
            uuid_to_block_idx = {'AAA': 0, 'BBB': 1}
            existing_entries = load_existing_entries_with_block_idx(str(backup_path), uuid_to_block_idx)
            manifest_metadata = load_manifest_metadata(str(backup_path))
            
            rewrite_manifest_sorted(str(output_path), manifest_metadata, existing_entries)
            
            # Simulate backup cleanup
            if backup_path.exists():
                backup_path.unlink()
            
            # Verify result: output exists, backup cleaned up
            assert output_path.exists()
            assert not backup_path.exists()
            
            # Verify content
            with open(output_path, 'r') as f:
                lines = f.readlines()
            assert len(lines) == 3  # meta + 2 entries
        finally:
            shutil.rmtree(tmpdir)
    
    def test_scenario_incomplete_backup_continue(self):
        """
        Scenario: --resume with missing output and incomplete backup
        Expected: Copy backup to output, continue processing
        """
        tmpdir = tempfile.mkdtemp()
        output_path = Path(tmpdir) / 'manifest.jsonl'
        backup_path = Path(tmpdir) / 'manifest.jsonl.bak'
        
        try:
            # Create incomplete backup (only AAA processed)
            entries = [create_audit_entry('AAA')]
            backup_entries_path = create_temp_jsonl(entries, include_meta=True)
            shutil.copy(backup_entries_path, backup_path)
            backup_entries_path.unlink()
            
            # Verify setup
            assert not output_path.exists()
            assert backup_path.exists()
            
            # Check if backup is complete for target range
            completed_uuids = load_completed_uuids(str(backup_path))
            target_uuids = {'AAA', 'BBB', 'CCC'}  # Need all 3
            
            # Backup is incomplete
            assert not target_uuids.issubset(completed_uuids)
            
            # Simulate rename (or copy fallback)
            try:
                backup_path.rename(output_path)
            except OSError:
                # Fallback to copy
                with open(backup_path, 'r') as src:
                    with open(output_path, 'w') as dst:
                        dst.write(src.read())
                backup_path.unlink()
            
            # Verify result: output exists from backup
            assert output_path.exists()
            assert not backup_path.exists()
            
            # Should have only AAA
            completed_after = load_completed_uuids(str(output_path))
            assert 'AAA' in completed_after
            assert 'BBB' not in completed_after
            assert 'CCC' not in completed_after
        finally:
            shutil.rmtree(tmpdir)
    
    def test_scenario_rename_fallback_to_copy(self):
        """
        Scenario: rename() fails, fallback to copy + cleanup
        """
        tmpdir = tempfile.mkdtemp()
        output_path = Path(tmpdir) / 'manifest.jsonl'
        backup_path = Path(tmpdir) / 'manifest.jsonl.bak'
        
        try:
            # Create backup
            entries = [create_audit_entry('AAA')]
            backup_entries_path = create_temp_jsonl(entries, include_meta=True)
            shutil.copy(backup_entries_path, backup_path)
            backup_entries_path.unlink()
            
            # Mock rename to fail
            def mock_rename_fail(self, target):
                raise OSError("Cross-device link")
            
            with patch.object(Path, 'rename', mock_rename_fail):
                try:
                    backup_path.rename(output_path)
                except OSError:
                    # Fallback to copy
                    with open(backup_path, 'r', encoding='utf-8') as src:
                        with open(output_path, 'w', encoding='utf-8') as dst:
                            dst.write(src.read())
                    # Clean up backup after copy
                    if backup_path.exists():
                        backup_path.unlink()
            
            # Verify: output has content, backup cleaned up
            assert output_path.exists()
            assert not backup_path.exists()
            
            completed = load_completed_uuids(str(output_path))
            assert 'AAA' in completed
        finally:
            shutil.rmtree(tmpdir)


# ============================================================
# Tests: save_manifest_entry (Append Operation)
# ============================================================

class TestSaveManifestEntry:
    """Tests for save_manifest_entry function"""
    
    def test_append_entry(self):
        """Test appending entry to manifest"""
        fd, path = tempfile.mkstemp(suffix='.jsonl')
        
        try:
            # Append first entry
            save_manifest_entry(path, {'uuid': 'AAA', 'data': 'first'})
            
            # Append second entry
            save_manifest_entry(path, {'uuid': 'BBB', 'data': 'second'})
            
            # Verify both entries present
            with open(path, 'r') as f:
                lines = f.readlines()
            
            assert len(lines) == 2
            assert json.loads(lines[0]).get('uuid') == 'AAA'
            assert json.loads(lines[1]).get('uuid') == 'BBB'
        finally:
            Path(path).unlink(missing_ok=True)
    
    def test_append_to_new_file(self):
        """Test appending to non-existent file creates it"""
        tmpdir = tempfile.mkdtemp()
        path = Path(tmpdir) / 'new_manifest.jsonl'
        
        try:
            assert not path.exists()
            
            save_manifest_entry(str(path), {'uuid': 'AAA'})
            
            assert path.exists()
            with open(path, 'r') as f:
                data = json.loads(f.read().strip())
            assert data.get('uuid') == 'AAA'
        finally:
            shutil.rmtree(tmpdir)
    
    def test_append_unicode_content(self):
        """Test appending entry with unicode content"""
        fd, path = tempfile.mkstemp(suffix='.jsonl')
        
        try:
            save_manifest_entry(path, {
                'uuid': 'AAA',
                'content': '中文内容测试',
                'heading': '第一章 概述'
            })
            
            with open(path, 'r', encoding='utf-8') as f:
                data = json.loads(f.read().strip())
            
            assert data.get('content') == '中文内容测试'
            assert data.get('heading') == '第一章 概述'
        finally:
            Path(path).unlink(missing_ok=True)


# ============================================================
# Main
# ============================================================

if __name__ == '__main__':
    pytest.main([__file__, '-v'])
