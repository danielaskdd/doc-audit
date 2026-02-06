#!/usr/bin/env python3
"""
Test that run_audit.py correctly exits with error code when blocks fail.

This test verifies the fix for the issue where run_audit.py would exit with
code 0 even when some blocks failed during audit, causing agents to incorrectly
assume the task completed successfully.
"""

import json
import subprocess
import sys
import tempfile
from pathlib import Path

# Add script directory to path
_SCRIPT_DIR = Path(__file__).resolve().parent.parent / "skills" / "doc-audit" / "scripts"
sys.path.insert(0, str(_SCRIPT_DIR))


def test_exit_code_on_block_failure():
    """
    Test that run_audit.py exits with code 1 when blocks fail.
    
    This simulates a scenario where LLM API calls fail and verifies that
    the script properly reports the failure via exit code.
    """
    # Create temporary files
    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir = Path(tmpdir)
        
        # Create a minimal blocks file
        blocks_file = tmpdir / "test_blocks.jsonl"
        with open(blocks_file, 'w', encoding='utf-8') as f:
            # Metadata
            f.write(json.dumps({
                "type": "meta",
                "source_file": "test.docx",
                "source_hash": "abc123"
            }, ensure_ascii=False) + '\n')
            # Single test block
            f.write(json.dumps({
                "uuid": "12345678",
                "uuid_end": "12345678",
                "heading": "Test Heading",
                "content": "Test content",
                "type": "text",
                "parent_headings": []
            }, ensure_ascii=False) + '\n')
        
        # Create a minimal rules file
        rules_file = tmpdir / "test_rules.json"
        with open(rules_file, 'w', encoding='utf-8') as f:
            json.dump({
                "rules": [
                    {
                        "id": "R001",
                        "description": "Test rule",
                        "severity": "high",
                        "category": "test",
                        "type": "block"
                    }
                ]
            }, f, ensure_ascii=False)
        
        manifest_file = tmpdir / "manifest.jsonl"
        
        # Test Case 1: Simulate failure by using invalid API key (should exit with code 1)
        print("Test Case 1: Invalid API key (expecting exit code 1)")
        env = {"GOOGLE_API_KEY": "invalid_key_that_will_fail"}
        
        result = subprocess.run(
            [
                sys.executable,
                str(_SCRIPT_DIR / "run_audit.py"),
                "--document", str(blocks_file),
                "--rules", str(rules_file),
                "--output", str(manifest_file),
                "--max-retries", "0"  # Don't retry to fail fast
            ],
            env=env,
            capture_output=True,
            text=True
        )
        
        print(f"  Exit code: {result.returncode}")
        print(f"  Stderr (last 500 chars): {result.stderr[-500:]}")
        
        # The script should exit with code 1 when API calls fail
        # Note: This test may not work perfectly if both GOOGLE_API_KEY and OPENAI_API_KEY
        # are missing from the environment, as the script will exit early with a different error.
        # For now, we're just documenting the expected behavior.
        
        if result.returncode == 0:
            print("  ❌ FAIL: Script exited with code 0 despite failures")
            return False
        else:
            print("  ✅ PASS: Script exited with non-zero code as expected")
        
        print("\nTest Case 2: Dry-run mode (expecting exit code 0)")
        result = subprocess.run(
            [
                sys.executable,
                str(_SCRIPT_DIR / "run_audit.py"),
                "--document", str(blocks_file),
                "--rules", str(rules_file),
                "--output", str(manifest_file),
                "--dry-run"
            ],
            capture_output=True,
            text=True
        )
        
        print(f"  Exit code: {result.returncode}")
        
        if result.returncode != 0:
            print("  ❌ FAIL: Dry-run should exit with code 0")
            return False
        else:
            print("  ✅ PASS: Dry-run exited with code 0")
        
        return True


if __name__ == "__main__":
    print("Testing run_audit.py exit code behavior")
    print("=" * 60)
    
    success = test_exit_code_on_block_failure()
    
    print("\n" + "=" * 60)
    if success:
        print("✅ All tests passed!")
        sys.exit(0)
    else:
        print("❌ Some tests failed")
        sys.exit(1)
