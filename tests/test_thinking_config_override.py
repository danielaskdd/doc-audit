#!/usr/bin/env python3
"""
Test that CLI thinking/reasoning parameters override environment variables correctly.

This test verifies the fix for the issue where CLI args should override env vars
instead of causing validation errors.
"""

import os
import subprocess
import sys
from pathlib import Path

# Get the path to run_audit.py
SCRIPT_DIR = Path(__file__).parent.parent / "skills" / "doc-audit" / "scripts"
RUN_AUDIT = SCRIPT_DIR / "run_audit.py"

def run_with_env(env_vars: dict, cli_args: list) -> tuple:
    """
    Run run_audit.py with specified environment variables and CLI args.
    
    Returns:
        (return_code, stdout, stderr)
    """
    # Merge with current environment
    env = os.environ.copy()
    env.update(env_vars)
    
    # Build command
    cmd = [sys.executable, str(RUN_AUDIT), "--dry-run"] + cli_args
    
    # Run command
    result = subprocess.run(
        cmd,
        env=env,
        capture_output=True,
        text=True
    )
    
    return result.returncode, result.stdout, result.stderr


def test_cli_thinking_level_overrides_env_budget():
    """Test that --thinking-level CLI arg overrides GEMINI_THINKING_BUDGET env var."""
    env_vars = {
        "GEMINI_THINKING_BUDGET": "1024",  # Env var for Gemini 2.5
        "GOOGLE_API_KEY": "fake_key_for_test"
    }
    
    cli_args = [
        "--document", "dummy.jsonl",
        "--rules", "dummy.json",
        "--thinking-level", "medium",  # CLI arg for Gemini 3
        "--provider", "gemini"
    ]
    
    returncode, stdout, stderr = run_with_env(env_vars, cli_args)
    
    # Should NOT error - CLI should override env
    assert "Error: Both thinking_level and thinking_budget are set" not in stderr, \
        "Got validation error (should override, not error)"


def test_cli_thinking_budget_overrides_env_level():
    """Test that --thinking-budget CLI arg overrides GEMINI_THINKING_LEVEL env var."""
    env_vars = {
        "GEMINI_THINKING_LEVEL": "high",  # Env var for Gemini 3
        "GOOGLE_API_KEY": "fake_key_for_test"
    }
    
    cli_args = [
        "--document", "dummy.jsonl",
        "--rules", "dummy.json",
        "--thinking-budget", "2048",  # CLI arg for Gemini 2.5
        "--provider", "gemini"
    ]
    
    returncode, stdout, stderr = run_with_env(env_vars, cli_args)
    
    # Should NOT error - CLI should override env
    assert "Error: Both thinking_level and thinking_budget are set" not in stderr, \
        "Got validation error (should override, not error)"


def test_both_env_vars_still_errors():
    """Test that setting both env vars without CLI still errors (expected behavior)."""
    env_vars = {
        "GEMINI_THINKING_LEVEL": "high",
        "GEMINI_THINKING_BUDGET": "1024",
        "GOOGLE_API_KEY": "fake_key_for_test"
    }
    
    cli_args = [
        "--document", "dummy.jsonl",
        "--rules", "dummy.json",
        "--provider", "gemini"
    ]
    
    returncode, stdout, stderr = run_with_env(env_vars, cli_args)
    
    # SHOULD error when both env vars are set
    assert "Error: Both thinking_level and thinking_budget are set" in stderr, \
        "Should have errored on conflicting env vars"


def main():
    """Run all tests (standalone mode)."""
    print("=" * 60)
    print("Testing CLI override behavior for thinking configuration")
    print("=" * 60)
    
    results = []
    
    # Test 1
    print("\n1. Testing: CLI --thinking-level overrides env GEMINI_THINKING_BUDGET")
    try:
        test_cli_thinking_level_overrides_env_budget()
        print("  ✅ PASSED: No validation error (CLI override worked)")
        results.append(True)
    except AssertionError as e:
        print(f"  ❌ FAILED: {e}")
        results.append(False)
    
    # Test 2
    print("\n2. Testing: CLI --thinking-budget overrides env GEMINI_THINKING_LEVEL")
    try:
        test_cli_thinking_budget_overrides_env_level()
        print("  ✅ PASSED: No validation error (CLI override worked)")
        results.append(True)
    except AssertionError as e:
        print(f"  ❌ FAILED: {e}")
        results.append(False)
    
    # Test 3
    print("\n3. Testing: Both env vars set without CLI still errors (expected)")
    try:
        test_both_env_vars_still_errors()
        print("  ✅ PASSED: Got expected validation error")
        results.append(True)
    except AssertionError as e:
        print(f"  ❌ FAILED: {e}")
        results.append(False)
    
    # Summary
    print("\n" + "=" * 60)
    passed = sum(results)
    total = len(results)
    print(f"Results: {passed}/{total} tests passed")
    print("=" * 60)
    
    if passed == total:
        print("\n✅ All tests passed!")
        return 0
    else:
        print(f"\n❌ {total - passed} test(s) failed")
        return 1


if __name__ == "__main__":
    sys.exit(main())
