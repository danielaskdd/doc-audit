#!/usr/bin/env python3
"""
ABOUTME: Unit tests for fixlevel-aware flush behavior in parse_document.py
ABOUTME: Covers _flush_current_block, _build_unsplit_block, and oversized warnings
"""

import sys
from pathlib import Path

_scripts_dir = Path(__file__).parent.parent / 'skills' / 'doc-audit' / 'scripts'
sys.path.insert(0, str(_scripts_dir))

import pytest  # noqa: E402

from parse_document import (  # noqa: E402  # type: ignore[import-not-found]
    _build_unsplit_block,
    _flush_current_block,
    estimate_tokens,
    MAX_BLOCK_CONTENT_TOKENS,
)


def _make_para(text: str, para_id: str, is_table: bool = False) -> dict:
    return {'text': text, 'para_id': para_id, 'is_table': is_table}


# ============================================================
# _build_unsplit_block
# ============================================================

def test_build_unsplit_block_concatenates_text_and_picks_uuids():
    paras = [
        _make_para("first paragraph", "AAAAAAAA"),
        _make_para("second paragraph", "BBBBBBBB"),
        _make_para("third paragraph", "CCCCCCCC"),
    ]
    block = _build_unsplit_block("Heading", paras, ["Parent"], 2)

    assert block['uuid'] == "AAAAAAAA"
    assert block['uuid_end'] == "CCCCCCCC"
    assert block['heading'] == "Heading"
    assert block['parent_headings'] == ["Parent"]
    assert block['level'] == 2
    assert block['type'] == "text"
    assert block['content'] == "first paragraph\nsecond paragraph\nthird paragraph"


def test_build_unsplit_block_uses_para_id_end_when_present():
    """Tables carry para_id_end (last cell paraId); should be used as uuid_end."""
    paras = [
        _make_para("text", "AAAAAAAA"),
        {
            'text': "<table>[[\"x\"]]</table>",
            'para_id': "BBBBBBBB",
            'para_id_end': "CCCCCCCC",
            'is_table': True,
        },
    ]
    block = _build_unsplit_block("H", paras, [], 1)
    assert block['uuid_end'] == "CCCCCCCC"


# ============================================================
# _flush_current_block — fixlevel mode
# ============================================================

def test_flush_in_fixlevel_mode_appends_single_block():
    blocks = []
    paras = [_make_para("short content", "AAAAAAAA")]
    _flush_current_block(blocks, "H1", paras, [], 1, fixlevel=1, debug=False)

    assert len(blocks) == 1
    assert blocks[0]['uuid'] == "AAAAAAAA"
    assert blocks[0]['content'] == "short content"
    # fixlevel path doesn't add table_chunk_role; main loop adds it later
    assert 'table_chunk_role' not in blocks[0]


def test_flush_with_empty_paragraphs_is_noop():
    blocks = []
    _flush_current_block(blocks, "H", [], [], 1, fixlevel=1, debug=False)
    _flush_current_block(blocks, "H", [], [], 1, fixlevel=None, debug=False)
    assert blocks == []


def test_flush_in_fixlevel_warns_on_oversized_block(capsys):
    # Build a paragraph whose token estimate exceeds MAX_BLOCK_CONTENT_TOKENS.
    # estimate_tokens returns ~ len/4 for ASCII; we need > 8000 tokens.
    huge_text = "x " * (MAX_BLOCK_CONTENT_TOKENS * 5)  # ~5x safety margin
    assert estimate_tokens(huge_text) > MAX_BLOCK_CONTENT_TOKENS

    blocks = []
    paras = [_make_para(huge_text, "AAAAAAAA")]
    _flush_current_block(blocks, "Big chapter", paras, [], 1, fixlevel=1, debug=False)

    captured = capsys.readouterr()
    assert len(blocks) == 1  # still emitted as one block
    assert "fixlevel block exceeds" in captured.err
    assert "Big chapter" in captured.err


def test_flush_in_fixlevel_no_warning_when_within_limit(capsys):
    blocks = []
    paras = [_make_para("a small paragraph", "AAAAAAAA")]
    _flush_current_block(blocks, "Small chapter", paras, [], 1, fixlevel=1, debug=False)

    captured = capsys.readouterr()
    assert captured.err == ""
    assert len(blocks) == 1


def test_flush_truncates_long_heading_in_warning_preview(capsys):
    huge_text = "x " * (MAX_BLOCK_CONTENT_TOKENS * 5)
    long_heading = "H" * 200
    blocks = []
    _flush_current_block(
        blocks, long_heading, [_make_para(huge_text, "AAAAAAAA")],
        [], 1, fixlevel=1, debug=False,
    )
    captured = capsys.readouterr()
    # Heading preview should be truncated to 80 chars + "..."
    assert "H" * 80 + "..." in captured.err


# ============================================================
# _flush_current_block — default mode (regression: must invoke split_long_block)
# ============================================================

def test_flush_in_default_mode_uses_split_long_block(monkeypatch):
    """In default mode, flush must delegate to split_long_block."""
    import parse_document  # type: ignore[import-not-found]

    captured_calls = []

    def fake_split(heading, paragraphs, parent_headings, level, debug):
        captured_calls.append((heading, len(paragraphs), level))
        return [{
            'uuid': paragraphs[0]['para_id'],
            'uuid_end': paragraphs[-1]['para_id'],
            'heading': heading,
            'content': '\n'.join(p['text'] for p in paragraphs),
            'type': 'text',
            'parent_headings': parent_headings,
            'level': level,
        }]

    monkeypatch.setattr(parse_document, 'split_long_block', fake_split)

    blocks = []
    paras = [_make_para("p1", "AAAAAAAA"), _make_para("p2", "BBBBBBBB")]
    _flush_current_block(blocks, "H", paras, [], 2, fixlevel=None, debug=False)

    assert captured_calls == [("H", 2, 2)]
    assert len(blocks) == 1


def test_flush_in_default_mode_does_not_emit_size_warning(capsys, monkeypatch):
    """Default mode delegates to split_long_block; the fixlevel-specific warning must not fire."""
    import parse_document  # type: ignore[import-not-found]

    monkeypatch.setattr(
        parse_document, 'split_long_block',
        lambda heading, paragraphs, parent_headings, level, debug: [
            _build_unsplit_block(heading, paragraphs, parent_headings, level)
        ],
    )

    huge_text = "x " * (MAX_BLOCK_CONTENT_TOKENS * 5)
    blocks = []
    _flush_current_block(
        blocks, "H", [_make_para(huge_text, "AAAAAAAA")], [], 1,
        fixlevel=None, debug=False,
    )
    captured = capsys.readouterr()
    assert "fixlevel block exceeds" not in captured.err


# ============================================================
# Sanity check warning (single-block fixlevel result)
# ============================================================

def test_extract_audit_blocks_warns_when_fixlevel_misses_all_headings(capsys, tmp_path, monkeypatch):
    """
    When --fixlevel=N>0 but no heading at level <= N exists, the document
    collapses into a single block. extract_audit_blocks should warn.
    """
    import parse_document  # type: ignore[import-not-found]

    # Stub Document, NumberingResolver, parse_styles_outline_levels to avoid real docx parsing.
    class FakeBody:
        def __iter__(self):
            return iter([])

    class FakeElement:
        body = FakeBody()

    class FakeDoc:
        _element = FakeElement()

    class FakeResolver:
        def __init__(self, *a, **kw): pass
        def reset_tracking_state(self): pass
        def get_label(self, p): return ""

    monkeypatch.setattr(parse_document, 'Document', lambda path: FakeDoc())
    monkeypatch.setattr(parse_document, 'NumberingResolver', FakeResolver)
    monkeypatch.setattr(parse_document, 'parse_styles_outline_levels', lambda path: {})

    # Empty document with fixlevel=1 -> 0 blocks (no body content). Warning still fires (<=1).
    blocks = parse_document.extract_audit_blocks("dummy.docx", fixlevel=1)
    captured = capsys.readouterr()
    assert blocks == []
    assert "--fixlevel=1 produced 0 block(s)" in captured.err
    assert "Try a higher --fixlevel value" in captured.err


def test_extract_audit_blocks_no_sanity_warning_for_fixlevel_zero(capsys, monkeypatch):
    """fixlevel=0 means 'split at all heading levels' — zero/one block result is not a misconfiguration."""
    import parse_document  # type: ignore[import-not-found]

    class FakeBody:
        def __iter__(self): return iter([])

    class FakeElement:
        body = FakeBody()

    class FakeDoc:
        _element = FakeElement()

    class FakeResolver:
        def __init__(self, *a, **kw): pass
        def reset_tracking_state(self): pass
        def get_label(self, p): return ""

    monkeypatch.setattr(parse_document, 'Document', lambda path: FakeDoc())
    monkeypatch.setattr(parse_document, 'NumberingResolver', FakeResolver)
    monkeypatch.setattr(parse_document, 'parse_styles_outline_levels', lambda path: {})

    blocks = parse_document.extract_audit_blocks("dummy.docx", fixlevel=0)
    captured = capsys.readouterr()
    assert blocks == []
    assert "produced" not in captured.err  # no sanity warning


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
