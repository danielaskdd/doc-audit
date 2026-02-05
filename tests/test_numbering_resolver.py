"""
ABOUTME: Tests for isLgl (legal numbering) support in NumberingResolver
ABOUTME: Verifies parent-level decimal coercion when isLgl is set
"""

import sys
import os

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'skills', 'doc-audit', 'scripts'))
from numbering_resolver import NumberingResolver  # noqa: E402  # type: ignore


def _make_resolver_with_state(levels: dict, counters: dict, num_id: str = '1'):
    """
    Create a NumberingResolver with manually injected state for _format_label testing.

    Args:
        levels: {ilvl: {start, numFmt, lvlText, isLgl}} level definitions
        counters: {ilvl: current_count} counter values
        num_id: numbering ID string
    """
    # Create instance without parsing any file
    resolver = object.__new__(NumberingResolver)
    resolver.abstract_nums = {}
    resolver.num_to_abstract = {}
    resolver.counters = {num_id: counters}
    resolver.start_overrides = {}
    resolver.style_numpr = {}
    resolver.style_numpr_overrides = {}
    resolver.style_based_on = {}
    resolver.last_numId = None
    resolver.last_abstract_id = None
    resolver.last_style_id = None
    return resolver


class TestIsLglBasic:
    """isLgl=True on level 1 coerces level 0's chineseCountingThousand to decimal."""

    def test_isLgl_basic(self):
        levels = {
            0: {'start': 1, 'numFmt': 'chineseCountingThousand', 'lvlText': '%1', 'isLgl': False},
            1: {'start': 1, 'numFmt': 'decimal', 'lvlText': '%1.%2', 'isLgl': True},
        }
        counters = {0: 5, 1: 1}
        resolver = _make_resolver_with_state(levels, counters)
        result = resolver._format_label('1', 1, levels)
        assert result == '5.1'


class TestIsLglDisabled:
    """isLgl val='0' means no override — parent keeps its native format."""

    def test_isLgl_disabled(self):
        levels = {
            0: {'start': 1, 'numFmt': 'chineseCountingThousand', 'lvlText': '%1', 'isLgl': False},
            1: {'start': 1, 'numFmt': 'decimal', 'lvlText': '%1.%2', 'isLgl': False},
        }
        counters = {0: 5, 1: 1}
        resolver = _make_resolver_with_state(levels, counters)
        result = resolver._format_label('1', 1, levels)
        assert result == '五.1'


class TestIsLglAbsent:
    """No isLgl element (key missing) — backward compatible, no override."""

    def test_isLgl_absent(self):
        levels = {
            0: {'start': 1, 'numFmt': 'chineseCountingThousand', 'lvlText': '%1'},
            1: {'start': 1, 'numFmt': 'decimal', 'lvlText': '%1.%2'},
        }
        counters = {0: 5, 1: 1}
        resolver = _make_resolver_with_state(levels, counters)
        result = resolver._format_label('1', 1, levels)
        assert result == '五.1'


class TestIsLglLevel0NoEffect:
    """isLgl on level 0 has no parent levels to override — renders normally."""

    def test_isLgl_level0_no_effect(self):
        levels = {
            0: {'start': 1, 'numFmt': 'chineseCountingThousand', 'lvlText': '第%1条', 'isLgl': True},
        }
        counters = {0: 3}
        resolver = _make_resolver_with_state(levels, counters)
        result = resolver._format_label('1', 0, levels)
        assert result == '第三条'


class TestIsLglMultiParent:
    """isLgl on level 2 coerces both parent levels (0 and 1) to decimal."""

    def test_isLgl_multi_parent(self):
        levels = {
            0: {'start': 1, 'numFmt': 'chineseCountingThousand', 'lvlText': '%1', 'isLgl': False},
            1: {'start': 1, 'numFmt': 'upperRoman', 'lvlText': '%1.%2', 'isLgl': False},
            2: {'start': 1, 'numFmt': 'decimal', 'lvlText': '%1.%2.%3', 'isLgl': True},
        }
        counters = {0: 3, 1: 2, 2: 1}
        resolver = _make_resolver_with_state(levels, counters)
        result = resolver._format_label('1', 2, levels)
        # Level 0 chinese→decimal "3", level 1 upperRoman→decimal "2", level 2 stays decimal "1"
        assert result == '3.2.1'


class TestIsLglCurrentLevelPreserved:
    """isLgl only overrides parent levels — current level keeps its own numFmt."""

    def test_isLgl_current_level_preserved(self):
        levels = {
            0: {'start': 1, 'numFmt': 'chineseCountingThousand', 'lvlText': '%1', 'isLgl': False},
            1: {'start': 1, 'numFmt': 'lowerLetter', 'lvlText': '%1.%2', 'isLgl': True},
        }
        counters = {0: 2, 1: 3}
        resolver = _make_resolver_with_state(levels, counters)
        result = resolver._format_label('1', 1, levels)
        # Level 0 coerced to decimal "2", level 1 keeps lowerLetter "c"
        assert result == '2.c'
