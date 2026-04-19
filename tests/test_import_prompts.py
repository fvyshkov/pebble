"""Prompt-content regression tests for the Excel import Claude prompts.

Rationale: the Excel import flow depends on several hard-won instructions
in `SHEET_ANALYSIS_PROMPT` (routers/import_excel.py). When someone rewrites
the prompt they tend to silently drop rules that cost real debugging time.
These tests assert the presence of those rules as literal substrings, so a
regression is caught instantly without invoking Claude.
"""
from __future__ import annotations

from backend.routers.import_excel import SHEET_ANALYSIS_PROMPT


def test_prompt_has_name_pattern_grouping_rule():
    """Rule 12 (added 2026-04): rows like "... в т.ч.:" on sheet 0 are
    group headers even without a formula — LLM must mark them as
    sum_children, NOT manual. User caught this regression on sheet 0
    ("общее количество партнеров, в т.ч.:")."""
    p = SHEET_ANALYSIS_PROMPT
    # The key phrases that encode the rule.
    assert "в т.ч" in p, (
        "SHEET_ANALYSIS_PROMPT must list 'в т.ч.' as a grouping pattern. "
        "Without this rule, the LLM treats rows like «общее количество "
        "партнеров, в т.ч.:» as manual input instead of sum_children."
    )
    assert "в том числе" in p, (
        "SHEET_ANALYSIS_PROMPT must list 'в том числе' as a grouping pattern."
    )
    assert "sum_children" in p, (
        "SHEET_ANALYSIS_PROMPT must explicitly reference rule 'sum_children' "
        "so the LLM returns it for parent rows with no formula."
    )
    # Warning about the absence-of-formula pitfall.
    assert "NOT" in p and "manual" in p, (
        "Prompt must warn that missing formula ≠ manual input for "
        "grouping rows (see rule 12)."
    )


def test_prompt_has_total_header_patterns():
    """Parents starting with 'Итого'/'Всего'/'Общее' are headers."""
    p = SHEET_ANALYSIS_PROMPT
    for needle in ("Итого", "Всего", "Общее"):
        assert needle in p, (
            f"SHEET_ANALYSIS_PROMPT must list '{needle}' as a grouping "
            f"header marker."
        )


def test_prompt_keeps_critical_disambiguation_rule():
    """Rule 9 — disambiguation of duplicate names across groups. If this
    goes away, cross-group formulas become self-references."""
    p = SHEET_ANALYSIS_PROMPT
    assert "disambiguating suffix" in p or "disambiguate" in p.lower()
    assert "(KGS)" in p and "(RUB)" in p


def test_prompt_keeps_cross_sheet_separator():
    """Cross-sheet refs must use `::` separator."""
    p = SHEET_ANALYSIS_PROMPT
    assert "::" in p
