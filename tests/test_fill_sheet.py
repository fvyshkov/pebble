"""Tests for the direct fill-sheet endpoint and related chat flow.

`POST /api/chat/fill_sheet/{sheet_id}` fills every cartesian-leaf cell on
a sheet (skipping formula/sum cells). Runs without hitting Anthropic, so
it's fast and reliable.
"""
import pytest
import requests

BASE = "http://localhost:8000/api"


def _first_sheet():
    models = requests.get(f"{BASE}/models", timeout=5).json()
    if not models:
        pytest.skip("No models in DB")
    sheets = requests.get(f"{BASE}/sheets/by-model/{models[0]['id']}", timeout=5).json()
    if not sheets:
        pytest.skip("No sheets in DB")
    # Pick any sheet with >= 2 analytics (needed for real cartesian)
    for s in sheets:
        sa = requests.get(f"{BASE}/sheets/{s['id']}/analytics", timeout=5).json()
        if len(sa) >= 2:
            return s["id"]
    pytest.skip("No sheet with >= 2 analytics")


def test_fill_sheet_with_constant():
    sheet_id = _first_sheet()
    resp = requests.post(
        f"{BASE}/chat/fill_sheet/{sheet_id}",
        json={"mode": "value", "value": "42"},
        timeout=30,
    )
    assert resp.status_code == 200, resp.text
    body = resp.json()
    assert body["ok"] is True
    assert body["cells_written"] > 0
    # Verify at least one cell actually has value 42
    cells = requests.get(f"{BASE}/cells/by-sheet/{sheet_id}", timeout=10).json()
    assert any(str(c.get("value")) == "42" for c in cells), \
        "No cell with value 42 after fill_sheet(value=42)"


def test_fill_sheet_random():
    sheet_id = _first_sheet()
    resp = requests.post(
        f"{BASE}/chat/fill_sheet/{sheet_id}",
        json={"mode": "random", "min": 10, "max": 20},
        timeout=30,
    )
    assert resp.status_code == 200, resp.text
    body = resp.json()
    assert body["ok"] is True
    assert body["cells_written"] > 0
    # Every newly-written manual cell should fall in [10, 20]
    cells = requests.get(f"{BASE}/cells/by-sheet/{sheet_id}", timeout=10).json()
    manual_values = [int(c["value"]) for c in cells
                     if c.get("rule") == "manual" and c.get("value", "").lstrip("-").isdigit()]
    assert manual_values, "Expected at least one manual-rule cell with numeric value"
    in_range = [v for v in manual_values if 10 <= v <= 20]
    # Allow a few pre-existing out-of-range values; but the majority should be in
    # the just-written range.
    assert len(in_range) >= len(manual_values) * 0.8, (
        f"Expected most manual cells in [10,20]; got {manual_values[:10]}..."
    )


def test_fill_sheet_skips_formulas():
    """A cell with rule != 'manual' must not be overwritten by fill_sheet."""
    sheet_id = _first_sheet()
    # Pick an existing cell and force its rule to 'formula'.
    cells = requests.get(f"{BASE}/cells/by-sheet/{sheet_id}", timeout=10).json()
    if not cells:
        # Bootstrap: fill once so we have cells to protect
        requests.post(f"{BASE}/chat/fill_sheet/{sheet_id}",
                      json={"mode": "value", "value": "1"}, timeout=30)
        cells = requests.get(f"{BASE}/cells/by-sheet/{sheet_id}", timeout=10).json()
    target = cells[0]
    # Promote to formula with a sentinel value
    r = requests.put(
        f"{BASE}/cells/by-sheet/{sheet_id}?no_recalc=true",
        json={"cells": [{
            "coord_key": target["coord_key"],
            "value": "999",
            "rule": "formula",
            "formula": "=1+1",
        }]},
        timeout=10,
    )
    assert r.status_code == 200, r.text
    # Now try to fill with value=7 and verify the formula cell is untouched.
    requests.post(f"{BASE}/chat/fill_sheet/{sheet_id}",
                  json={"mode": "value", "value": "7"}, timeout=30)
    after = requests.get(f"{BASE}/cells/by-sheet/{sheet_id}", timeout=10).json()
    match = [c for c in after if c["coord_key"] == target["coord_key"]]
    assert match, "formula cell disappeared"
    assert match[0]["rule"] == "formula", "formula rule was overwritten"
    assert match[0]["value"] == "999", (
        f"formula cell value was overwritten to {match[0]['value']} (expected 999)"
    )
