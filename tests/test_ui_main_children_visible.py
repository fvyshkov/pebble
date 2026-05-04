"""Visual smoke test for MAIN.xlsx import: after import, manual values in
CHILD rows under expanded parents must be visible in the AG Grid.

This test fails if the lazy-load logic in PivotGridAG.tsx forgets to fetch
cells for children rows — exactly the bug that compare_excel_exact.py
cannot detect (because that script reads the DB, not the rendered UI).

Pre-requisites:
  - backend on :8000 with a MAIN model already imported
  - that model contains sheet "BaaS - параметры модели" with parent row
    "общее количество партнеров, в т.ч.:" having children r10..r24

Run:  pytest tests/test_ui_main_children_visible.py --headed
"""
from __future__ import annotations

import os
import re
import sqlite3
import sys
from pathlib import Path

import pytest
import requests
from playwright.sync_api import Page, expect

BASE = os.environ.get("PEBBLE_BASE", "http://localhost:8000")
DB_PATH = Path(__file__).resolve().parent.parent / "pebble.db"


def _find_main_model_id() -> str:
    """Locate the most recent model that has the MAIN-shape sheet
    'BaaS - параметры модели'. Streaming import names models 'Imported Model'
    by default, so a name LIKE 'MAIN%' filter is too narrow."""
    db = sqlite3.connect(str(DB_PATH))
    rows = db.execute(
        """SELECT m.id FROM models m
           JOIN sheets s ON s.model_id = m.id
           WHERE s.name = 'BaaS - параметры модели'
           ORDER BY m.created_at DESC LIMIT 1"""
    ).fetchall()
    db.close()
    if not rows:
        pytest.skip("No model with MAIN sheets imported — run "
                    "tests/test_main_full_accuracy.py first")
    return rows[0][0]


def _read_db_value(model_id: str, sheet_name: str, indicator_excel_row: int,
                   period_name: str) -> str:
    """Look up the cell value for (sheet, indicator_row, period_name) directly
    from cell_data so we know what the UI ought to show."""
    db = sqlite3.connect(str(DB_PATH))
    cur = db.cursor()
    sheet_id = cur.execute(
        "SELECT id FROM sheets WHERE model_id=? AND name=?",
        (model_id, sheet_name),
    ).fetchone()
    assert sheet_id, f"Sheet {sheet_name!r} not found"
    sheet_id = sheet_id[0]
    ind_seq = cur.execute(
        """SELECT ar.seq_id FROM analytic_records ar
           JOIN sheet_analytics sa ON sa.analytic_id=ar.analytic_id
           JOIN analytics a ON a.id=ar.analytic_id
           WHERE sa.sheet_id=? AND a.name LIKE 'Показатели%' AND ar.excel_row=?""",
        (sheet_id, indicator_excel_row),
    ).fetchone()
    assert ind_seq, f"Indicator row {indicator_excel_row} not found on {sheet_name}"
    per_seq = cur.execute(
        """SELECT ar.seq_id FROM analytic_records ar
           JOIN sheet_analytics sa ON sa.analytic_id=ar.analytic_id
           JOIN analytics a ON a.id=ar.analytic_id
           WHERE sa.sheet_id=? AND a.name LIKE 'Периоды%'
             AND json_extract(ar.data_json,'$.name')=?""",
        (sheet_id, period_name),
    ).fetchone()
    assert per_seq, f"Period {period_name!r} not found"
    val = cur.execute(
        "SELECT value FROM cell_data WHERE sheet_id=? AND coord_key=?",
        (sheet_id, f"{per_seq[0]}|{ind_seq[0]}"),
    ).fetchone()
    db.close()
    assert val and val[0] not in (None, ""), \
        f"DB has no value for r{indicator_excel_row} × {period_name} on {sheet_name}"
    return val[0]


def _normalize(s: str) -> str:
    """'15.0' -> '15'; trim whitespace; drop NBSP thousand separators."""
    s = s.replace(" ", "").replace(" ", "").strip()
    if re.fullmatch(r"-?\d+\.0+", s):
        s = s.split(".")[0]
    return s


@pytest.fixture(scope="module")
def page(browser):
    ctx = browser.new_context(viewport={"width": 1600, "height": 900})
    p = ctx.new_page()
    p.goto(BASE)
    p.wait_for_load_state("networkidle")
    # If login form is showing, log in.
    login = p.locator('input[name="username"]')
    if login.is_visible():
        login.fill("admin")
        p.locator('input[name="password"]').fill("admin")
        p.locator('button:has-text("Войти")').click()
        p.wait_for_load_state("networkidle")
        p.wait_for_timeout(800)
    yield p
    ctx.close()


def test_child_rows_show_values_after_expansion(page: Page):
    """The bug reported by the user (2026-05-03): after MAIN import, when
    you expand "общее количество партнеров, в т.ч.:" on the
    "BaaS - параметры модели" sheet, the child rows
    (кредитование свои/чужие/депозиты/...) are blank in the UI even though
    the DB has values for them. This test asserts the DB values appear
    rendered after expansion."""
    mid = _find_main_model_id()
    sheet_name = "BaaS - параметры модели"
    expected_parent_jan = _read_db_value(mid, sheet_name, 9, "Январь 2026")
    expected_child_jan = _read_db_value(mid, sheet_name, 10, "Январь 2026")
    expected_child_feb = _read_db_value(mid, sheet_name, 10, "Февраль 2026")
    print(f"\nDB says: parent r9 Jan={expected_parent_jan}, "
          f"child r10 Jan={expected_child_jan} Feb={expected_child_feb}")

    # 1) Open the BaaS - параметры модели sheet from the left tree.
    page.wait_for_selector(".tree-item-label", timeout=10000)
    sheet_label = page.locator(".tree-item-label", has_text="параметры модели").first
    sheet_label.click()
    page.wait_for_selector(".ag-root-wrapper", timeout=15000)
    page.wait_for_timeout(1500)
    page.wait_for_function("() => !!window.__pebbleGridApi", timeout=10000)

    def _find_row(needle: str) -> dict | None:
        return page.evaluate(
            """(needle) => {
              const api = window.__pebbleGridApi;
              if (!api) return null;
              let hit = null;
              api.forEachNode(n => {
                if (hit || !n.data) return;
                const lbl = String(n.data.label || '');
                if (lbl.includes(needle)) {
                  hit = { expanded: !!n.expanded, label: lbl, data: { ...n.data } };
                }
              });
              return hit;
            }""",
            needle,
        )

    def _expand(needle: str):
        page.evaluate(
            """(needle) => {
              const api = window.__pebbleGridApi;
              api.forEachNode(n => {
                if (!n.data) return;
                if (String(n.data.label || '').includes(needle) && !n.expanded) {
                  n.setExpanded(true);
                }
              });
            }""",
            needle,
        )

    # 2) Sanity: parent row carries DB value in some p_<jan-leaf> field.
    parent = _find_row("общее количество партнеров")
    assert parent, "Parent row 'общее количество партнеров' not found in grid"
    parent_data = parent["data"]
    p_keys = [k for k in parent_data.keys() if k.startswith("p_")]
    print(f"Parent has {len(p_keys)} p_ fields")
    # Find the p_ field whose value matches the DB Jan value (15) AND another
    # whose value matches Feb (16) — these point to Jan and Feb columns.
    jan_field = next((k for k in p_keys
                      if _normalize(str(parent_data.get(k) or "")) == _normalize(expected_parent_jan)),
                     None)
    feb_field = next((k for k in p_keys
                      if _normalize(str(parent_data.get(k) or "")) == _normalize(expected_child_feb)),
                     None)
    assert jan_field, (
        f"Parent has no p_ field with value {expected_parent_jan}. "
        f"Sample p_ values: { {k: parent_data[k] for k in p_keys[:8]} }"
    )
    print(f"Jan field={jan_field}  Feb field={feb_field}")

    # 3) Expand parent and wait for lazy-load.
    _expand("общее количество партнеров")
    page.wait_for_timeout(2500)

    # 4) Read child row data — the bug: child p_<jan_field> stays blank
    #    because lazy-load is skipped when initial populate set p_=''.
    child = _find_row("свои кредиты")
    page.screenshot(path="/tmp/main_children_after_expand.png", full_page=False)
    assert child, "Child row 'свои кредиты' not found in grid (parent not expanded?)"
    child_jan = str(child["data"].get(jan_field) or "").strip()
    print(f"Child r10 Jan UI={child_jan!r}  DB={expected_child_jan}")
    assert child_jan, (
        f"Child 'свои кредиты' has BLANK value for Jan column ({jan_field}). "
        f"DB has {expected_child_jan}. Lazy-load bug — children cells not "
        f"fetched after expansion. Screenshot: /tmp/main_children_after_expand.png"
    )
    assert _normalize(child_jan) == _normalize(expected_child_jan), (
        f"Child Jan UI={child_jan!r} != DB={expected_child_jan!r}"
    )

    if feb_field:
        child_feb = str(child["data"].get(feb_field) or "").strip()
        assert child_feb and _normalize(child_feb) == _normalize(expected_child_feb), (
            f"Child Feb UI={child_feb!r} != DB={expected_child_feb!r}"
        )
