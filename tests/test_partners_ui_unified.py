"""Playwright UI test: «общее количество партнеров, в т.ч.» indicator must
display per-cell formulas consistently across the three surfaces:

  1. Indicator settings panel (Аналитики → Показатели → row click → right pane)
  2. Cell tooltip in the pivot grid
  3. Formula mode of the pivot grid

The bug we're guarding against: panel showed ONE leaf formula + "SUM"
consolidation; tooltip showed "ƒ SUM" on every cell; formula-mode grid
showed "SUM" everywhere — even though Excel cells carry distinct per-cell
formulas (and Pebble's cell_data stores them correctly).

After fix: panel shows "разные формулы у клеток"; tooltip and grid show the
actual per-cell formula text for formula cells, and the manual-input label
for the literal cells.

Requires backend on :8000 with the latest MAIN.xlsx imported.
"""
import pytest
from playwright.sync_api import Page, expect

BASE = "http://localhost:8000"
ADMIN_USER = "admin"
ADMIN_PASS = "admin"


@pytest.fixture(scope="module")
def page(browser):
    ctx = browser.new_context()
    p = ctx.new_page()
    p.goto(BASE)
    p.wait_for_load_state("networkidle")
    login = p.locator('input[name="username"]')
    if login.is_visible():
        login.fill(ADMIN_USER)
        p.locator('input[name="password"]').fill(ADMIN_PASS)
        p.locator('button:has-text("Войти")').click()
        p.wait_for_load_state("networkidle")
        p.wait_for_timeout(800)
    yield p
    ctx.close()


def _switch_to_settings_mode(page: Page) -> None:
    """Click the gear (settings) toggle so the left tree exposes Аналитики."""
    page.locator('button[value="settings"]').first.click()
    page.wait_for_timeout(500)


def _open_partners_panel(page: Page) -> None:
    """Idempotently navigate to Аналитики → Показатели (0).

    Settings mode is required because data/formulas modes hide the analytics
    folder (sheetsOnly=true). Clicks toggle expand state, so we only click
    when the target node isn't already visible/active.
    """
    _switch_to_settings_mode(page)
    # If "Показатели (0)" is already visible, we're already on or close to it.
    pokazateli = page.locator(".tree-item-label", has_text="Показатели (0)").first
    if not pokazateli.is_visible():
        # Expand the Analytics folder first.
        page.locator(".tree-item-label", has_text="Аналитики").first.click()
        page.wait_for_timeout(500)
    pokazateli.wait_for(state="visible", timeout=10000)
    pokazateli.click()
    page.wait_for_timeout(1500)


def test_panel_shows_synthesized_scoped_rules(page: Page):
    """Surface 1: indicator settings panel must list each distinct cell-formula
    as its own scoped-rule entry (one rule per formula, with multiple period
    chips when several cells share that formula)."""
    _open_partners_panel(page)
    row = page.locator("tr", has_text="общее количество партнеров").first
    row.click()
    page.wait_for_timeout(800)
    panel = page.locator('[data-testid="indicator-formulas-panel"]')
    expect(panel).to_be_visible()
    # Panel must NOT show the old placeholder text; it must show real rules.
    expect(panel).not_to_contain_text("разные формулы у клеток")
    # Partners has at least 4 distinct formulas → ≥4 scoped entries.
    scoped_slots = panel.locator('[data-testid="formula-slot-scoped"]')
    assert scoped_slots.count() >= 4, (
        f"expected ≥4 synthesized scoped rules, got {scoped_slots.count()}"
    )
    # The +1 / +2 formula text must appear in the panel somewhere.
    expect(panel).to_contain_text("+1")
    expect(panel).to_contain_text("+2")


def test_grid_formula_column_shows_per_cell(page: Page):
    """Surface 1b: formula column in indicators grid (kept compact — the
    column is too narrow for 6 separate formulas, so the sentinel summary
    «разные формулы у клеток» still surfaces here)."""
    _open_partners_panel(page)
    cell = page.locator("tr", has_text="общее количество партнеров").locator(
        '[data-testid^="formula-cell-"]'
    ).first
    expect(cell).to_have_text("разные формулы у клеток")


def test_tooltip_and_formula_mode_per_cell(page: Page):
    """Surfaces 2 + 3: cell tooltip + grid formula mode show per-cell formula."""
    # Switch back to data mode so the sheets list appears in the left tree.
    page.locator('button[value="data"]').first.click()
    page.wait_for_timeout(700)
    sheet_label = page.locator(".tree-item-label", has_text="параметры модели").first
    sheet_label.wait_for(state="visible", timeout=10000)
    sheet_label.click()
    page.wait_for_timeout(1500)
    page.locator(".ag-root-wrapper").first.wait_for(state="visible", timeout=20000)

    # Switch to formulas mode (Σ icon in the toolbar)
    formula_toggle = page.locator('button[value="formulas"]').first
    formula_toggle.click()
    page.wait_for_timeout(1500)

    # The label "общее количество партнеров" lives in the LEFT-pinned slice;
    # the data cells we care about live in the CENTER slice on the SAME logical
    # row index. Find the partners row's index via its label, then locate the
    # corresponding center-slice row to read formula-mode cell texts.
    label_row = page.locator(
        '.ag-pinned-left-cols-container .ag-row',
        has_text="общее количество партнеров",
    ).first
    label_row.wait_for(state="visible", timeout=15000)
    row_index = label_row.get_attribute("row-index")
    assert row_index is not None, "could not read row-index of partners row"
    data_row = page.locator(
        f'.ag-center-cols-container .ag-row[row-index="{row_index}"]',
    ).first
    cells = data_row.locator(".ag-cell")
    n = cells.count()
    seen_texts: list[str] = []
    for i in range(min(6, n)):
        seen_texts.append((cells.nth(i).inner_text() or "").strip())
    # Bug under guard: partners row showed plain "SUM" on every month cell.
    # First 3 cells = Jan/Feb/Mar 2026 — Excel has D9=15 (manual) and E9/F9 as
    # +1 formulas. None of these may be the bare SUM placeholder.
    months = seen_texts[:3]
    plain_sum_in_months = [t for t in months if t == "ƒ SUM" or t == "SUM"]
    assert not plain_sum_in_months, (
        f"Formula mode still shows bare 'SUM' on partners month cells: {seen_texts}"
    )
    # The Jan cell must surface the manual-input label (D9 is a literal 15).
    assert "SUM" not in months[0] and ("ввод" in months[0] or "[" in months[0]), (
        f"Expected manual label or formula on Jan cell, got {months[0]!r}; row: {seen_texts}"
    )


def test_tooltip_in_data_mode_uses_same_source(page: Page):
    """Surface 2 (data-mode tooltip): hovering a parent-indicator cell in the
    regular data view must show the same per-cell formula / manual label as
    formula mode — never the misleading 'ƒ SUM' fallback."""
    page.locator('button[value="data"]').first.click()
    page.wait_for_timeout(700)
    sheet_label = page.locator(".tree-item-label", has_text="параметры модели").first
    sheet_label.wait_for(state="visible", timeout=10000)
    sheet_label.click()
    page.wait_for_timeout(2000)
    page.locator(".ag-root-wrapper").first.wait_for(state="visible", timeout=20000)
    # Make sure we're in DATA mode (not formulas)
    page.locator('button[value="data"]').first.click()
    page.wait_for_timeout(1500)

    label_row = page.locator(
        '.ag-pinned-left-cols-container .ag-row', has_text="общее количество партнеров",
    ).first
    label_row.wait_for(state="visible", timeout=15000)
    row_index = label_row.get_attribute("row-index")
    data_row = page.locator(
        f'.ag-center-cols-container .ag-row[row-index="{row_index}"]',
    ).first
    cells = data_row.locator(".ag-cell")

    def hover_tooltip(cell) -> str:
        page.mouse.move(0, 0)
        page.wait_for_timeout(300)
        cell.hover()
        page.wait_for_timeout(1500)
        tip = page.locator(".ag-tooltip, [class*='tooltip']").first
        return (tip.inner_text() or "").strip() if tip.is_visible() else ""

    jan_tip = hover_tooltip(cells.nth(0))
    feb_tip = hover_tooltip(cells.nth(1))
    assert "SUM" not in jan_tip, f"Jan tooltip still shows SUM: {jan_tip!r}"
    assert "ввод" in jan_tip, f"Jan (manual=15) tooltip must say 'ввод', got {jan_tip!r}"
    assert "SUM" not in feb_tip, f"Feb tooltip still shows SUM: {feb_tip!r}"
    assert "+1" in feb_tip or "[" in feb_tip, (
        f"Feb (formula) tooltip must echo the actual formula, got {feb_tip!r}"
    )
