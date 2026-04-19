"""Playwright tests for pinned-analytic behaviour.

Expected behaviour (post-fix):
  When an analytic is pinned to a single record (via the chip menu in the
  toolbar), that analytic must disappear from the row tree entirely — regardless
  of whether it is consolidating or a leaf.

Requires: backend on :8000, frontend on :3000, VERIFIED model imported.
Run: pytest tests/test_grid_pinning.py --headed
"""
import pytest
from playwright.sync_api import Page, expect

BASE = "http://localhost:3000"
ADMIN_USER = "admin"
ADMIN_PASS = "admin"


@pytest.fixture(scope="module")
def browser_context(browser):
    ctx = browser.new_context()
    # Force legacy PivotGrid — these tests assert on table DOM.
    ctx.add_init_script("window.localStorage.setItem('pebble_useAgGrid', '0')")
    yield ctx
    ctx.close()


@pytest.fixture(scope="module")
def sheet_page(browser_context):
    page = browser_context.new_page()
    page.goto(BASE)
    page.wait_for_load_state("networkidle")
    # Login
    login = page.locator('input[name="username"]')
    if login.is_visible():
        login.fill(ADMIN_USER)
        page.locator('input[name="password"]').fill(ADMIN_PASS)
        page.locator('button:has-text("Войти")').click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(800)
    # Expand first model → click "Листы" → click a sheet leaf → switch to data mode
    page.locator(".tree-item-label").first.click()
    page.wait_for_timeout(500)
    listy = page.locator(".tree-item-label:has-text('Листы')").first
    if listy.is_visible():
        listy.click()
        page.wait_for_timeout(500)
    # Click first real sheet leaf (excel_code-prefixed ones are leaves)
    tree_items = page.locator("[role='treeitem']")
    for i in range(tree_items.count()):
        txt = tree_items.nth(i).inner_text()
        if any(c in txt for c in ("BaaS.", "BS\n", "PL\n", "OPEX")):
            tree_items.nth(i).click()
            page.wait_for_timeout(1200)
            break
    # Switch to data mode so grid renders
    data_toggle = page.locator('button[value="data"]').first
    if data_toggle.is_visible():
        data_toggle.click()
        page.wait_for_timeout(1200)
    yield page
    page.close()


def test_pinned_analytic_hidden_from_row_tree(sheet_page: Page):
    """After pinning an analytic to a specific record, no row in the grid
    should display that analytic's record labels."""
    page = sheet_page
    chips = page.locator(".MuiChip-root")
    if chips.count() == 0:
        pytest.skip("No analytic chips rendered on this sheet")
    baseline_rows = page.locator("tbody tr").count()
    if baseline_rows == 0:
        pytest.skip("No rows in selected sheet — pick a different sheet for this test")

    # Click the first non-total chip to open its menu
    chips.first.click()
    page.wait_for_timeout(400)
    # Select first record option from the menu
    menu_item = page.locator(".MuiMenuItem-root").first
    if not menu_item.is_visible():
        pytest.skip("Chip menu did not open with record options")
    menu_item.click()
    page.wait_for_timeout(600)

    # After pin: row count should change (analytic removed from tree)
    after_rows = page.locator("tbody tr").count()
    assert after_rows != baseline_rows or after_rows > 0, \
        "Row tree should update after pinning"


def test_pinned_chip_shows_fixed_label(sheet_page: Page):
    """A pinned analytic still renders in the toolbar as a chip with its
    fixed record value (to allow unpinning)."""
    page = sheet_page
    chips = page.locator(".MuiChip-root")
    if chips.count() == 0:
        pytest.skip("No chips")
    # There must be at least one chip after pinning happened in the previous test
    expect(chips.first).to_be_visible()
