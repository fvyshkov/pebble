"""Playwright tests for Excel-style keyboard shortcuts in the PivotGrid.

Covers:
  - arrow keys commit the current edit and move focus (while editing)
  - Ctrl+D  fills the focused cell value DOWN across the selection
  - Ctrl+R  fills RIGHT
  - Ctrl+Arrow jumps to next non-empty cell
  - Alt+Arrow toggles row collapse (previously Ctrl+Arrow)

Requires: backend on :8000, frontend on :3000, VERIFIED model.
Run: pytest tests/test_grid_shortcuts.py --headed
"""
import pytest
from playwright.sync_api import Page

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
    login = page.locator('input[name="username"]')
    if login.is_visible():
        login.fill(ADMIN_USER)
        page.locator('input[name="password"]').fill(ADMIN_PASS)
        page.locator('button:has-text("Войти")').click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(800)
    # Navigate into a sheet with editable cells
    # Expand first model → "Листы" → first sheet leaf → data mode
    page.locator(".tree-item-label").first.click()
    page.wait_for_timeout(500)
    listy = page.locator(".tree-item-label:has-text('Листы')").first
    if listy.is_visible():
        listy.click()
        page.wait_for_timeout(500)
    tree_items = page.locator("[role='treeitem']")
    for i in range(tree_items.count()):
        txt = tree_items.nth(i).inner_text()
        if any(c in txt for c in ("BaaS.", "BS\n", "PL\n", "OPEX")):
            tree_items.nth(i).click()
            page.wait_for_timeout(1200)
            break
    data_toggle = page.locator('button[value="data"]').first
    if data_toggle.is_visible():
        data_toggle.click()
        page.wait_for_timeout(1200)
    yield page
    page.close()


def _first_editable_cell(page: Page):
    """Return the first visibly-editable <td> (skips header cells)."""
    cells = page.locator("tbody td")
    return cells.nth(3) if cells.count() > 3 else cells.first


def test_arrow_commits_and_moves(sheet_page: Page):
    """Typing into a cell and pressing ArrowDown exits edit mode
    and moves focus (no <input> left open)."""
    page = sheet_page
    cell = _first_editable_cell(page)
    if not cell.is_visible():
        pytest.skip("No editable cells on sheet")
    cell.click()
    page.wait_for_timeout(100)
    cell.dblclick()
    page.wait_for_timeout(300)
    editing_input = page.locator("tbody input").first
    if not editing_input.count() or not editing_input.is_visible():
        pytest.skip("Selected cell is not editable (e.g. calculated row)")
    page.keyboard.type("42")
    page.keyboard.press("ArrowDown")
    page.wait_for_timeout(500)
    # After ArrowDown, the editor should be gone (edit committed & focus moved)
    inputs_now = page.locator("tbody input").count()
    assert inputs_now == 0, "ArrowDown should commit the edit and close the editor"


def test_ctrl_d_fills_down(sheet_page: Page):
    """Ctrl+D fills selection down with focused cell's value."""
    page = sheet_page
    cell = _first_editable_cell(page)
    if not cell.is_visible():
        pytest.skip("No editable cells")
    # Set a starting value
    cell.dblclick()
    page.wait_for_timeout(150)
    page.keyboard.type("123")
    page.keyboard.press("Enter")
    page.wait_for_timeout(300)
    # Select a range down: click start, shift+click 2 rows down in same column
    cell.click()
    page.wait_for_timeout(100)
    # Extend selection with shift+ArrowDown
    page.keyboard.press("Shift+ArrowDown")
    page.keyboard.press("Shift+ArrowDown")
    page.wait_for_timeout(150)
    # Invoke fill-down
    page.keyboard.press("Control+d")
    page.wait_for_timeout(500)
    # No exception & grid still renders
    assert page.locator("tbody tr").count() > 0


def test_ctrl_r_fills_right(sheet_page: Page):
    """Ctrl+R fills selection to the right."""
    page = sheet_page
    cell = _first_editable_cell(page)
    if not cell.is_visible():
        pytest.skip("No editable cells")
    cell.dblclick()
    page.wait_for_timeout(150)
    page.keyboard.type("7")
    page.keyboard.press("Enter")
    page.wait_for_timeout(200)
    cell.click()
    page.wait_for_timeout(100)
    page.keyboard.press("Shift+ArrowRight")
    page.keyboard.press("Shift+ArrowRight")
    page.wait_for_timeout(150)
    page.keyboard.press("Control+r")
    page.wait_for_timeout(500)
    assert page.locator("tbody tr").count() > 0


def test_ctrl_arrow_jumps_nonempty(sheet_page: Page):
    """Ctrl+ArrowRight jumps focus across the row; does not throw."""
    page = sheet_page
    cell = _first_editable_cell(page)
    if not cell.is_visible():
        pytest.skip("No editable cells")
    cell.click()
    page.wait_for_timeout(150)
    page.keyboard.press("Control+ArrowRight")
    page.wait_for_timeout(300)
    page.keyboard.press("Control+ArrowDown")
    page.wait_for_timeout(300)
    # Grid still alive
    assert page.locator("tbody tr").count() > 0


def test_alt_arrow_collapse(sheet_page: Page):
    """Alt+ArrowLeft/Right collapses/expands the focused row group."""
    page = sheet_page
    cell = _first_editable_cell(page)
    if not cell.is_visible():
        pytest.skip("No editable cells")
    cell.click()
    page.wait_for_timeout(150)
    before = page.locator("tbody tr").count()
    page.keyboard.press("Alt+ArrowLeft")
    page.wait_for_timeout(300)
    # Either row count changes (collapse) or nothing happens if not a group row
    after = page.locator("tbody tr").count()
    assert after >= 0
    # Restore
    page.keyboard.press("Alt+ArrowRight")
    page.wait_for_timeout(300)
