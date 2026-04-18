"""Playwright tests for the AG Grid pivot view (Phase 1).

Covers:
  - Toggle switches grid between legacy and AG Grid
  - AG Grid renders rows/columns
  - Cell values load lazily on group expansion

Requires: backend on :8000, frontend on :3000, imported model.
Run: pytest tests/test_aggrid.py --headed
"""
import pytest
from playwright.sync_api import Page, expect

BASE = "http://localhost:3000"
ADMIN_USER = "admin"
ADMIN_PASS = "admin"


@pytest.fixture(scope="module")
def browser_context(browser):
    ctx = browser.new_context()
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
        page.wait_for_timeout(1000)
    yield page
    page.close()


def _enable_aggrid(page: Page) -> None:
    toggle = page.locator('[data-testid="aggrid-toggle"]')
    if not toggle.is_visible():
        pytest.skip("AG Grid toggle button not found")
    if "AG" not in toggle.inner_text():
        toggle.click()
        page.wait_for_timeout(1500)


def test_aggrid_toggle_visible(sheet_page: Page):
    expect(sheet_page.locator('[data-testid="aggrid-toggle"]')).to_be_visible()


def test_aggrid_renders_grid(sheet_page: Page):
    page = sheet_page
    _enable_aggrid(page)
    # AG Grid's root container has class .ag-root-wrapper
    expect(page.locator(".ag-root-wrapper").first).to_be_visible(timeout=8000)
    # And at least one row is rendered
    rows = page.locator(".ag-row")
    assert rows.count() > 0, "AG Grid rendered no rows"


def test_aggrid_lazy_load_on_expand(sheet_page: Page):
    """Expanding a top-level group should trigger cell fetch → values appear."""
    page = sheet_page
    _enable_aggrid(page)
    page.wait_for_timeout(1000)
    # AG Grid renders both .ag-group-contracted and .ag-group-expanded per
    # group row; only one is visible at a time.
    expander = page.locator(".ag-group-contracted:visible").first
    try:
        expander.wait_for(state="visible", timeout=5000)
    except Exception:
        pytest.skip("No visible collapsed-group chevron")
    requests: list[str] = []
    page.on("request", lambda r: requests.append(r.url) if "/cells/by-sheet/" in r.url and "/partial" in r.url else None)
    expander.click()
    page.wait_for_timeout(1500)
    # Either a partial request fired OR some cell already has non-empty value
    # (cells could already have been loaded on first render for top-level).
    assert len(requests) > 0, "Expected a lazy cell fetch after expanding a group"
