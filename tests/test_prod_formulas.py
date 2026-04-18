"""Playwright test — run against production (pebble-cu5s.onrender.com).

Verifies: no cell displays a raw Excel formula (text starting with `=`, like
`=E20*E21/1000` or `=SUM(E22:E24)` or `=BaaS.1!D20/BaaS.1!D16`).

Run:
  pytest tests/test_prod_formulas.py --headed -s
"""
import re
import pytest
from playwright.sync_api import Page

PROD = "https://pebble-cu5s.onrender.com"
ADMIN_USER = "admin"
ADMIN_PASS = "admin"

# Patterns of raw Excel formulas that should never leak into the grid
RAW_EXCEL_RE = re.compile(
    r"^\s*=\s*("
    r"[A-Z]+\$?\d+"                     # =E20 or =E$20
    r"|SUM\("                           # =SUM(...)
    r"|[A-Za-z0-9_.]+![A-Z]+\$?\d+"    # =BaaS.1!D20
    r"|\d+\*\d+"                        # =15*35000
    r")"
)


@pytest.fixture(scope="module")
def prod_page(browser):
    ctx = browser.new_context()
    page = ctx.new_page()
    page.goto(PROD, wait_until="networkidle")
    login = page.locator('input[name="username"]')
    if login.is_visible():
        login.fill(ADMIN_USER)
        page.locator('input[name="password"]').fill(ADMIN_PASS)
        page.locator('button:has-text("Войти")').click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(1500)
    yield page
    ctx.close()


def _open_sheet(page: Page, sheet_hint: str):
    """Expand first model → Листы → click the sheet whose label contains hint."""
    # Expand first model
    page.locator(".tree-item-label").first.click()
    page.wait_for_timeout(400)
    listy = page.locator(".tree-item-label:has-text('Листы')").first
    if listy.is_visible():
        listy.click()
        page.wait_for_timeout(400)
    # Click specific sheet
    tree_items = page.locator("[role='treeitem']")
    for i in range(tree_items.count()):
        if sheet_hint in tree_items.nth(i).inner_text():
            tree_items.nth(i).click()
            page.wait_for_timeout(1500)
            return True
    return False


def _raw_formula_cells(page: Page) -> list[str]:
    """Return list of cell texts that look like raw Excel formulas."""
    bad = []
    cells = page.locator("tbody td")
    n = cells.count()
    for i in range(n):
        try:
            text = cells.nth(i).inner_text().strip()
        except Exception:
            continue
        if RAW_EXCEL_RE.match(text):
            bad.append(text)
    return bad


def test_opex_capex_sheet_no_raw_excel_formulas(prod_page: Page):
    """Sheet 'Операционные расходы и Инвестиции в BaaS' must not show raw
    Excel formulas like =E20*E21/1000 or =SUM(E22:E24)."""
    page = prod_page
    found = _open_sheet(page, "OPEX")
    if not found:
        pytest.skip("OPEX+CAPEX sheet not present on prod")
    # Switch to data mode
    data_toggle = page.locator('button[value="data"]').first
    if data_toggle.is_visible():
        data_toggle.click()
        page.wait_for_timeout(1200)

    bad = _raw_formula_cells(page)
    assert not bad, (
        f"Found {len(bad)} cells displaying raw Excel formulas. "
        f"First few: {bad[:5]}"
    )


def test_baas1_sheet_no_raw_excel_formulas(prod_page: Page):
    """Sheet 'BaaS - Онлайн кредитование' must not show raw Excel formulas."""
    page = prod_page
    found = _open_sheet(page, "BaaS.1")
    if not found:
        pytest.skip("BaaS.1 sheet not present on prod")
    data_toggle = page.locator('button[value="data"]').first
    if data_toggle.is_visible():
        data_toggle.click()
        page.wait_for_timeout(1200)

    bad = _raw_formula_cells(page)
    assert not bad, (
        f"Found {len(bad)} cells displaying raw Excel formulas. "
        f"First few: {bad[:5]}"
    )
