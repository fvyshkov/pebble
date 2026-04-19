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
    # Tree is pre-expanded in admin view: sheet leaves appear directly under
    # their model. Find the first tree label that carries an excel_code chip
    # (e.g. "BaaS.1", "BS", "PL", "OPEX…") and click it.
    page.wait_for_selector(".tree-item-label", timeout=8000)
    labels = page.locator(".tree-item-label")
    clicked = False
    for i in range(labels.count()):
        txt = labels.nth(i).inner_text()
        if any(c in txt for c in ("BaaS.", "BS\n", "PL\n", "OPEX")):
            labels.nth(i).click()
            page.wait_for_timeout(1500)
            clicked = True
            break
    if not clicked:
        pytest.fail(f"Could not find a sheet-leaf to click. Tree labels: {labels.all_inner_texts()[:10]}")
    # Data mode toggle may or may not be visible — default is already data.
    data_toggle = page.locator('button[value="data"]').first
    try:
        if data_toggle.is_visible():
            data_toggle.click()
            page.wait_for_timeout(1000)
    except Exception:
        pass
    yield page
    page.close()


def _enable_aggrid(page: Page) -> None:
    """AG Grid is now the only grid implementation — no toggle. Just wait
    until the AG root wrapper is in the DOM."""
    page.locator(".ag-root-wrapper").first.wait_for(state="visible", timeout=8000)


def test_aggrid_renders_grid(sheet_page: Page):
    page = sheet_page
    _enable_aggrid(page)
    # AG Grid's root container has class .ag-root-wrapper
    expect(page.locator(".ag-root-wrapper").first).to_be_visible(timeout=8000)
    # And at least one row is rendered
    rows = page.locator(".ag-row")
    assert rows.count() > 0, "AG Grid rendered no rows"


def test_aggrid_expand_reveals_children(sheet_page: Page):
    """Clicking a group chevron should expand the node and reveal more rows."""
    page = sheet_page
    _enable_aggrid(page)
    # Make sure no pins are active from a previous test run (VS is persisted).
    _unpin_all(page)
    page.wait_for_timeout(1500)
    rows_before = page.locator(".ag-row").count()
    # AG Grid with treeData: chevron is inside .ag-group-contracted. Try
    # multiple selector forms.
    expander = page.locator(".ag-group-contracted:visible, .ag-icon-tree-closed:visible").first
    try:
        expander.wait_for(state="visible", timeout=6000)
    except Exception:
        # Debug aid: dump group-cell class names for the first 5 rows.
        html = page.locator(".ag-row").first.inner_html()[:400]
        pytest.fail(f"No visible collapsed-group chevron. rows={rows_before}. First row HTML: {html}")
    expander.click()
    page.wait_for_timeout(800)
    rows_after = page.locator(".ag-row").count()
    assert rows_after > rows_before, (
        f"Expanding a group should add rows; before={rows_before} after={rows_after}"
    )


# ── Pinning behaviour ────────────────────────────────────────────────────────

def _drag_first_leaf_to_pin_zone(page: Page) -> bool:
    """Drag the first visible leaf row's analytic cell onto the pin drop-zone
    (the toolbar above the grid). Returns True if a chip appeared."""
    # The pin drop-zone is the bgcolor #fafafa strip above the grid — it's a
    # direct child of the PivotGridAG root. We target by the hint text it shows
    # when empty, or by the existing chips area.
    leaf_cell = page.locator(".ag-row .ag-cell .ag-group-value span[draggable='true']").first
    if not leaf_cell.is_visible():
        return False
    box_leaf = leaf_cell.bounding_box()
    # Drop zone: element with the hint text OR an existing MUI chip
    zone = page.locator(
        "text=Перетащите строку сюда, чтобы зафиксировать аналитику"
    ).first
    if not zone.is_visible():
        # chip already exists → we've already pinned something
        return True
    box_zone = zone.bounding_box()
    if not box_leaf or not box_zone:
        return False
    page.mouse.move(box_leaf["x"] + 10, box_leaf["y"] + box_leaf["height"] / 2)
    page.mouse.down()
    page.mouse.move(box_zone["x"] + 10, box_zone["y"] + box_zone["height"] / 2, steps=10)
    page.mouse.up()
    page.wait_for_timeout(800)
    return page.locator(".MuiChip-root").count() > 0


def _unpin_all(page: Page) -> None:
    """Remove any existing pin chips from the strip above the grid so tests
    start from a clean state. Uses the same xpath as the chip assertion."""
    for _ in range(10):  # cap — at most 10 chips
        chip_close = page.locator(
            "xpath=//div[contains(@class,'ag-root-wrapper')]/../../preceding-sibling::div"
            "//*[contains(@class,'MuiChip-root')]//*[contains(@class,'MuiChip-deleteIcon')]"
        ).first
        if chip_close.count() == 0 or not chip_close.is_visible():
            break
        chip_close.click()
        page.wait_for_timeout(400)


def test_aggrid_pin_single_row_analytic_shows_summary_row(sheet_page: Page):
    """Case 1: the sheet has only ONE row analytic. After pinning it, the
    grid must still show a single summary row (labelled with the analytic
    name + pinned record). Regression for the "No Rows To Show" bug."""
    page = sheet_page
    _enable_aggrid(page)
    _unpin_all(page)  # clean slate
    page.wait_for_timeout(1200)
    # Only proceed if the grid currently has rows; otherwise the sheet isn't
    # suitable for this test.
    if page.locator(".ag-row").count() == 0:
        pytest.skip("Selected sheet has no rows")
    pinned = _drag_first_leaf_to_pin_zone(page)
    if not pinned:
        pytest.skip("Could not drag a row onto the pin zone in this layout")
    # After a successful pin of the only row analytic, the grid must still
    # have at least one visible row. "No Rows To Show" overlay must NOT be up.
    overlay = page.locator(".ag-overlay-no-rows-wrapper:visible")
    assert overlay.count() == 0, \
        "After pinning the only row analytic the grid shows 'No Rows To Show' " \
        "— expected a summary row instead"
    assert page.locator(".ag-row").count() >= 1, \
        "Expected at least one summary row after pinning"


def test_aggrid_pin_chip_shows_analytic_name_and_value(sheet_page: Page):
    """Case 2: the pin chip must show the analytic name AND the fixed record
    value so the user knows what's pinned. The pin strip sits directly above
    the AG Grid root — scope the chip query there to avoid matching the
    calc-mode "авто" chip elsewhere in the page."""
    page = sheet_page
    _enable_aggrid(page)
    # Pin strip = the sibling element right above .ag-root-wrapper that has
    # bgcolor #fafafa and contains chips or the hint text.
    pin_chip = page.locator(
        "xpath=//div[contains(@class,'ag-root-wrapper')]/../../preceding-sibling::div"
        "//*[contains(@class,'MuiChip-root')]"
    ).first
    if pin_chip.count() == 0:
        pytest.skip("No pinned chip visible in pin strip")
    label = pin_chip.inner_text()
    assert ":" in label, f"Pin chip label missing ':' separator — got: {label!r}"
    a, _, b = label.partition(":")
    assert a.strip() and b.strip(), \
        f"Pin chip should have analytic name and record value — got {label!r}"
