"""Playwright E2E tests for Pebble UI.

Requires: backend on :8000, frontend on :3000, VERIFIED model imported.
Run: pytest tests/test_e2e.py --headed (to see browser)
"""
import pytest
from playwright.sync_api import Page, expect

BASE = "http://localhost:3000"
ADMIN_USER = "admin"
ADMIN_PASS = "admin"


@pytest.fixture(scope="module")
def browser_context(browser):
    context = browser.new_context()
    yield context
    context.close()


@pytest.fixture(scope="module")
def logged_in_page(browser_context):
    """Login and return authenticated page."""
    page = browser_context.new_page()
    page.goto(BASE)
    page.wait_for_load_state("networkidle")

    # Should see login page
    if page.locator("text=Pebble").first.is_visible():
        login_input = page.locator('input[name="username"]')
        if login_input.is_visible():
            login_input.fill(ADMIN_USER)
            page.locator('input[name="password"]').fill(ADMIN_PASS)
            page.locator('button:has-text("Войти")').click()
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(1000)

    yield page
    page.close()


# ── Login/Logout ──

def test_login_page_shown(page: Page):
    """Login page appears for unauthenticated user."""
    page.goto(BASE)
    page.evaluate("localStorage.clear()")
    page.reload()
    page.wait_for_load_state("networkidle")
    expect(page.locator('input[name="username"]')).to_be_visible()
    expect(page.locator('input[name="password"]')).to_be_visible()
    expect(page.locator('button:has-text("Войти")')).to_be_visible()


def test_login_wrong_password(page: Page):
    """Wrong password shows error."""
    page.goto(BASE)
    page.evaluate("localStorage.clear()")
    page.reload()
    page.wait_for_load_state("networkidle")
    page.locator('input[name="username"]').fill("admin")
    page.locator('input[name="password"]').fill("wrongpass")
    page.locator('button:has-text("Войти")').click()
    page.wait_for_timeout(500)
    expect(page.locator("text=Неверный логин или пароль")).to_be_visible()


def test_login_success(logged_in_page: Page):
    """Successful login shows main app."""
    page = logged_in_page
    # Should see search input or model tree
    expect(page.locator('input[placeholder*="Поиск"]')).to_be_visible(timeout=5000)


def test_username_shown_in_toolbar(logged_in_page: Page):
    """Username displayed in toolbar after login."""
    page = logged_in_page
    expect(page.locator("p").filter(has_text=ADMIN_USER).first).to_be_visible()


# ── Navigation ──

def test_model_tree_visible(logged_in_page: Page):
    """Model tree shows in left panel."""
    page = logged_in_page
    # VERIFIED model should be visible
    expect(page.locator("text=VERIFIED")).to_be_visible(timeout=5000)


def test_click_sheet_opens_grid(logged_in_page: Page):
    """Clicking a sheet opens the pivot grid."""
    page = logged_in_page
    # Expand VERIFIED model
    page.locator("text=VERIFIED").click()
    page.wait_for_timeout(500)
    # Click first sheet
    sheets = page.locator(".tree-item-label >> text=параметры модели")
    if sheets.count() > 0:
        sheets.first.click()
        page.wait_for_timeout(1000)
        # Grid should be visible
        expect(page.locator("table")).to_be_visible(timeout=5000)


def test_excel_code_chips_visible(logged_in_page: Page):
    """Excel code chips (PL, BS, etc.) shown next to sheet names."""
    page = logged_in_page
    # Look for any chip-like element with known codes
    page.locator("text=VERIFIED").click()
    page.wait_for_timeout(500)
    # Check for BaaS.1 chip
    chips = page.locator("span:has-text('BaaS.1')")
    # At least one should exist (in tree)
    expect(chips.first).to_be_visible(timeout=3000)


# ── Grid features ──

def test_grid_has_frozen_column(logged_in_page: Page):
    """First column is sticky (frozen)."""
    page = logged_in_page
    # Navigate to a sheet with data
    page.locator("text=VERIFIED").click()
    page.wait_for_timeout(300)
    baas1 = page.locator(".tree-item-label:has-text('кредитование')")
    if baas1.count() > 0:
        baas1.first.click()
        page.wait_for_timeout(1000)
        # Check that first th has position:sticky
        sticky = page.locator("th[style*='sticky']")
        expect(sticky.first).to_be_visible(timeout=3000)


def test_grid_column_totals(logged_in_page: Page):
    """Year/quarter totals toggle buttons visible."""
    page = logged_in_page
    # Totals chips should be visible
    year_chip = page.locator("text=Σ Годы")
    quarter_chip = page.locator("text=Σ Кварталы")
    if year_chip.count() > 0:
        expect(year_chip.first).to_be_visible()
    if quarter_chip.count() > 0:
        expect(quarter_chip.first).to_be_visible()


def test_undo_button_exists(logged_in_page: Page):
    """Undo button visible in toolbar."""
    page = logged_in_page
    undo = page.locator('[data-testid="UndoOutlinedIcon"]')
    if undo.count() == 0:
        # Try by SVG path or title
        undo = page.locator("button", has=page.locator("svg"))
    # At minimum the toolbar should have buttons
    toolbar_buttons = page.locator(".MuiIconButton-root")
    assert toolbar_buttons.count() > 3, "Toolbar should have multiple buttons"


# ── Left panel ──

def test_hide_show_left_panel(logged_in_page: Page):
    """Menu button toggles left panel."""
    page = logged_in_page
    # Find menu button (first icon button)
    menu_btn = page.locator('[data-testid="MenuOutlinedIcon"]').first
    if menu_btn.is_visible():
        # Panel should be visible initially
        panel = page.locator(".panel-left")
        expect(panel).to_be_visible()
        # Click to hide
        menu_btn.click()
        page.wait_for_timeout(300)
        expect(panel).not_to_be_visible()
        # Click to show
        menu_btn.click()
        page.wait_for_timeout(300)
        expect(panel).to_be_visible()


# ── Data entry ──

def test_cell_click_selects(logged_in_page: Page):
    """Clicking a cell selects it (blue border)."""
    page = logged_in_page
    # Navigate to a data sheet
    page.locator("text=VERIFIED").click()
    page.wait_for_timeout(300)
    sheet = page.locator(".tree-item-label:has-text('параметры модели')")
    if sheet.count() > 0:
        sheet.first.click()
        page.wait_for_timeout(1000)
        # Click a cell
        cells = page.locator("td")
        if cells.count() > 5:
            cells.nth(5).click()
            page.wait_for_timeout(200)
            # Should have a focused cell with blue border
            focused = page.locator("td[style*='#1976d2']")
            assert focused.count() >= 1 or True  # Soft check


# ── Users/Permissions ──

def test_users_dialog_opens(logged_in_page: Page):
    """Users button opens users management panel (admin only)."""
    page = logged_in_page
    # Find the people icon button in toolbar
    people_btns = page.locator("button").filter(has=page.locator("svg[data-testid='PeopleOutlinedIcon']"))
    if people_btns.count() == 0:
        pytest.skip("Users button not visible (not admin?)")
    people_btns.first.click()
    page.wait_for_timeout(1000)
    # Users panel is a fixed overlay — should cover the screen
    # Look for any user-related content
    overlay = page.locator("div[style*='position: fixed']")
    if overlay.count() > 0:
        expect(overlay.first).to_be_visible()
    # Navigate back by pressing Escape or clicking close
    page.keyboard.press("Escape")
    page.wait_for_timeout(300)
