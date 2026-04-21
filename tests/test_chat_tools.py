"""Playwright E2E tests for chat tool invocations (AI-driven actions).

These tests exercise real Anthropic API calls through the chat panel,
verifying that tool-use commands (create model, rename, fill sheet,
build chart, build presentation) produce visible UI results.

Requires: backend on :8000, ANTHROPIC_API_KEY set, admin/admin credentials.
Run: pytest tests/test_chat_tools.py --headed -x
"""
import os
import pytest
from playwright.sync_api import Page, expect, BrowserContext

BASE = "http://localhost:8000"
ADMIN_USER = "admin"
ADMIN_PASS = "admin"

# LLM responses can be slow
LLM_TIMEOUT = 60_000  # 60s max wait for assistant response


@pytest.fixture(scope="module")
def browser_context(browser):
    context = browser.new_context(viewport={"width": 1400, "height": 900})
    yield context
    context.close()


@pytest.fixture(scope="module")
def logged_in_page(browser_context: BrowserContext):
    page = browser_context.new_page()
    page.goto(BASE)
    page.wait_for_load_state("networkidle")
    login_input = page.locator('input[name="username"]')
    if login_input.is_visible(timeout=5000):
        login_input.fill(ADMIN_USER)
        page.locator('input[name="password"]').fill(ADMIN_PASS)
        page.locator('button:has-text("ВОЙТИ")').click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(2000)
    yield page
    page.close()


def _open_chat(page: Page):
    """Ensure chat panel is open via Cmd+J."""
    panel = page.locator('[data-testid="chat-panel"]')
    if not panel.is_visible():
        page.keyboard.press("Meta+j")
        page.wait_for_timeout(500)
    expect(panel).to_be_visible(timeout=3000)
    return panel


def _send_message(page: Page, text: str):
    """Type a message into chat and send it."""
    _open_chat(page)
    # The chat input is a MUI TextField textarea inside data-testid="chat-input"
    input_box = page.locator('[data-testid="chat-input"]')
    input_box.fill(text)
    page.locator('[data-testid="chat-send"]').click()
    page.wait_for_timeout(500)


def _wait_for_response(page: Page, timeout: int = LLM_TIMEOUT):
    """Wait until the assistant finishes responding (CircularProgress gone)."""
    page.wait_for_timeout(2000)  # give LLM a head start
    # Wait for the loading spinner to appear then disappear
    spinner = page.locator('[data-testid="chat-panel"] .MuiCircularProgress-root')
    try:
        spinner.wait_for(state="visible", timeout=5000)
    except Exception:
        pass  # May have already finished
    try:
        spinner.wait_for(state="hidden", timeout=timeout)
    except Exception:
        pass
    page.wait_for_timeout(1000)


# ── Tests ──


@pytest.mark.skipif(not os.environ.get("ANTHROPIC_API_KEY"),
                    reason="ANTHROPIC_API_KEY not set")
def test_clear_chat_history(logged_in_page: Page):
    """Typing 'очисти историю чата' clears the chat messages."""
    page = logged_in_page

    # Send a seed message first so there's something to clear
    _send_message(page, "привет")
    _wait_for_response(page)

    # Now clear
    _send_message(page, "очисти историю чата")
    page.wait_for_timeout(1000)

    # After clearing, the welcome placeholder should reappear
    placeholder = page.locator('[data-testid="chat-panel"] >> text=Задавайте вопросы')
    expect(placeholder).to_be_visible(timeout=5000)


@pytest.mark.skipif(not os.environ.get("ANTHROPIC_API_KEY"),
                    reason="ANTHROPIC_API_KEY not set")
def test_create_model_via_chat(logged_in_page: Page):
    """Typing 'создай модель Тест-чат' creates a model visible in the tree."""
    page = logged_in_page

    _send_message(page, "создай модель Тест-чат-автотест")
    _wait_for_response(page)

    # The model should now appear in the left navigation tree
    model_node = page.locator('text=Тест-чат-автотест').first
    expect(model_node).to_be_visible(timeout=10000)


@pytest.mark.skipif(not os.environ.get("ANTHROPIC_API_KEY"),
                    reason="ANTHROPIC_API_KEY not set")
def test_rename_model_via_chat(logged_in_page: Page):
    """Renaming a model via chat updates the tree."""
    page = logged_in_page

    # Ensure model exists (may already exist from previous test)
    model_before = page.locator('text=Тест-чат-автотест')
    if model_before.count() == 0:
        _send_message(page, "создай модель Тест-чат-автотест")
        _wait_for_response(page)
        page.wait_for_timeout(2000)

    _send_message(page, "переименуй модель Тест-чат-автотест в Переименованная")
    _wait_for_response(page)

    # New name should appear in the tree
    expect(page.locator('text=Переименованная').first).to_be_visible(timeout=10000)


@pytest.mark.skipif(not os.environ.get("ANTHROPIC_API_KEY"),
                    reason="ANTHROPIC_API_KEY not set")
def test_fill_sheet_via_chat(logged_in_page: Page):
    """Asking to fill a sheet with random numbers populates cells."""
    page = logged_in_page

    # Navigate to a sheet with data — click on Sheet1
    sheet_link = page.locator('text=Sheet1').first
    if sheet_link.is_visible(timeout=3000):
        sheet_link.click()
        page.wait_for_timeout(2000)

    _send_message(page, "заполни случайными числами")
    _wait_for_response(page)

    # Wait extra for grid to re-render with values
    page.wait_for_timeout(3000)

    # AG Grid cells with numeric content should be present
    grid_cells = page.locator('.ag-cell-value')
    cell_count = grid_cells.count()
    non_empty = 0
    for i in range(min(cell_count, 30)):
        text = grid_cells.nth(i).inner_text().strip()
        if text and text not in ("", "—", "-"):
            non_empty += 1
    assert non_empty > 0, "No cells with values found after fill command"


@pytest.mark.skipif(not os.environ.get("ANTHROPIC_API_KEY"),
                    reason="ANTHROPIC_API_KEY not set")
def test_build_chart_via_chat(logged_in_page: Page):
    """Asking for a chart produces a chart panel with amCharts canvas."""
    page = logged_in_page

    _send_message(page, "сделай столбчатую диаграмму количества партнеров по периодам")
    _wait_for_response(page)

    # amCharts renders into a canvas element
    chart = page.locator('canvas').first
    expect(chart).to_be_visible(timeout=15000)


@pytest.mark.skipif(not os.environ.get("ANTHROPIC_API_KEY"),
                    reason="ANTHROPIC_API_KEY not set")
def test_build_presentation_via_chat(logged_in_page: Page):
    """Asking for a presentation produces an iframe with HTML report."""
    page = logged_in_page

    # Close chart if open (use first() to avoid strict mode violation with multiple close btns)
    close_btn = page.locator('button:has([data-testid="CloseOutlinedIcon"])').first
    if close_btn.is_visible(timeout=1000):
        close_btn.click()
        page.wait_for_timeout(500)

    _send_message(page, "сделай презентацию по данным")
    _wait_for_response(page, timeout=90_000)  # presentations take longer

    # Presentation renders as an iframe
    pres = page.locator('iframe').first
    expect(pres).to_be_visible(timeout=15000)

    # PDF button should be visible
    pdf_btn = page.locator('button:has-text("PDF")')
    expect(pdf_btn).to_be_visible(timeout=5000)
