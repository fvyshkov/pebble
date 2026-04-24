"""Playwright E2E tests for chat-driven workflows.

Four scenarios that also serve as benchmarks for comparing LLM providers:
1. Open a sheet and set cell values via chat
2. Import a model from Excel via UI dialog
3. Create a chart by indicator names
4. Generate a report/presentation for a model

Requires: backend on :8000, frontend on :5173 (or served by backend),
ANTHROPIC_API_KEY set, admin/admin credentials.

Run:  pytest tests/test_chat_benchmark.py --headed -x -v
"""
import json
import os
import pathlib
import time
import urllib.request

import pytest
from playwright.sync_api import Page, expect, BrowserContext

BASE = "http://localhost:8000"
ADMIN_USER = "admin"
ADMIN_PASS = "admin"

LLM_TIMEOUT = 90_000  # 90s — import + presentation can be slow
TEST_EXCEL = str(pathlib.Path(__file__).resolve().parent.parent / "test-avg.xlsx")

# ── Fixtures ──


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


# ── Helpers ──


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
    input_box = page.locator('[data-testid="chat-input"]')
    input_box.fill(text)
    page.locator('[data-testid="chat-send"]').click()
    page.wait_for_timeout(500)


def _wait_for_response(page: Page, timeout: int = LLM_TIMEOUT):
    """Wait until the assistant finishes responding (spinner gone)."""
    page.wait_for_timeout(2000)
    spinner = page.locator('[data-testid="chat-panel"] .MuiCircularProgress-root')
    try:
        spinner.wait_for(state="visible", timeout=5000)
    except Exception:
        pass
    try:
        spinner.wait_for(state="hidden", timeout=timeout)
    except Exception:
        pass
    page.wait_for_timeout(1000)


def _get_last_assistant_message(page: Page) -> str:
    """Return text of the last assistant message in chat."""
    msgs = page.locator('[data-testid="chat-panel"] .MuiPaper-root')
    if msgs.count() == 0:
        return ""
    return msgs.last.inner_text()


def _clear_chat(page: Page):
    """Clear chat history."""
    _send_message(page, "очисти историю чата")
    page.wait_for_timeout(2000)


# ── Test 1: Import model from Excel ──


@pytest.mark.skipif(not os.environ.get("ANTHROPIC_API_KEY"),
                    reason="ANTHROPIC_API_KEY not set")
def test_import_excel_model(logged_in_page: Page):
    """Import test-avg.xlsx via the UI import dialog and verify model appears."""
    page = logged_in_page
    t0 = time.time()

    # Click the import button in the toolbar (FileUploadOutlined icon)
    import_btn = page.locator('button:has([data-testid="FileUploadOutlinedIcon"])')
    if not import_btn.is_visible(timeout=3000):
        # Fallback: look for tooltip-based button
        import_btn = page.locator('[aria-label*="мпорт"]').first
    import_btn.click()
    page.wait_for_timeout(500)

    # Import dialog should be visible
    dialog = page.locator('.MuiDialog-root')
    expect(dialog).to_be_visible(timeout=3000)

    # Upload the file via hidden file input
    file_input = dialog.locator('input[type="file"]')
    file_input.set_input_files(TEST_EXCEL)
    page.wait_for_timeout(500)

    # Model name should be auto-filled from filename
    name_input = dialog.locator('input').filter(has_text="").last
    # The name field should have been populated
    page.wait_for_timeout(300)

    # Click the import/start button
    start_btn = dialog.locator('button:has-text("мпорт")').last
    if not start_btn.is_visible(timeout=2000):
        start_btn = dialog.locator('button').filter(has_text="Start").last
    start_btn.click()

    # Wait for import to complete — look for "done" or progress reaching end
    # The import shows streaming logs, and eventually a "done" state
    page.wait_for_timeout(3000)

    # Wait for the dialog to show completion (close button becomes active, or log says done)
    # We'll wait for up to 120s for the import to finish
    done_text = dialog.locator('text=/готово|done|✓|завершен/i')
    try:
        done_text.first.wait_for(state="visible", timeout=120_000)
    except Exception:
        # Check if dialog closed automatically
        pass

    # Close dialog if still open
    close_btn = dialog.locator('button:has-text("Закрыть"), button:has-text("Close")')
    if close_btn.first.is_visible(timeout=2000):
        close_btn.first.click()
        page.wait_for_timeout(500)

    elapsed = time.time() - t0

    # The imported model should appear in the left tree
    model_node = page.locator('.MuiTreeItem-root', has_text="test-avg")
    expect(model_node.first).to_be_visible(timeout=10000)

    print(f"\n[BENCH] test_import_excel_model: {elapsed:.1f}s")


# ── Test 2: Open sheet and set values via chat ──


@pytest.mark.skipif(not os.environ.get("ANTHROPIC_API_KEY"),
                    reason="ANTHROPIC_API_KEY not set")
def test_fill_sheet_with_random_values(logged_in_page: Page):
    """Ask chat to fill a sheet with random values and verify cells are populated."""
    page = logged_in_page
    t0 = time.time()

    # Click on the first sheet in the imported model's tree
    sheet_items = page.locator('.MuiTreeItem-root .MuiTreeItem-root')
    if sheet_items.count() > 0:
        sheet_items.first.click()
        page.wait_for_timeout(2000)

    # Ask chat to fill the sheet with random numbers
    _send_message(page, "заполни текущий лист случайными числами от 100 до 999")
    _wait_for_response(page)

    elapsed = time.time() - t0

    # Wait for grid to re-render with values
    page.wait_for_timeout(3000)

    # AG Grid cells with numeric content should be present
    grid_cells = page.locator('.ag-cell-value')
    cell_count = grid_cells.count()
    non_empty = 0
    for i in range(min(cell_count, 50)):
        text = grid_cells.nth(i).inner_text().strip().replace(" ", "").replace(",", ".")
        if text and text.replace(".", "", 1).replace("-", "", 1).isdigit():
            non_empty += 1

    assert non_empty > 0, f"No numeric cells found after fill command (checked {min(cell_count, 50)} cells)"
    print(f"\n[BENCH] test_fill_sheet_with_random_values: {elapsed:.1f}s, {non_empty} numeric cells")


# ── Test 3: Create chart by indicators ──


@pytest.mark.skipif(not os.environ.get("ANTHROPIC_API_KEY"),
                    reason="ANTHROPIC_API_KEY not set")
def test_create_chart_via_chat(logged_in_page: Page):
    """Ask chat to build a chart — verify canvas/SVG element appears."""
    page = logged_in_page
    t0 = time.time()

    # Make sure a sheet is open (from previous test)
    sheet_items = page.locator('.MuiTreeItem-root .MuiTreeItem-root')
    if sheet_items.count() > 0:
        sheet_items.first.click()
        page.wait_for_timeout(2000)

    _send_message(page, "построй столбчатую диаграмму по первому показателю")
    _wait_for_response(page)

    elapsed = time.time() - t0

    # amCharts renders to canvas
    chart = page.locator('canvas').first
    expect(chart).to_be_visible(timeout=15000)

    print(f"\n[BENCH] test_create_chart_via_chat: {elapsed:.1f}s")


# ── Test 4: Generate report/presentation ──


@pytest.mark.skipif(not os.environ.get("ANTHROPIC_API_KEY"),
                    reason="ANTHROPIC_API_KEY not set")
def test_generate_report(logged_in_page: Page):
    """Ask chat to build a presentation/report — verify iframe appears."""
    page = logged_in_page
    t0 = time.time()

    # Close any open chart/overlay first
    for _ in range(3):
        close_btn = page.locator('button:has([data-testid="CloseOutlinedIcon"])').first
        if close_btn.is_visible(timeout=1000):
            close_btn.click()
            page.wait_for_timeout(500)
        else:
            break

    # Clear chat history for clean context
    _clear_chat(page)

    # Make sure a sheet is selected
    sheet_items = page.locator('.MuiTreeItem-root .MuiTreeItem-root')
    if sheet_items.count() > 0:
        sheet_items.first.click()
        page.wait_for_timeout(1000)

    _send_message(page, "сделай презентацию по текущей модели")
    _wait_for_response(page, timeout=120_000)

    elapsed = time.time() - t0

    # Presentation renders as an iframe — may take extra time after LLM responds
    iframe = page.locator('iframe').first
    expect(iframe).to_be_visible(timeout=60_000)

    # PDF export button should be present
    pdf_btn = page.locator('button:has-text("PDF")')
    expect(pdf_btn).to_be_visible(timeout=10000)

    print(f"\n[BENCH] test_generate_report: {elapsed:.1f}s")


# ── Cleanup ──


def test_cleanup_imported_model(logged_in_page: Page):
    """Remove the test model created during import test via API."""
    # Find model by name via API
    resp = urllib.request.urlopen(f"{BASE}/api/models")
    models = json.loads(resp.read())
    test_models = [m for m in models if "test-avg" in m.get("name", "").lower()]

    for m in test_models:
        req = urllib.request.Request(f"{BASE}/api/models/{m['id']}", method="DELETE")
        urllib.request.urlopen(req)

    # Refresh the page so tree updates
    page = logged_in_page
    page.reload()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(2000)

    # Model should no longer be in tree
    model_node = page.locator('.MuiTreeItem-root', has_text="test-avg")
    expect(model_node).to_have_count(0, timeout=10000)
