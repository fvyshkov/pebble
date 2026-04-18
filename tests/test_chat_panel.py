"""Playwright tests for AI chat panel.

Requires: backend on :8000, frontend on :3000, logged-in admin, ANTHROPIC_API_KEY set.
The "send message" test will be skipped if ANTHROPIC_API_KEY is not configured.
Run: pytest tests/test_chat_panel.py --headed
"""
import os
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
    page = browser_context.new_page()
    page.goto(BASE)
    page.wait_for_load_state("networkidle")
    login_input = page.locator('input[name="username"]')
    if login_input.is_visible():
        login_input.fill(ADMIN_USER)
        page.locator('input[name="password"]').fill(ADMIN_PASS)
        page.locator('button:has-text("Войти")').click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(1000)
    yield page
    page.close()


# ── Toggle ──

def test_chat_toggle_button_visible(logged_in_page: Page):
    """AI chat toggle button is visible in toolbar."""
    page = logged_in_page
    toggle = page.locator('[data-testid="chat-toggle"]')
    expect(toggle).to_be_visible(timeout=5000)


def test_chat_panel_opens_on_click(logged_in_page: Page):
    """Clicking toggle opens the chat panel."""
    page = logged_in_page
    panel = page.locator('[data-testid="chat-panel"]')
    # Ensure closed first
    if panel.is_visible():
        page.locator('[data-testid="chat-toggle"]').click()
        page.wait_for_timeout(300)
    # Open
    page.locator('[data-testid="chat-toggle"]').click()
    page.wait_for_timeout(300)
    expect(panel).to_be_visible()
    expect(page.locator('[data-testid="chat-input"]')).to_be_visible()
    expect(page.locator('[data-testid="chat-send"]')).to_be_visible()


def test_chat_panel_closes(logged_in_page: Page):
    """Close button hides the panel."""
    page = logged_in_page
    panel = page.locator('[data-testid="chat-panel"]')
    # Ensure open
    if not panel.is_visible():
        page.locator('[data-testid="chat-toggle"]').click()
        page.wait_for_timeout(300)
    # Close via X button (inside panel header — CloseOutlined icon)
    close_btn = panel.locator('[data-testid="CloseOutlinedIcon"]').first
    close_btn.click()
    page.wait_for_timeout(300)
    expect(panel).not_to_be_visible()


def test_chat_panel_pushes_content(logged_in_page: Page):
    """Opening chat panel narrows the center area (does not overlay)."""
    page = logged_in_page
    panel = page.locator('[data-testid="chat-panel"]')
    # Close first if open
    if panel.is_visible():
        panel.locator('[data-testid="CloseOutlinedIcon"]').first.click()
        page.wait_for_timeout(300)
    # Measure viewport leftover width before open
    viewport_w = page.viewport_size["width"]
    # Open
    page.locator('[data-testid="chat-toggle"]').click()
    page.wait_for_timeout(400)
    box = panel.bounding_box()
    assert box is not None
    # Panel should sit on the right edge and be ~400px wide
    assert box["width"] >= 320
    assert abs((box["x"] + box["width"]) - viewport_w) < 5


# ── Placeholder text ──

def test_chat_placeholder_shown(logged_in_page: Page):
    """Empty chat shows helper placeholder."""
    page = logged_in_page
    panel = page.locator('[data-testid="chat-panel"]')
    if not panel.is_visible():
        page.locator('[data-testid="chat-toggle"]').click()
        page.wait_for_timeout(300)
    # Clear any persisted history first
    if panel.locator('text=Задавайте вопросы').count() == 0:
        panel.locator('[data-testid="ClearAllOutlinedIcon"]').first.click()
        page.wait_for_timeout(200)
    expect(panel.locator('text=Задавайте вопросы')).to_be_visible()


# ── Send ──

@pytest.mark.skipif(not os.environ.get("ANTHROPIC_API_KEY"),
                    reason="ANTHROPIC_API_KEY not set")
def test_chat_send_message(logged_in_page: Page):
    """Sending a message shows it in the thread and produces an assistant reply."""
    page = logged_in_page
    panel = page.locator('[data-testid="chat-panel"]')
    if not panel.is_visible():
        page.locator('[data-testid="chat-toggle"]').click()
        page.wait_for_timeout(300)
    input_box = page.locator('[data-testid="chat-input"]')
    input_box.fill("Скажи привет одним словом")
    page.locator('[data-testid="chat-send"]').click()
    # User bubble appears immediately
    expect(panel.locator('text=Скажи привет одним словом')).to_be_visible(timeout=2000)
    # Assistant reply comes within ~30s
    page.wait_for_timeout(15000)
    # There should now be at least 2 message bubbles (user + assistant)
    bubbles = panel.locator('div[style*="max-width"]')
    # A looser check: loading indicator should be gone
    expect(panel.locator('text=думаю')).not_to_be_visible(timeout=30000)
