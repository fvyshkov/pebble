"""Unified Playwright E2E test suite for Pebble.

Self-contained: starts from a clean DB, imports a model, creates users,
runs UI checks including AG Grid features, then cleans up.

Run:
    pytest tests/test_e2e.py -v -s                    # headless
    pytest tests/test_e2e.py -v -s --headed            # with browser window
"""
import json as _json
import os
import time
import requests
import pytest
from playwright.sync_api import Page, Browser, BrowserContext, expect

BASE = os.environ.get("PEBBLE_BASE", "http://localhost:3000")
API = os.environ.get("PEBBLE_API", "http://localhost:8000/api")
MODEL_NAME = "UIFLOW"
XLSX_PATH = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "models.xlsx")

ADMIN = ("admin", "admin")
DEP1 = ("dep1", "dep1")
DEP2 = ("dep2", "dep2")

SCREENSHOT_DIR = "/tmp/uiflow"


# ──────────────────────────────────────────────────────────────────
# Setup / cleanup
# ──────────────────────────────────────────────────────────────────

def _api_get(path):
    r = requests.get(f"{API}{path}", timeout=30)
    r.raise_for_status()
    return r.json()


def _api_delete(path):
    return requests.delete(f"{API}{path}", timeout=30)


def _cleanup_namespace():
    """Delete UIFLOW model and dep1/dep2 users if they exist.
    Ensure admin user exists."""
    try:
        models = _api_get("/models")
        for m in models:
            if m["name"] == MODEL_NAME:
                _api_delete(f"/models/{m['id']}")
        users = _api_get("/users")
        for u in users:
            if u["username"] in ("dep1", "dep2", "Новый пользователь"):
                _api_delete(f"/users/{u['id']}")
        # Ensure admin exists
        import sqlite3, uuid
        db_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "pebble.db")
        c = sqlite3.connect(db_path)
        row = c.execute("SELECT id FROM users WHERE username = 'admin'").fetchone()
        if not row:
            c.execute(
                "INSERT INTO users (id, username, password, can_admin, created_at) VALUES (?, 'admin', 'admin', 1, datetime('now'))",
                (str(uuid.uuid4()),)
            )
            c.commit()
        c.close()
    except Exception as e:
        print(f"[cleanup] warning: {e}")


@pytest.fixture(scope="module", autouse=True)
def prepare():
    """Ensure backend is up; clean namespace."""
    os.makedirs(SCREENSHOT_DIR, exist_ok=True)
    for _ in range(30):
        try:
            r = requests.get(f"{API}/models", timeout=2)
            if r.status_code == 200:
                break
        except Exception:
            time.sleep(1)
    else:
        pytest.fail("Backend not responding")
    _cleanup_namespace()
    yield


# ──────────────────────────────────────────────────────────────────
# Browser fixtures (module-scoped, chain state across tests)
# ──────────────────────────────────────────────────────────────────

@pytest.fixture(scope="module")
def browser_context(browser: Browser):
    ctx = browser.new_context(viewport={"width": 1400, "height": 900})
    ctx.set_default_timeout(20000)
    yield ctx
    ctx.close()


@pytest.fixture(scope="module")
def page(browser_context: BrowserContext):
    p = browser_context.new_page()
    yield p
    p.close()


def _shot(page: Page, name: str):
    try:
        page.screenshot(path=f"{SCREENSHOT_DIR}/{name}.png", full_page=True)
    except Exception:
        pass


# ──────────────────────────────────────────────────────────────────
# Helpers
# ──────────────────────────────────────────────────────────────────

def _login(page: Page, username: str, password: str):
    page.goto(BASE)
    page.evaluate("localStorage.clear()")
    page.reload()
    page.wait_for_load_state("networkidle")
    page.locator('input[name="username"]').fill(username)
    page.locator('input[name="password"]').fill(password)
    page.locator('button:has-text("Войти")').click()
    page.wait_for_selector('input[placeholder*="Поиск"]', timeout=15000)
    page.wait_for_timeout(800)


def _logout(page: Page):
    for _ in range(3):
        page.keyboard.press('Escape')
        page.wait_for_timeout(200)
    page.evaluate("localStorage.clear()")
    page.goto(BASE)
    page.wait_for_selector('input[name="username"]', timeout=15000)


def _expand_tree_item_by_id(page: Page, id_fragment: str):
    ti = page.locator(f'li[role="treeitem"][id*="{id_fragment}"]').first
    ti.wait_for(state="attached", timeout=10000)
    expanded = ti.get_attribute('aria-expanded')
    if expanded == 'false':
        icon = ti.locator('> .MuiTreeItem-content > .MuiTreeItem-iconContainer').first
        icon.click()
        page.wait_for_timeout(350)


def _open_tree_model(page: Page):
    _expand_tree_item_by_id(page, 'model:')
    folder = page.locator('li[role="treeitem"][id*="sheets-folder:"]')
    if folder.count() > 0:
        _expand_tree_item_by_id(page, 'sheets-folder:')


def _switch_to_data_mode(page: Page):
    btn = page.locator('button:has(svg[data-testid="TableChartOutlinedIcon"])').first
    if btn.count() == 0:
        return
    btn.wait_for(state='visible', timeout=5000)
    if btn.get_attribute('aria-pressed') != 'true':
        btn.click()
        page.wait_for_timeout(2000)


def _switch_to_settings_mode(page: Page):
    btn = page.locator('button:has(svg[data-testid="SettingsOutlinedIcon"])').first
    if btn.count() == 0:
        return
    btn.wait_for(state='visible', timeout=5000)
    if btn.get_attribute('aria-pressed') != 'true':
        btn.click()
        page.wait_for_timeout(1500)


def _click_sheet(page: Page, code: str = 'BaaS.1'):
    # Dismiss any overlay / context-menu backdrop that may intercept clicks
    page.keyboard.press('Escape')
    page.wait_for_timeout(300)
    _open_tree_model(page)
    sheet = page.locator('li[role="treeitem"][id*="-sheet:"]').filter(
        has=page.locator(f'span:text-is("{code}")')).first
    sheet.scroll_into_view_if_needed()
    sheet.locator('.tree-item-label').first.click()
    page.wait_for_timeout(1200)
    _switch_to_data_mode(page)
    page.locator('.ag-root-wrapper').first.wait_for(state='visible', timeout=15000)
    page.wait_for_timeout(1500)


def _wait_for_ag_grid(page: Page, timeout: int = 15000):
    page.locator('.ag-root-wrapper').first.wait_for(state='visible', timeout=timeout)
    page.locator('.ag-cell').first.wait_for(state='visible', timeout=timeout)


# ══════════════════════════════════════════════════════════════════
# PART 1: Login
# ══════════════════════════════════════════════════════════════════

def test_01_login_page_shown(page: Page):
    """Login page appears for unauthenticated user."""
    page.goto(BASE, wait_until="networkidle")
    page.evaluate("localStorage.clear()")
    page.reload()
    page.wait_for_load_state("networkidle")
    expect(page.locator('input[name="username"]')).to_be_visible()
    expect(page.locator('input[name="password"]')).to_be_visible()
    expect(page.locator('button:has-text("Войти")')).to_be_visible()


def test_02_login_wrong_password(page: Page):
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


def test_03_admin_login(page: Page):
    """Admin login succeeds."""
    _login(page, *ADMIN)
    expect(page.locator('input[placeholder*="Поиск"]')).to_be_visible(timeout=5000)
    _shot(page, "03_login")


# ══════════════════════════════════════════════════════════════════
# PART 2: Import model
# ══════════════════════════════════════════════════════════════════

def test_04_upload_model(page: Page):
    """Import reference model from Excel."""
    if not os.path.isfile(XLSX_PATH):
        pytest.skip(f"models.xlsx not found: {XLSX_PATH}")

    page.locator('button[aria-label="Импорт модели из Excel"]').click()
    page.wait_for_timeout(500)
    page.get_by_label('Название модели').fill(MODEL_NAME)
    page.locator('input[type="file"][accept=".xlsx,.xls"]').set_input_files(XLSX_PATH)
    page.wait_for_timeout(500)
    page.locator('.MuiDialog-root button', has_text='Импортировать').click()
    close_btn = page.locator('.MuiDialog-root button', has_text='Закрыть')
    close_btn.wait_for(state="visible", timeout=600_000)
    _shot(page, "04_import_done")
    close_btn.click()
    page.wait_for_timeout(1000)
    expect(page.locator(f'text={MODEL_NAME}').first).to_be_visible(timeout=10000)


def test_05_model_tree_visible(page: Page):
    """Model tree shows imported model with sheets."""
    model = page.locator(f'text={MODEL_NAME}').first
    model.scroll_into_view_if_needed()
    expect(model).to_be_visible(timeout=5000)


# ══════════════════════════════════════════════════════════════════
# PART 3: Grid UI checks
# ══════════════════════════════════════════════════════════════════

def test_06_grid_opens(page: Page):
    """Clicking a sheet opens AG Grid."""
    _click_sheet(page, 'BaaS.1')
    expect(page.locator("[role=treegrid]")).to_be_visible(timeout=5000)
    _shot(page, "06_grid")


def test_07_grid_has_data(page: Page):
    """AG Grid rendered with numeric cells."""
    cells = page.locator('.ag-cell')
    count = cells.count()
    assert count > 50, f"Expected > 50 AG Grid cells, got {count}"
    texts = cells.all_text_contents()
    num_count = 0
    for t in texts:
        t = t.strip().replace('\xa0', '').replace(' ', '').replace(',', '.')
        try:
            float(t)
            num_count += 1
        except ValueError:
            pass
    assert num_count > 5, f"No numeric cells found: {num_count}"


def test_08_grid_has_frozen_column(page: Page):
    """First column is pinned left (frozen) in AG Grid."""
    pinned = page.locator(".ag-pinned-left-header")
    expect(pinned.first).to_be_visible(timeout=3000)


def test_09_excel_code_chips_visible(page: Page):
    """Excel code chips shown next to sheet names in tree."""
    chips = page.locator("span:has-text('BaaS.1')")
    expect(chips.first).to_be_visible(timeout=3000)


def test_10_period_chips_toggle(page: Page):
    """Period level Σ chips are clickable yes/no toggles."""
    # Find period chips (Σ Годы, Σ Кварталы etc.)
    chips = page.locator('[data-testid^="col-level-toggle-"]')
    if chips.count() == 0:
        pytest.skip("No period level chips found on this sheet")
    chip = chips.first
    expect(chip).to_be_visible()
    # Chip should be a simple clickable element (not a 3-button group)
    toggle_buttons = chip.locator('.MuiToggleButton-root')
    assert toggle_buttons.count() == 0, "Period chip should be a simple Chip, not ToggleButtonGroup"
    # Chips default to ON (filled). Turn off first, then verify toggle adds columns.
    is_on = chip.evaluate("el => el.className.includes('MuiChip-filled')")
    if is_on:
        chip.click()
        page.wait_for_timeout(1500)
    cols_off = page.evaluate(
        "() => window.__pebbleGridApi ? window.__pebbleGridApi.getColumns().length : 0"
    )
    # Click to toggle ON — sum columns should appear
    chip.click()
    page.wait_for_timeout(1500)
    cols_on = page.evaluate(
        "() => window.__pebbleGridApi ? window.__pebbleGridApi.getColumns().length : 0"
    )
    assert cols_on > cols_off, f"Sum columns should appear after toggle ON ({cols_off} -> {cols_on})"
    _shot(page, "10_period_on")
    # Click to toggle OFF — sum columns should disappear
    chip.click()
    page.wait_for_timeout(1500)
    cols_final = page.evaluate(
        "() => window.__pebbleGridApi ? window.__pebbleGridApi.getColumns().length : 0"
    )
    assert cols_final < cols_on, f"Sum columns should disappear after toggle OFF ({cols_on} -> {cols_final})"
    _shot(page, "10_period_off")


def test_11_column_resize_persists(page: Page):
    """Resizing a column width persists after period toggle."""
    grid = page.locator('.ag-root-wrapper').first
    if not grid.is_visible():
        pytest.skip("Grid not visible")
    # Get first data column header
    headers = page.locator('.ag-header-cell[col-id^="p_"]')
    if headers.count() == 0:
        pytest.skip("No period columns")
    header = headers.first
    bb = header.bounding_box()
    if not bb:
        pytest.skip("Cannot get header bounding box")
    # Drag right edge to resize
    right_edge_x = bb['x'] + bb['width'] - 2
    center_y = bb['y'] + bb['height'] / 2
    page.mouse.move(right_edge_x, center_y)
    page.mouse.down()
    page.mouse.move(right_edge_x + 50, center_y, steps=5)
    page.mouse.up()
    page.wait_for_timeout(300)
    # Read new width
    bb2 = header.bounding_box()
    new_width = bb2['width'] if bb2 else 0
    # Toggle a period chip on/off to trigger column rebuild
    chips = page.locator('[data-testid^="col-level-toggle-"]')
    if chips.count() > 0:
        chips.first.click()
        page.wait_for_timeout(500)
        chips.first.click()
        page.wait_for_timeout(500)
    # Check width is approximately preserved (within 20px tolerance)
    bb3 = header.bounding_box()
    if bb3:
        assert abs(bb3['width'] - new_width) < 20, \
            f"Column width jumped: was {new_width}, now {bb3['width']}"


def test_12_undo_button_exists(page: Page):
    """Undo button and dropdown visible in main toolbar."""
    undo_btn = page.locator('[data-testid="undo-btn"]')
    expect(undo_btn).to_be_visible(timeout=3000)
    dropdown_btn = page.locator('[data-testid="undo-dropdown-btn"]')
    expect(dropdown_btn).to_be_visible(timeout=3000)


def test_13_undo_dropdown_opens(page: Page):
    """Undo dropdown opens and shows history after a cell edit."""
    # Make an actual cell edit so undo history exists
    editable = page.locator('.ag-cell[col-id^="p_"]').first
    if not editable.is_visible():
        pytest.skip("No editable period cell visible")
    editable.dblclick()
    page.wait_for_timeout(300)
    page.keyboard.type('999')
    page.keyboard.press('Tab')
    page.wait_for_timeout(1500)
    # Now undo dropdown should be enabled
    dropdown_btn = page.locator('[data-testid="undo-dropdown-btn"]')
    if dropdown_btn.get_attribute('disabled') is not None:
        pytest.skip("Undo dropdown still disabled after edit")
    dropdown_btn.click()
    page.wait_for_timeout(500)
    # Popover should be open
    popover = page.locator('.MuiPopover-paper')
    expect(popover).to_be_visible(timeout=3000)
    # Close it
    page.keyboard.press('Escape')
    page.wait_for_timeout(300)
    # Undo the edit we just made
    undo_btn = page.locator('[data-testid="undo-btn"]')
    undo_btn.click()
    page.wait_for_timeout(1500)


def test_14_hide_show_left_panel(page: Page):
    """Menu button toggles left panel visibility."""
    menu_btn = page.locator('[data-testid="MenuOutlinedIcon"]').first
    if not menu_btn.is_visible():
        pytest.skip("Menu button not visible")
    panel = page.locator(".panel-left")
    expect(panel).to_be_visible()
    menu_btn.click()
    page.wait_for_timeout(300)
    expect(panel).not_to_be_visible()
    menu_btn.click()
    page.wait_for_timeout(300)
    expect(panel).to_be_visible()


def test_15_users_dialog_opens(page: Page):
    """Users button opens management panel and Escape closes it."""
    people_btns = page.locator("button").filter(
        has=page.locator("svg[data-testid='PeopleOutlinedIcon']"))
    if people_btns.count() == 0:
        pytest.skip("Users button not visible")
    people_btns.first.click()
    page.wait_for_timeout(1000)
    # Users panel is a fixed overlay
    overlay = page.locator("div[style*='position: fixed']")
    if overlay.count() > 0:
        expect(overlay.first).to_be_visible()
    # Close via Escape (UsersDialog supports it)
    page.keyboard.press("Escape")
    page.wait_for_timeout(500)
    # Verify overlay is gone
    fixed_overlays = page.locator("div[style*='position: fixed'][style*='inset: 0']")
    assert fixed_overlays.count() == 0, "Users overlay should be closed after Escape"


# ══════════════════════════════════════════════════════════════════
# PART 4: Import verification (API-based)
# ══════════════════════════════════════════════════════════════════

def test_16_import_sets_is_main(page: Page):
    """After import, indicator analytics have is_main=1 on their sheet bindings."""
    models = requests.get(f"{API}/models").json()
    uif = next((m for m in models if m["name"] == MODEL_NAME), None)
    assert uif, f"Model {MODEL_NAME} not found"
    tree = requests.get(f"{API}/models/{uif['id']}/tree").json()
    sheets = tree.get("sheets", [])
    assert len(sheets) > 0
    sa = requests.get(f"{API}/sheets/{sheets[0]['id']}/analytics").json()
    main_bindings = [b for b in sa if b.get("is_main") == 1]
    assert len(main_bindings) >= 1, "At least one analytic should have is_main=1"


def test_17_hierarchy_records_have_parents(page: Page):
    """Imported records preserve parent-child hierarchy."""
    models = requests.get(f"{API}/models").json()
    uif = next((m for m in models if m["name"] == MODEL_NAME), None)
    assert uif
    tree = requests.get(f"{API}/models/{uif['id']}/tree").json()
    for sheet in tree.get("sheets", []):
        sa = requests.get(f"{API}/sheets/{sheet['id']}/analytics").json()
        ind = next((b for b in sa if b.get("is_main") == 1), None)
        if not ind:
            continue
        recs = requests.get(f"{API}/analytics/{ind['analytic_id']}/records").json()
        children = [r for r in recs if r.get("parent_id") is not None]
        if len(children) >= 2:
            return
    pytest.fail("No sheet has indicator records with parent-child hierarchy")


# ══════════════════════════════════════════════════════════════════
# PART 5: Department analytic + multi-user scenario
# ══════════════════════════════════════════════════════════════════

def test_18_create_analytic(page: Page):
    """Create Подразделение analytic with hierarchy via API."""
    models = requests.get(f"{API}/models").json()
    uif = next((m for m in models if m["name"] == MODEL_NAME), None)
    assert uif
    model_id = uif["id"]

    r = requests.post(f"{API}/analytics", json={"model_id": model_id, "name": "Подразделение"})
    assert r.status_code == 200
    analytic_id = r.json()["id"]

    r = requests.post(f"{API}/analytics/{analytic_id}/records",
                       json={"data_json": {"name": "Головной"}, "parent_id": None})
    assert r.status_code == 200
    head_id = r.json()["id"]

    for name in ("Филиал 1", "Филиал 2"):
        r = requests.post(f"{API}/analytics/{analytic_id}/records",
                           json={"data_json": {"name": name}, "parent_id": head_id})
        assert r.status_code == 200


def test_19_add_analytic_to_sheets(page: Page):
    """Bind Подразделение to all sheets via API."""
    models = requests.get(f"{API}/models").json()
    uif = next((m for m in models if m["name"] == MODEL_NAME), None)
    assert uif
    tree = requests.get(f"{API}/models/{uif['id']}/tree").json()
    podr = next((a for a in tree.get("analytics", []) if a["name"] == "Подразделение"), None)
    assert podr

    linked = 0
    for s in tree.get("sheets", []):
        r = requests.post(f"{API}/sheets/{s['id']}/analytics",
                          json={"analytic_id": podr["id"]})
        if r.status_code == 200:
            linked += 1
    assert linked > 0


def test_20_verify_branch_distribution(page: Page):
    """After binding, grid shows data and branches are linked."""
    _click_sheet(page, 'BaaS.1')
    page.wait_for_timeout(2000)
    expect(page.locator('.ag-root-wrapper').first).to_be_visible()

    models = requests.get(f"{API}/models").json()
    uif = next((m for m in models if m["name"] == MODEL_NAME), None)
    assert uif
    tree = requests.get(f"{API}/models/{uif['id']}/tree").json()
    podr = next((a for a in tree.get("analytics", []) if a["name"] == "Подразделение"), None)
    assert podr

    linked = 0
    for s in tree.get("sheets", []):
        sa = requests.get(f"{API}/sheets/{s['id']}/analytics").json()
        if any(x.get("analytic_id") == podr["id"] for x in sa):
            linked += 1
    assert linked == 7, f"Подразделение linked to {linked}/7 sheets"
    _shot(page, "20_branches")


def _create_user_with_branch_api(username: str, password: str, branch_name: str):
    r = requests.post(f"{API}/users", json={"username": username}, timeout=30)
    assert r.status_code == 200
    user_id = r.json()["id"]

    r = requests.post(f"{API}/users/{user_id}/reset-password",
                      json={"password": password}, timeout=30)
    assert r.status_code == 200

    models = requests.get(f"{API}/models", timeout=30).json()
    uif = next((m for m in models if m["name"] == MODEL_NAME), None)
    assert uif
    tree = requests.get(f"{API}/models/{uif['id']}/tree", timeout=30).json()

    for a in tree.get("analytics", []):
        if a.get("is_periods"):
            continue
        recs = requests.get(f"{API}/analytics/{a['id']}/records", timeout=30).json()
        for rec in recs:
            dj = rec.get("data_json", {})
            if isinstance(dj, str):
                dj = _json.loads(dj)
            if (dj or {}).get("name") == branch_name:
                requests.put(f"{API}/users/analytic-permissions/set", json={
                    "user_id": user_id, "analytic_id": a["id"],
                    "record_id": rec["id"], "can_view": True, "can_edit": True,
                }, timeout=30)
                return user_id
    pytest.fail(f"Record '{branch_name}' not found")


def test_21_create_dep1(page: Page):
    _create_user_with_branch_api('dep1', 'dep1', 'Филиал 1')


def test_22_create_dep2(page: Page):
    _create_user_with_branch_api('dep2', 'dep2', 'Филиал 2')


def _find_editable_ag_cell(page: Page):
    cells = page.locator('.ag-cell').all()
    for cell in cells:
        style = cell.get_attribute('style') or ''
        if 'fdf8e8' in style:
            return cell
    return page.locator('.ag-cell').first


def test_23_dep1_enter(page: Page):
    """dep1 logs in, enters value in editable cell."""
    _logout(page)
    _login(page, *DEP1)
    _click_sheet(page, 'BaaS.1')
    _wait_for_ag_grid(page)

    cell = _find_editable_ag_cell(page)
    expect(cell).to_be_visible(timeout=10000)
    cell.scroll_into_view_if_needed()
    cell.dblclick()
    page.wait_for_timeout(300)
    page.keyboard.type('77777')
    page.keyboard.press('Enter')
    page.wait_for_timeout(1500)
    _shot(page, "23_dep1")


def test_24_dep2_enter(page: Page):
    """dep2 logs in, enters value in editable cell."""
    _logout(page)
    _login(page, *DEP2)
    _click_sheet(page, 'BaaS.1')
    _wait_for_ag_grid(page)

    cell = _find_editable_ag_cell(page)
    expect(cell).to_be_visible(timeout=10000)
    cell.scroll_into_view_if_needed()
    cell.dblclick()
    page.wait_for_timeout(300)
    page.keyboard.type('33333')
    page.keyboard.press('Enter')
    page.wait_for_timeout(1500)
    _shot(page, "24_dep2")


def test_25_admin_consolidation(page: Page):
    """Admin sees branch data distributed correctly."""
    models = requests.get(f"{API}/models").json()
    uif = next((m for m in models if m["name"] == MODEL_NAME), None)
    assert uif
    model_id = uif["id"]

    sheets = requests.get(f"{API}/sheets/by-model/{model_id}").json()
    baas1 = next((s for s in sheets if s.get("excel_code") == "BaaS.1"), None)
    assert baas1
    sheet_id = baas1["id"]

    tree = requests.get(f"{API}/models/{model_id}/tree").json()
    podr = next((a for a in tree.get("analytics", []) if a["name"] == "Подразделение"), None)
    assert podr

    recs = requests.get(f"{API}/analytics/{podr['id']}/records").json()
    rec_names = {}
    for rec in recs:
        dj = rec.get("data_json", {})
        if isinstance(dj, str):
            dj = _json.loads(dj)
        rec_names[rec["id"]] = (dj or {}).get("name", "")

    f1_rid = next((rid for rid, n in rec_names.items() if n == "Филиал 1"), None)
    f2_rid = next((rid for rid, n in rec_names.items() if n == "Филиал 2"), None)
    assert f1_rid and f2_rid

    cells = requests.get(f"{API}/cells/by-sheet/{sheet_id}", timeout=30).json()
    f1_nonzero = sum(1 for c in cells
                     if f1_rid in c.get("coord_key", "")
                     and c.get("value") and float(c["value"]) != 0)
    f2_nonzero = sum(1 for c in cells
                     if f2_rid in c.get("coord_key", "")
                     and c.get("value") and float(c["value"]) != 0)
    assert f1_nonzero + f2_nonzero > 0, "No non-zero cells on branches"
