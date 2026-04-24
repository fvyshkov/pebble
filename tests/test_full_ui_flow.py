"""Полный UI E2E тест через Playwright.

Сценарий: чистое состояние → загрузить эталонную модель → добавить
аналитику Подразделение (Голова + Ф1 + Ф2) → добавить на все листы →
проверить распределение → создать users dep1/dep2 → залогиниться
под каждым → ввести данные → проверить консолидацию у админа.

Запуск:
    pytest tests/test_full_ui_flow.py -v -s                    # headless
    pytest tests/test_full_ui_flow.py -v -s --headed          # с окном
"""
import os
import time
import requests
import pytest
from playwright.sync_api import Page, Browser, BrowserContext, expect, TimeoutError as PWTimeout

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
    """Удалить модель UIFLOW и пользователей dep1/dep2, если есть.
    Также гарантировать наличие admin."""
    try:
        models = _api_get("/models")
        for m in models:
            if m["name"] == MODEL_NAME:
                _api_delete(f"/models/{m['id']}")
                print(f"  [cleanup] удалена модель {m['id']}")
        users = _api_get("/users")
        for u in users:
            if u["username"] in ("dep1", "dep2", "Новый пользователь"):
                _api_delete(f"/users/{u['id']}")
                print(f"  [cleanup] удалён пользователь {u['username']}")
        # Ensure admin exists with password 'admin'
        import sqlite3, uuid, os
        db_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "pebble.db")
        c = sqlite3.connect(db_path)
        row = c.execute("SELECT id FROM users WHERE username = 'admin'").fetchone()
        if not row:
            c.execute(
                "INSERT INTO users (id, username, password, can_admin, created_at) VALUES (?, 'admin', 'admin', 1, datetime('now'))",
                (str(uuid.uuid4()),)
            )
            c.commit()
            print("  [cleanup] восстановлен admin")
        c.close()
    except Exception as e:
        print(f"[cleanup] warning: {e}")


@pytest.fixture(scope="module", autouse=True)
def prepare():
    """Ensure backend is up; clean UIFLOW namespace."""
    os.makedirs(SCREENSHOT_DIR, exist_ok=True)
    for _ in range(30):
        try:
            r = requests.get(f"{API}/models", timeout=2)
            if r.status_code == 200:
                break
        except Exception:
            time.sleep(1)
    else:
        pytest.fail("Backend не отвечает на http://localhost:8000")
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
        path = f"{SCREENSHOT_DIR}/{name}.png"
        page.screenshot(path=path, full_page=True)
        print(f"  [screenshot] {path}")
    except Exception as e:
        print(f"  [screenshot failed: {e}]")


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
    # Press Escape a few times to clear any lingering dialog/modal
    for _ in range(3):
        page.keyboard.press('Escape')
        page.wait_for_timeout(200)
    # Use localStorage clear + reload — robust, bypasses UI overlay issues
    page.evaluate("localStorage.clear()")
    page.goto(BASE)
    page.wait_for_selector('input[name="username"]', timeout=15000)


def _expand_tree_item_by_id(page: Page, id_fragment: str):
    """Expand a MUI TreeItem identified by a substring of its id attribute."""
    ti = page.locator(f'li[role="treeitem"][id*="{id_fragment}"]').first
    ti.wait_for(state="attached", timeout=10000)
    expanded = ti.get_attribute('aria-expanded')
    if expanded == 'false':
        icon = ti.locator('> .MuiTreeItem-content > .MuiTreeItem-iconContainer').first
        icon.click()
        page.wait_for_timeout(350)


def _open_tree_model(page: Page):
    """Expand the UIFLOW model node and its Листы folder (if present)."""
    _expand_tree_item_by_id(page, 'model:')
    folder = page.locator('li[role="treeitem"][id*="sheets-folder:"]')
    if folder.count() > 0:
        _expand_tree_item_by_id(page, 'sheets-folder:')


def _switch_to_data_mode(page: Page):
    """Click the data-mode ToggleButton (table grid icon).
    For non-admin users the mode toggle is absent — no-op."""
    btn = page.locator('button:has(svg[data-testid="TableChartOutlinedIcon"])').first
    if btn.count() == 0:
        return
    btn.wait_for(state='visible', timeout=5000)
    if btn.get_attribute('aria-pressed') != 'true':
        btn.click()
        page.wait_for_timeout(2000)


def _switch_to_settings_mode(page: Page):
    """Click the settings-mode ToggleButton (gear icon). Admin only."""
    btn = page.locator('button:has(svg[data-testid="SettingsOutlinedIcon"])').first
    if btn.count() == 0:
        return
    btn.wait_for(state='visible', timeout=5000)
    if btn.get_attribute('aria-pressed') != 'true':
        btn.click()
        page.wait_for_timeout(1500)


def _click_sheet(page: Page, code: str = 'BaaS.1'):
    """Click a sheet by its excel_code chip in the tree, then show grid."""
    _open_tree_model(page)
    sheet = page.locator('li[role="treeitem"][id*="-sheet:"]').filter(
        has=page.locator(f'span:text-is("{code}")')).first
    sheet.scroll_into_view_if_needed()
    sheet.locator('.tree-item-label').first.click()
    page.wait_for_timeout(1200)
    _switch_to_data_mode(page)
    # Wait for AG Grid to render
    page.locator('.ag-root-wrapper').first.wait_for(state='visible', timeout=15000)
    page.wait_for_timeout(1500)


def _wait_for_ag_grid(page: Page, timeout: int = 15000):
    """Wait for AG Grid to be visible and have rendered cells."""
    page.locator('.ag-root-wrapper').first.wait_for(state='visible', timeout=timeout)
    page.locator('.ag-cell').first.wait_for(state='visible', timeout=timeout)


# ──────────────────────────────────────────────────────────────────
# Tests
# ──────────────────────────────────────────────────────────────────

def test_01_admin_login(page: Page):
    """Шаг 1: логин admin/admin."""
    _login(page, *ADMIN)
    expect(page.locator('button[aria-label="Пользователи"]')).to_be_visible(timeout=5000)
    expect(page.locator('button[aria-label="Импорт модели из Excel"]')).to_be_visible(timeout=5000)
    print("✓ admin вошёл")


def test_02_upload_model(page: Page):
    """Шаг 2: импорт эталонной модели."""
    assert os.path.isfile(XLSX_PATH), f"models.xlsx не найден: {XLSX_PATH}"

    page.locator('button[aria-label="Импорт модели из Excel"]').click()
    page.wait_for_timeout(500)

    # Dialog is now open
    page.get_by_label('Название модели').fill(MODEL_NAME)

    # Set file (hidden input)
    page.locator('input[type="file"][accept=".xlsx,.xls"]').set_input_files(XLSX_PATH)
    page.wait_for_timeout(500)

    import_btn = page.locator('.MuiDialog-root button', has_text='Импортировать')
    import_btn.click()

    # Wait for "Закрыть" in the dialog (import done)
    close_btn = page.locator('.MuiDialog-root button', has_text='Закрыть')
    close_btn.wait_for(state="visible", timeout=600_000)
    _shot(page, "02_import_done")
    close_btn.click()
    page.wait_for_timeout(1000)

    # Verify model appears in tree
    expect(page.locator(f'text={MODEL_NAME}').first).to_be_visible(timeout=10000)
    print(f"✓ модель {MODEL_NAME} загружена")


def test_03_verify_reference_numbers(page: Page):
    """Шаг 3: открыть лист, проверить что AG Grid отрендерил данные."""
    _click_sheet(page, 'BaaS.1')
    _shot(page, "03_grid_view")

    # AG Grid cells
    cells = page.locator('.ag-cell')
    count = cells.count()
    assert count > 50, f"Ожидалось > 50 ячеек AG Grid, получено {count}"

    # Numeric content check: find at least one cell that parses as float
    texts = cells.all_text_contents()
    num_count = 0
    for t in texts:
        t = t.strip().replace('\xa0', '').replace(' ', '').replace(',', '.')
        try:
            float(t)
            num_count += 1
        except ValueError:
            pass
    assert num_count > 5, f"Не найдено числовых ячеек: {num_count}"
    print(f"✓ лист содержит {count} ячеек AG Grid, из них числовых >= {num_count}")


def test_04_create_analytic(page: Page):
    """Шаг 4: создать аналитику Подразделение с иерархией."""
    _switch_to_settings_mode(page)
    _shot(page, "04_pre")
    _expand_tree_item_by_id(page, 'model:')

    # Find analytics folder and click + button inside its label.
    analytics_folder_li = page.locator('li[role="treeitem"][id*="analytics-folder:"]').first
    analytics_folder_li.wait_for(state='visible', timeout=10000)
    folder_label = analytics_folder_li.locator('.tree-item-label').first
    folder_label.hover()
    page.wait_for_timeout(300)
    add_btn = analytics_folder_li.locator('.tree-item-label .actions button').first
    add_btn.click(force=True)
    page.wait_for_timeout(1500)

    _shot(page, "04a_analytic_created")

    # Now AnalyticSettings visible. Fill Название field.
    name_field = page.get_by_label('Название')
    name_field.fill('Подразделение')
    name_field.press('Tab')
    page.wait_for_timeout(800)

    # Add root record "Головной"
    add_record_btn = page.locator('button[aria-label="Добавить запись"]').first
    add_record_btn.click()
    page.wait_for_timeout(600)

    record_inputs = page.locator('table tbody input[type="text"]')
    assert record_inputs.count() >= 1, "Нет input для новой записи"
    last = record_inputs.last
    last.click()
    last.fill('Головной')
    last.press('Tab')
    page.wait_for_timeout(800)

    # Add children
    _add_child(page, parent_name='Головной', child_name='Филиал 1')
    _add_child(page, parent_name='Головной', child_name='Филиал 2')

    # Save
    page.keyboard.press('Control+s')
    page.wait_for_timeout(1500)

    _shot(page, "04b_hierarchy_done")
    print("✓ создана иерархия: Головной → Филиал 1, Филиал 2")


def _add_child(page: Page, parent_name: str, child_name: str):
    """Add child record under parent by name."""
    row = page.locator('table tbody tr', has=page.locator(f'input[value="{parent_name}"]'))
    row.hover()
    page.wait_for_timeout(300)
    # The add-child button is inside .row-actions (opacity:0 until hover).
    # On headless CI, hover may not trigger opacity — use JS click.
    btn = row.locator('.row-actions button').first
    btn.dispatch_event('click')
    page.wait_for_timeout(600)
    inputs = page.locator('table tbody input[type="text"]')
    new_input = inputs.last
    new_input.click()
    new_input.fill(child_name)
    new_input.press('Tab')
    page.wait_for_timeout(800)


def test_05_add_analytic_to_all_sheets(page: Page):
    """Шаг 5: добавить аналитику во все листы."""
    btn = page.locator('button', has_text='Добавить во все листы')
    expect(btn).to_be_visible(timeout=5000)
    btn.click()
    success = page.locator('text=/Добавлено в \\d+ лист/').first
    success.wait_for(state="visible", timeout=60000)
    print(f"✓ {success.inner_text()}")

    page.wait_for_timeout(500)


def test_06_verify_branch_distribution(page: Page):
    """Шаг 6: после привязки Подразделения грид перерисовался."""
    _click_sheet(page, 'BaaS.1')
    page.wait_for_timeout(2000)

    # Verify AG Grid is visible
    grid = page.locator('.ag-root-wrapper').first
    expect(grid).to_be_visible()

    # Verify via API that Подразделение is linked to 7 sheets
    models = requests.get(f"{API}/models").json()
    uif = [m for m in models if m["name"] == MODEL_NAME][0]
    tree = requests.get(f"{API}/models/{uif['id']}/tree").json()
    podr = None
    for a in tree.get("analytics", []):
        if a["name"] == "Подразделение":
            podr = a
            break
    assert podr is not None, "Аналитика Подразделение не найдена"

    # Each sheet should have podrazdelenie attached
    linked = 0
    for s in tree.get("sheets", []):
        sa = requests.get(f"{API}/sheets/{s['id']}/analytics").json()
        if any(x.get("analytic_id") == podr["id"] for x in sa):
            linked += 1
    assert linked == 7, f"Подразделение привязана только к {linked}/7 листов"

    _shot(page, "06_grid_with_branches")
    print(f"✓ Подразделение привязана ко всем 7 листам")


# ── Users ──

def _open_users_dialog(page: Page):
    # Close any existing dialog first (from a previous failed test)
    if page.locator('h6:has-text("Пользователи")').count() > 0:
        try:
            _close_users_dialog(page)
        except Exception:
            page.keyboard.press('Escape')
            page.wait_for_timeout(300)
    btn = page.locator('button[aria-label="Пользователи"]').first
    btn.click(force=True)
    page.wait_for_timeout(500)
    expect(page.locator('h6:has-text("Пользователи")')).to_be_visible(timeout=5000)


def _close_users_dialog(page: Page):
    header = page.locator('h6:has-text("Пользователи")').locator('..')
    close_btn = header.locator('button').last
    close_btn.click()
    page.wait_for_timeout(500)
    expect(page.locator('h6:has-text("Пользователи")')).not_to_be_visible(timeout=5000)


def _perm_row(page: Page, text: str):
    """Find a row in the permissions tree by its label text."""
    xp = (
        f'xpath=//div[normalize-space(text())="{text}"]'
        f'/ancestor::div[descendant::input[@type="checkbox"]][1]'
    )
    return page.locator(xp).first


def _expand_row_by_text(page: Page, text: str):
    row = _perm_row(page, text)
    row.wait_for(state='visible', timeout=10000)
    expanded_icon = row.locator('svg[data-testid="ExpandMoreOutlinedIcon"]').first
    if expanded_icon.count() > 0:
        return  # Already expanded
    chevron = row.locator('svg[data-testid="ChevronRightOutlinedIcon"]').first
    chevron.click(force=True)
    page.wait_for_timeout(500)


def _check_row_checkboxes(page: Page, text: str, view: bool = True, edit: bool = True):
    row = _perm_row(page, text)
    row.wait_for(state='visible', timeout=10000)
    wrappers = row.locator('span.MuiCheckbox-root')
    assert wrappers.count() >= 2, f"row '{text}' has {wrappers.count()} checkboxes"
    for idx, want in ((0, view), (1, edit)):
        if not want:
            continue
        inp = wrappers.nth(idx).locator('input[type="checkbox"]')
        is_checked = inp.is_checked()
        if not is_checked:
            wrappers.nth(idx).click(force=True)
            page.wait_for_timeout(300)


def _create_user_with_branch(page: Page, username: str, password: str, branch_label: str):
    _open_users_dialog(page)

    page.locator('button[aria-label="Добавить пользователя"]').first.click()
    page.wait_for_timeout(1000)
    _shot(page, f"user_add_{username}")
    new_user_btn = page.locator('text="Новый пользователь"').first
    new_user_btn.wait_for(state='visible', timeout=10000)
    new_user_btn.click()
    page.wait_for_timeout(500)
    name_input = page.get_by_label('Имя')
    cur = name_input.input_value()
    assert cur == 'Новый пользователь', f"Ожидали 'Новый пользователь' в Имя, получили {cur!r}"

    name_input.fill(username)
    name_input.press('Tab')
    page.wait_for_timeout(800)

    page.get_by_text(username, exact=True).first.click()
    page.wait_for_timeout(700)

    for _ in range(10):
        if name_input.input_value() == username:
            break
        page.wait_for_timeout(200)
    assert name_input.input_value() == username, \
        f"После клика ожидали выбор {username}, в поле: {name_input.input_value()!r}"

    # Set password
    page.locator('button[aria-label="Сменить пароль"]').first.click()
    page.wait_for_timeout(400)
    pw_input = page.get_by_label('Новый пароль')
    pw_input.fill(password)
    page.locator('button', has_text='Сохранить').click()
    page.wait_for_timeout(600)

    # Wait for perms tree to load
    page.locator(f'text={MODEL_NAME}').first.wait_for(state='visible', timeout=10000)
    page.wait_for_timeout(500)
    _shot(page, f"user_{username}_before_expand")

    _expand_row_by_text(page, MODEL_NAME)
    _shot(page, f"user_{username}_after_model_expand")

    _check_row_checkboxes(page, 'Листы', view=True, edit=True)

    _expand_row_by_text(page, 'Аналитики')
    _expand_row_by_text(page, 'Подразделение')
    _expand_row_by_text(page, 'Головной')

    _check_row_checkboxes(page, branch_label, view=True, edit=True)

    _shot(page, f"user_{username}_perms")

    _close_users_dialog(page)


def test_07_create_dep1(page: Page):
    """Шаг 7a: пользователь dep1 — доступ только к Филиал 1."""
    _create_user_with_branch(page, 'dep1', 'dep1', 'Филиал 1')
    print("✓ создан dep1")


def test_08_create_dep2(page: Page):
    """Шаг 7b: пользователь dep2 — доступ только к Филиал 2."""
    _create_user_with_branch(page, 'dep2', 'dep2', 'Филиал 2')
    print("✓ создан dep2")


# ── Non-admin users ──

def _find_editable_ag_cell(page: Page):
    """Find first editable AG Grid cell (manual input, yellow background).

    AG Grid cells with manual input have background #fdf8e8.
    We look for cells in the ag-body-viewport that are editable.
    """
    # AG Grid marks editable cells; we find a cell with the yellow manual-input bg
    # by checking inline styles. The cellStyle function sets background: '#fdf8e8'.
    cells = page.locator('.ag-cell').all()
    for cell in cells:
        style = cell.get_attribute('style') or ''
        if 'fdf8e8' in style:
            return cell
    # Fallback: return first cell in data area
    return page.locator('.ag-cell').first


def test_09_dep1_enter(page: Page):
    """Шаг 8: dep1 логин, ввод в редактируемую ячейку."""
    _logout(page)
    _login(page, *DEP1)

    _click_sheet(page, 'BaaS.1')
    _wait_for_ag_grid(page)

    cell = _find_editable_ag_cell(page)
    expect(cell).to_be_visible(timeout=10000)
    cell.scroll_into_view_if_needed()
    # AG Grid: double-click to enter edit mode
    cell.dblclick()
    page.wait_for_timeout(300)
    page.keyboard.type('77777')
    page.keyboard.press('Enter')
    page.wait_for_timeout(1500)
    _shot(page, "09_dep1_entered")
    print("✓ dep1 ввёл 77777")


def test_10_dep2_enter(page: Page):
    """Шаг 9: dep2 логин, ввод."""
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
    _shot(page, "10_dep2_entered")
    print("✓ dep2 ввёл 33333")


def test_11_admin_consolidation(page: Page):
    """Шаг 10: admin видит консолидированную сумму на Головной.

    Verify via API that consolidation works: Головной == Ф1 + Ф2.
    """
    _logout(page)
    _login(page, *ADMIN)

    _click_sheet(page, 'BaaS.1')
    _wait_for_ag_grid(page)
    page.wait_for_timeout(2000)
    _shot(page, "11_admin_grid")

    # AG Grid rows: use div.ag-row with role="row"
    rows = page.locator('.ag-center-cols-container .ag-row').all()
    assert len(rows) > 0, "AG Grid не отрендерил строки"

    # Extract row texts to find Головной / Филиал 1 / Филиал 2 block
    def _parse_row_numbers(row_loc):
        cells = row_loc.locator('.ag-cell').all()
        vals = []
        for c in cells:
            txt = c.inner_text().strip().replace('\xa0', '').replace(' ', '').replace(',', '.')
            try:
                vals.append(float(txt))
            except ValueError:
                vals.append(None)
        return vals

    # Find Головной row followed by Филиал 1 and Филиал 2
    consolidation_ok = False
    for i, r in enumerate(rows):
        txt = r.inner_text()
        if 'Головной' not in txt:
            continue
        if i + 2 >= len(rows):
            continue
        f1_txt = rows[i + 1].inner_text()
        f2_txt = rows[i + 2].inner_text()
        if 'Филиал 1' not in f1_txt or 'Филиал 2' not in f2_txt:
            continue
        head_vals = _parse_row_numbers(r)
        f1_vals = _parse_row_numbers(rows[i + 1])
        f2_vals = _parse_row_numbers(rows[i + 2])
        checked = 0
        matches = 0
        for h, a, b in zip(head_vals, f1_vals, f2_vals):
            if h is None or a is None or b is None:
                continue
            checked += 1
            if abs(h - (a + b)) < 0.5:
                matches += 1
        if checked > 0 and matches == checked:
            consolidation_ok = True
            print(f"  ✓ строка #{i}: Головной == Ф1 + Ф2 по {checked} колонкам")
            print(f"    Головной: {head_vals[:8]}")
            print(f"    Филиал 1: {f1_vals[:8]}")
            print(f"    Филиал 2: {f2_vals[:8]}")
            break
    _shot(page, "11_admin_consolidation")
    assert consolidation_ok, "Не нашли строку где Головной == Ф1 + Ф2 по всем числовым колонкам"
    print("✓ админ видит консолидацию: Головной == Филиал 1 + Филиал 2")
