## Active (in progress / next up)

1. [x] Flash ONLY parent/group rows on recalc (not leaf cells). After recalc-diff updates leaves, recompute sum_children on parents and flash green only for parents whose value changed.
2. [x] Chat tool: `add_analytic_to_all_sheets` + `remove_analytic_from_all_sheets` (+ `find_analytic_by_name`) in `backend/routers/chat.py` — so ИИ делает это сам, не отсылает в UI.
3. [x] Voice auto-send fix: `ChatPanel.tsx` — schedule pause-commit on every `onresult` (not only on `isFinal`), plus watchdog at mic start.
4. [x] Double-space inside chat input should still toggle voice. `App.tsx` — allow double-space when `data-testid="chat-input"` (strip the two spaces from buffer).
5. [x] Rerun Playwright pin tests — all 4 pass (`test_aggrid_pin_single_row_analytic_shows_summary_row`, `test_aggrid_pin_chip_shows_analytic_name_and_value`) after the above fixes.
6. [x] Parent sum flash: `recomputeParentsForField` used `api.getRowNode(path)` but no `getRowId` is defined → node was always null, flash never fired. Fixed by building a path→node index via `forEachNode`. Test `test_aggrid_edit_flashes_parent_sum_green` passes.
7. [x] Default-expand всех листов: `groupDefaultExpanded={-1}`. Новый тест `test_aggrid_groups_expanded_by_default` + переписан прежний expand-test.
8. [x] Дабл-спейс как триггер голосового ввода работает ВЕЗДЕ в приложении — в т.ч. внутри редактора ячейки AG Grid и любых инпутов. В `App.tsx` убрана проверка `isEditable() && !isChatInput()`, и добавлена общая `stripTrailingSpace()` (INPUT/TEXTAREA через native value setter, contenteditable через textContent).

## Next up

- [x] **Баг (подразделения):** formula-ячейка листа теперь принимает ручной ввод (Excel/legacy-параллель): `editable` возвращает true и для `formula`, а `onCellValueChanged` переключает локальный `__rule` на `manual` (и сохраняет с `rule: 'manual'`).
- [x] **СРОЧНО:** Отображение формул в `PivotGridAG`. Добавлен `mode` prop ('data' | 'formulas'), `formulaMapRef` грузится из `CellData.formula`, `valueFormatter` показывает формулу или rule-лейбл (`✎ ввод`, `Σ сумма`, `ƒ <текст>`, `∅`), `cellStyle` переключается на пастельную палитру + left-align, `tooltipValueGetter` всплывает формулу. Mode in `App.tsx` прокидывается через `key`, чтобы форсировать пересборку columnDefs при переключении.
- [x] Чипы с итогами по периодной аналитике (Годы / Кварталы / Месяцы) в `PivotGridAG`: детект уровней по именам записей, чипы над гридом, `colLevelToggles` state + rebuild columnDefs при переключении, `makeSumColDef` вставляет суммирующую колонку после группы на включённом уровне. MVP: пересобирает только period-колонки, сохраняет auto-group колонку первой.
- [x] Кнопки «добавить во все листы» / «удалить из всех листов» — анимировать. `AnalyticSettings.tsx`: `bulkBusy` + `bulkProgress`, лейбл превращается в «Добавляется… 3/12», CircularProgress вместо иконки, `disabled` пока идёт; обе кнопки блокируют друг друга.

## Backlog

- [x] Playwright E2E: resize колонок (`test_aggrid_column_resize_persists_in_dom`) и period-level totals (`test_aggrid_period_totals_toggle_adds_sum_column`). Ввод данных + автопересчёт покрыт `test_aggrid_edit_flashes_parent_sum_green`. Grid API выставлен на `window.__pebbleGridApi` для E2E. Level-чипам добавлен `data-testid="col-level-chip-{level}"`.
- [x] Серверная фильтрация ячеек по правам — E2E тест `test_cell_filtering_actually_hides_records` в `tests/test_permissions.py` реально создаёт grant на 1 запись и проверяет, что `/cells/by-sheet?user_id=...` возвращает строго меньше строк чем baseline.
- [ ] Rule на group-level ячейках (суммы кварталов могут быть формулой или вводом)
- [ ] Когда модели станут большие (>1с расчёт): режим "по запросу" + пометка неактуальных красным

## Active (current batch)

1. [x] **D11 не принимал ручной ввод**: `editable: (p) => rule === 'manual' || rule === 'formula'` работал, но keyboard handler на `onCellKeyDown` (`PivotGridAG.tsx:1192`) стартовал `startEditingCell` только для `rule === 'manual'`. Формулы молча игнорировали ввод. Fix: `if (rule === 'manual' || rule === 'formula')`.
2. [x] **Мягкий numeric parser**: вместо rejectа нечислового ввода `valueParser` теперь вырезает всё кроме цифр/первой точки/ведущего минуса. `"11у12.а12"` → `"1112.12"`. Тест `test_aggrid_lenient_numeric_parser_strips_letters`.
3a. [x] **Вертикальные границы между клетками** в AG Grid. v33 использует Theming API вместо CSS-класса `.ag-theme-alpine`, поэтому старые `border-right` в `App.css` не срабатывали. Создал `themeAlpineWithBorders = themeAlpine.withParams({ columnBorder, rowBorder, headerColumnBorder })` и передаю в `<AgGridReact theme={...} />`.
3. [x] **Σ-колонки (итоги по кварталам/годам) тоже флешат** на recalc. `makeSumColDef` теперь задаёт стабильный `colId` и регистрирует `leafIds` в `sumColLeavesRef`. `recomputeParentsForField` после пересчёта родителей рефрешит и флешит Σ-колонки, чьи `leafIds` содержат изменившийся лист, на самом листе И на всех родительских строках.
4. [x] **Супер-тест**: End-to-end сценарий `tests/test_super_scenario.py::test_super_scenario_permissions_and_aggregation` — админ пишет значения в 4 клетки (Jan × {PL_A,PL_B} × {D12,D13}), дв. user'ам выдаются per-record view-only права (D12/D13), проверяется: a) dep12 не видит ни одной D13-клетки, b) dep13 не видит ни одной D12-клетки, c) 3-part Dep-tagged coord_keys не пересекаются между пользователями, d) каждый видит именно свои значения, e) админ без фильтра видит обе суммы и D1 = D12+D13. End-of-test rollback возвращает исходные значения. Passing.

   Оригинальный план-сценарий целиком через chat-prompts (импорт → create_analytic → set_record_permission → login через UI → typing в клетках) остаётся backlog-ом:
   - импорт модели → `import_excel_from_path` через chat
   - добавление аналитики «Подразделения» → уже есть в модели; иначе `create_model` + API для `analytics` (нет chat tool для создания analytic — написать? **side note**)
   - раздача прав на 2 терминальных record → **нет chat tool**; нужно добавить `set_record_permission(user_id, analytic_id, record_id, can_view, can_edit)` в `backend/routers/chat.py`
   - фиксация аналитики → `pin_analytic` ✓ (есть)
   - логин под 2 юзеров поочередно (без prompts, напрямую через `/api/auth/login`), ввод чисел в клетки (UI typing), проверка: каждый видит только свою запись, show показатели по дефолту
   - админ → сумма корректная
   
   Шаги реализации:
   - [ ] Добавить chat tools: `create_analytic`, `add_record_to_analytic`, `set_record_permission` в `backend/routers/chat.py`
   - [ ] Написать helper `chat_prompt(page, text)` в тестах (шлёт промпт, ждёт ответ, возвращает actions)
   - [ ] Написать `login_as(page, username, password)` helper
   - [ ] Собственно `tests/test_super_scenario.py` с единственным `test_super_scenario_permissions_and_aggregation`
