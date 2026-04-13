# Тесты Pebble

## Запуск

```bash
# Требования: backend на :8000, models.xlsx в корне проекта
# VERIFIED модель создаётся автоматически если отсутствует

pytest tests/ -v            # все 26 тестов
pytest tests/test_import_verify.py -v   # только верификация
pytest tests/test_hide_model.py::test_hide_sheet_from_user -v  # один тест
```

---

## test_import_verify.py — Верификация импорта (10 тестов)

Поячеечное сравнение импортированной модели с Excel data_only значениями.

| Тест | Что проверяет |
|------|--------------|
| `test_all_7_sheets_present` | Все 7 листов импортированы (0, BaaS.1-3, BS, PL, OPEX) |
| `test_sheets_have_excel_code` | У каждого листа есть код (PL, BS, BaaS.1 и т.д.) |
| `test_sheet_order_matches_excel` | Порядок листов = порядок табов в Excel |
| `test_sheet_cell_values_match_excel[параметры модели]` | Sheet 0: 648/648 = 100% |
| `test_sheet_cell_values_match_excel[кредитование]` | BaaS.1: 4032/4032 = 100% |
| `test_sheet_cell_values_match_excel[депозит]` | BaaS.2: 2480/2480 = 100% |
| `test_sheet_cell_values_match_excel[транзакционный]` | BaaS.3: 4893/4893 = 100% |
| `test_sheet_cell_values_match_excel[Баланс]` | BS: 924/924 = 100% |
| `test_sheet_cell_values_match_excel[Финансовый результат]` | PL: 4500/4500 = 100% |
| `test_sheet_cell_values_match_excel[Операционные расходы]` | OPEX: 3095/3095 = 100% |

Метод: для каждой записи с `excel_row` берёт значение из БД по coord_key (period_rid|record_id) и сравнивает с Excel cell(excel_row, col). Допуск: 1% от значения + 0.01. При дублях excel_row берётся запись с большим числом ячеек.

---

## test_hide_model.py — Скрытие листов от пользователя (4 теста)

| Тест | Что проверяет |
|------|--------------|
| `test_user_sees_all_sheets_by_default` | Новый пользователь видит все 7 листов |
| `test_hide_sheet_from_user` | Скрытие одного листа → 6 из 7 видны |
| `test_hide_all_sheets_hides_model` | Все листы скрыты → модель исчезает из списка |
| `test_read_only_sheet` | can_view=true + can_edit=false → видимый, но нередактируемый |

---

## test_permissions.py — Права доступа (5 тестов)

| Тест | Что проверяет |
|------|--------------|
| `test_list_users` | GET /users возвращает список |
| `test_sheet_permissions` | Установка can_view/can_edit, проверка через accessible-sheets |
| `test_analytic_permissions` | Права на запись аналитики (подразделение) |
| `test_cell_filtering_by_user` | GET /cells?user_id= фильтрует по разрешённым записям |
| `test_accessible_sheets_have_excel_code` | excel_code присутствует в ответе API |

---

## test_z_api.py — CRUD, расчёты, экспорт (7 тестов)

Запускается последним (z-prefix) чтобы не портить данные для verify.

| Тест | Что проверяет |
|------|--------------|
| `test_list_models` | Список моделей |
| `test_create_update_delete_model` | Полный цикл: создание → переименование → каскадное удаление |
| `test_create_sheet` | Создание листа |
| `test_create_analytic` | Создание аналитики |
| `test_recalculate` | Пересчёт формул (>0 ячеек пересчитано) |
| `test_export_model` | Экспорт в xlsx (проверка content-type и размера) |
| `test_cell_save_triggers_recalc` | Сохранение ячейки → автопересчёт → восстановление значения |

---

## Итого: 26 тестов, ~4 секунды

- 10 верификация (20572 ячейки поячеечно × 100%)
- 4 скрытие моделей/листов
- 5 права доступа
- 7 API/CRUD/расчёты
