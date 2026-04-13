# Тесты Pebble

## Backend (pytest)
- `test_import_verify.py` — Импорт Excel + полная верификация 20572 ячеек
- `test_api.py` — CRUD моделей, листов, аналитик, пересчёт, экспорт
- `test_permissions.py` — Права доступа на листы и аналитики

## E2E (Playwright)
- `test_e2e.py` — Логин, навигация, ввод данных, UI фичи

## Запуск
```bash
# Backend тесты
pytest tests/ -v

# E2E тесты (нужен запущенный сервер)
npx playwright test tests/test_e2e.py
```
