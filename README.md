# Pebble

Мини-Anaplan: многомерное планирование с pivot-таблицами.

## Возможности

- **Модели** — контейнеры для листов и аналитик
- **Аналитики** — иерархические справочники (продукты, регионы, периоды и т.д.) с произвольными полями
- **Периоды** — автогенерация год/квартал/месяц с иерархией
- **Листы** — привязка любых аналитик, настройка порядка
- **Pivot-таблица** — многомерный ввод данных на пересечении аналитик, суммирование по иерархии, формулы (если/то/иначе), фиксация аналитик
- **Копирование/вставка** — Ctrl+C/V, совместимо с Excel, выделение диапазонов Shift+стрелки
- **Excel импорт/экспорт** — выгрузка/загрузка записей аналитик с иерархией
- **Пользователи и права** — управление доступом к листам (просмотр/редактирование)
- **История изменений** — аудит ячеек с логом кто/когда/что менял

## Стек

- **Backend**: Python, FastAPI, SQLite (aiosqlite), openpyxl
- **Frontend**: React 18, TypeScript, Vite, Material UI 5

## Запуск

```bash
# Backend
cd backend
pip install -r requirements.txt
uvicorn backend.main:app --reload --port 8000

# Frontend
cd frontend
npm install
npm run dev
```

Приложение: http://localhost:5173, API: http://localhost:8000/docs

## Тесты

```bash
# Backend (pytest + httpx, in-memory SQLite)
source .venv/bin/activate  # или создать: python3 -m venv .venv && pip install -r backend/requirements.txt pytest pytest-asyncio httpx
pytest backend/ -v

# Frontend (vitest)
cd frontend
npx vitest run
```
