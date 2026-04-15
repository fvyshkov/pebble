# Pebble

Mini-Anaplan: multidimensional planning with pivot tables.

## Установка и запуск

### 1. Установить Python 3.10+

**Windows** (PowerShell):
```powershell
winget install Python.Python.3.12
```

**macOS:**
```bash
brew install python@3.12
```

**Linux (Ubuntu/Debian):**
```bash
sudo apt-get install python3 python3-venv python3-pip
```

### 2. Запустить

```bash
python start.py
```

Скрипт автоматически:
- создаст виртуальное окружение и установит Python-зависимости
- установит Node.js (если нужно), соберёт фронтенд
- запустит сервер и откроет браузер на http://localhost:8000

## Features

- **Models** — containers for sheets and analytics
- **Analytics** — hierarchical dimensions (products, regions, periods, etc.) with custom fields
- **Periods** — auto-generation of year/quarter/month hierarchies
- **Sheets** — bind any analytics, configure order
- **Pivot table** — multidimensional data entry at analytic intersections, hierarchy aggregation, formulas (if/then/else), analytic pinning
- **Copy/paste** — Ctrl+C/V, Excel-compatible, Shift+arrows range selection
- **Excel import/export** — upload/download analytic records with hierarchy
- **Users & permissions** — per-sheet access control (view/edit)
- **Change history** — cell audit log (who/when/what)

## Stack

- **Backend**: Python, FastAPI, SQLite (aiosqlite), openpyxl
- **Frontend**: React 18, TypeScript, Vite, Material UI 5

## Testing

```bash
# Backend (pytest + httpx, in-memory SQLite)
python3 -m venv .venv && source .venv/bin/activate
pip install -r backend/requirements.txt pytest pytest-asyncio httpx
pytest backend/ -v

# Frontend (vitest)
cd frontend
npx vitest run
```
