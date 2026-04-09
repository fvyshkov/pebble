# Pebble

Mini-Anaplan: multidimensional planning with pivot tables.

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

## Running

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

App: http://localhost:5173, API docs: http://localhost:8000/docs

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
