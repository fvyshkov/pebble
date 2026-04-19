import os
from pathlib import Path
from contextlib import asynccontextmanager
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse
from backend.db import init_db, close_db
from backend.routers import models, analytics, sheets, cells, excel_io, users, import_excel, auth, chat, indicator_rules


@asynccontextmanager
async def lifespan(app: FastAPI):
    await init_db()
    yield
    await close_db()


app = FastAPI(title="Pebble", lifespan=lifespan)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

app.include_router(models.router)
app.include_router(analytics.router)
app.include_router(sheets.router)
app.include_router(cells.router)
app.include_router(excel_io.router)
app.include_router(users.router)
app.include_router(import_excel.router)
app.include_router(auth.router)
app.include_router(chat.router)
app.include_router(indicator_rules.router)


@app.get("/api/health")
async def health():
    return {"status": "ok"}


# Serve frontend static files (production build)
_dist = Path(__file__).resolve().parent.parent / "frontend" / "dist"
if _dist.is_dir():
    app.mount("/assets", StaticFiles(directory=str(_dist / "assets")), name="static")

    @app.get("/{full_path:path}")
    async def serve_spa(full_path: str):
        # Try static file first, fallback to index.html (SPA routing)
        file = _dist / full_path
        if file.is_file():
            return FileResponse(str(file))
        return FileResponse(str(_dist / "index.html"))
