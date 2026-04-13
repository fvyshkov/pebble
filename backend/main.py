from contextlib import asynccontextmanager
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from backend.db import init_db, close_db
from backend.routers import models, analytics, sheets, cells, excel_io, users, import_excel, auth


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


@app.get("/api/health")
async def health():
    return {"status": "ok"}
