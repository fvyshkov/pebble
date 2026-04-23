"""API router for translations and language support."""
from __future__ import annotations

from fastapi import APIRouter, Query
from pydantic import BaseModel

from backend.db import get_db
from backend.translation_service import (
    SUPPORTED_LANGS,
    DEFAULT_LANG,
    get_translations,
    save_translations,
    batch_translate,
    MONTH_NAMES,
    QUARTER_NAMES,
    HALFYEAR_NAMES,
    FIELD_LABELS,
)

router = APIRouter(prefix="/api/i18n", tags=["i18n"])


@router.get("/languages")
async def list_languages():
    """Return supported languages."""
    return {
        "languages": [
            {"code": "ru", "name": "Русский", "native_name": "Русский"},
            {"code": "en", "name": "English", "native_name": "English"},
            {"code": "ky", "name": "Кыргызча", "native_name": "Кыргызча"},
        ],
        "default": DEFAULT_LANG,
    }


class TranslationBatch(BaseModel):
    entity_type: str
    entities: list[dict]  # [{id, field, translations: {lang: value}}]


@router.put("/translations")
async def upsert_translations(batch: TranslationBatch):
    """Bulk upsert translations."""
    db = get_db()
    for ent in batch.entities:
        await save_translations(
            entity_type=batch.entity_type,
            entity_id=ent["id"],
            field=ent.get("field", "name"),
            translations=ent["translations"],
            db=db,
        )
    await db.commit()
    return {"ok": True, "count": len(batch.entities)}


@router.get("/translations/{entity_type}")
async def get_entity_translations(
    entity_type: str,
    ids: str = Query(..., description="Comma-separated entity IDs"),
    field: str = "name",
    lang: str | None = None,
):
    """Get translations for entities."""
    id_list = [x.strip() for x in ids.split(",") if x.strip()]
    result = await get_translations(entity_type, id_list, field=field, lang=lang)
    return result


@router.get("/translations/{entity_type}/{entity_id}")
async def get_single_entity_translations(entity_type: str, entity_id: str):
    """Get all translations for a single entity."""
    db = get_db()
    rows = await db.execute_fetchall(
        "SELECT field, lang, value FROM translations WHERE entity_type = ? AND entity_id = ?",
        (entity_type, entity_id),
    )
    result: dict[str, dict[str, str]] = {}
    for row in rows:
        result.setdefault(row["field"], {})[row["lang"]] = row["value"]
    return result


class TranslateRequest(BaseModel):
    texts: list[str]
    target_langs: list[str] | None = None


@router.post("/translate")
async def translate_texts(req: TranslateRequest):
    """Translate texts using Claude API. Returns {text: {lang: translation}}."""
    result = await batch_translate(req.texts, req.target_langs)
    return result


@router.get("/period-names")
async def get_period_names(lang: str = DEFAULT_LANG):
    """Get localized period names (months, quarters, half-years)."""
    l = lang if lang in SUPPORTED_LANGS else DEFAULT_LANG
    return {
        "months": MONTH_NAMES[l],
        "quarters": QUARTER_NAMES[l],
        "halfyears": HALFYEAR_NAMES[l],
    }


@router.get("/labels")
async def get_field_labels(lang: str = DEFAULT_LANG):
    """Get localized common field labels."""
    l = lang if lang in SUPPORTED_LANGS else DEFAULT_LANG
    return FIELD_LABELS[l]


@router.get("/model/{model_id}")
async def get_model_translations(model_id: str, lang: str = DEFAULT_LANG):
    """Get all translations for a model and its children (sheets, analytics, records).
    Returns a flat dict keyed by entity_type:entity_id:field → value.

    Strategy: collect entity IDs first (fast indexed lookups), then batch-fetch
    translations. This avoids slow correlated subqueries on large DBs.
    """
    db = get_db()

    # Collect all entity IDs belonging to this model
    entity_ids: list[str] = [model_id]

    sheet_rows = await db.execute_fetchall(
        "SELECT id FROM sheets WHERE model_id = ?", (model_id,))
    entity_ids.extend(r["id"] for r in sheet_rows)

    analytic_rows = await db.execute_fetchall(
        "SELECT id FROM analytics WHERE model_id = ?", (model_id,))
    analytic_ids = [r["id"] for r in analytic_rows]
    entity_ids.extend(analytic_ids)

    if analytic_ids:
        ph = ",".join("?" * len(analytic_ids))
        rec_rows = await db.execute_fetchall(
            f"SELECT id FROM analytic_records WHERE analytic_id IN ({ph})",
            analytic_ids,
        )
        entity_ids.extend(r["id"] for r in rec_rows)

    if not entity_ids:
        return {}

    # Batch fetch translations for all entity_ids in one query
    ph = ",".join("?" * len(entity_ids))
    rows = await db.execute_fetchall(
        f"""SELECT entity_type, entity_id, field, value FROM translations
            WHERE lang = ? AND entity_id IN ({ph})""",
        [lang] + entity_ids,
    )
    result: dict[str, str] = {}
    for row in rows:
        key = f"{row['entity_type']}:{row['entity_id']}:{row['field']}"
        result[key] = row["value"]
    return result
