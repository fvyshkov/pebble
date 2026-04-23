"""Multilingual translation service for Pebble.

Supports three languages: Russian (ru), English (en), Kyrgyz (ky).
Uses Claude API to detect language and translate to the other two.
Translations are stored in the `translations` table.
"""
from __future__ import annotations

import json
import os
import uuid
from typing import Sequence

import aiosqlite
import anthropic

from backend.db import get_db

SUPPORTED_LANGS = ("ru", "en", "ky")
DEFAULT_LANG = "ru"

LANG_NAMES = {"ru": "Russian", "en": "English", "ky": "Kyrgyz"}

# Month / quarter / half-year names per language
MONTH_NAMES = {
    "ru": ["Январь", "Февраль", "Март", "Апрель", "Май", "Июнь",
            "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"],
    "en": ["January", "February", "March", "April", "May", "June",
            "July", "August", "September", "October", "November", "December"],
    "ky": ["Январь", "Февраль", "Март", "Апрель", "Май", "Июнь",
            "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"],
}

QUARTER_NAMES = {
    "ru": ["1-й квартал", "2-й квартал", "3-й квартал", "4-й квартал"],
    "en": ["Q1", "Q2", "Q3", "Q4"],
    "ky": ["1-чейрек", "2-чейрек", "3-чейрек", "4-чейрек"],
}

HALFYEAR_NAMES = {
    "ru": ["1-е полугодие", "2-е полугодие"],
    "en": ["H1", "H2"],
    "ky": ["1-жарым жылдык", "2-жарым жылдык"],
}

# Common field labels per language
FIELD_LABELS = {
    "ru": {"name": "Наименование", "start": "Начало", "end": "Окончание",
            "unit": "Единица измерения", "periods": "Периоды",
            "indicators": "Показатели"},
    "en": {"name": "Name", "start": "Start", "end": "End",
            "unit": "Unit of measurement", "periods": "Periods",
            "indicators": "Indicators"},
    "ky": {"name": "Аталышы", "start": "Башталышы", "end": "Аякталышы",
            "unit": "Өлчөө бирдиги", "periods": "Мезгилдер",
            "indicators": "Көрсөткүчтөр"},
}


def _get_client() -> anthropic.AsyncAnthropic:
    return anthropic.AsyncAnthropic(
        api_key=os.environ.get("ANTHROPIC_API_KEY", "")
    )


async def batch_translate(texts: list[str], target_langs: list[str] | None = None) -> dict[str, dict[str, str]]:
    """Translate a batch of texts to all target languages.

    Args:
        texts: list of strings to translate
        target_langs: if None, translates to all SUPPORTED_LANGS

    Returns:
        {original_text: {lang: translated_text, ...}, ...}
    """
    if not texts:
        return {}

    if target_langs is None:
        target_langs = list(SUPPORTED_LANGS)

    # Deduplicate
    unique_texts = list(dict.fromkeys(texts))
    if not unique_texts:
        return {}

    client = _get_client()

    prompt = f"""You are a professional translator. Translate each text below into the requested languages.
The texts are names of financial indicators, analytics categories, and similar business terms.

IMPORTANT RULES:
- Detect the source language of each text automatically
- If text is already in a target language, keep it as-is
- Keep numbers, abbreviations, and proper nouns unchanged
- Translations should be natural and professional
- For Kyrgyz (ky): use standard Kyrgyz terminology for financial/business terms

Target languages: {', '.join(f'{l} ({LANG_NAMES[l]})' for l in target_langs)}

Texts to translate (one per line, numbered):
{chr(10).join(f'{i+1}. {t}' for i, t in enumerate(unique_texts))}

Respond with ONLY a JSON object. Keys are the original texts (exactly as given), values are objects mapping language code to translation.
Example: {{"Выручка": {{"ru": "Выручка", "en": "Revenue", "ky": "Киреше"}}}}
"""

    try:
        resp = await client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=4096,
            messages=[{"role": "user", "content": prompt}],
        )
        text = resp.content[0].text.strip()
        # Extract JSON from response (may be wrapped in ```json ... ```)
        if text.startswith("```"):
            text = text.split("\n", 1)[1]
            if text.endswith("```"):
                text = text[: text.rfind("```")]
        result = json.loads(text)
        # Ensure all original texts are in result
        for t in unique_texts:
            if t not in result:
                result[t] = {lang: t for lang in target_langs}
        return result
    except Exception as e:
        print(f"[translation] batch_translate failed: {e}")
        # Fallback: return originals for all langs
        return {t: {lang: t for lang in target_langs} for t in unique_texts}


async def save_translations(
    entity_type: str,
    entity_id: str,
    field: str,
    translations: dict[str, str],
    db: aiosqlite.Connection | None = None,
) -> None:
    """Save translations for an entity field.

    Args:
        entity_type: 'model' | 'analytic' | 'sheet' | 'analytic_record'
        entity_id: the entity's ID
        field: which field ('name', 'description', etc.)
        translations: {lang: value, ...}
    """
    if db is None:
        db = get_db()
    for lang, value in translations.items():
        if lang not in SUPPORTED_LANGS:
            continue
        tid = str(uuid.uuid4())
        await db.execute(
            """INSERT INTO translations (id, entity_type, entity_id, field, lang, value)
               VALUES (?, ?, ?, ?, ?, ?)
               ON CONFLICT(entity_type, entity_id, field, lang)
               DO UPDATE SET value = excluded.value""",
            (tid, entity_type, entity_id, field, lang, value),
        )


async def get_translations(
    entity_type: str,
    entity_ids: list[str],
    field: str = "name",
    lang: str | None = None,
    db: aiosqlite.Connection | None = None,
) -> dict[str, dict[str, str]]:
    """Get translations for multiple entities.

    Returns:
        {entity_id: {lang: value, ...}, ...}
    """
    if db is None:
        db = get_db()
    if not entity_ids:
        return {}

    placeholders = ",".join("?" * len(entity_ids))
    params: list = [entity_type, field] + entity_ids
    sql = f"""SELECT entity_id, lang, value FROM translations
              WHERE entity_type = ? AND field = ? AND entity_id IN ({placeholders})"""
    if lang:
        sql += " AND lang = ?"
        params.append(lang)

    rows = await db.execute_fetchall(sql, params)
    result: dict[str, dict[str, str]] = {}
    for row in rows:
        eid = row["entity_id"]
        result.setdefault(eid, {})[row["lang"]] = row["value"]
    return result


async def get_translated_name(
    entity_type: str,
    entity_id: str,
    lang: str,
    fallback: str = "",
    db: aiosqlite.Connection | None = None,
) -> str:
    """Get a single translated name, falling back to the original."""
    if db is None:
        db = get_db()
    row = await db.execute_fetchall(
        """SELECT value FROM translations
           WHERE entity_type = ? AND entity_id = ? AND field = 'name' AND lang = ?""",
        (entity_type, entity_id, lang),
    )
    if row:
        return row[0]["value"]
    return fallback


async def delete_entity_translations(
    entity_type: str,
    entity_id: str,
    db: aiosqlite.Connection | None = None,
) -> None:
    """Delete all translations for an entity (used when entity is deleted)."""
    if db is None:
        db = get_db()
    await db.execute(
        "DELETE FROM translations WHERE entity_type = ? AND entity_id = ?",
        (entity_type, entity_id),
    )
