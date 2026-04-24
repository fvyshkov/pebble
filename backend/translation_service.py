"""Multilingual translation service for Pebble.

Supports four languages: Russian (ru), English (en), Kyrgyz (ky), Vietnamese (vi).
Uses Google Translate (via deep-translator) for fast free translation.
Translations are stored in the `translations` table.
"""
from __future__ import annotations

import json
import os
import re
import uuid
from typing import Sequence

import aiosqlite

from backend.db import get_db

SUPPORTED_LANGS = ("ru", "en", "ky", "vi")
DEFAULT_LANG = "ru"

LANG_NAMES = {"ru": "Russian", "en": "English", "ky": "Kyrgyz", "vi": "Vietnamese"}

# Month / quarter / half-year names per language
MONTH_NAMES = {
    "ru": ["Январь", "Февраль", "Март", "Апрель", "Май", "Июнь",
            "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"],
    "en": ["January", "February", "March", "April", "May", "June",
            "July", "August", "September", "October", "November", "December"],
    "ky": ["Январь", "Февраль", "Март", "Апрель", "Май", "Июнь",
            "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"],
    "vi": ["Tháng 1", "Tháng 2", "Tháng 3", "Tháng 4", "Tháng 5", "Tháng 6",
            "Tháng 7", "Tháng 8", "Tháng 9", "Tháng 10", "Tháng 11", "Tháng 12"],
}

QUARTER_NAMES = {
    "ru": ["1-й квартал", "2-й квартал", "3-й квартал", "4-й квартал"],
    "en": ["Q1", "Q2", "Q3", "Q4"],
    "ky": ["1-чейрек", "2-чейрек", "3-чейрек", "4-чейрек"],
    "vi": ["Quý 1", "Quý 2", "Quý 3", "Quý 4"],
}

HALFYEAR_NAMES = {
    "ru": ["1-е полугодие", "2-е полугодие"],
    "en": ["H1", "H2"],
    "ky": ["1-жарым жылдык", "2-жарым жылдык"],
    "vi": ["Nửa năm 1", "Nửa năm 2"],
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
    "vi": {"name": "Tên", "start": "Bắt đầu", "end": "Kết thúc",
            "unit": "Đơn vị đo", "periods": "Kỳ",
            "indicators": "Chỉ tiêu"},
}


def _detect_source_lang(texts: str | list[str]) -> str:
    """Detect dominant language from text(s) based on Cyrillic character frequency."""
    if isinstance(texts, str):
        texts = [texts]
    # Count texts with Cyrillic vs Latin
    cyrillic_count = sum(1 for t in texts if re.search(r'[а-яА-ЯёЁ]', t))
    return "ru" if cyrillic_count > len(texts) // 2 else "en"


def _build_local_dict() -> dict[str, dict[str, str]]:
    """Build a dictionary of known period/field translations (no API needed)."""
    d: dict[str, dict[str, str]] = {}
    # Months: "Январь 2026" etc.
    for i in range(12):
        for lang_src in SUPPORTED_LANGS:
            name = MONTH_NAMES[lang_src][i]
            tr = {lang: MONTH_NAMES[lang][i] for lang in SUPPORTED_LANGS}
            d[name] = tr
            # With year suffix: "Январь 2026" → "January 2026"
            for year in range(2020, 2041):
                key = f"{name} {year}"
                d[key] = {lang: f"{MONTH_NAMES[lang][i]} {year}" for lang in SUPPORTED_LANGS}
    # Quarters (with and without year)
    for i in range(4):
        for lang_src in SUPPORTED_LANGS:
            d[QUARTER_NAMES[lang_src][i]] = {lang: QUARTER_NAMES[lang][i] for lang in SUPPORTED_LANGS}
            for year in range(2020, 2041):
                key = f"{QUARTER_NAMES[lang_src][i]} {year}"
                d[key] = {lang: f"{QUARTER_NAMES[lang][i]} {year}" for lang in SUPPORTED_LANGS}
    # Half-years (with and without year)
    for i in range(2):
        for lang_src in SUPPORTED_LANGS:
            d[HALFYEAR_NAMES[lang_src][i]] = {lang: HALFYEAR_NAMES[lang][i] for lang in SUPPORTED_LANGS}
            for year in range(2020, 2041):
                key = f"{HALFYEAR_NAMES[lang_src][i]} {year}"
                d[key] = {lang: f"{HALFYEAR_NAMES[lang][i]} {year}" for lang in SUPPORTED_LANGS}
    # Bare years
    for year in range(2020, 2041):
        d[str(year)] = {lang: str(year) for lang in SUPPORTED_LANGS}
    # Field labels
    for lang_src in SUPPORTED_LANGS:
        for field, label in FIELD_LABELS[lang_src].items():
            d[label] = {lang: FIELD_LABELS[lang][field] for lang in SUPPORTED_LANGS}
    return d


_LOCAL_DICT = _build_local_dict()


async def batch_translate(texts: list[str], target_langs: list[str] | None = None) -> dict[str, dict[str, str]]:
    """Translate a batch of texts to all target languages.

    Uses local dictionary for known period/field names, Google Translate for the rest.
    Free, fast (~2s for 600 names), supports ru/en/ky.
    """
    if not texts:
        return {}

    if target_langs is None:
        target_langs = list(SUPPORTED_LANGS)

    unique_texts = list(dict.fromkeys(texts))
    if not unique_texts:
        return {}

    result: dict[str, dict[str, str]] = {}
    needs_translate: list[str] = []

    # Phase 1: resolve locally (periods, years, known labels)
    for t in unique_texts:
        if t in _LOCAL_DICT:
            result[t] = _LOCAL_DICT[t]
        elif t.isdigit() or t.replace(".", "", 1).replace("-", "", 1).isdigit():
            result[t] = {lang: t for lang in target_langs}
        else:
            needs_translate.append(t)

    if not needs_translate:
        return result

    # Phase 2: check DB cache for previously translated texts
    db = get_db()
    still_need: list[str] = []
    if needs_translate:
        placeholders = ",".join("?" * len(needs_translate))
        cached_rows = await db.execute_fetchall(
            f"SELECT source_text, lang, translated FROM translation_cache WHERE source_text IN ({placeholders})",
            needs_translate,
        )
        cache_map: dict[str, dict[str, str]] = {}
        for row in cached_rows:
            cache_map.setdefault(row["source_text"], {})[row["lang"]] = row["translated"]

        for t in needs_translate:
            cached = cache_map.get(t, {})
            # Source lang doesn't need to be cached — it's the original text
            non_src_langs = [lang for lang in target_langs if lang != _detect_source_lang(t)]
            if all(lang in cached for lang in non_src_langs):
                # Fill in source lang + cached translations
                tr = {lang: t for lang in target_langs}  # default all to original
                tr.update(cached)
                result[t] = tr
            else:
                still_need.append(t)

    if not still_need:
        return result

    # Phase 3: translate remaining via Google Translate (free, fast)
    try:
        from deep_translator import GoogleTranslator
    except ImportError:
        print("[translation] deep-translator not installed, keeping originals")
        for t in still_need:
            result[t] = {lang: t for lang in target_langs}
        return result

    src_lang = _detect_source_lang(still_need) if still_need else "ru"

    # Split texts into chunks of ~1000 chars
    chunks: list[list[str]] = []
    current: list[str] = []
    current_len = 0
    for t in still_need:
        if current_len + len(t) + 1 > 1000 and current:
            chunks.append(current)
            current = [t]
            current_len = len(t)
        else:
            current.append(t)
            current_len += len(t) + 1
    if current:
        chunks.append(current)

    api_result: dict[str, dict[str, str]] = {}
    for lang in target_langs:
        if lang == src_lang:
            for t in still_need:
                api_result.setdefault(t, {})[lang] = t
            continue
        try:
            translator = GoogleTranslator(source="auto", target=lang)
            for chunk in chunks:
                combined = "\n".join(chunk)
                translated_text = translator.translate(combined)
                translated_parts = translated_text.split("\n") if translated_text else []
                for i, orig in enumerate(chunk):
                    tr = translated_parts[i].strip() if i < len(translated_parts) else orig
                    api_result.setdefault(orig, {})[lang] = tr or orig
        except Exception as e:
            print(f"[translation] Google Translate {src_lang}→{lang} failed: {e}")
            for t in still_need:
                api_result.setdefault(t, {})[lang] = t

    # Ensure all langs present
    for t in still_need:
        for lang in target_langs:
            if lang not in api_result.get(t, {}):
                api_result.setdefault(t, {})[lang] = t

    result.update(api_result)

    # Phase 4: save new translations to cache (only actually translated ones)
    for t in still_need:
        tr = api_result.get(t, {})
        for lang, value in tr.items():
            if not value or value == t:
                continue  # don't cache untranslated
            try:
                await db.execute(
                    "INSERT OR REPLACE INTO translation_cache (source_text, lang, translated) VALUES (?, ?, ?)",
                    (t, lang, value),
                )
            except Exception:
                pass
    await db.commit()

    # Ensure all texts covered
    for t in unique_texts:
        if t not in result:
            result[t] = {lang: t for lang in target_langs}

    return result


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
