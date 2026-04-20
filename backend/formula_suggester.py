"""Shared LLM helper: suggest indicator consolidation formulas.

Used both by chat.py (add_analytic_to_all_sheets) and import_excel.py
(post-import period-consolidation pass).
"""
from __future__ import annotations

import json
import os
import uuid


async def suggest_consolidations_for_sheet(
    db, sheet_id: str, new_analytic_name: str
) -> int:
    """Ask Claude for per-indicator consolidation formulas.

    Called when a new analytic axis is added to a sheet (e.g. «Подразделения»)
    or after import with the periods analytic name (so that averages/rates are
    not blindly summed across time periods).

    Writes non-SUM answers as `consolidation`-kind rows into
    `indicator_formula_rules`. Tolerant of Claude/API failures — returns the
    number of rules written (0 on any error or when no API key is set).
    """
    # Main analytic (where indicators live)
    rows = await db.execute_fetchall(
        "SELECT sa.analytic_id FROM sheet_analytics sa "
        "JOIN analytics a ON a.id = sa.analytic_id "
        "WHERE sa.sheet_id = ? AND sa.is_main = 1 AND a.is_periods = 0 LIMIT 1",
        (sheet_id,),
    )
    if not rows:
        return 0
    main_aid = rows[0]["analytic_id"]

    # All indicator records (with names + units)
    recs = await db.execute_fetchall(
        "SELECT id, data_json FROM analytic_records WHERE analytic_id = ?",
        (main_aid,),
    )
    indicators: list[dict] = []
    for r in recs:
        try:
            d = json.loads(r["data_json"]) if isinstance(r["data_json"], str) else (r["data_json"] or {})
        except Exception:
            d = {}
        nm = (d.get("name") or "").strip()
        if nm:
            indicators.append({"id": r["id"], "name": nm, "unit": d.get("unit", "")})
    if not indicators:
        return 0

    # Skip indicators that already have a consolidation rule — don't overwrite.
    existing = await db.execute_fetchall(
        "SELECT indicator_id FROM indicator_formula_rules "
        "WHERE sheet_id = ? AND kind = 'consolidation'",
        (sheet_id,),
    )
    already = {r["indicator_id"] for r in existing}
    todo = [i for i in indicators if i["id"] not in already]
    if not todo:
        return 0

    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        return 0

    try:
        import anthropic
        kwargs: dict = {"api_key": api_key}
        base_url = os.environ.get("ANTHROPIC_BASE_URL")
        if base_url:
            kwargs["base_url"] = base_url
        client = anthropic.Anthropic(**kwargs)

        all_lines = "\n".join(
            f"- {i['name']}" + (f" ({i['unit']})" if i["unit"] else "")
            for i in indicators
        )
        # Use names (not UUIDs) in prompt so cache key is stable across re-imports
        todo_lines = "\n".join(f"- {i['name']}" for i in todo)
        name_to_id = {i["name"]: i["id"] for i in todo}

        prompt = f"""На лист финмодели добавляется разрез «{new_analytic_name}».

Показатели этого листа:
{all_lines}

Для каждого показателя из списка ниже определи, как консолидируется значение \
по этому разрезу на строке-итоге (или при свёртке периодов).

Варианты:
- SUM — обычная сумма (по умолчанию для большинства абсолютных: выручка, количество).
- Формула — в синтаксисе Pebble: `[имя показателя] / [имя другого]`. \
Использовать только имена из списка показателей этого листа. Типичные случаи:
  * среднее/на одного: `[сумма] / [количество]`
  * доля/процент: `[числитель] / [знаменатель]`
  * ставки/коэффициенты/средние: формулой, а НЕ суммой.
  * если показатель называется «средний», «ср.», «на 1 ...», «% ...», «ставка», \
«доля» — почти всегда формула.

Ответь ТОЛЬКО JSON массивом без пояснений:
[{{"name": "<имя>", "kind": "sum"}}, {{"name": "<имя>", "kind": "formula", "formula": "[a] / [b]"}}]

Показатели, по которым нужен ответ:
{todo_lines}
"""

        import asyncio
        from backend.llm_cache import cached_messages_create
        loop = asyncio.get_event_loop()
        resp = await loop.run_in_executor(None, lambda: cached_messages_create(
            client,
            model="claude-sonnet-4-20250514",
            max_tokens=2000,
            messages=[{"role": "user", "content": prompt}],
        ))
        text = "".join(
            b.text for b in resp.content if getattr(b, "type", "") == "text"
        ).strip()
        if text.startswith("```"):
            text = text.strip("`")
            if text.lower().startswith("json"):
                text = text[4:]
            text = text.strip()
        suggestions = json.loads(text)
        if not isinstance(suggestions, list):
            return 0
    except Exception as e:
        print(f"[suggest_consolidations] LLM failed for sheet {sheet_id}: {e}")
        return 0

    written = 0
    for s in suggestions:
        if not isinstance(s, dict):
            continue
        # Support both name-based and legacy id-based responses
        iid = s.get("id") or name_to_id.get(s.get("name", ""))
        if not iid or iid not in name_to_id.values():
            continue
        if s.get("kind") != "formula":
            continue
        formula = (s.get("formula") or "").strip()
        if not formula:
            continue
        await db.execute(
            "INSERT INTO indicator_formula_rules "
            "(id, sheet_id, indicator_id, kind, scope_json, priority, formula) "
            "VALUES (?, ?, ?, 'consolidation', '{}', 0, ?)",
            (str(uuid.uuid4()), sheet_id, iid, formula),
        )
        written += 1
    if written:
        await db.commit()
    return written
