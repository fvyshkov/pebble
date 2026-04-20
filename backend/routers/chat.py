"""AI chat — proxies to Claude with tool-use to perform actions in the app.

Tools that only need the database run server-side (list_models, create_model,
set_cell, recalc, ...). Tools that change UI state (open_sheet, pin_analytic)
are returned as `client_actions` for the frontend to execute after display.
"""
import os
import uuid
import json
from typing import Any
from fastapi import APIRouter, HTTPException
from pydantic import BaseModel
from backend.db import get_db


router = APIRouter(prefix="/api/chat", tags=["chat"])


class ChatMessage(BaseModel):
    role: str  # "user" | "assistant"
    content: Any  # string (user text) or list of blocks (assistant/tool results)


class ChatContext(BaseModel):
    current_model_id: str | None = None
    current_sheet_id: str | None = None
    user_id: str | None = None


class ChatRequest(BaseModel):
    messages: list[ChatMessage]
    context: ChatContext = ChatContext()


SYSTEM_PROMPT = """Ты помощник в приложении Pebble — финансовом моделировщике (аналог Excel с pivot-таблицами).

Ты умеешь:
- Отвечать на общие вопросы о Pebble и финансовом моделировании.
- Выполнять действия с моделями: перечислять, создавать, открывать, импортировать Excel.
- Работать с листами: перечислять, открывать, фиксировать/снимать аналитики, вводить значения.
- Добавлять/убирать аналитики сразу со всех листов модели.
- Запускать пересчёт формул.
- Управлять правами пользователей на записи аналитики: `list_users`, `list_analytic_records` (найти id записи по имени), `set_record_permission(user_id, analytic_id, record_id, can_view, can_edit)`. Когда хотя бы одно разрешение для пары (user, analytic) задано — пользователь видит только разрешённые записи этой аналитики.

Важно: когда пользователь просит "добавь/убери аналитику X со всех листов" — НЕ отсылай его в UI; найди analytic_id через list_model_analytics и вызови add_analytic_to_all_sheets / remove_analytic_from_all_sheets сам.

Контекст пользователя передаётся в каждом запросе (текущая модель, текущий лист, id пользователя).
Когда пользователь просит сделать что-то в приложении, используй инструменты. Если нужен идентификатор (model_id, sheet_id), сначала вызови list_models или list_sheets чтобы его узнать.

Отвечай коротко и по-русски. После выполнения действий — одно-два предложения подтверждения."""


TOOLS: list[dict] = [
    {
        "name": "list_models",
        "description": "Получить список всех моделей. Возвращает массив {id, name, description}.",
        "input_schema": {"type": "object", "properties": {}, "required": []},
    },
    {
        "name": "create_model",
        "description": "Создать новую (пустую) модель.",
        "input_schema": {
            "type": "object",
            "properties": {
                "name": {"type": "string", "description": "Название модели"},
                "description": {"type": "string", "description": "Описание (опционально)"},
            },
            "required": ["name"],
        },
    },
    {
        "name": "list_sheets",
        "description": "Получить список листов указанной модели.",
        "input_schema": {
            "type": "object",
            "properties": {"model_id": {"type": "string"}},
            "required": ["model_id"],
        },
    },
    {
        "name": "list_analytics",
        "description": "Получить аналитики листа (колонки + строки).",
        "input_schema": {
            "type": "object",
            "properties": {"sheet_id": {"type": "string"}},
            "required": ["sheet_id"],
        },
    },
    {
        "name": "read_cell",
        "description": "Прочитать значение ячейки по её координатному ключу.",
        "input_schema": {
            "type": "object",
            "properties": {
                "sheet_id": {"type": "string"},
                "coord_key": {"type": "string", "description": "Ключ вида 'analytic1=rec1;analytic2=rec2'"},
            },
            "required": ["sheet_id", "coord_key"],
        },
    },
    {
        "name": "set_cell",
        "description": "Записать значение в ячейку (manual rule).",
        "input_schema": {
            "type": "object",
            "properties": {
                "sheet_id": {"type": "string"},
                "coord_key": {"type": "string"},
                "value": {"type": "string"},
                "data_type": {"type": "string", "description": "number/string/currency/percent"},
            },
            "required": ["sheet_id", "coord_key", "value"],
        },
    },
    {
        "name": "recalc",
        "description": "Пересчитать все формулы модели. Возвращает когда готово.",
        "input_schema": {
            "type": "object",
            "properties": {"model_id": {"type": "string"}},
            "required": ["model_id"],
        },
    },
    {
        "name": "open_sheet",
        "description": "Открыть лист в интерфейсе (клиентское действие — после ответа).",
        "input_schema": {
            "type": "object",
            "properties": {
                "model_id": {"type": "string"},
                "sheet_id": {"type": "string"},
            },
            "required": ["model_id", "sheet_id"],
        },
    },
    {
        "name": "switch_mode",
        "description": "Переключить режим работы: 'settings' (настройки модели), 'data' (ввод данных), 'formulas' (формулы).",
        "input_schema": {
            "type": "object",
            "properties": {"mode": {"type": "string", "enum": ["settings", "data", "formulas"]}},
            "required": ["mode"],
        },
    },
    {
        "name": "pin_analytic",
        "description": "Зафиксировать аналитику на конкретной записи (скрывает из дерева строк).",
        "input_schema": {
            "type": "object",
            "properties": {
                "analytic_id": {"type": "string"},
                "record_id": {"type": "string"},
            },
            "required": ["analytic_id", "record_id"],
        },
    },
    {
        "name": "unpin_analytic",
        "description": "Снять фиксацию с аналитики.",
        "input_schema": {
            "type": "object",
            "properties": {"analytic_id": {"type": "string"}},
            "required": ["analytic_id"],
        },
    },
    {
        "name": "list_excel_in_folder",
        "description": (
            "Найти все Excel-файлы (.xlsx / .xls) в указанной папке на "
            "локальной машине. Путь может начинаться с '~'. Возвращает "
            "массив объектов {path, name, size, mtime}. Если файлов несколько, "
            "ВЫЗОВИ этот инструмент перед import_excel_from_path и уточни у "
            "пользователя, какой именно импортировать (или все подряд)."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "folder_path": {"type": "string", "description": "Путь к папке"},
            },
            "required": ["folder_path"],
        },
    },
    {
        "name": "import_excel_from_path",
        "description": (
            "Импортировать модель из Excel-файла, лежащего на локальной машине. "
            "Используй ПОЛНЫЙ путь, полученный из list_excel_in_folder."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "file_path": {"type": "string"},
                "model_name": {"type": "string", "description": "Имя новой модели (по умолчанию — имя файла)"},
            },
            "required": ["file_path"],
        },
    },
    {
        "name": "list_model_analytics",
        "description": (
            "Получить все аналитики модели (не привязанные к конкретному листу). "
            "Возвращает массив {id, name, code}. Используй чтобы найти analytic_id "
            "по имени аналитики."
        ),
        "input_schema": {
            "type": "object",
            "properties": {"model_id": {"type": "string"}},
            "required": ["model_id"],
        },
    },
    {
        "name": "add_analytic_to_all_sheets",
        "description": (
            "Добавить аналитику во ВСЕ листы модели (идемпотентно — листы, где "
            "она уже есть, пропускаются). Возвращает {added, skipped}."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "model_id": {"type": "string"},
                "analytic_id": {"type": "string"},
            },
            "required": ["model_id", "analytic_id"],
        },
    },
    {
        "name": "remove_analytic_from_all_sheets",
        "description": "Убрать аналитику со ВСЕХ листов модели. Возвращает {removed}.",
        "input_schema": {
            "type": "object",
            "properties": {
                "model_id": {"type": "string"},
                "analytic_id": {"type": "string"},
            },
            "required": ["model_id", "analytic_id"],
        },
    },
    {
        "name": "fill_sheet",
        "description": (
            "Заполнить ВСЕ ячейки листа (декартово произведение листовых записей "
            "всех аналитик). mode='value' пишет константу value; mode='random' — "
            "случайные целые в диапазоне [min, max]. Только ячейки с правилом "
            "manual (и новые) перезаписываются; формулы и суммы не трогаются."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "sheet_id": {"type": "string"},
                "mode": {"type": "string", "enum": ["value", "random"]},
                "value": {"type": "string", "description": "Константа для mode=value"},
                "min": {"type": "number", "description": "Минимум для mode=random (по умолчанию 1)"},
                "max": {"type": "number", "description": "Максимум для mode=random (по умолчанию 100)"},
            },
            "required": ["sheet_id", "mode"],
        },
    },
    {
        "name": "list_users",
        "description": "Получить список всех пользователей. Возвращает [{id, username, can_admin}].",
        "input_schema": {"type": "object", "properties": {}, "required": []},
    },
    {
        "name": "list_analytic_records",
        "description": (
            "Получить плоский список записей аналитики с именами и parent_id. "
            "Нужно для поиска нужной терминальной записи по имени (D11, D12 и т.п.). "
            "Возвращает [{id, name, parent_id, has_children}]."
        ),
        "input_schema": {
            "type": "object",
            "properties": {"analytic_id": {"type": "string"}},
            "required": ["analytic_id"],
        },
    },
    {
        "name": "set_record_permission",
        "description": (
            "Установить право пользователя на конкретную запись аналитики. "
            "Когда установлено хотя бы одно разрешение для пары (user, analytic), "
            "пользователь видит ТОЛЬКО разрешённые записи этой аналитики во всех "
            "листах. Идемпотентно — можно вызывать многократно."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "user_id": {"type": "string"},
                "analytic_id": {"type": "string"},
                "record_id": {"type": "string"},
                "can_view": {"type": "boolean"},
                "can_edit": {"type": "boolean"},
            },
            "required": ["user_id", "analytic_id", "record_id"],
        },
    },
]


# ── Shared sheet-fill implementation (used by chat tool + HTTP endpoint) ──

async def _fill_sheet_impl(
    db, *, sheet_id: str, mode: str, value=None, vmin: float = 1, vmax: float = 100,
    user_id: str | None = None,
) -> dict:
    """Fill every cartesian-leaf cell of a sheet.

    Skips cells whose stored rule is NOT 'manual' (formulas / sum_children).
    For new cells (no row yet) inserts with rule='manual'.
    Returns {ok, cells_written, skipped_non_manual}.
    """
    import random
    # Binding order
    sa_rows = await db.execute_fetchall(
        "SELECT analytic_id FROM sheet_analytics WHERE sheet_id = ? ORDER BY sort_order",
        (sheet_id,),
    )
    if not sa_rows:
        return {"error": "sheet has no analytics"}
    analytic_ids = [r["analytic_id"] for r in sa_rows]
    # Leaf records per analytic (records that are NOT anyone's parent)
    leaves_by_a: dict[str, list[str]] = {}
    for aid in analytic_ids:
        recs = await db.execute_fetchall(
            "SELECT id, parent_id FROM analytic_records WHERE analytic_id = ?", (aid,),
        )
        parent_set = {r["parent_id"] for r in recs if r["parent_id"]}
        leaves_by_a[aid] = [r["id"] for r in recs if r["id"] not in parent_set]
    for aid in analytic_ids:
        if not leaves_by_a[aid]:
            return {"error": f"analytic {aid} has no leaf records"}
    # Cartesian product
    combos: list[list[str]] = [[]]
    for aid in analytic_ids:
        combos = [c + [lid] for c in combos for lid in leaves_by_a[aid]]
    # Existing cells keyed by coord_key for rule/id lookup
    existing_rows = await db.execute_fetchall(
        "SELECT id, coord_key, rule FROM cell_data WHERE sheet_id = ?", (sheet_id,),
    )
    existing = {r["coord_key"]: (r["id"], r["rule"]) for r in existing_rows}
    cells_written = 0
    skipped = 0
    for combo in combos:
        coord_key = "|".join(combo)
        prev = existing.get(coord_key)
        if prev and prev[1] != "manual":
            skipped += 1
            continue
        if mode == "random":
            v = str(random.randint(int(vmin), int(vmax)))
        else:
            v = "" if value is None else str(value)
        if prev:
            await db.execute(
                "UPDATE cell_data SET value = ?, data_type = 'number' WHERE id = ?",
                (v, prev[0]),
            )
        else:
            await db.execute(
                """INSERT INTO cell_data (id, sheet_id, coord_key, value, data_type, rule, formula)
                   VALUES (?, ?, ?, ?, 'number', 'manual', '')""",
                (str(uuid.uuid4()), sheet_id, coord_key, v),
            )
        cells_written += 1
    await db.commit()
    return {"ok": True, "cells_written": cells_written, "skipped_non_manual": skipped}


class FillSheetRequest(BaseModel):
    mode: str = "random"  # "random" | "value"
    value: str | None = None
    min: float = 1
    max: float = 100
    user_id: str | None = None


@router.post("/fill_sheet/{sheet_id}")
async def fill_sheet_direct(sheet_id: str, req: FillSheetRequest):
    """Non-LLM endpoint for filling every manual cell on a sheet.

    Useful for scripting and for tests that don't want to hit Anthropic.
    """
    db = get_db()
    result = await _fill_sheet_impl(
        db, sheet_id=sheet_id, mode=req.mode, value=req.value,
        vmin=req.min, vmax=req.max, user_id=req.user_id,
    )
    if result.get("error"):
        raise HTTPException(status_code=400, detail=result["error"])
    return result


# ── LLM helper: auto-detect consolidation formulas on add-analytic ────────
# Implementation lives in backend.formula_suggester (shared with import_excel).
from backend.formula_suggester import suggest_consolidations_for_sheet as _suggest_consolidations_for_sheet
from backend.formula_suggester import propagate_consolidations_across_sheets as _propagate_consolidations


# ── Server-side tool execution ─────────────────────────────────────────────

async def _exec_tool(name: str, inp: dict, ctx: ChatContext, client_actions: list[dict]) -> str:
    """Execute a tool call and return a string result for Claude.

    Client-side actions are appended to client_actions list; the tool returns
    an acknowledgement. Data tools hit the DB directly.
    """
    db = get_db()
    try:
        if name == "list_models":
            rows = await db.execute_fetchall(
                "SELECT id, name, description FROM models ORDER BY created_at"
            )
            return json.dumps([dict(r) for r in rows], ensure_ascii=False)

        if name == "create_model":
            mid = str(uuid.uuid4())
            await db.execute(
                "INSERT INTO models (id, name, description) VALUES (?, ?, ?)",
                (mid, inp.get("name", ""), inp.get("description", "")),
            )
            await db.commit()
            return json.dumps({"id": mid, "name": inp.get("name")}, ensure_ascii=False)

        if name == "list_sheets":
            rows = await db.execute_fetchall(
                "SELECT id, name, excel_code FROM sheets WHERE model_id = ? ORDER BY created_at",
                (inp["model_id"],),
            )
            return json.dumps([dict(r) for r in rows], ensure_ascii=False)

        if name == "list_analytics":
            rows = await db.execute_fetchall(
                """SELECT sa.analytic_id, sa.sort_order, a.name
                   FROM sheet_analytics sa JOIN analytics a ON a.id = sa.analytic_id
                   WHERE sa.sheet_id = ? ORDER BY sa.sort_order""",
                (inp["sheet_id"],),
            )
            return json.dumps([dict(r) for r in rows], ensure_ascii=False)

        if name == "read_cell":
            rows = await db.execute_fetchall(
                "SELECT value FROM cells WHERE sheet_id = ? AND coord_key = ?",
                (inp["sheet_id"], inp["coord_key"]),
            )
            return json.dumps({"value": rows[0]["value"] if rows else None}, ensure_ascii=False)

        if name == "set_cell":
            # Upsert manual cell value; use existing cells endpoint logic
            existing = await db.execute_fetchall(
                "SELECT id FROM cells WHERE sheet_id = ? AND coord_key = ?",
                (inp["sheet_id"], inp["coord_key"]),
            )
            if existing:
                await db.execute(
                    "UPDATE cells SET value = ?, data_type = COALESCE(?, data_type), user_id = ? WHERE id = ?",
                    (inp["value"], inp.get("data_type"), ctx.user_id, existing[0]["id"]),
                )
            else:
                await db.execute(
                    """INSERT INTO cells (id, sheet_id, coord_key, value, data_type, rule, user_id)
                       VALUES (?, ?, ?, ?, ?, 'manual', ?)""",
                    (str(uuid.uuid4()), inp["sheet_id"], inp["coord_key"],
                     inp["value"], inp.get("data_type", "number"), ctx.user_id),
                )
            await db.commit()
            # Let the frontend know so it can reload
            client_actions.append({"type": "reload_sheet", "sheet_id": inp["sheet_id"]})
            return json.dumps({"ok": True}, ensure_ascii=False)

        if name == "recalc":
            from backend.formula_engine import calculate_model
            await calculate_model(db, inp["model_id"])
            client_actions.append({"type": "reload_model", "model_id": inp["model_id"]})
            return json.dumps({"ok": True}, ensure_ascii=False)

        if name == "open_sheet":
            client_actions.append({
                "type": "open_sheet",
                "model_id": inp["model_id"],
                "sheet_id": inp["sheet_id"],
            })
            return json.dumps({"ok": True}, ensure_ascii=False)

        if name == "switch_mode":
            client_actions.append({"type": "switch_mode", "mode": inp["mode"]})
            return json.dumps({"ok": True}, ensure_ascii=False)

        if name == "pin_analytic":
            client_actions.append({
                "type": "pin_analytic",
                "analytic_id": inp["analytic_id"],
                "record_id": inp["record_id"],
            })
            return json.dumps({"ok": True}, ensure_ascii=False)

        if name == "unpin_analytic":
            client_actions.append({"type": "unpin_analytic", "analytic_id": inp["analytic_id"]})
            return json.dumps({"ok": True}, ensure_ascii=False)

        if name == "list_model_analytics":
            rows = await db.execute_fetchall(
                "SELECT id, name, code FROM analytics WHERE model_id = ? ORDER BY sort_order",
                (inp["model_id"],),
            )
            return json.dumps([dict(r) for r in rows], ensure_ascii=False)

        if name == "add_analytic_to_all_sheets":
            model_id = inp["model_id"]
            analytic_id = inp["analytic_id"]
            # Resolve analytic name for the consolidation-suggestion prompt.
            a_rows = await db.execute_fetchall(
                "SELECT name FROM analytics WHERE id = ?", (analytic_id,),
            )
            analytic_name = a_rows[0]["name"] if a_rows else "аналитика"
            sheets = await db.execute_fetchall(
                "SELECT id FROM sheets WHERE model_id = ?", (model_id,),
            )
            added = 0
            skipped = 0
            newly_added_sheet_ids: list[str] = []
            for s in sheets:
                sid = s["id"]
                existing = await db.execute_fetchall(
                    "SELECT id FROM sheet_analytics WHERE sheet_id = ? AND analytic_id = ?",
                    (sid, analytic_id),
                )
                if existing:
                    skipped += 1
                    continue
                cnt_rows = await db.execute_fetchall(
                    "SELECT COUNT(*) AS n FROM sheet_analytics WHERE sheet_id = ?", (sid,),
                )
                sort_order = cnt_rows[0]["n"] if cnt_rows else 0
                await db.execute(
                    "INSERT INTO sheet_analytics (id, sheet_id, analytic_id, sort_order, is_fixed, fixed_record_id) VALUES (?, ?, ?, ?, 0, NULL)",
                    (str(uuid.uuid4()), sid, analytic_id, sort_order),
                )
                added += 1
                newly_added_sheet_ids.append(sid)
            await db.commit()
            # P4: ask Claude to fill in consolidation formulas (ratios/avgs)
            # per indicator on each sheet where we just added a new analytic.
            # Tolerant of failures — returns 0 if Claude/API is unavailable.
            formulas_written = 0
            for sid in newly_added_sheet_ids:
                try:
                    formulas_written += await _suggest_consolidations_for_sheet(
                        db, sid, analytic_name,
                    )
                except Exception as e:
                    print(f"[add_analytic_to_all_sheets] suggest failed on {sid}: {e}")
            # Propagate: if same-named indicator got formula on one sheet but not
            # another, copy it so all sheets are consistent.
            try:
                formulas_written += await _propagate_consolidations(db, model_id)
            except Exception as e:
                print(f"[add_analytic_to_all_sheets] propagate failed: {e}")
            client_actions.append({"type": "reload_model", "model_id": model_id})
            return json.dumps(
                {"added": added, "skipped": skipped, "formulas_suggested": formulas_written},
                ensure_ascii=False,
            )

        if name == "remove_analytic_from_all_sheets":
            model_id = inp["model_id"]
            analytic_id = inp["analytic_id"]
            sheets = await db.execute_fetchall(
                "SELECT id FROM sheets WHERE model_id = ?", (model_id,),
            )
            removed = 0
            for s in sheets:
                existing = await db.execute_fetchall(
                    "SELECT id FROM sheet_analytics WHERE sheet_id = ? AND analytic_id = ?",
                    (s["id"], analytic_id),
                )
                if existing:
                    await db.execute(
                        "DELETE FROM sheet_analytics WHERE sheet_id = ? AND analytic_id = ?",
                        (s["id"], analytic_id),
                    )
                    removed += 1
            await db.commit()
            client_actions.append({"type": "reload_model", "model_id": model_id})
            return json.dumps({"removed": removed}, ensure_ascii=False)

        if name == "fill_sheet":
            result = await _fill_sheet_impl(
                db,
                sheet_id=inp["sheet_id"],
                mode=inp.get("mode", "value"),
                value=inp.get("value"),
                vmin=inp.get("min", 1),
                vmax=inp.get("max", 100),
                user_id=ctx.user_id,
            )
            client_actions.append({"type": "reload_sheet", "sheet_id": inp["sheet_id"]})
            return json.dumps(result, ensure_ascii=False)

        if name == "list_users":
            rows = await db.execute_fetchall(
                "SELECT id, username, can_admin FROM users ORDER BY username"
            )
            return json.dumps([dict(r) for r in rows], ensure_ascii=False)

        if name == "list_analytic_records":
            rows = await db.execute_fetchall(
                """SELECT id, parent_id, data_json FROM analytic_records
                   WHERE analytic_id = ? ORDER BY sort_order""",
                (inp["analytic_id"],),
            )
            # Detect has_children by checking if this id appears as parent_id
            parent_ids = {r["parent_id"] for r in rows if r["parent_id"]}
            out = []
            for r in rows:
                data = json.loads(r["data_json"] or "{}")
                out.append({
                    "id": r["id"],
                    "name": data.get("name", ""),
                    "parent_id": r["parent_id"],
                    "has_children": r["id"] in parent_ids,
                })
            return json.dumps(out, ensure_ascii=False)

        if name == "set_record_permission":
            user_id = inp["user_id"]
            analytic_id = inp["analytic_id"]
            record_id = inp["record_id"]
            can_view = 1 if inp.get("can_view", True) else 0
            can_edit = 1 if inp.get("can_edit", False) else 0
            existing = await db.execute_fetchall(
                """SELECT id FROM analytic_record_permissions
                   WHERE user_id = ? AND analytic_id = ? AND record_id = ?""",
                (user_id, analytic_id, record_id),
            )
            if existing:
                await db.execute(
                    """UPDATE analytic_record_permissions
                       SET can_view = ?, can_edit = ? WHERE id = ?""",
                    (can_view, can_edit, existing[0]["id"]),
                )
            else:
                await db.execute(
                    """INSERT INTO analytic_record_permissions
                       (id, user_id, analytic_id, record_id, can_view, can_edit)
                       VALUES (?, ?, ?, ?, ?, ?)""",
                    (str(uuid.uuid4()), user_id, analytic_id, record_id,
                     can_view, can_edit),
                )
            await db.commit()
            return json.dumps({"ok": True}, ensure_ascii=False)

        if name == "list_excel_in_folder":
            import os as _os
            folder = _os.path.expanduser(inp.get("folder_path", "")).strip()
            if not folder or not _os.path.isdir(folder):
                return json.dumps({"error": f"not a directory: {folder}"}, ensure_ascii=False)
            files = []
            for fn in sorted(_os.listdir(folder)):
                if fn.startswith("~$"):  # Excel lock files
                    continue
                if not fn.lower().endswith((".xlsx", ".xls")):
                    continue
                full = _os.path.join(folder, fn)
                try:
                    st = _os.stat(full)
                    files.append({
                        "path": full,
                        "name": fn,
                        "size": st.st_size,
                        "mtime": int(st.st_mtime),
                    })
                except OSError:
                    continue
            return json.dumps({"folder": folder, "files": files}, ensure_ascii=False)

        if name == "import_excel_from_path":
            import os as _os, io as _io
            from fastapi import UploadFile as _UploadFile
            from backend.routers.import_excel import import_excel as _do_import
            path = _os.path.expanduser(inp.get("file_path", "")).strip()
            if not path or not _os.path.isfile(path):
                return json.dumps({"error": f"file not found: {path}"}, ensure_ascii=False)
            fname = _os.path.basename(path)
            model_name = inp.get("model_name") or _os.path.splitext(fname)[0]
            with open(path, "rb") as fh:
                data = fh.read()
            upload = _UploadFile(file=_io.BytesIO(data), filename=fname)
            try:
                result = await _do_import(file=upload, model_name=model_name)
            except Exception as ex:
                return json.dumps({"error": f"import failed: {ex}"}, ensure_ascii=False)
            client_actions.append({"type": "reload_model", "model_id": result.get("model_id", "")})
            return json.dumps(result, ensure_ascii=False)

        return json.dumps({"error": f"unknown tool {name}"}, ensure_ascii=False)
    except Exception as e:
        return json.dumps({"error": str(e)}, ensure_ascii=False)


# ── Main endpoint ──────────────────────────────────────────────────────────

@router.post("/message")
async def chat_message(req: ChatRequest):
    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        raise HTTPException(status_code=500, detail="ANTHROPIC_API_KEY not set on server")

    import anthropic
    kwargs = {"api_key": api_key}
    base_url = os.environ.get("ANTHROPIC_BASE_URL")
    if base_url:
        kwargs["base_url"] = base_url
    client = anthropic.Anthropic(**kwargs)

    # Build context string injected into system prompt
    ctx = req.context
    ctx_lines = []
    if ctx.current_model_id:
        ctx_lines.append(f"Текущая модель: {ctx.current_model_id}")
    if ctx.current_sheet_id:
        ctx_lines.append(f"Текущий лист: {ctx.current_sheet_id}")
    if ctx.user_id:
        ctx_lines.append(f"Пользователь: {ctx.user_id}")
    system = SYSTEM_PROMPT
    if ctx_lines:
        system += "\n\nКонтекст:\n" + "\n".join(ctx_lines)

    # Convert messages to Anthropic format (pass through; frontend sends correct shape)
    messages = [m.model_dump() for m in req.messages]
    client_actions: list[dict] = []

    # Tool-use loop (cap at 8 iterations to prevent runaway)
    from backend.llm_cache import cached_messages_create
    for _ in range(8):
        resp = cached_messages_create(
            client,
            model="claude-sonnet-4-20250514",
            max_tokens=2000,
            system=system,
            tools=TOOLS,
            messages=messages,
        )
        if resp.stop_reason != "tool_use":
            # Final response — extract text
            text = "".join(b.text for b in resp.content if b.type == "text")
            return {"message": text, "actions": client_actions}

        # Execute tool calls
        tool_results = []
        for block in resp.content:
            if block.type == "tool_use":
                result = await _exec_tool(block.name, block.input, ctx, client_actions)
                tool_results.append({
                    "type": "tool_result",
                    "tool_use_id": block.id,
                    "content": result,
                })
        messages.append({"role": "assistant", "content": [b.model_dump() for b in resp.content]})
        messages.append({"role": "user", "content": tool_results})

    return {"message": "(ограничение: слишком много итераций)", "actions": client_actions}
