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
- Запускать пересчёт формул.

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
]


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
    for _ in range(8):
        resp = client.messages.create(
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
