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
- Заполнять лист случайными или фиксированными значениями (fill_sheet).
- Добавлять/убирать аналитики сразу со всех листов модели.
- Запускать пересчёт формул.
- Строить графики по данным модели (build_chart).
- Управлять правами пользователей на записи аналитики.
- Переключать режимы интерфейса (switch_mode) и навигировать (open_sheet).

КРИТИЧЕСКОЕ ПРАВИЛО: НИКОГДА не давай текстовые инструкции вроде "перейди в настройки", "нажми кнопку" и т.п.
Вместо этого ВСЕГДА используй соответствующие инструменты:
- Надо переключить режим → вызови switch_mode
- Надо открыть лист → вызови open_sheet
- Надо заполнить данные → вызови fill_sheet или set_cell
- Надо создать аналитику → вызови create_analytic, потом create_records для записей, потом add_analytic_to_all_sheets
- Надо добавить записи в аналитику → вызови create_records
- Надо добавить аналитику в листы → вызови add_analytic_to_all_sheets или add_analytic_to_sheet
- Надо создать лист → вызови create_sheet
- Надо удалить модель → спроси подтверждение, затем вызови delete_model
- Надо пересчитать → вызови recalc
- Надо построить график → вызови build_chart с названиями показателей и группировки (данные бэкенд достанет сам)
- Надо сделать презентацию / отчёт → вызови build_presentation (бэкенд сам соберёт данные и сгенерирует HTML)
- Надо импортировать Excel с компьютера пользователя → вызови import_excel_from_browser (откроет файлопикер в браузере)
- import_excel_from_path и list_excel_in_folder — только если Pebble развёрнут ЛОКАЛЬНО и пользователь указал путь к файлу на сервере.
  Если пользователь просто хочет импортировать Excel — используй import_excel_from_browser.
Ты — агент, который ВЫПОЛНЯЕТ действия, а не инструктор, который рассказывает, что делать.

Контекст пользователя передаётся в каждом запросе (текущая модель, текущий лист, id пользователя).
Если нужен идентификатор (model_id, sheet_id), сначала вызови list_models или list_sheets.

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
        "name": "delete_model",
        "description": "Удалить модель со всеми её листами, аналитиками, ячейками и правилами. Необратимо! Перед удалением спроси подтверждение у пользователя.",
        "input_schema": {
            "type": "object",
            "properties": {
                "model_id": {"type": "string", "description": "ID модели для удаления"},
            },
            "required": ["model_id"],
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
        "name": "query_data",
        "description": (
            "Запросить данные листа с серверной агрегацией. НЕ возвращает сырые ячейки — "
            "делает SQL-агрегацию на бэкенде и возвращает компактный результат. "
            "Указываешь показатель (indicator) и измерение для группировки (group_by). "
            "Бэкенд найдёт нужные ячейки, сгруппирует и вернёт [{label, value}]. "
            "Для графиков используй build_chart_from_data — он сам вызывает query_data."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "sheet_id": {"type": "string"},
                "indicator": {"type": "string", "description": "Название показателя (запись из аналитики-справочника)"},
                "group_by": {"type": "string", "description": "Название аналитики для группировки (например 'Периоды')"},
                "filter_records": {
                    "type": "array", "items": {"type": "string"},
                    "description": "Ограничить только этими записями группировки (опционально)",
                },
            },
            "required": ["sheet_id", "indicator"],
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
        "name": "import_excel_from_browser",
        "description": (
            "Открыть диалог выбора файла в браузере пользователя для импорта "
            "Excel-файла. Используй этот инструмент, когда пользователь хочет "
            "импортировать Excel с локального компьютера. После выбора файла "
            "импорт запустится автоматически."
        ),
        "input_schema": {
            "type": "object",
            "properties": {},
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
    {
        "name": "build_chart",
        "description": (
            "Построить график по данным листа. Указываешь показатель и группировку — "
            "бэкенд сам достаёт данные из БД, агрегирует и отправляет график в интерфейс. "
            "НЕ нужно предварительно читать данные — всё делается за один вызов. "
            "Несколько показателей = несколько серий на одном графике."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "sheet_id": {"type": "string", "description": "ID листа"},
                "title": {"type": "string", "description": "Заголовок графика"},
                "chart_type": {
                    "type": "string",
                    "enum": ["line", "bar", "pie", "area"],
                    "description": "Тип графика",
                },
                "indicators": {
                    "type": "array",
                    "items": {"type": "string"},
                    "description": "Названия показателей (серии графика)",
                },
                "group_by": {
                    "type": "string",
                    "description": "Название аналитики для оси X / категорий (например 'Периоды')",
                },
                "filter_records": {
                    "type": "array", "items": {"type": "string"},
                    "description": "Показывать только эти записи группировки (опционально, например только месяцы без кварталов)",
                },
            },
            "required": ["sheet_id", "title", "chart_type", "indicators", "group_by"],
        },
    },
    {
        "name": "create_analytic",
        "description": (
            "Создать новую аналитику в модели. Возвращает {id, name, code}. "
            "После создания можно добавить записи через create_records и привязать к листам через add_analytic_to_all_sheets."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "model_id": {"type": "string", "description": "ID модели"},
                "name": {"type": "string", "description": "Название аналитики"},
                "is_periods": {"type": "boolean", "description": "Это периоды? (год/квартал/месяц)", "default": False},
                "period_types": {
                    "type": "array", "items": {"type": "string"},
                    "description": "Типы периодов: ['year','quarter','month']",
                    "default": [],
                },
                "period_start": {"type": "string", "description": "Начало периодов (YYYY-MM-DD)"},
                "period_end": {"type": "string", "description": "Окончание периодов (YYYY-MM-DD)"},
            },
            "required": ["model_id", "name"],
        },
    },
    {
        "name": "create_records",
        "description": (
            "Создать одну или несколько записей в аналитике. "
            "Каждая запись — это строка (например, подразделение, продукт, показатель). "
            "Для иерархии укажи parent_id у дочерних записей."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "analytic_id": {"type": "string", "description": "ID аналитики"},
                "records": {
                    "type": "array",
                    "items": {
                        "type": "object",
                        "properties": {
                            "name": {"type": "string", "description": "Название записи"},
                            "parent_id": {"type": "string", "description": "ID родительской записи (для иерархии)"},
                        },
                        "required": ["name"],
                    },
                    "description": "Массив записей [{name, parent_id?}]",
                },
            },
            "required": ["analytic_id", "records"],
        },
    },
    {
        "name": "add_analytic_to_sheet",
        "description": "Привязать аналитику к одному конкретному листу.",
        "input_schema": {
            "type": "object",
            "properties": {
                "sheet_id": {"type": "string", "description": "ID листа"},
                "analytic_id": {"type": "string", "description": "ID аналитики"},
            },
            "required": ["sheet_id", "analytic_id"],
        },
    },
    {
        "name": "create_sheet",
        "description": "Создать новый лист в модели.",
        "input_schema": {
            "type": "object",
            "properties": {
                "model_id": {"type": "string", "description": "ID модели"},
                "name": {"type": "string", "description": "Название листа"},
            },
            "required": ["model_id", "name"],
        },
    },
    {
        "name": "rename_model",
        "description": "Переименовать модель.",
        "input_schema": {
            "type": "object",
            "properties": {
                "model_id": {"type": "string"},
                "name": {"type": "string", "description": "Новое название"},
            },
            "required": ["model_id", "name"],
        },
    },
    {
        "name": "rename_sheet",
        "description": "Переименовать лист.",
        "input_schema": {
            "type": "object",
            "properties": {
                "sheet_id": {"type": "string"},
                "name": {"type": "string", "description": "Новое название"},
            },
            "required": ["sheet_id", "name"],
        },
    },
    {
        "name": "rename_analytic",
        "description": "Переименовать аналитику.",
        "input_schema": {
            "type": "object",
            "properties": {
                "analytic_id": {"type": "string"},
                "name": {"type": "string", "description": "Новое название"},
            },
            "required": ["analytic_id", "name"],
        },
    },
    {
        "name": "update_record",
        "description": "Переименовать запись аналитики.",
        "input_schema": {
            "type": "object",
            "properties": {
                "record_id": {"type": "string"},
                "name": {"type": "string", "description": "Новое название записи"},
            },
            "required": ["record_id", "name"],
        },
    },
    {
        "name": "delete_record",
        "description": "Удалить запись из аналитики.",
        "input_schema": {
            "type": "object",
            "properties": {
                "record_id": {"type": "string"},
                "analytic_id": {"type": "string"},
            },
            "required": ["record_id", "analytic_id"],
        },
    },
    {
        "name": "delete_analytic",
        "description": "Удалить аналитику со всеми записями. Необратимо.",
        "input_schema": {
            "type": "object",
            "properties": {
                "analytic_id": {"type": "string"},
                "model_id": {"type": "string", "description": "ID модели (для reload)"},
            },
            "required": ["analytic_id", "model_id"],
        },
    },
    {
        "name": "delete_sheet",
        "description": "Удалить лист со всеми ячейками.",
        "input_schema": {
            "type": "object",
            "properties": {
                "sheet_id": {"type": "string"},
                "model_id": {"type": "string"},
            },
            "required": ["sheet_id", "model_id"],
        },
    },
    {
        "name": "build_presentation",
        "description": (
            "Создать аналитическую презентацию по данным модели. "
            "Бэкенд сам соберёт все данные, отправит на анализ и сгенерирует красивый HTML "
            "с графиками (Chart.js), трендами, выводами и ключевыми показателями. "
            "Результат отобразится в центральной панели с возможностью скачать как PDF."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "sheet_id": {"type": "string", "description": "ID листа для анализа"},
                "title": {"type": "string", "description": "Заголовок презентации"},
                "focus": {"type": "string", "description": "На чём сфокусировать анализ (опционально)"},
            },
            "required": ["sheet_id"],
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


class BulkAnalyticRequest(BaseModel):
    model_id: str
    analytic_id: str


@router.post("/bulk_add_analytic")
async def bulk_add_analytic(req: BulkAnalyticRequest):
    """Add analytic to all sheets + suggest consolidation formulas."""
    db = get_db()
    a_rows = await db.execute_fetchall(
        "SELECT name FROM analytics WHERE id = ?", (req.analytic_id,),
    )
    analytic_name = a_rows[0]["name"] if a_rows else "аналитика"
    sheets = await db.execute_fetchall(
        "SELECT id FROM sheets WHERE model_id = ?", (req.model_id,),
    )
    added = 0
    newly_added_sheet_ids: list[str] = []
    for s in sheets:
        sid = s["id"]
        existing = await db.execute_fetchall(
            "SELECT id FROM sheet_analytics WHERE sheet_id = ? AND analytic_id = ?",
            (sid, req.analytic_id),
        )
        if existing:
            continue
        cnt_rows = await db.execute_fetchall(
            "SELECT COUNT(*) AS n FROM sheet_analytics WHERE sheet_id = ?", (sid,),
        )
        sort_order = cnt_rows[0]["n"] if cnt_rows else 0
        await db.execute(
            "INSERT INTO sheet_analytics (id, sheet_id, analytic_id, sort_order, is_fixed, fixed_record_id) VALUES (?, ?, ?, ?, 0, NULL)",
            (str(uuid.uuid4()), sid, req.analytic_id, sort_order),
        )
        added += 1
        newly_added_sheet_ids.append(sid)
    await db.commit()
    formulas_written = 0
    for sid in newly_added_sheet_ids:
        try:
            formulas_written += await _suggest_consolidations_for_sheet(
                db, sid, analytic_name,
            )
        except Exception as e:
            print(f"[bulk_add_analytic] suggest failed on {sid}: {e}")
    try:
        formulas_written += await _propagate_consolidations(db, req.model_id)
    except Exception as e:
        print(f"[bulk_add_analytic] propagate failed: {e}")
    return {"added": added, "total_sheets": len(sheets), "formulas_suggested": formulas_written}


@router.post("/bulk_remove_analytic")
async def bulk_remove_analytic(req: BulkAnalyticRequest):
    """Remove analytic from all sheets."""
    db = get_db()
    sheets = await db.execute_fetchall(
        "SELECT id FROM sheets WHERE model_id = ?", (req.model_id,),
    )
    removed = 0
    for s in sheets:
        existing = await db.execute_fetchall(
            "SELECT id FROM sheet_analytics WHERE sheet_id = ? AND analytic_id = ?",
            (s["id"], req.analytic_id),
        )
        if existing:
            await db.execute(
                "DELETE FROM sheet_analytics WHERE sheet_id = ? AND analytic_id = ?",
                (s["id"], req.analytic_id),
            )
            removed += 1
    await db.commit()
    return {"removed": removed, "total_sheets": len(sheets)}


# ── Presentation builder ───────────────────────────────────────────────────

async def _build_presentation(db, sheet_id: str, title: str | None, focus: str | None) -> str:
    """Gather sheet data, send to Claude for HTML presentation with analysis."""
    # 1. Get sheet info
    sheet_rows = await db.execute_fetchall("SELECT name, model_id FROM sheets WHERE id = ?", (sheet_id,))
    if not sheet_rows:
        return "<h1>Лист не найден</h1>"
    sheet_name = sheet_rows[0]["name"]
    model_id = sheet_rows[0]["model_id"]
    model_rows = await db.execute_fetchall("SELECT name FROM models WHERE id = ?", (model_id,))
    model_name = model_rows[0]["name"] if model_rows else ""

    # 2. Get analytics and records
    sa_rows = await db.execute_fetchall(
        """SELECT a.id, a.name, a.is_periods FROM sheet_analytics sa
           JOIN analytics a ON a.id = sa.analytic_id
           WHERE sa.sheet_id = ? ORDER BY sa.sort_order""",
        (sheet_id,),
    )
    analytics_info = []
    all_records: dict[str, list[dict]] = {}  # analytic_id -> [{id, name, parent_id}]
    for a in sa_rows:
        recs = await db.execute_fetchall(
            "SELECT id, data_json, parent_id FROM analytic_records WHERE analytic_id = ? ORDER BY sort_order",
            (a["id"],),
        )
        rec_list = []
        for r in recs:
            try:
                d = json.loads(r["data_json"] or "{}")
                rec_list.append({"id": r["id"], "name": d.get("name", ""), "parent_id": r["parent_id"]})
            except Exception:
                pass
        all_records[a["id"]] = rec_list
        analytics_info.append({
            "name": a["name"],
            "is_periods": bool(a["is_periods"]),
            "records": [r["name"] for r in rec_list],
        })

    # 3. Get cell data — build a structured table
    cells = await db.execute_fetchall(
        "SELECT coord_key, value FROM cell_data WHERE sheet_id = ? AND value IS NOT NULL AND value != ''",
        (sheet_id,),
    )
    # Build a lookup: record_id -> name
    rid_to_name: dict[str, str] = {}
    for recs in all_records.values():
        for r in recs:
            rid_to_name[r["id"]] = r["name"]

    # Convert cells to readable rows
    table_rows = []
    for c in cells:
        parts = (c["coord_key"] or "").split("|")
        labels = [rid_to_name.get(p, p) for p in parts]
        try:
            val = float(c["value"])
            val_str = f"{val:,.2f}" if val != int(val) else str(int(val))
        except (ValueError, TypeError):
            val_str = str(c["value"])
        table_rows.append({"coords": " | ".join(labels), "value": val_str})

    # 4. Build compact data summary for LLM (limit to ~2000 rows)
    data_text = "\n".join(f"{r['coords']}: {r['value']}" for r in table_rows[:2000])

    # 5. Send to Claude for HTML generation
    pres_title = title or f"Анализ: {sheet_name}"
    prompt = f"""Создай HTML-презентацию для аналитического отчёта.

Модель: {model_name}
Лист: {sheet_name}
Заголовок: {pres_title}
{f'Фокус анализа: {focus}' if focus else ''}

Аналитики:
{json.dumps(analytics_info, ensure_ascii=False, indent=2)}

Данные (формат: координата | координата: значение):
{data_text}

ТРЕБОВАНИЯ к HTML:
1. Полностью самодостаточный HTML (inline CSS, Chart.js через CDN)
2. Используй <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
3. Создай 3-5 графиков (bar, line, pie по уместности) по ключевым показателям
4. Добавь аналитические выводы: тренды, аномалии, ключевые цифры
5. Структура: заголовок → ключевые метрики (KPI карточки) → графики с пояснениями → выводы и рекомендации
6. Стиль: современный, чистый, с цветовой схемой синий/серый. Шрифт sans-serif.
7. Адаптивный для печати (@media print)
8. Язык: русский
9. НЕ используй внешние ресурсы кроме Chart.js CDN
10. Каждый график в отдельном <canvas> с уникальным id

Верни ТОЛЬКО HTML код, без markdown обёртки, без ```html блоков."""

    import anthropic as _anthropic
    _aclient = _anthropic.AsyncAnthropic(api_key=os.environ.get("ANTHROPIC_API_KEY", ""))
    resp = await _aclient.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=8000,
        messages=[{"role": "user", "content": prompt}],
    )
    html = "".join(b.text for b in resp.content if b.type == "text")
    # Strip markdown code fences if present
    if html.strip().startswith("```"):
        lines = html.strip().split("\n")
        if lines[0].startswith("```"):
            lines = lines[1:]
        if lines and lines[-1].strip() == "```":
            lines = lines[:-1]
        html = "\n".join(lines)
    return html


# ── Server-side data query (no LLM context needed) ────────────────────────

async def _query_sheet_data(
    db, sheet_id: str, indicator: str,
    group_by: str | None = None, filter_records: list[str] | None = None,
) -> dict:
    """Query cell data for a specific indicator, grouped by an analytic dimension.

    Returns {"data": [{"label": "...", "value": N}, ...]} — compact, never large.
    All heavy lifting (joins, aggregation) happens in Python/SQL, not in LLM context.
    """
    # 1. Find all analytics on this sheet
    sa_rows = await db.execute_fetchall(
        """SELECT a.id, a.name, a.is_periods FROM sheet_analytics sa
           JOIN analytics a ON a.id = sa.analytic_id
           WHERE sa.sheet_id = ? ORDER BY sa.sort_order""",
        (sheet_id,),
    )
    if not sa_rows:
        return {"error": "На листе нет аналитик"}

    # 2. Find the indicator record (search by name across all non-period analytics)
    indicator_record_id = None
    indicator_analytic_id = None
    for a in sa_rows:
        if a["is_periods"]:
            continue
        recs = await db.execute_fetchall(
            "SELECT id, data_json FROM analytic_records WHERE analytic_id = ?", (a["id"],),
        )
        for r in recs:
            try:
                d = json.loads(r["data_json"] or "{}")
                if d.get("name", "").lower().strip() == indicator.lower().strip():
                    indicator_record_id = r["id"]
                    indicator_analytic_id = a["id"]
                    break
            except Exception:
                pass
        if indicator_record_id:
            break

    if not indicator_record_id:
        return {"error": f"Показатель '{indicator}' не найден"}

    # 3. Find the grouping analytic
    group_analytic = None
    if group_by:
        for a in sa_rows:
            if a["name"].lower().strip() == group_by.lower().strip():
                group_analytic = a
                break
        if not group_analytic:
            # Try partial match
            for a in sa_rows:
                if group_by.lower() in a["name"].lower():
                    group_analytic = a
                    break
    if not group_analytic:
        # Default: use the first analytic that isn't the indicator's analytic
        for a in sa_rows:
            if a["id"] != indicator_analytic_id:
                group_analytic = a
                break

    if not group_analytic:
        return {"error": "Не найдена аналитика для группировки"}

    # 4. Get grouping records (leaf only — no parents)
    group_recs = await db.execute_fetchall(
        "SELECT id, data_json, parent_id FROM analytic_records WHERE analytic_id = ? ORDER BY sort_order",
        (group_analytic["id"],),
    )
    parent_ids = {r["parent_id"] for r in group_recs if r["parent_id"]}
    leaf_recs = [r for r in group_recs if r["id"] not in parent_ids]

    # 5. Filter if requested
    if filter_records:
        filter_lower = {f.lower().strip() for f in filter_records}
        filtered = []
        for r in leaf_recs:
            try:
                d = json.loads(r["data_json"] or "{}")
                if d.get("name", "").lower().strip() in filter_lower:
                    filtered.append(r)
            except Exception:
                pass
        leaf_recs = filtered

    # 6. Query cells
    # coord_key format is pipe-separated record IDs: "rec_id1|rec_id2|rec_id3"
    # We need cells containing both the indicator record ID and the group record ID
    data = []
    for rec in leaf_recs:
        try:
            rec_name = json.loads(rec["data_json"] or "{}").get("name", "")
        except Exception:
            rec_name = ""
        cells = await db.execute_fetchall(
            "SELECT value FROM cell_data WHERE sheet_id = ? AND coord_key LIKE ? AND coord_key LIKE ?",
            (sheet_id, f"%{indicator_record_id}%", f"%{rec['id']}%"),
        )
        # Sum all matching cells (e.g. across periods)
        total = 0.0
        count = 0
        for c in cells:
            try:
                v = float(c["value"]) if c["value"] else 0
                total += v
                count += 1
            except (ValueError, TypeError):
                pass
        if count > 0:
            data.append({"label": rec_name, "value": round(total, 2)})

    return {"data": data, "indicator": indicator, "group_by": group_analytic["name"]}


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

        if name == "delete_model":
            mid = inp["model_id"]
            print(f"[delete_model] Deleting model {mid}")
            # Gather all sheets
            sheets = await db.execute_fetchall("SELECT id FROM sheets WHERE model_id = ?", (mid,))
            sheet_ids = [s["id"] for s in sheets]
            # Delete cells, formula rules, sheet_analytics, view settings for each sheet
            for sid in sheet_ids:
                await db.execute("DELETE FROM cell_data WHERE sheet_id = ?", (sid,))
                await db.execute("DELETE FROM indicator_formula_rules WHERE sheet_id = ?", (sid,))
                await db.execute("DELETE FROM sheet_analytics WHERE sheet_id = ?", (sid,))
                await db.execute("DELETE FROM sheet_view_settings WHERE sheet_id = ?", (sid,))
            # Delete sheets
            await db.execute("DELETE FROM sheets WHERE model_id = ?", (mid,))
            # Delete analytics and their records
            analytics = await db.execute_fetchall("SELECT id FROM analytics WHERE model_id = ?", (mid,))
            for a in analytics:
                await db.execute("DELETE FROM analytic_records WHERE analytic_id = ?", (a["id"],))
            await db.execute("DELETE FROM analytics WHERE model_id = ?", (mid,))
            # Delete model
            await db.execute("DELETE FROM models WHERE id = ?", (mid,))
            await db.commit()
            print(f"[delete_model] Done — deleted {len(sheet_ids)} sheets")
            client_actions.append({"type": "reload_model", "model_id": mid})
            return json.dumps({"ok": True, "deleted_sheets": len(sheet_ids)}, ensure_ascii=False)

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
                "SELECT value FROM cell_data WHERE sheet_id = ? AND coord_key = ?",
                (inp["sheet_id"], inp["coord_key"]),
            )
            return json.dumps({"value": rows[0]["value"] if rows else None}, ensure_ascii=False)

        if name == "query_data":
            result = await _query_sheet_data(
                db, inp["sheet_id"], inp["indicator"],
                inp.get("group_by"), inp.get("filter_records"),
            )
            return json.dumps(result, ensure_ascii=False)

        if name == "set_cell":
            # Upsert manual cell value
            existing = await db.execute_fetchall(
                "SELECT id FROM cell_data WHERE sheet_id = ? AND coord_key = ?",
                (inp["sheet_id"], inp["coord_key"]),
            )
            if existing:
                await db.execute(
                    "UPDATE cell_data SET value = ?, data_type = COALESCE(?, data_type) WHERE id = ?",
                    (inp["value"], inp.get("data_type"), existing[0]["id"]),
                )
            else:
                await db.execute(
                    """INSERT INTO cell_data (id, sheet_id, coord_key, value, data_type, rule, formula)
                       VALUES (?, ?, ?, ?, ?, 'manual', '')""",
                    (str(uuid.uuid4()), inp["sheet_id"], inp["coord_key"],
                     inp["value"], inp.get("data_type", "number")),
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

        if name == "import_excel_from_browser":
            client_actions.append({"type": "pick_excel_file"})
            return json.dumps({"ok": True, "message": "Диалог выбора файла открыт в браузере."}, ensure_ascii=False)

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

        if name == "build_chart":
            sheet_id = inp["sheet_id"]
            indicators = inp.get("indicators", [])
            group_by = inp.get("group_by")
            filter_recs = inp.get("filter_records")
            chart_type = inp.get("chart_type", "bar")
            title = inp.get("title", "")
            # Query data for each indicator
            all_labels: list[str] = []
            series_data: dict[str, dict[str, float]] = {}
            for ind_name in indicators:
                result = await _query_sheet_data(db, sheet_id, ind_name, group_by, filter_recs)
                if "error" in result:
                    return json.dumps(result, ensure_ascii=False)
                for item in result.get("data", []):
                    label = item["label"]
                    if label not in all_labels:
                        all_labels.append(label)
                    series_data.setdefault(ind_name, {})[label] = item["value"]
            # Build chart data array
            data = []
            for label in all_labels:
                row: dict = {"category": label}
                for ind_name in indicators:
                    field = ind_name.replace(" ", "_")
                    row[field] = series_data.get(ind_name, {}).get(label, 0)
                data.append(row)
            series = [{"field": ind.replace(" ", "_"), "name": ind} for ind in indicators]
            client_actions.append({
                "type": "show_chart",
                "title": title,
                "chart_type": chart_type,
                "data": data,
                "series": series,
                "category_field": "category",
            })
            return json.dumps({"ok": True, "points": len(data), "series": len(series)}, ensure_ascii=False)

        if name == "create_analytic":
            from backend.transliterate import transliterate
            model_id = inp["model_id"]
            aname = inp["name"]
            code = transliterate(aname)
            is_periods = inp.get("is_periods", False)
            period_types = inp.get("period_types", [])
            aid = str(uuid.uuid4())
            await db.execute(
                """INSERT INTO analytics (id, model_id, name, code, icon, is_periods, data_type,
                   period_types, period_start, period_end, sort_order)
                   VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                (aid, model_id, aname, code, "", int(is_periods), "sum",
                 json.dumps(period_types), inp.get("period_start", ""), inp.get("period_end", ""), 0),
            )
            if not is_periods:
                fid = str(uuid.uuid4())
                await db.execute(
                    "INSERT INTO analytic_fields (id, analytic_id, name, code, data_type, sort_order) VALUES (?,?,?,?,?,?)",
                    (fid, aid, "Наименование", "name", "string", 0),
                )
            else:
                # Create period fields
                from backend.routers.analytics import _ensure_period_fields
                await _ensure_period_fields(db, aid)
            await db.commit()
            if is_periods and period_types:
                from backend.routers.analytics import generate_periods as _gen_periods
                await _gen_periods(aid)
            client_actions.append({"type": "reload_model", "model_id": model_id})
            return json.dumps({"id": aid, "name": aname, "code": code}, ensure_ascii=False)

        if name == "create_records":
            analytic_id = inp["analytic_id"]
            records = inp.get("records", [])
            created = []
            for rec in records:
                rid = str(uuid.uuid4())
                data_json = {"name": rec["name"]}
                await db.execute(
                    "INSERT INTO analytic_records (id, analytic_id, parent_id, sort_order, data_json) VALUES (?,?,?,?,?)",
                    (rid, analytic_id, rec.get("parent_id"), len(created), json.dumps(data_json, ensure_ascii=False)),
                )
                created.append({"id": rid, "name": rec["name"]})
            await db.commit()
            return json.dumps({"created": len(created), "records": created}, ensure_ascii=False)

        if name == "add_analytic_to_sheet":
            sheet_id = inp["sheet_id"]
            analytic_id = inp["analytic_id"]
            existing = await db.execute_fetchall(
                "SELECT id FROM sheet_analytics WHERE sheet_id = ? AND analytic_id = ?",
                (sheet_id, analytic_id),
            )
            if existing:
                return json.dumps({"ok": True, "skipped": True}, ensure_ascii=False)
            said = str(uuid.uuid4())
            max_ord = await db.execute_fetchall(
                "SELECT COALESCE(MAX(sort_order),0) as m FROM sheet_analytics WHERE sheet_id = ?",
                (sheet_id,),
            )
            await db.execute(
                "INSERT INTO sheet_analytics (id, sheet_id, analytic_id, sort_order) VALUES (?,?,?,?)",
                (said, sheet_id, analytic_id, (max_ord[0]["m"] if max_ord else 0) + 1),
            )
            await db.commit()
            client_actions.append({"type": "reload_sheet", "sheet_id": sheet_id})
            return json.dumps({"ok": True, "added": True}, ensure_ascii=False)

        if name == "create_sheet":
            model_id = inp["model_id"]
            sname = inp["name"]
            sid = str(uuid.uuid4())
            await db.execute(
                "INSERT INTO sheets (id, model_id, name) VALUES (?,?,?)",
                (sid, model_id, sname),
            )
            await db.commit()
            client_actions.append({"type": "reload_model", "model_id": model_id})
            return json.dumps({"id": sid, "name": sname}, ensure_ascii=False)

        if name == "build_presentation":
            print(f"[PEBBLE] build_presentation called, sheet_id={inp.get('sheet_id')}")
            try:
                html = await _build_presentation(db, inp["sheet_id"], inp.get("title"), inp.get("focus"))
                print(f"[PEBBLE] build_presentation OK, html length={len(html)}")
            except Exception as pres_err:
                print(f"[PEBBLE] build_presentation FAILED: {pres_err}")
                import traceback; traceback.print_exc()
                return json.dumps({"error": f"Ошибка генерации: {pres_err}"}, ensure_ascii=False)
            client_actions.append({
                "type": "show_presentation",
                "html": html,
                "title": inp.get("title", "Презентация"),
            })
            return json.dumps({"ok": True, "message": "Презентация создана"}, ensure_ascii=False)

        if name == "rename_model":
            await db.execute("UPDATE models SET name = ? WHERE id = ?", (inp["name"], inp["model_id"]))
            await db.commit()
            client_actions.append({"type": "reload_model", "model_id": inp["model_id"]})
            return json.dumps({"ok": True}, ensure_ascii=False)

        if name == "rename_sheet":
            await db.execute("UPDATE sheets SET name = ? WHERE id = ?", (inp["name"], inp["sheet_id"]))
            await db.commit()
            # Find model_id for reload
            rows = await db.execute_fetchall("SELECT model_id FROM sheets WHERE id = ?", (inp["sheet_id"],))
            if rows:
                client_actions.append({"type": "reload_model", "model_id": rows[0]["model_id"]})
            return json.dumps({"ok": True}, ensure_ascii=False)

        if name == "rename_analytic":
            from backend.transliterate import transliterate
            new_name = inp["name"]
            new_code = transliterate(new_name)
            await db.execute(
                "UPDATE analytics SET name = ?, code = ? WHERE id = ?",
                (new_name, new_code, inp["analytic_id"]),
            )
            await db.commit()
            rows = await db.execute_fetchall("SELECT model_id FROM analytics WHERE id = ?", (inp["analytic_id"],))
            if rows:
                client_actions.append({"type": "reload_model", "model_id": rows[0]["model_id"]})
            return json.dumps({"ok": True}, ensure_ascii=False)

        if name == "update_record":
            # Read current data_json, update name
            rows = await db.execute_fetchall("SELECT data_json FROM analytic_records WHERE id = ?", (inp["record_id"],))
            if not rows:
                return json.dumps({"error": "record not found"}, ensure_ascii=False)
            data = json.loads(rows[0]["data_json"] or "{}")
            data["name"] = inp["name"]
            await db.execute(
                "UPDATE analytic_records SET data_json = ? WHERE id = ?",
                (json.dumps(data, ensure_ascii=False), inp["record_id"]),
            )
            await db.commit()
            return json.dumps({"ok": True}, ensure_ascii=False)

        if name == "delete_record":
            # Delete record and its children
            await db.execute("DELETE FROM analytic_records WHERE id = ? OR parent_id = ?", (inp["record_id"], inp["record_id"]))
            await db.commit()
            return json.dumps({"ok": True}, ensure_ascii=False)

        if name == "delete_analytic":
            aid = inp["analytic_id"]
            mid = inp["model_id"]
            await db.execute("DELETE FROM analytic_records WHERE analytic_id = ?", (aid,))
            await db.execute("DELETE FROM analytic_fields WHERE analytic_id = ?", (aid,))
            await db.execute("DELETE FROM sheet_analytics WHERE analytic_id = ?", (aid,))
            await db.execute("DELETE FROM indicator_formula_rules WHERE analytic_id = ?", (aid,))
            await db.execute("DELETE FROM analytics WHERE id = ?", (aid,))
            await db.commit()
            client_actions.append({"type": "reload_model", "model_id": mid})
            return json.dumps({"ok": True}, ensure_ascii=False)

        if name == "delete_sheet":
            sid = inp["sheet_id"]
            mid = inp["model_id"]
            await db.execute("DELETE FROM cell_data WHERE sheet_id = ?", (sid,))
            await db.execute("DELETE FROM indicator_formula_rules WHERE sheet_id = ?", (sid,))
            await db.execute("DELETE FROM sheet_analytics WHERE sheet_id = ?", (sid,))
            await db.execute("DELETE FROM sheet_view_settings WHERE sheet_id = ?", (sid,))
            await db.execute("DELETE FROM sheets WHERE id = ?", (sid,))
            await db.commit()
            client_actions.append({"type": "reload_model", "model_id": mid})
            return json.dumps({"ok": True}, ensure_ascii=False)

        return json.dumps({"error": f"unknown tool {name}"}, ensure_ascii=False)
    except Exception as e:
        import traceback
        traceback.print_exc()
        return json.dumps({"error": str(e)}, ensure_ascii=False)


# ── Friendly tool names for thinking display ─────────────────────────────

TOOL_LABELS = {
    "list_models": "Получаю список моделей",
    "list_sheets": "Получаю список листов",
    "list_analytics": "Получаю аналитики листа",
    "list_model_analytics": "Получаю аналитики модели",
    "list_analytic_records": "Получаю записи аналитики",
    "read_cell": "Читаю ячейку",
    "read_sheet_data": "Читаю данные листа",
    "query_data": "Запрашиваю данные",
    "set_cell": "Записываю значение",
    "fill_sheet": "Заполняю лист",
    "create_model": "Создаю модель",
    "create_sheet": "Создаю лист",
    "create_analytic": "Создаю аналитику",
    "create_records": "Создаю записи",
    "delete_model": "Удаляю модель",
    "delete_sheet": "Удаляю лист",
    "delete_analytic": "Удаляю аналитику",
    "delete_record": "Удаляю запись",
    "rename_model": "Переименовываю модель",
    "rename_sheet": "Переименовываю лист",
    "rename_analytic": "Переименовываю аналитику",
    "recalc": "Пересчитываю формулы",
    "build_chart": "Строю график",
    "build_presentation": "Готовлю презентацию",
    "import_excel_from_path": "Импортирую Excel",
    "import_excel_from_browser": "Открываю выбор файла",
    "add_analytic_to_all_sheets": "Добавляю аналитику на листы",
    "add_analytic_to_sheet": "Добавляю аналитику на лист",
    "remove_analytic_from_all_sheets": "Убираю аналитику с листов",
    "open_sheet": "Открываю лист",
    "switch_mode": "Переключаю режим",
    "pin_analytic": "Фиксирую аналитику",
    "unpin_analytic": "Снимаю фиксацию",
}


# ── Main endpoint ──────────────────────────────────────────────────────────

@router.post("/message")
async def chat_message(req: ChatRequest):
    try:
        return await _chat_message_impl(req)
    except HTTPException:
        raise
    except Exception as e:
        import traceback
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=str(e))


async def _chat_message_impl(req: ChatRequest):
    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        raise HTTPException(status_code=500, detail="ANTHROPIC_API_KEY not set on server")

    import anthropic
    kwargs = {"api_key": api_key}
    base_url = os.environ.get("ANTHROPIC_BASE_URL")
    if base_url:
        kwargs["base_url"] = base_url
    client = anthropic.AsyncAnthropic(**kwargs)

    # Build context string injected into system prompt
    db = get_db()
    ctx = req.context
    ctx_lines = []
    if ctx.current_model_id:
        model_rows = await db.execute_fetchall(
            "SELECT name FROM models WHERE id = ?", (ctx.current_model_id,))
        mname = model_rows[0]["name"] if model_rows else ctx.current_model_id
        ctx_lines.append(f"Текущая модель: {mname} (id={ctx.current_model_id})")
    if ctx.current_sheet_id:
        sheet_rows = await db.execute_fetchall(
            "SELECT name FROM sheets WHERE id = ?", (ctx.current_sheet_id,))
        sname = sheet_rows[0]["name"] if sheet_rows else ctx.current_sheet_id
        ctx_lines.append(f"Текущий лист: {sname} (id={ctx.current_sheet_id})")
        # Include analytics bound to this sheet
        sa_rows = await db.execute_fetchall(
            """SELECT a.id, a.name, a.is_periods FROM sheet_analytics sa
               JOIN analytics a ON a.id = sa.analytic_id
               WHERE sa.sheet_id = ? ORDER BY sa.sort_order""",
            (ctx.current_sheet_id,),
        )
        if sa_rows:
            parts = []
            for r in sa_rows:
                label = 'периоды' if r['is_periods'] else 'справочник'
                part = f"  - {r['name']} (id={r['id']}, {label})"
                # Add record names (up to 30) so agent knows the actual values
                recs = await db.execute_fetchall(
                    "SELECT id, data_json FROM analytic_records WHERE analytic_id = ? ORDER BY sort_order LIMIT 30",
                    (r['id'],),
                )
                if recs:
                    import json as _json
                    rec_items = []
                    for rec in recs:
                        try:
                            d = _json.loads(rec['data_json'] or '{}')
                            n = d.get('name', '')
                            if n:
                                rec_items.append(f"{n} (id={rec['id']})")
                        except Exception:
                            pass
                    if rec_items:
                        part += f"\n    Записи: {', '.join(rec_items)}"
                parts.append(part)
            ctx_lines.append("Аналитики на листе:\n" + "\n".join(parts))
    if ctx.user_id:
        ctx_lines.append(f"Пользователь: {ctx.user_id}")
    system = SYSTEM_PROMPT
    if ctx_lines:
        system += "\n\nКонтекст:\n" + "\n".join(ctx_lines)

    # Convert messages to Anthropic format; keep only last 20 to avoid context overflow
    messages = [m.model_dump() for m in req.messages]
    if len(messages) > 20:
        messages = messages[-20:]
        # Ensure first message is from user (Anthropic requirement)
        while messages and messages[0].get("role") != "user":
            messages.pop(0)
    client_actions: list[dict] = []

    # Tool-use loop (cap at 8 iterations to prevent runaway)
    for _ in range(15):
        resp = await client.messages.create(
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


@router.post("/message-stream")
async def chat_message_stream(req: ChatRequest):
    """Streaming version — sends SSE events for each thinking step."""
    from starlette.responses import StreamingResponse

    async def generate():
        try:
            api_key = os.environ.get("ANTHROPIC_API_KEY")
            if not api_key:
                yield f"data: {json.dumps({'type': 'error', 'text': 'ANTHROPIC_API_KEY not set'})}\n\n"
                return

            import anthropic
            kwargs = {"api_key": api_key}
            base_url = os.environ.get("ANTHROPIC_BASE_URL")
            if base_url:
                kwargs["base_url"] = base_url
            client = anthropic.AsyncAnthropic(**kwargs)

            db = get_db()
            ctx = req.context
            ctx_lines = []
            if ctx.current_model_id:
                model_rows = await db.execute_fetchall(
                    "SELECT name FROM models WHERE id = ?", (ctx.current_model_id,))
                mname = model_rows[0]["name"] if model_rows else ctx.current_model_id
                ctx_lines.append(f"Текущая модель: {mname} (id={ctx.current_model_id})")
            if ctx.current_sheet_id:
                sheet_rows = await db.execute_fetchall(
                    "SELECT name FROM sheets WHERE id = ?", (ctx.current_sheet_id,))
                sname = sheet_rows[0]["name"] if sheet_rows else ctx.current_sheet_id
                ctx_lines.append(f"Текущий лист: {sname} (id={ctx.current_sheet_id})")
                sa_rows = await db.execute_fetchall(
                    """SELECT a.id, a.name, a.is_periods FROM sheet_analytics sa
                       JOIN analytics a ON a.id = sa.analytic_id
                       WHERE sa.sheet_id = ? ORDER BY sa.sort_order""",
                    (ctx.current_sheet_id,),
                )
                if sa_rows:
                    parts = []
                    for r in sa_rows:
                        label = 'периоды' if r['is_periods'] else 'справочник'
                        part = f"  - {r['name']} (id={r['id']}, {label})"
                        recs = await db.execute_fetchall(
                            "SELECT id, data_json FROM analytic_records WHERE analytic_id = ? ORDER BY sort_order LIMIT 30",
                            (r['id'],),
                        )
                        if recs:
                            rec_items = []
                            for rec in recs:
                                try:
                                    d = json.loads(rec['data_json'] or '{}')
                                    n = d.get('name', '')
                                    if n:
                                        rec_items.append(f"{n} (id={rec['id']})")
                                except Exception:
                                    pass
                            if rec_items:
                                part += f"\n    Записи: {', '.join(rec_items)}"
                        parts.append(part)
                    ctx_lines.append("Аналитики на листе:\n" + "\n".join(parts))
            if ctx.user_id:
                ctx_lines.append(f"Пользователь: {ctx.user_id}")
            system = SYSTEM_PROMPT
            if ctx_lines:
                system += "\n\nКонтекст:\n" + "\n".join(ctx_lines)

            messages = [m.model_dump() for m in req.messages]
            if len(messages) > 20:
                messages = messages[-20:]
                while messages and messages[0].get("role") != "user":
                    messages.pop(0)
            client_actions: list[dict] = []

            for iteration in range(15):
                resp = await client.messages.create(
                    model="claude-sonnet-4-20250514",
                    max_tokens=2000,
                    system=system,
                    tools=TOOLS,
                    messages=messages,
                )
                if resp.stop_reason != "tool_use":
                    text = "".join(b.text for b in resp.content if b.type == "text")
                    yield f"data: {json.dumps({'type': 'done', 'text': text, 'actions': client_actions}, ensure_ascii=False)}\n\n"
                    return

                # Execute tool calls and emit thinking events
                tool_results = []
                for block in resp.content:
                    if block.type == "tool_use":
                        label = TOOL_LABELS.get(block.name, block.name)
                        yield f"data: {json.dumps({'type': 'thinking', 'text': label + '…'}, ensure_ascii=False)}\n\n"
                        result = await _exec_tool(block.name, block.input, ctx, client_actions)
                        tool_results.append({
                            "type": "tool_result",
                            "tool_use_id": block.id,
                            "content": result,
                        })
                messages.append({"role": "assistant", "content": [b.model_dump() for b in resp.content]})
                messages.append({"role": "user", "content": tool_results})

            yield f"data: {json.dumps({'type': 'done', 'text': '(ограничение: слишком много итераций)', 'actions': client_actions}, ensure_ascii=False)}\n\n"
        except Exception as e:
            import traceback
            traceback.print_exc()
            yield f"data: {json.dumps({'type': 'error', 'text': str(e)}, ensure_ascii=False)}\n\n"

    return StreamingResponse(generate(), media_type="text/event-stream")
