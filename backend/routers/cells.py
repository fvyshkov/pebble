import uuid
import json
import asyncio
from fastapi import APIRouter, Query
from fastapi.responses import StreamingResponse
from pydantic import BaseModel
from backend.db import get_db

router = APIRouter(prefix="/api/cells", tags=["cells"])


async def _get_allowed_records(db, user_id: str | None, sheet_id: str) -> dict[str, set[str]] | None:
    """Return {analytic_id: set(record_ids)} for restricted analytics, or None if no restrictions."""
    if not user_id:
        return None
    bindings = await db.execute_fetchall(
        "SELECT sa.analytic_id FROM sheet_analytics sa WHERE sa.sheet_id = ? ORDER BY sa.sort_order",
        (sheet_id,),
    )
    restrictions: dict[str, set[str]] = {}
    for b in bindings:
        aid = b["analytic_id"]
        perms = await db.execute_fetchall(
            "SELECT record_id FROM analytic_record_permissions WHERE user_id = ? AND analytic_id = ? AND can_view = 1",
            (user_id, aid),
        )
        if perms:
            restrictions[aid] = {p["record_id"] for p in perms}
    return restrictions if restrictions else None


async def _get_editable_records(db, user_id: str | None, sheet_id: str) -> dict[str, set[str]] | None:
    """Return {analytic_id: set(record_ids)} where user can_edit, or None if no restrictions."""
    if not user_id:
        return None
    bindings = await db.execute_fetchall(
        "SELECT sa.analytic_id FROM sheet_analytics sa WHERE sa.sheet_id = ? ORDER BY sa.sort_order",
        (sheet_id,),
    )
    restrictions: dict[str, set[str]] = {}
    for b in bindings:
        aid = b["analytic_id"]
        perms = await db.execute_fetchall(
            "SELECT record_id FROM analytic_record_permissions WHERE user_id = ? AND analytic_id = ? AND can_edit = 1",
            (user_id, aid),
        )
        if perms:
            restrictions[aid] = {p["record_id"] for p in perms}
    return restrictions if restrictions else None


def _coord_allowed(coord_key: str, restrictions: dict[str, set[str]], order: list[str]) -> bool:
    """Check if a coord_key is allowed given restrictions.
    coord_key = "rid1|rid2|..." matching order of analytics.
    """
    parts = coord_key.split("|")
    for i, aid in enumerate(order):
        if aid in restrictions and i < len(parts):
            if parts[i] not in restrictions[aid]:
                return False
    return True


class CellIn(BaseModel):
    coord_key: str
    value: str | None = None
    data_type: str = "number"
    user_id: str | None = None
    rule: str | None = None
    formula: str | None = None


class BulkCellsIn(BaseModel):
    cells: list[CellIn]


@router.get("/by-sheet/{sheet_id}")
async def get_cells(sheet_id: str, user_id: str | None = Query(None)):
    db = get_db()
    rows = await db.execute_fetchall(
        "SELECT * FROM cell_data WHERE sheet_id = ?", (sheet_id,)
    )
    restrictions = await _get_allowed_records(db, user_id, sheet_id)
    if not restrictions:
        return [dict(r) for r in rows]

    # Get analytic order for coord_key parsing
    order = [b["analytic_id"] for b in await db.execute_fetchall(
        "SELECT analytic_id FROM sheet_analytics WHERE sheet_id = ? ORDER BY sort_order", (sheet_id,)
    )]
    return [dict(r) for r in rows if _coord_allowed(r["coord_key"], restrictions, order)]


class PartialCellsIn(BaseModel):
    coord_keys: list[str]


@router.post("/by-sheet/{sheet_id}/partial")
async def get_cells_partial(sheet_id: str, body: PartialCellsIn, user_id: str | None = Query(None)):
    """Return cells only for the requested coord_keys (lazy loading)."""
    db = get_db()
    if not body.coord_keys:
        return []
    # Fetch in batches of 500 using IN clause
    results = []
    keys = body.coord_keys
    for i in range(0, len(keys), 500):
        batch = keys[i:i+500]
        placeholders = ",".join("?" for _ in batch)
        rows = await db.execute_fetchall(
            f"SELECT * FROM cell_data WHERE sheet_id = ? AND coord_key IN ({placeholders})",
            (sheet_id, *batch),
        )
        results.extend(rows)

    restrictions = await _get_allowed_records(db, user_id, sheet_id)
    if not restrictions:
        return [dict(r) for r in results]

    order = [b["analytic_id"] for b in await db.execute_fetchall(
        "SELECT analytic_id FROM sheet_analytics WHERE sheet_id = ? ORDER BY sort_order", (sheet_id,)
    )]
    return [dict(r) for r in results if _coord_allowed(r["coord_key"], restrictions, order)]


async def _save_cell(db, sheet_id: str, cell: CellIn):
    existing = await db.execute_fetchall(
        "SELECT id, value FROM cell_data WHERE sheet_id = ? AND coord_key = ?",
        (sheet_id, cell.coord_key),
    )
    old_value = existing[0]["value"] if existing else None

    if existing:
        # Build dynamic update
        fields = ["value=?", "data_type=?"]
        params: list = [cell.value, cell.data_type]
        if cell.rule is not None:
            fields.append("rule=?")
            params.append(cell.rule)
        if cell.formula is not None:
            fields.append("formula=?")
            params.append(cell.formula)
        params.extend([sheet_id, cell.coord_key])
        await db.execute(
            f"UPDATE cell_data SET {', '.join(fields)} WHERE sheet_id=? AND coord_key=?",
            params,
        )
    else:
        cid = str(uuid.uuid4())
        await db.execute(
            "INSERT INTO cell_data (id, sheet_id, coord_key, value, data_type, rule, formula) VALUES (?,?,?,?,?,?,?)",
            (cid, sheet_id, cell.coord_key, cell.value, cell.data_type,
             cell.rule or "manual", cell.formula or ""),
        )

    # Record history
    if old_value != cell.value:
        hid = str(uuid.uuid4())
        await db.execute(
            "INSERT INTO cell_history (id, sheet_id, coord_key, user_id, old_value, new_value) VALUES (?,?,?,?,?,?)",
            (hid, sheet_id, cell.coord_key, cell.user_id, old_value, cell.value),
        )


async def _recalc_model(db, sheet_id: str) -> int:
    """Recalculate all formula cells in the model containing this sheet."""
    from backend.formula_engine import calculate_model
    sheet = await db.execute_fetchall("SELECT model_id FROM sheets WHERE id = ?", (sheet_id,))
    if not sheet:
        return 0
    result = await calculate_model(db, sheet[0]["model_id"])
    total = 0
    for sid, changes in result.items():
        for ck, val in changes.items():
            await db.execute(
                """INSERT INTO cell_data (id, sheet_id, coord_key, value, rule)
                   VALUES (?, ?, ?, ?, 'formula')
                   ON CONFLICT(sheet_id, coord_key) DO UPDATE SET value = excluded.value""",
                (str(__import__('uuid').uuid4()), sid, ck, val),
            )
        total += len(changes)
    return total


@router.put("/by-sheet/{sheet_id}")
async def save_cells(sheet_id: str, body: BulkCellsIn, no_recalc: bool = Query(False)):
    db = get_db()
    # Check edit permissions if user_id provided
    user_id = body.cells[0].user_id if body.cells else None
    edit_restrictions = await _get_editable_records(db, user_id, sheet_id)
    if edit_restrictions:
        order = [b["analytic_id"] for b in await db.execute_fetchall(
            "SELECT analytic_id FROM sheet_analytics WHERE sheet_id = ? ORDER BY sort_order", (sheet_id,)
        )]
        for cell in body.cells:
            if not _coord_allowed(cell.coord_key, edit_restrictions, order):
                return {"error": f"No edit permission for {cell.coord_key}"}
    for cell in body.cells:
        await _save_cell(db, sheet_id, cell)
    computed = 0
    if not no_recalc:
        computed = await _recalc_model(db, sheet_id)
    await db.commit()
    return {"ok": True, "computed": computed}


@router.put("/by-sheet/{sheet_id}/single")
async def save_single_cell(sheet_id: str, body: CellIn):
    db = get_db()
    edit_restrictions = await _get_editable_records(db, body.user_id, sheet_id)
    if edit_restrictions:
        order = [b["analytic_id"] for b in await db.execute_fetchall(
            "SELECT analytic_id FROM sheet_analytics WHERE sheet_id = ? ORDER BY sort_order", (sheet_id,)
        )]
        if not _coord_allowed(body.coord_key, edit_restrictions, order):
            return {"error": "No edit permission"}
    await _save_cell(db, sheet_id, body)
    computed = await _recalc_model(db, sheet_id)
    await db.commit()
    return {"ok": True, "computed": computed}


@router.post("/calculate/{sheet_id}")
async def calculate(sheet_id: str):
    """Recalculate all formula cells in the model (lazy pull, cross-sheet)."""
    db = get_db()
    computed = await _recalc_model(db, sheet_id)
    await db.commit()
    return {"computed": computed}


@router.post("/calculate-model/{model_id}/stream")
async def calculate_model_stream(model_id: str):
    """Recalculate all formula cells with SSE streaming progress."""
    from backend.formula_engine import calculate_model

    async def event_stream():
        db = get_db()
        sheets = await db.execute_fetchall(
            "SELECT id, name FROM sheets WHERE model_id = ? ORDER BY created_at", (model_id,))
        total_sheets = len(sheets)
        # Count formula cells across the model so the UI can show X/Y progress.
        total_cells_rows = await db.execute_fetchall(
            """SELECT COUNT(*) AS n FROM cell_data
               WHERE sheet_id IN (SELECT id FROM sheets WHERE model_id = ?)
                 AND rule = 'formula'""",
            (model_id,),
        )
        total_cells = total_cells_rows[0]["n"] if total_cells_rows else 0
        yield f"data: {json.dumps({'phase': 'start', 'total_sheets': total_sheets, 'total_cells': total_cells})}\n\n"

        result = await calculate_model(db, model_id)
        total = 0
        done_sheets = 0
        for sid, changes in result.items():
            for ck, val in changes.items():
                await db.execute(
                    """INSERT INTO cell_data (id, sheet_id, coord_key, value, rule)
                       VALUES (?, ?, ?, ?, 'formula')
                       ON CONFLICT(sheet_id, coord_key) DO UPDATE SET value = excluded.value""",
                    (str(__import__('uuid').uuid4()), sid, ck, val),
                )
            total += len(changes)
            done_sheets += 1
            sheet_name = next((s["name"] for s in sheets if s["id"] == sid), sid)
            yield f"data: {json.dumps({'phase': 'sheet_done', 'sheet': sheet_name, 'done': done_sheets, 'total_sheets': total_sheets, 'computed': total})}\n\n"
            await asyncio.sleep(0)  # yield control

        await db.commit()
        yield f"data: {json.dumps({'phase': 'done', 'computed': total})}\n\n"

    return StreamingResponse(event_stream(), media_type="text/event-stream")


@router.get("/history/{sheet_id}/{coord_key}")
async def get_cell_history(sheet_id: str, coord_key: str):
    db = get_db()
    rows = await db.execute_fetchall(
        """SELECT h.*, u.username FROM cell_history h
           LEFT JOIN users u ON u.id = h.user_id
           WHERE h.sheet_id = ? AND h.coord_key = ?
           ORDER BY h.created_at DESC LIMIT 50""",
        (sheet_id, coord_key),
    )
    return [dict(r) for r in rows]


@router.get("/model-history/{model_id}")
async def get_model_history(model_id: str, limit: int = 10):
    """Recent changes across all sheets in a model."""
    db = get_db()
    # Two-step: get sheet IDs, then filter history
    sheet_rows = await db.execute_fetchall(
        "SELECT id, name FROM sheets WHERE model_id = ?", (model_id,))
    if not sheet_rows:
        return []
    sheet_names = {r["id"]: r["name"] for r in sheet_rows}

    # Get ALL recent history and filter in Python
    # Fetch history for each sheet individually (workaround for aiosqlite query issues)
    rows = []
    for sid in sheet_names:
        sheet_rows = await db.execute_fetchall(
            """SELECT h.*, u.username FROM cell_history h
               LEFT JOIN users u ON u.id = h.user_id
               WHERE h.sheet_id = ?
               ORDER BY h.created_at DESC LIMIT ?""",
            (sid, limit),
        )
        rows.extend(sheet_rows)
    result = []
    for r in rows:
        if r["sheet_id"] in sheet_names:
            d = dict(r)
            d["sheet_name"] = sheet_names[r["sheet_id"]]
            result.append(d)
            if len(result) >= limit:
                break
    return result


class UndoIn(BaseModel):
    history_id: str  # undo up to and including this history entry


@router.post("/undo/{model_id}")
async def undo(model_id: str, body: UndoIn):
    """Undo changes from most recent back to history_id (inclusive)."""
    db = get_db()
    # Get target timestamp
    target = await db.execute_fetchall("SELECT created_at FROM cell_history WHERE id = ?", (body.history_id,))
    if not target:
        return {"error": "History entry not found"}
    target_ts = target[0]["created_at"]
    # Get sheet IDs for model
    sheet_rows = await db.execute_fetchall("SELECT id FROM sheets WHERE model_id = ?", (model_id,))
    rows = []
    for sr in sheet_rows:
        sheet_hist = await db.execute_fetchall(
            """SELECT id, sheet_id, coord_key, old_value FROM cell_history
               WHERE sheet_id = ? AND created_at >= ?
               ORDER BY created_at DESC""",
            (sr["id"], target_ts),
        )
        rows.extend(sheet_hist)
    if not rows:
        return {"error": "No changes to undo"}

    undone = 0
    for r in rows:
        await db.execute(
            "UPDATE cell_data SET value = ? WHERE sheet_id = ? AND coord_key = ?",
            (r["old_value"], r["sheet_id"], r["coord_key"]),
        )
        await db.execute("DELETE FROM cell_history WHERE id = ?", (r["id"],))
        undone += 1

    # Recalc
    sheet_id = rows[0]["sheet_id"]
    computed = await _recalc_model(db, sheet_id)
    await db.commit()
    return {"undone": undone, "computed": computed}


@router.delete("/model-history/{model_id}")
async def clear_history(model_id: str):
    """Clear all history for a model."""
    db = get_db()
    await db.execute(
        "DELETE FROM cell_history WHERE sheet_id IN (SELECT id FROM sheets WHERE model_id = ?)",
        (model_id,),
    )
    await db.commit()
    return {"ok": True}
