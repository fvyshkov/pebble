import uuid
from fastapi import APIRouter, Query
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
            await db.execute("UPDATE cell_data SET value = ? WHERE sheet_id = ? AND coord_key = ?", (val, sid, ck))
        total += len(changes)
    return total


@router.put("/by-sheet/{sheet_id}")
async def save_cells(sheet_id: str, body: BulkCellsIn):
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


@router.get("/history/model/{model_id}")
async def get_model_history(model_id: str, limit: int = 10):
    """Recent changes across all sheets in a model."""
    db = get_db()
    rows = await db.execute_fetchall(
        """SELECT h.id, h.sheet_id, h.coord_key, h.old_value, h.new_value, h.created_at,
                  s.name as sheet_name, u.username
           FROM cell_history h
           JOIN sheets s ON s.id = h.sheet_id AND s.model_id = ?
           LEFT JOIN users u ON u.id = h.user_id
           ORDER BY h.created_at DESC LIMIT ?""",
        (model_id, limit),
    )
    return [dict(r) for r in rows]


class UndoIn(BaseModel):
    history_id: str  # undo up to and including this history entry


@router.post("/undo/{model_id}")
async def undo(model_id: str, body: UndoIn):
    """Undo changes from most recent back to history_id (inclusive)."""
    db = get_db()
    # Get all history entries from newest to the target
    rows = await db.execute_fetchall(
        """SELECT h.id, h.sheet_id, h.coord_key, h.old_value, h.created_at
           FROM cell_history h
           JOIN sheets s ON s.id = h.sheet_id AND s.model_id = ?
           WHERE h.created_at >= (SELECT created_at FROM cell_history WHERE id = ?)
           ORDER BY h.created_at DESC""",
        (model_id, body.history_id),
    )
    if not rows:
        return {"error": "History entry not found"}

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


@router.delete("/history/model/{model_id}")
async def clear_history(model_id: str):
    """Clear all history for a model."""
    db = get_db()
    await db.execute(
        "DELETE FROM cell_history WHERE sheet_id IN (SELECT id FROM sheets WHERE model_id = ?)",
        (model_id,),
    )
    await db.commit()
    return {"ok": True}
