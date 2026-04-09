import uuid
from fastapi import APIRouter
from pydantic import BaseModel
from backend.db import get_db

router = APIRouter(prefix="/api/cells", tags=["cells"])


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
async def get_cells(sheet_id: str):
    db = get_db()
    rows = await db.execute_fetchall(
        "SELECT * FROM cell_data WHERE sheet_id = ?", (sheet_id,)
    )
    return [dict(r) for r in rows]


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


@router.put("/by-sheet/{sheet_id}")
async def save_cells(sheet_id: str, body: BulkCellsIn):
    db = get_db()
    for cell in body.cells:
        await _save_cell(db, sheet_id, cell)
    await db.commit()
    return {"ok": True}


@router.put("/by-sheet/{sheet_id}/single")
async def save_single_cell(sheet_id: str, body: CellIn):
    db = get_db()
    await _save_cell(db, sheet_id, body)
    await db.commit()
    return {"ok": True}


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
