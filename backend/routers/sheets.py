import uuid
import json as _json
from fastapi import APIRouter, HTTPException, Query
from pydantic import BaseModel
from backend.db import get_db

router = APIRouter(prefix="/api/sheets", tags=["sheets"])


async def _find_first_leaf(db, analytic_id: str) -> str | None:
    """Find the first leaf (terminal) record of an analytic."""
    recs = await db.execute_fetchall(
        "SELECT id, parent_id FROM analytic_records WHERE analytic_id = ? ORDER BY sort_order",
        (analytic_id,),
    )
    if not recs:
        return None
    parent_ids = {r["id"] for r in recs if any(c["parent_id"] == r["id"] for c in recs)}
    for r in recs:
        if r["id"] not in parent_ids:
            return r["id"]
    return recs[0]["id"]


async def _find_root_record(db, analytic_id: str) -> str | None:
    """Find the root (top-level parent) record of an analytic.
    If there's a single root with children, return it. Otherwise return first leaf."""
    recs = await db.execute_fetchall(
        "SELECT id, parent_id FROM analytic_records WHERE analytic_id = ? ORDER BY sort_order",
        (analytic_id,),
    )
    if not recs:
        return None
    roots = [r for r in recs if not r["parent_id"]]
    parent_ids = {r["id"] for r in recs if any(c["parent_id"] == r["id"] for c in recs)}
    # If there's exactly one root that has children, use it (HEAD)
    root_parents = [r for r in roots if r["id"] in parent_ids]
    if len(root_parents) == 1:
        return root_parents[0]["id"]
    # Fallback to first leaf
    for r in recs:
        if r["id"] not in parent_ids:
            return r["id"]
    return recs[0]["id"]


class SheetIn(BaseModel):
    model_id: str | None = None
    name: str = ""
    excel_code: str = ""


class SheetAnalyticIn(BaseModel):
    analytic_id: str
    sort_order: int = 0
    is_fixed: bool = False
    fixed_record_id: str | None = None


class ReorderIn(BaseModel):
    ordered_ids: list[str]


# --- Sheets CRUD ---

@router.get("/by-model/{model_id}")
async def list_sheets(model_id: str):
    db = get_db()
    rows = await db.execute_fetchall(
        "SELECT * FROM sheets WHERE model_id = ? ORDER BY sort_order, created_at", (model_id,)
    )
    return [dict(r) for r in rows]


@router.put("/reorder/{model_id}")
async def reorder_sheets(model_id: str, body: ReorderIn):
    db = get_db()
    for i, sheet_id in enumerate(body.ordered_ids):
        await db.execute(
            "UPDATE sheets SET sort_order=? WHERE id=? AND model_id=?",
            (i, sheet_id, model_id),
        )
    await db.commit()
    return {"ok": True}


@router.post("")
async def create_sheet(body: SheetIn):
    db = get_db()
    sid = str(uuid.uuid4())
    await db.execute(
        "INSERT INTO sheets (id, model_id, name, excel_code) VALUES (?, ?, ?, ?)",
        (sid, body.model_id, body.name, body.excel_code),
    )
    # Auto-grant permissions to all existing users
    users = await db.execute_fetchall("SELECT id FROM users")
    for u in users:
        spid = str(uuid.uuid4())
        await db.execute(
            "INSERT OR IGNORE INTO sheet_permissions (id, sheet_id, user_id, can_view, can_edit) VALUES (?,?,?,1,1)",
            (spid, sid, u["id"]),
        )
    await db.commit()
    from backend.formula_engine import invalidate_engine
    await invalidate_engine(db, body.model_id)
    row = await db.execute_fetchall("SELECT * FROM sheets WHERE id = ?", (sid,))
    return dict(row[0])


@router.put("/{sheet_id}")
async def update_sheet(sheet_id: str, body: SheetIn):
    db = get_db()
    fields = ["name=?", "updated_at=datetime('now')"]
    params: list = [body.name]
    if body.excel_code:
        fields.append("excel_code=?")
        params.append(body.excel_code)
    params.append(sheet_id)
    await db.execute(f"UPDATE sheets SET {', '.join(fields)} WHERE id=?", params)
    await db.commit()
    row = await db.execute_fetchall("SELECT * FROM sheets WHERE id = ?", (sheet_id,))
    return dict(row[0])


@router.patch("/{sheet_id}/lock")
async def toggle_lock(sheet_id: str):
    """Toggle the locked state of a sheet. Only admins should call this."""
    db = get_db()
    row = await db.execute_fetchall("SELECT locked FROM sheets WHERE id = ?", (sheet_id,))
    if not row:
        raise HTTPException(404, "Sheet not found")
    new_val = 0 if row[0]["locked"] else 1
    await db.execute("UPDATE sheets SET locked = ?, updated_at = datetime('now') WHERE id = ?", (new_val, sheet_id))
    await db.commit()
    row2 = await db.execute_fetchall("SELECT * FROM sheets WHERE id = ?", (sheet_id,))
    return dict(row2[0])


@router.delete("/{sheet_id}")
async def delete_sheet(sheet_id: str):
    db = get_db()
    # Get model_id before deletion for invalidation
    model_row = await db.execute_fetchall("SELECT model_id FROM sheets WHERE id = ?", (sheet_id,))
    await db.execute("DELETE FROM sheets WHERE id = ?", (sheet_id,))
    await db.commit()
    if model_row:
        from backend.formula_engine import invalidate_engine
        await invalidate_engine(db, model_row[0]["model_id"])
    return {"ok": True}


# --- Sheet-Analytic bindings ---

@router.get("/{sheet_id}/analytics")
async def list_sheet_analytics(sheet_id: str):
    db = get_db()
    rows = await db.execute_fetchall(
        """SELECT sa.*, a.name as analytic_name, a.icon as analytic_icon
           FROM sheet_analytics sa
           JOIN analytics a ON a.id = sa.analytic_id
           WHERE sa.sheet_id = ? ORDER BY sa.sort_order""",
        (sheet_id,),
    )
    return [dict(r) for r in rows]


@router.post("/{sheet_id}/analytics")
async def add_sheet_analytic(sheet_id: str, body: SheetAnalyticIn):
    db = get_db()
    said = str(uuid.uuid4())
    await db.execute(
        "INSERT INTO sheet_analytics (id, sheet_id, analytic_id, sort_order, is_fixed, fixed_record_id) VALUES (?,?,?,?,?,?)",
        (said, sheet_id, body.analytic_id, body.sort_order, int(body.is_fixed), body.fixed_record_id),
    )

    # Migrate existing cell data: append first leaf record of new analytic to coord_keys.
    # Values AND formulas move to the first leaf; other leaves start empty.
    # Per-cell formulas (e.g. [выдачи]/[партнёры]) keep working because they
    # reference indicators by name — same-context resolution with the new dimension.
    # HEAD = SUM(leaves) is computed by the formula engine's consolidation phase.
    first_leaf = await _find_first_leaf(db, body.analytic_id)
    if first_leaf:
        from backend.coord_key import intern as _ck_intern
        first_leaf_seq = str(await _ck_intern(db, first_leaf))
        cells = await db.execute_fetchall(
            "SELECT id, coord_key FROM cell_data WHERE sheet_id = ?", (sheet_id,)
        )
        for c in cells:
            new_key = c["coord_key"] + "|" + first_leaf_seq
            await db.execute(
                "UPDATE cell_data SET coord_key = ? WHERE id = ?",
                (new_key, c["id"])
            )

    # Invalidate engine cache — model structure changed (new dimension)
    model_row = await db.execute_fetchall(
        "SELECT model_id FROM sheets WHERE id = ?", (sheet_id,))
    if model_row:
        from backend.formula_engine import invalidate_engine
        await invalidate_engine(db, model_row[0]["model_id"])

    await db.commit()
    row = await db.execute_fetchall("SELECT * FROM sheet_analytics WHERE id = ?", (said,))
    return dict(row[0])


@router.put("/{sheet_id}/analytics/{sa_id}")
async def update_sheet_analytic(sheet_id: str, sa_id: str, body: SheetAnalyticIn):
    db = get_db()
    await db.execute(
        "UPDATE sheet_analytics SET sort_order=?, is_fixed=?, fixed_record_id=? WHERE id=? AND sheet_id=?",
        (body.sort_order, int(body.is_fixed), body.fixed_record_id, sa_id, sheet_id),
    )
    # Invalidate engine cache — sort order change affects coord_key structure
    model_row = await db.execute_fetchall(
        "SELECT model_id FROM sheets WHERE id = ?", (sheet_id,))
    if model_row:
        from backend.formula_engine import invalidate_engine
        await invalidate_engine(db, model_row[0]["model_id"])
    await db.commit()
    row = await db.execute_fetchall("SELECT * FROM sheet_analytics WHERE id = ?", (sa_id,))
    return dict(row[0])


class PeriodLevelIn(BaseModel):
    min_period_level: str | None = None   # 'M', 'Q', 'H', 'Y' or null
    visible_record_ids: list[str] | None = None


@router.patch("/{sheet_id}/analytics/{sa_id}/period-level")
async def set_period_level(sheet_id: str, sa_id: str, body: PeriodLevelIn):
    """Set the minimum period level and/or visible record IDs for a period-analytic binding."""
    import json
    db = get_db()
    valid = {None, 'M', 'Q', 'H', 'Y'}
    if body.min_period_level not in valid:
        raise HTTPException(400, f"Invalid level: {body.min_period_level}")
    vis_json = json.dumps(body.visible_record_ids) if body.visible_record_ids else None
    await db.execute(
        "UPDATE sheet_analytics SET min_period_level = ?, visible_record_ids = ? WHERE id = ? AND sheet_id = ?",
        (body.min_period_level, vis_json, sa_id, sheet_id),
    )
    await db.commit()
    row = await db.execute_fetchall("SELECT * FROM sheet_analytics WHERE id = ?", (sa_id,))
    return dict(row[0]) if row else {}


@router.delete("/{sheet_id}/analytics/{sa_id}")
async def remove_sheet_analytic(sheet_id: str, sa_id: str):
    db = get_db()

    # Get the analytic being removed and current binding order
    sa_row = await db.execute_fetchall("SELECT analytic_id FROM sheet_analytics WHERE id = ?", (sa_id,))
    if not sa_row:
        return {"ok": True}
    removing_aid = sa_row[0]["analytic_id"]

    # Get all analytic IDs in order for this sheet
    bindings = await db.execute_fetchall(
        "SELECT analytic_id FROM sheet_analytics WHERE sheet_id = ? ORDER BY sort_order", (sheet_id,)
    )
    order = [b["analytic_id"] for b in bindings]

    # Find the position of the analytic being removed
    if removing_aid in order:
        pos = order.index(removing_aid)
        # Get all record seq_ids for this analytic (coord_key parts are seq_id strings).
        rec_ids: set[str] = set()
        recs = await db.execute_fetchall(
            "SELECT seq_id FROM analytic_records WHERE analytic_id = ? AND seq_id IS NOT NULL",
            (removing_aid,)
        )
        for r in recs:
            rec_ids.add(str(r["seq_id"]))

        # Strip the part from coord_keys.
        # Multiple cells (HEAD, F1, F2) may collapse to the same key — keep the
        # one with the most data (longest value) and delete duplicates.
        cells = await db.execute_fetchall(
            "SELECT id, coord_key, value FROM cell_data WHERE sheet_id = ?", (sheet_id,)
        )
        # Group by new_key to detect collisions
        new_key_map: dict[str, list] = {}  # new_key → [(cell_id, value_len)]
        delete_ids: list[str] = []
        for c in cells:
            parts = c["coord_key"].split("|")
            if pos < len(parts) and parts[pos] in rec_ids:
                new_parts = parts[:pos] + parts[pos + 1:]
                if not new_parts:
                    delete_ids.append(c["id"])
                else:
                    nk = "|".join(new_parts)
                    val_len = len(c["value"] or "") if c["value"] and c["value"] != "0" else 0
                    new_key_map.setdefault(nk, []).append((c["id"], val_len, nk))

        for nk, entries in new_key_map.items():
            # Keep the entry with the longest value, delete the rest
            entries.sort(key=lambda x: x[1], reverse=True)
            keep_id, _, _ = entries[0]
            await db.execute(
                "UPDATE cell_data SET coord_key = ? WHERE id = ?", (nk, keep_id)
            )
            for cid, _, _ in entries[1:]:
                delete_ids.append(cid)

        for cid in delete_ids:
            await db.execute("DELETE FROM cell_data WHERE id = ?", (cid,))

    await db.execute("DELETE FROM sheet_analytics WHERE id = ? AND sheet_id = ?", (sa_id, sheet_id))
    # Invalidate engine cache — model structure changed
    model_row = await db.execute_fetchall(
        "SELECT model_id FROM sheets WHERE id = ?", (sheet_id,))
    if model_row:
        from backend.formula_engine import invalidate_engine
        await invalidate_engine(db, model_row[0]["model_id"])
    await db.commit()
    return {"ok": True}


# ── Main analytic per sheet ──

class MainAnalyticIn(BaseModel):
    analytic_id: str


@router.get("/{sheet_id}/main-analytic")
async def get_main_analytic(sheet_id: str):
    db = get_db()
    rows = await db.execute_fetchall(
        """SELECT sa.analytic_id FROM sheet_analytics sa
           JOIN analytics a ON a.id = sa.analytic_id
           WHERE sa.sheet_id = ? AND sa.is_main = 1 AND a.is_periods = 0
           LIMIT 1""",
        (sheet_id,),
    )
    return {"analytic_id": rows[0]["analytic_id"] if rows else None}


@router.put("/{sheet_id}/main-analytic")
async def set_main_analytic(sheet_id: str, body: MainAnalyticIn):
    """Set is_main=1 on one non-period analytic, clear it on the rest."""
    db = get_db()
    # Validate target is bound to this sheet and is not a period analytic
    target = await db.execute_fetchall(
        """SELECT sa.id, a.is_periods FROM sheet_analytics sa
           JOIN analytics a ON a.id = sa.analytic_id
           WHERE sa.sheet_id = ? AND sa.analytic_id = ?""",
        (sheet_id, body.analytic_id),
    )
    if not target:
        return {"error": "analytic not bound to this sheet"}
    if target[0]["is_periods"]:
        return {"error": "cannot mark period analytic as main"}
    await db.execute(
        "UPDATE sheet_analytics SET is_main = 0 WHERE sheet_id = ?", (sheet_id,)
    )
    await db.execute(
        "UPDATE sheet_analytics SET is_main = 1 WHERE sheet_id = ? AND analytic_id = ?",
        (sheet_id, body.analytic_id),
    )
    # Invalidate engine cache — main axis change affects consolidation
    model_row = await db.execute_fetchall(
        "SELECT model_id FROM sheets WHERE id = ?", (sheet_id,))
    if model_row:
        from backend.formula_engine import invalidate_engine
        await invalidate_engine(db, model_row[0]["model_id"])
    await db.commit()
    return {"ok": True, "analytic_id": body.analytic_id}


@router.put("/{sheet_id}/analytics-reorder")
async def reorder_sheet_analytics(sheet_id: str, body: ReorderIn):
    db = get_db()
    for i, sa_id in enumerate(body.ordered_ids):
        await db.execute(
            "UPDATE sheet_analytics SET sort_order=? WHERE id=? AND sheet_id=?",
            (i, sa_id, sheet_id),
        )
    await db.commit()
    return {"ok": True}


# ── View settings ──

class ViewSettingsIn(BaseModel):
    settings: dict


@router.get("/{sheet_id}/view-settings")
async def get_view_settings(sheet_id: str, user_id: str = ""):
    db = get_db()
    import json
    rows = await db.execute_fetchall(
        "SELECT settings FROM sheet_view_settings WHERE sheet_id = ? AND user_id = ?",
        (sheet_id, user_id),
    )
    if rows:
        return json.loads(rows[0]["settings"])
    return {}


@router.put("/{sheet_id}/view-settings")
async def save_view_settings(sheet_id: str, body: ViewSettingsIn):
    db = get_db()
    import json
    user_id = body.settings.pop("_user_id", "") if isinstance(body.settings, dict) else ""
    settings_json = json.dumps(body.settings, ensure_ascii=False)
    await db.execute(
        """INSERT INTO sheet_view_settings (sheet_id, user_id, settings) VALUES (?, ?, ?)
           ON CONFLICT(sheet_id, user_id) DO UPDATE SET settings = excluded.settings""",
        (sheet_id, user_id, settings_json),
    )
    await db.commit()
    return {"ok": True}


@router.get("/{sheet_id}/load-bundle")
async def load_bundle(sheet_id: str, user_id: str = Query("")):
    """Single endpoint returning everything the grid needs to render:
    sheet info, sheet_analytics, analytics details, all records, and view settings.
    Replaces ~15 separate HTTP calls with one."""
    db = get_db()

    # Sheet info (locked state, model_id)
    sheet_rows = await db.execute_fetchall("SELECT * FROM sheets WHERE id = ?", (sheet_id,))
    if not sheet_rows:
        raise HTTPException(404, "Sheet not found")
    sheet = dict(sheet_rows[0])

    # Sheet analytics (bindings)
    sa_rows = await db.execute_fetchall(
        """SELECT sa.*, a.name as analytic_name, a.icon as analytic_icon
           FROM sheet_analytics sa
           JOIN analytics a ON a.id = sa.analytic_id
           WHERE sa.sheet_id = ? ORDER BY sa.sort_order""",
        (sheet_id,),
    )
    sheet_analytics = [dict(r) for r in sa_rows]
    analytic_ids = [r["analytic_id"] for r in sa_rows]

    # All analytics details + all records in 2 bulk queries
    if analytic_ids:
        ph = ",".join("?" for _ in analytic_ids)
        a_rows = await db.execute_fetchall(
            f"SELECT * FROM analytics WHERE id IN ({ph})", analytic_ids
        )
        r_rows = await db.execute_fetchall(
            f"SELECT * FROM analytic_records WHERE analytic_id IN ({ph}) ORDER BY sort_order",
            analytic_ids,
        )
    else:
        a_rows, r_rows = [], []

    analytics = {r["id"]: dict(r) for r in a_rows}
    records: dict[str, list] = {aid: [] for aid in analytic_ids}
    for r in r_rows:
        aid = r["analytic_id"]
        if aid in records:
            records[aid].append(dict(r))

    # View settings
    vs_rows = await db.execute_fetchall(
        "SELECT settings FROM sheet_view_settings WHERE sheet_id = ? AND user_id = ?",
        (sheet_id, user_id),
    )
    view_settings = _json.loads(vs_rows[0]["settings"]) if vs_rows else {}

    return {
        "sheet": sheet,
        "sheet_analytics": sheet_analytics,
        "analytics": analytics,
        "records": records,
        "view_settings": view_settings,
    }
