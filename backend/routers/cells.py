import uuid
import json
import asyncio
from fastapi import APIRouter, HTTPException, Query
from fastapi.responses import StreamingResponse
from pydantic import BaseModel
from backend.db import get_db, get_read_db

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


def _cell_slim(r) -> dict:
    """Return only fields the frontend needs (skip id, sheet_id to reduce payload)."""
    d: dict = {"coord_key": r["coord_key"], "value": r["value"]}
    if r["rule"]:
        d["rule"] = r["rule"]
    if r["formula"]:
        d["formula"] = r["formula"]
    return d


@router.get("/by-sheet/{sheet_id}")
async def get_cells(sheet_id: str, user_id: str | None = Query(None)):
    async with get_read_db() as db:
        restrictions = await _get_allowed_records(db, user_id, sheet_id)
        order = None
        if restrictions:
            order = [b["analytic_id"] for b in await db.execute_fetchall(
                "SELECT analytic_id FROM sheet_analytics WHERE sheet_id = ? ORDER BY sort_order", (sheet_id,)
            )]

        rows = await db.execute_fetchall(
            """SELECT cd.coord_key, cd.value, cd.rule,
                      COALESCE(NULLIF(cd.formula,''), f.text, '') AS formula
               FROM cell_data cd LEFT JOIN formulas f ON f.id = cd.formula_id
               WHERE cd.sheet_id = ?""", (sheet_id,),
        )

    # Build JSON manually to avoid per-row dict creation + json.dumps overhead.
    # For 350K rows this saves ~1-2s compared to returning list-of-dicts.
    chunks: list[str] = []
    _dumps = json.dumps
    for r in rows:
        if restrictions and order and not _coord_allowed(r["coord_key"], restrictions, order):
            continue
        ck = r["coord_key"]
        v = r["value"]
        s = '{"coord_key":"' + ck + '","value":' + (_dumps(v) if v is not None else '""')
        rule = r["rule"]
        if rule:
            s += ',"rule":"' + rule + '"'
        formula = r["formula"]
        if formula:
            s += ',"formula":' + _dumps(formula)
        s += "}"
        chunks.append(s)

    from fastapi.responses import Response
    content = "[" + ",".join(chunks) + "]"
    return Response(content=content, media_type="application/json")


class PartialCellsIn(BaseModel):
    coord_keys: list[str]


@router.post("/by-sheet/{sheet_id}/partial")
async def get_cells_partial(sheet_id: str, body: PartialCellsIn, user_id: str | None = Query(None)):
    """Return cells only for the requested coord_keys (lazy loading)."""
    if not body.coord_keys:
        return []
    async with get_read_db() as db:
        # Fetch in batches of 500 using IN clause
        results = []
        keys = body.coord_keys
        for i in range(0, len(keys), 500):
            batch = keys[i:i+500]
            placeholders = ",".join("?" for _ in batch)
            rows = await db.execute_fetchall(
                f"""SELECT cd.coord_key, cd.value, cd.rule,
                           COALESCE(NULLIF(cd.formula,''), f.text, '') AS formula
                    FROM cell_data cd LEFT JOIN formulas f ON f.id = cd.formula_id
                    WHERE cd.sheet_id = ? AND cd.coord_key IN ({placeholders})""",
                (sheet_id, *batch),
            )
            results.extend(rows)

        restrictions = await _get_allowed_records(db, user_id, sheet_id)
        if not restrictions:
            return [_cell_slim(r) for r in results]

        order = [b["analytic_id"] for b in await db.execute_fetchall(
            "SELECT analytic_id FROM sheet_analytics WHERE sheet_id = ? ORDER BY sort_order", (sheet_id,)
        )]
        return [_cell_slim(r) for r in results if _coord_allowed(r["coord_key"], restrictions, order)]


async def _save_cell(db, sheet_id: str, cell: CellIn):
    existing = await db.execute_fetchall(
        "SELECT id, value FROM cell_data WHERE sheet_id = ? AND coord_key = ?",
        (sheet_id, cell.coord_key),
    )
    old_value = existing[0]["value"] if existing else None

    from backend.db import intern_formula
    fid = await intern_formula(db, cell.formula) if cell.formula else None
    if existing:
        # Build dynamic update
        fields = ["value=?", "data_type=?"]
        params: list = [cell.value, cell.data_type]
        if cell.rule is not None:
            fields.append("rule=?")
            params.append(cell.rule)
        if cell.formula is not None:
            fields.append("formula=?")
            params.append("")
            fields.append("formula_id=?")
            params.append(fid)
        params.extend([sheet_id, cell.coord_key])
        await db.execute(
            f"UPDATE cell_data SET {', '.join(fields)} WHERE sheet_id=? AND coord_key=?",
            params,
        )
    else:
        cid = str(uuid.uuid4())
        await db.execute(
            "INSERT INTO cell_data (id, sheet_id, coord_key, value, data_type, rule, formula, formula_id) VALUES (?,?,?,?,?,?,?,?)",
            (cid, sheet_id, cell.coord_key, cell.value, cell.data_type,
             cell.rule or "manual", "" if fid else (cell.formula or ""), fid),
        )

    # Record history
    if old_value != cell.value:
        hid = str(uuid.uuid4())
        await db.execute(
            "INSERT INTO cell_history (id, sheet_id, coord_key, user_id, old_value, new_value) VALUES (?,?,?,?,?,?)",
            (hid, sheet_id, cell.coord_key, cell.user_id, old_value, cell.value),
        )


async def _materialize_sums(db, model_id: str,
                            sheet_ids: list[str] | None = None) -> int:
    """Build bottom-up sum_children aggregates for every parent coord_key.

    Coord keys are pipe-separated record IDs.  The key length varies — a cell
    may have 2 parts (period|indicator) or more when additional analytics are
    pinned/expanded.  We look up each record ID to find which analytic it
    belongs to, then aggregate child→parent within each position.

    If *sheet_ids* is given, only those sheets are re-materialized (much
    faster for single-cell edits).
    """
    import uuid as _uuid
    if sheet_ids:
        sheets = [{"id": sid} for sid in sheet_ids]
    else:
        sheets = await db.execute_fetchall(
            "SELECT id FROM sheets WHERE model_id = ?", (model_id,))
    total = 0

    # Clear all existing sum_children entries for this model's sheets
    for s in sheets:
        await db.execute(
            "DELETE FROM cell_data WHERE sheet_id = ? AND rule = 'sum_children'",
            (s["id"],),
        )

    for s in sheets:
        sid = s["id"]
        bindings = await db.execute_fetchall(
            "SELECT analytic_id FROM sheet_analytics WHERE sheet_id = ? ORDER BY sort_order",
            (sid,),
        )
        all_aids = [b["analytic_id"] for b in bindings]
        if len(all_aids) < 2:
            continue

        # Build parent_of map: rid → parent_rid (None if root)
        parent_of: dict[str, str | None] = {}
        rid_to_aid: dict[str, str] = {}  # rid → analytic_id
        for aid in all_aids:
            recs = await db.execute_fetchall(
                "SELECT id, parent_id FROM analytic_records WHERE analytic_id = ?",
                (aid,),
            )
            for r in recs:
                parent_of[r["id"]] = r["parent_id"]
                rid_to_aid[r["id"]] = aid

        # Load ALL non-aggregate cells
        cells = await db.execute_fetchall(
            "SELECT coord_key, value FROM cell_data WHERE sheet_id = ? AND rule != 'sum_children'",
            (sid,),
        )

        # Parse into dict
        cell_vals: dict[tuple, float] = {}
        for c in cells:
            parts = tuple(c["coord_key"].split("|"))
            if len(parts) < 2:
                continue
            try:
                v = float(c["value"]) if c["value"] not in (None, '') else 0.0
            except (ValueError, TypeError):
                continue
            cell_vals[parts] = v

        if not cell_vals:
            continue

        # ── Phase 1: within-dimension aggregation (bottom-up tree) ──
        # For each position, group cells by context (everything except that
        # position), then do a single bottom-up tree walk per group.
        # This is O(cells × positions) instead of O(cells² × depth).
        from collections import defaultdict, deque
        import time as _time

        aggregated: dict[tuple, float] = dict(cell_vals)
        parent_keys: set[tuple] = set()
        max_len = max(len(k) for k in aggregated)

        t_start = _time.monotonic()

        def _within_dim(keys_iter, klen: int | None = None):
            """Within-dimension aggregation using bottom-up tree walk."""
            # Build context groups for each position
            for pos in range(klen if klen else max_len):
                ctx_rids: dict[tuple, dict[str, float]] = defaultdict(lambda: defaultdict(float))
                for key, val in keys_iter():
                    if pos >= len(key):
                        continue
                    if klen is not None and len(key) != klen:
                        continue
                    ctx = key[:pos] + key[pos+1:]
                    ctx_rids[ctx][key[pos]] += val

                for ctx, rid_vals in ctx_rids.items():
                    # Find ancestors of rids in this group
                    node_vals: dict[str, float] = dict(rid_vals)
                    for rid in list(rid_vals.keys()):
                        r = parent_of.get(rid)
                        while r is not None and r not in node_vals:
                            node_vals[r] = 0.0
                            r = parent_of.get(r)

                    # Build local children map
                    local_ch: dict[str, list[str]] = defaultdict(list)
                    pending: dict[str, int] = {}
                    for rid in node_vals:
                        pid = parent_of.get(rid)
                        if pid is not None and pid in node_vals:
                            local_ch[pid].append(rid)
                    for rid in node_vals:
                        pending[rid] = len(local_ch.get(rid, []))

                    # Bottom-up BFS: leaves first
                    queue = deque(r for r, cnt in pending.items() if cnt == 0)
                    while queue:
                        rid = queue.popleft()
                        pid = parent_of.get(rid)
                        if pid is not None and pid in node_vals:
                            node_vals[pid] += node_vals[rid]
                            pending[pid] -= 1
                            if pending[pid] == 0:
                                queue.append(pid)

                    # Write parent keys into aggregated
                    for rid, val in node_vals.items():
                        if rid not in rid_vals:
                            key = ctx[:pos] + (rid,) + ctx[pos:]
                            aggregated[key] = val
                            parent_keys.add(key)

        _within_dim(lambda: aggregated.items())
        t_dim1 = _time.monotonic()

        # ── Phase 2: cross-dimension collapse ──
        # Collapse root-position tails to produce shorter prefix keys.
        for drop_pos in range(max_len - 1, 1, -1):
            collapse_sums: dict[tuple, float] = {}
            for key, val in aggregated.items():
                if len(key) != drop_pos + 1:
                    continue
                rid = key[drop_pos]
                if parent_of.get(rid) is not None:
                    continue
                prefix = key[:drop_pos]
                collapse_sums[prefix] = collapse_sums.get(prefix, 0.0) + val
            for pk, sv in collapse_sums.items():
                aggregated[pk] = sv
                parent_keys.add(pk)
            if collapse_sums:
                _within_dim(lambda cs=collapse_sums: ((k, v) for k, v in aggregated.items() if len(k) == drop_pos), klen=drop_pos)

        t_end = _time.monotonic()
        print(f"[materialize] sheet {sid[:8]}: {len(cell_vals)} cells → {len(aggregated)} aggregated "
              f"(dim={t_dim1-t_start:.2f}s, collapse={t_end-t_dim1:.2f}s)")

        # NOTE: sum_children rows are no longer persisted — they're recomputed
        # on demand. Count them for the response but skip writeback.
        agg_count = sum(
            1 for key in aggregated
            if key not in cell_vals or key in parent_keys
        )
        total += agg_count

    await db.commit()
    return total


async def _apply_engine_result(db, model_id: str, sheet_id: str, result: dict) -> int:
    """Write formula engine results to DB and materialize sums."""
    total = 0
    import uuid as _uuid
    for sid, changes in result.items():
        if not changes:
            continue
        rows = []
        for ck, val in changes.items():
            rule = 'empty' if val == '__empty__' else 'formula'
            db_val = '' if val == '__empty__' else val
            rows.append((str(_uuid.uuid4()), sid, ck, db_val, rule))
        await db.executemany(
            """INSERT INTO cell_data (id, sheet_id, coord_key, value, rule)
               VALUES (?, ?, ?, ?, ?)
               ON CONFLICT(sheet_id, coord_key) DO UPDATE SET value = excluded.value, rule = excluded.rule
               WHERE cell_data.rule != 'manual'""",
            rows,
        )
        total += len(changes)

    affected_sheets = [sheet_id]
    for sid, changes in result.items():
        if changes and sid not in affected_sheets:
            affected_sheets.append(sid)
    total += await _materialize_sums(db, model_id, sheet_ids=affected_sheets)
    return total


async def _update_dependents(db, sheet_id: str,
                             changed_cells: list[tuple[str, str, str]]) -> int:
    """Incrementally update formula dependents using cached DAG.
    If no DAG is cached, only materializes sums (no full rebuild).
    DAG rebuild is only triggered explicitly via calculate endpoint."""
    from backend.formula_engine import calculate_model_incremental
    sheet = await db.execute_fetchall("SELECT model_id FROM sheets WHERE id = ?", (sheet_id,))
    if not sheet:
        return 0
    model_id = sheet[0]["model_id"]

    result = await calculate_model_incremental(db, model_id, changed_cells)
    return await _apply_engine_result(db, model_id, sheet_id, result)


async def _recalc_model(db, sheet_id: str) -> int:
    """Full model recalculation — rebuilds DAG from scratch.
    Called explicitly via calculate endpoint / generate button."""
    from backend.formula_engine import calculate_model
    sheet = await db.execute_fetchall("SELECT model_id FROM sheets WHERE id = ?", (sheet_id,))
    if not sheet:
        return 0
    model_id = sheet[0]["model_id"]
    result = await calculate_model(db, model_id)
    return await _apply_engine_result(db, model_id, sheet_id, result)


@router.put("/by-sheet/{sheet_id}")
async def save_cells(sheet_id: str, body: BulkCellsIn, no_recalc: bool = Query(False)):
    db = get_db()
    # Check sheet lock
    lock_row = await db.execute_fetchall("SELECT locked FROM sheets WHERE id = ?", (sheet_id,))
    if lock_row and lock_row[0]["locked"]:
        raise HTTPException(403, "Sheet is locked")
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
    changed = []
    for cell in body.cells:
        await _save_cell(db, sheet_id, cell)
        if cell.value is not None and (cell.rule is None or cell.rule == "manual"):
            changed.append((sheet_id, cell.coord_key, cell.value))
    computed = 0
    if not no_recalc and changed:
        computed = await _update_dependents(db, sheet_id, changed_cells=changed)
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
    computed = 0
    if body.value is not None and (body.rule is None or body.rule == "manual"):
        changed = [(sheet_id, body.coord_key, body.value)]
        computed = await _update_dependents(db, sheet_id, changed)
    await db.commit()
    return {"ok": True, "computed": computed}


@router.post("/calculate/{sheet_id}")
async def calculate(sheet_id: str):
    """Recalculate all formula cells in the model (lazy pull, cross-sheet)."""
    db = get_db()
    computed = await _recalc_model(db, sheet_id)
    await db.commit()
    return {"computed": computed}


class MarkDirtyChange(BaseModel):
    sheet_id: str
    coord_key: str


class MarkDirtyIn(BaseModel):
    changes: list[MarkDirtyChange]


@router.post("/mark-dirty/{model_id}")
async def mark_dirty(model_id: str, body: MarkDirtyIn):
    """Return all cells transitively affected by the given changes (using cached DAG)."""
    from backend.formula_engine import get_dirty_cells
    db = get_db()
    changes = [(c.sheet_id, c.coord_key) for c in body.changes]
    dirty = await get_dirty_cells(db, model_id, changes)
    return {"dirty": [{"sheet_id": s, "coord_key": c} for s, c in dirty]}


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
                rule = 'empty' if val == '__empty__' else 'formula'
                db_val = '' if val == '__empty__' else val
                await db.execute(
                    """INSERT INTO cell_data (id, sheet_id, coord_key, value, rule)
                       VALUES (?, ?, ?, ?, ?)
                       ON CONFLICT(sheet_id, coord_key) DO UPDATE SET value = excluded.value, rule = excluded.rule""",
                    (str(__import__('uuid').uuid4()), sid, ck, db_val, rule),
                )
            total += len(changes)
            done_sheets += 1
            sheet_name = next((s["name"] for s in sheets if s["id"] == sid), sid)
            yield f"data: {json.dumps({'phase': 'sheet_done', 'sheet': sheet_name, 'done': done_sheets, 'total_sheets': total_sheets, 'computed': total})}\n\n"
            await asyncio.sleep(0)  # yield control

        await db.commit()

        # Materialize sum_children aggregates
        sum_count = await _materialize_sums(db, model_id)
        total += sum_count

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

    # Build a lookup: record_id → record name for all records in this model
    rec_names: dict[str, str] = {}
    name_rows = await db.execute_fetchall(
        """SELECT ar.id, json_extract(ar.data_json, '$.name') as name
           FROM analytic_records ar
           JOIN analytics a ON a.id = ar.analytic_id
           WHERE a.model_id = ?""",
        (model_id,),
    )
    for nr in name_rows:
        rec_names[nr["id"]] = nr["name"] or nr["id"][:8]

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
    # Sort all rows by created_at DESC across sheets
    rows.sort(key=lambda r: r["created_at"], reverse=True)
    result = []
    for r in rows:
        if r["sheet_id"] in sheet_names:
            d = dict(r)
            d["sheet_name"] = sheet_names[r["sheet_id"]]
            # Resolve coord_key parts to human-readable names
            parts = (r["coord_key"] or "").split("|")
            d["description"] = " · ".join(
                rec_names.get(p, p[:8]) for p in parts if p
            )
            result.append(d)
            if len(result) >= limit:
                break
    return result


class UndoIn(BaseModel):
    history_id: str | None = None  # undo up to and including this entry; None = undo latest


@router.post("/undo/{model_id}")
async def undo(model_id: str, body: UndoIn):
    """Undo changes from most recent back to history_id (inclusive)."""
    db = get_db()

    # If no history_id given, find the latest entry for this model
    history_id = body.history_id
    if not history_id:
        sheet_ids = await db.execute_fetchall("SELECT id FROM sheets WHERE model_id = ?", (model_id,))
        latest = None
        for sr in sheet_ids:
            rows = await db.execute_fetchall(
                "SELECT id, created_at FROM cell_history WHERE sheet_id = ? ORDER BY created_at DESC LIMIT 1",
                (sr["id"],))
            if rows and (latest is None or rows[0]["created_at"] > latest["created_at"]):
                latest = rows[0]
        if not latest:
            return {"error": "No history to undo"}
        history_id = latest["id"]

    # Get target timestamp
    target = await db.execute_fetchall("SELECT created_at FROM cell_history WHERE id = ?", (history_id,))
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
    changes = []
    for r in rows:
        await db.execute(
            "UPDATE cell_data SET value = ? WHERE sheet_id = ? AND coord_key = ?",
            (r["old_value"], r["sheet_id"], r["coord_key"]),
        )
        await db.execute("DELETE FROM cell_history WHERE id = ?", (r["id"],))
        changes.append({"sheet_id": r["sheet_id"], "coord_key": r["coord_key"], "value": r["old_value"]})
        undone += 1

    # Incremental recalc for undone cells
    from backend.formula_engine import calculate_model_incremental
    sheet = await db.execute_fetchall("SELECT model_id FROM sheets WHERE id = ?", (rows[0]["sheet_id"],))
    model_id_for_recalc = sheet[0]["model_id"] if sheet else model_id
    changed_tuples = [(c["sheet_id"], c["coord_key"], c["value"] or "") for c in changes]
    recalc_result = await calculate_model_incremental(db, model_id_for_recalc, changed_tuples)
    computed = 0
    import uuid as _uuid
    for sid, recalc_changes in recalc_result.items():
        if not recalc_changes:
            continue
        rows_to_write = []
        for ck, val in recalc_changes.items():
            rule = 'empty' if val == '__empty__' else 'formula'
            db_val = '' if val == '__empty__' else val
            rows_to_write.append((str(_uuid.uuid4()), sid, ck, db_val, rule))
        await db.executemany(
            """INSERT INTO cell_data (id, sheet_id, coord_key, value, rule)
               VALUES (?, ?, ?, ?, ?)
               ON CONFLICT(sheet_id, coord_key) DO UPDATE SET value = excluded.value, rule = excluded.rule
               WHERE cell_data.rule != 'manual'""",
            rows_to_write,
        )
        computed += len(recalc_changes)
    # Materialize sums for affected sheets
    affected = list({c["sheet_id"] for c in changes})
    for sid, ch in recalc_result.items():
        if ch and sid not in affected:
            affected.append(sid)
    computed += await _materialize_sums(db, model_id_for_recalc, sheet_ids=affected)
    await db.commit()

    # Build minimal all_cells: only undone cells + recomputed cells
    all_cells: dict[str, str | None] = {}
    for c in changes:
        all_cells[c["coord_key"]] = c["value"]
    for sid, recalc_changes in recalc_result.items():
        for ck, val in recalc_changes.items():
            all_cells[ck] = '' if val == '__empty__' else val

    # Check if there's more history remaining
    has_more = False
    sheet_ids_check = await db.execute_fetchall("SELECT id FROM sheets WHERE model_id = ?", (model_id,))
    for sr in sheet_ids_check:
        remaining = await db.execute_fetchall(
            "SELECT 1 FROM cell_history WHERE sheet_id = ? LIMIT 1", (sr["id"],))
        if remaining:
            has_more = True
            break

    return {"undone": undone, "computed": computed, "changes": changes, "all_cells": all_cells, "has_more": has_more}


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
