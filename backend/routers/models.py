import uuid
import asyncio
import json
from fastapi import APIRouter
from fastapi.responses import StreamingResponse
from pydantic import BaseModel
from backend.db import get_db

router = APIRouter(prefix="/api/models", tags=["models"])


class ModelIn(BaseModel):
    name: str = ""
    description: str = ""


@router.get("")
async def list_models():
    db = get_db()
    rows = await db.execute_fetchall("SELECT * FROM models ORDER BY created_at")
    return [dict(r) for r in rows]


@router.post("")
async def create_model(body: ModelIn):
    db = get_db()
    mid = str(uuid.uuid4())
    await db.execute(
        "INSERT INTO models (id, name, description) VALUES (?, ?, ?)",
        (mid, body.name, body.description),
    )
    await db.commit()
    row = await db.execute_fetchall("SELECT * FROM models WHERE id = ?", (mid,))
    return dict(row[0])


@router.put("/{model_id}")
async def update_model(model_id: str, body: ModelIn):
    db = get_db()
    await db.execute(
        "UPDATE models SET name=?, description=?, updated_at=datetime('now') WHERE id=?",
        (body.name, body.description, model_id),
    )
    await db.commit()
    row = await db.execute_fetchall("SELECT * FROM models WHERE id = ?", (model_id,))
    return dict(row[0])


@router.delete("/{model_id}")
async def delete_model(model_id: str):
    db = get_db()
    # Speed up bulk deletion: disable synchronous writes and use in-memory journal.
    await db.execute("PRAGMA synchronous = OFF")
    await db.execute("PRAGMA journal_mode = MEMORY")
    try:
        sheet_sub = "SELECT id FROM sheets WHERE model_id = ?"
        analytic_sub = "SELECT id FROM analytics WHERE model_id = ?"
        await db.execute(f"DELETE FROM cell_data WHERE sheet_id IN ({sheet_sub})", (model_id,))
        await db.execute(f"DELETE FROM cell_history WHERE sheet_id IN ({sheet_sub})", (model_id,))
        await db.execute(f"DELETE FROM indicator_formula_rules WHERE sheet_id IN ({sheet_sub})", (model_id,))
        await db.execute(f"DELETE FROM dag_cache WHERE sheet_id IN ({sheet_sub})", (model_id,))
        await db.execute(f"DELETE FROM sheet_analytics WHERE sheet_id IN ({sheet_sub})", (model_id,))
        await db.execute(f"DELETE FROM sheet_permissions WHERE sheet_id IN ({sheet_sub})", (model_id,))
        await db.execute(f"DELETE FROM sheet_view_settings WHERE sheet_id IN ({sheet_sub})", (model_id,))
        await db.execute("DELETE FROM sheets WHERE model_id = ?", (model_id,))
        await db.execute(f"DELETE FROM analytic_record_permissions WHERE analytic_id IN ({analytic_sub})", (model_id,))
        await db.execute(f"DELETE FROM analytic_records WHERE analytic_id IN ({analytic_sub})", (model_id,))
        await db.execute(f"DELETE FROM analytic_fields WHERE analytic_id IN ({analytic_sub})", (model_id,))
        await db.execute("DELETE FROM analytics WHERE model_id = ?", (model_id,))
        await db.execute("DELETE FROM models WHERE id = ?", (model_id,))
        await db.commit()
    finally:
        await db.execute("PRAGMA synchronous = FULL")
        await db.execute("PRAGMA journal_mode = WAL")
    return {"ok": True}


@router.delete("")
async def delete_all_models():
    """Fast-path: drop all data by clearing tables (no per-row index updates)."""
    db = get_db()
    tables = [
        "cell_data", "cell_history", "indicator_formula_rules", "dag_cache",
        "sheet_analytics", "sheet_permissions", "sheet_view_settings", "sheets",
        "analytic_record_permissions", "analytic_records", "analytic_fields",
        "analytics", "models",
    ]
    for t in tables:
        await db.execute(f"DELETE FROM {t}")
    await db.commit()
    # Reclaim disk space in background (non-blocking for the response).
    try:
        await db.execute("VACUUM")
    except Exception:
        pass
    return {"ok": True}


@router.post("/{model_id}/generate")
async def generate_model(model_id: str):
    """Full DAG rebuild + recalculation. Called explicitly by user via Generate button.
    Returns SSE stream with progress."""
    from backend.formula_engine import calculate_model, invalidate_engine
    from backend.routers.cells import _materialize_sums

    async def event_stream():
        db = get_db()

        # Mark as generating
        await db.execute(
            "UPDATE models SET calc_status = 'generating' WHERE id = ?",
            (model_id,),
        )
        await db.commit()

        try:
            sheets = await db.execute_fetchall(
                "SELECT id, name FROM sheets WHERE model_id = ? ORDER BY created_at",
                (model_id,),
            )
            yield f"data: {json.dumps({'phase': 'start', 'total_sheets': len(sheets)})}\n\n"

            # Invalidate to force full rebuild
            await invalidate_engine(db, model_id)
            # Reset status back to generating (invalidate sets needs_generation)
            await db.execute(
                "UPDATE models SET calc_status = 'generating' WHERE id = ?",
                (model_id,),
            )
            await db.commit()

            yield f"data: {json.dumps({'phase': 'building_dag'})}\n\n"
            await asyncio.sleep(0)

            result = await calculate_model(db, model_id)

            # Write results to DB
            total = 0
            for sid, changes in result.items():
                if not changes:
                    continue
                rows = []
                for ck, val in changes.items():
                    rule = 'empty' if val == '__empty__' else 'formula'
                    db_val = '' if val == '__empty__' else val
                    rows.append((str(uuid.uuid4()), sid, ck, db_val, rule))
                await db.executemany(
                    """INSERT INTO cell_data (id, sheet_id, coord_key, value, rule)
                       VALUES (?, ?, ?, ?, ?)
                       ON CONFLICT(sheet_id, coord_key) DO UPDATE SET value = excluded.value, rule = excluded.rule
                       WHERE cell_data.rule != 'manual'""",
                    rows,
                )
                total += len(changes)
                sheet_name = next((s["name"] for s in sheets if s["id"] == sid), sid)
                yield f"data: {json.dumps({'phase': 'sheet_done', 'sheet': sheet_name, 'computed': total})}\n\n"
                await asyncio.sleep(0)

            # Materialize sums
            yield f"data: {json.dumps({'phase': 'materializing'})}\n\n"
            await asyncio.sleep(0)
            sum_count = await _materialize_sums(db, model_id)
            total += sum_count
            await db.commit()

            # Mark as ready
            await db.execute(
                "UPDATE models SET calc_status = 'ready' WHERE id = ?",
                (model_id,),
            )
            await db.commit()

            yield f"data: {json.dumps({'phase': 'done', 'computed': total})}\n\n"

        except Exception as e:
            # Mark as error
            await db.execute(
                "UPDATE models SET calc_status = 'error' WHERE id = ?",
                (model_id,),
            )
            await db.commit()
            yield f"data: {json.dumps({'phase': 'error', 'message': str(e)})}\n\n"

    return StreamingResponse(event_stream(), media_type="text/event-stream")


@router.get("/{model_id}/calc-status")
async def get_calc_status(model_id: str):
    db = get_db()
    rows = await db.execute_fetchall(
        "SELECT calc_status FROM models WHERE id = ?", (model_id,),
    )
    if not rows:
        return {"error": "not found"}
    return {"calc_status": rows[0]["calc_status"] or "ready"}


@router.get("/{model_id}/tree")
async def get_model_tree(model_id: str):
    db = get_db()
    model_rows = await db.execute_fetchall("SELECT * FROM models WHERE id = ?", (model_id,))
    if not model_rows:
        return {"error": "not found"}
    model = dict(model_rows[0])
    sheets = await db.execute_fetchall(
        "SELECT * FROM sheets WHERE model_id = ? ORDER BY sort_order, created_at", (model_id,)
    )
    analytics = await db.execute_fetchall(
        "SELECT * FROM analytics WHERE model_id = ? ORDER BY sort_order", (model_id,)
    )
    return {
        **model,
        "sheets": [dict(s) for s in sheets],
        "analytics": [dict(a) for a in analytics],
    }
