import uuid
import json
from datetime import date, timedelta
from calendar import monthrange
from fastapi import APIRouter
from pydantic import BaseModel
from backend.db import get_db
from backend.transliterate import transliterate

router = APIRouter(prefix="/api/analytics", tags=["analytics"])


async def _invalidate_engines_for_analytic(db, analytic_id: str):
    """Invalidate V4 engine cache for all models using this analytic."""
    from backend.formula_engine import invalidate_engine
    rows = await db.execute_fetchall(
        """SELECT DISTINCT s.model_id FROM sheet_analytics sa
           JOIN sheets s ON s.id = sa.sheet_id
           WHERE sa.analytic_id = ?""",
        (analytic_id,),
    )
    for r in rows:
        await invalidate_engine(db, r["model_id"])

MONTH_NAMES_RU = [
    "Январь", "Февраль", "Март", "Апрель", "Май", "Июнь",
    "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь",
]
QUARTER_NAMES_RU = ["1-й квартал", "2-й квартал", "3-й квартал", "4-й квартал"]


class AnalyticIn(BaseModel):
    model_id: str | None = None
    name: str = ""
    code: str | None = None
    icon: str = ""
    is_periods: bool = False
    data_type: str = "sum"  # sum | percent | string | quantity
    period_types: list[str] = []
    period_start: str | None = None
    period_end: str | None = None
    sort_order: int = 0


class FieldIn(BaseModel):
    name: str = ""
    code: str | None = None
    data_type: str = "string"
    sort_order: int = 0


class RecordIn(BaseModel):
    parent_id: str | None = None
    sort_order: int = 0
    data_json: dict = {}


# --- Analytics CRUD ---

@router.get("/by-model/{model_id}")
async def list_analytics(model_id: str):
    db = get_db()
    rows = await db.execute_fetchall(
        "SELECT * FROM analytics WHERE model_id = ? ORDER BY sort_order", (model_id,)
    )
    return [dict(r) for r in rows]


@router.post("")
async def create_analytic(body: AnalyticIn):
    db = get_db()
    aid = str(uuid.uuid4())
    code = body.code or transliterate(body.name)
    await db.execute(
        """INSERT INTO analytics (id, model_id, name, code, icon, is_periods, data_type, period_types,
           period_start, period_end, sort_order)
           VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
        (aid, body.model_id, body.name, code, body.icon, int(body.is_periods), body.data_type,
         json.dumps(body.period_types), body.period_start, body.period_end, body.sort_order),
    )
    if body.is_periods:
        await _ensure_period_fields(db, aid)
    else:
        # Add default "Наименование" field
        fid = str(uuid.uuid4())
        await db.execute(
            "INSERT INTO analytic_fields (id, analytic_id, name, code, data_type, sort_order) VALUES (?,?,?,?,?,?)",
            (fid, aid, "Наименование", "name", "string", 0),
        )
    await db.commit()
    row = await db.execute_fetchall("SELECT * FROM analytics WHERE id = ?", (aid,))
    return dict(row[0])


@router.get("/{analytic_id}")
async def get_analytic(analytic_id: str):
    db = get_db()
    rows = await db.execute_fetchall("SELECT * FROM analytics WHERE id = ?", (analytic_id,))
    if not rows:
        return {"error": "not found"}
    return dict(rows[0])


@router.put("/{analytic_id}")
async def update_analytic(analytic_id: str, body: AnalyticIn):
    db = get_db()
    code = body.code or transliterate(body.name)
    await db.execute(
        """UPDATE analytics SET name=?, code=?, icon=?, is_periods=?, data_type=?, period_types=?,
           period_start=?, period_end=?, sort_order=?, updated_at=datetime('now')
           WHERE id=?""",
        (body.name, code, body.icon, int(body.is_periods), body.data_type, json.dumps(body.period_types),
         body.period_start, body.period_end, body.sort_order, analytic_id),
    )
    if body.is_periods:
        await _ensure_period_fields(db, analytic_id)
    await db.commit()
    row = await db.execute_fetchall("SELECT * FROM analytics WHERE id = ?", (analytic_id,))
    return dict(row[0])


@router.delete("/{analytic_id}")
async def delete_analytic(analytic_id: str):
    db = get_db()
    await db.execute("DELETE FROM analytics WHERE id = ?", (analytic_id,))
    await db.commit()
    return {"ok": True}


# --- Fields ---

@router.get("/{analytic_id}/fields")
async def list_fields(analytic_id: str):
    db = get_db()
    rows = await db.execute_fetchall(
        "SELECT * FROM analytic_fields WHERE analytic_id = ? ORDER BY sort_order",
        (analytic_id,),
    )
    return [dict(r) for r in rows]


@router.post("/{analytic_id}/fields")
async def create_field(analytic_id: str, body: FieldIn):
    db = get_db()
    fid = str(uuid.uuid4())
    code = body.code or transliterate(body.name)
    await db.execute(
        "INSERT INTO analytic_fields (id, analytic_id, name, code, data_type, sort_order) VALUES (?,?,?,?,?,?)",
        (fid, analytic_id, body.name, code, body.data_type, body.sort_order),
    )
    await db.commit()
    row = await db.execute_fetchall("SELECT * FROM analytic_fields WHERE id = ?", (fid,))
    return dict(row[0])


@router.put("/{analytic_id}/fields/{field_id}")
async def update_field(analytic_id: str, field_id: str, body: FieldIn):
    db = get_db()
    code = body.code or transliterate(body.name)
    await db.execute(
        "UPDATE analytic_fields SET name=?, code=?, data_type=?, sort_order=? WHERE id=? AND analytic_id=?",
        (body.name, code, body.data_type, body.sort_order, field_id, analytic_id),
    )
    await db.commit()
    row = await db.execute_fetchall("SELECT * FROM analytic_fields WHERE id = ?", (field_id,))
    return dict(row[0])


@router.delete("/{analytic_id}/fields/{field_id}")
async def delete_field(analytic_id: str, field_id: str):
    db = get_db()
    await db.execute("DELETE FROM analytic_fields WHERE id = ? AND analytic_id = ?", (field_id, analytic_id))
    await db.commit()
    return {"ok": True}


# --- Records ---

@router.get("/{analytic_id}/records")
async def list_records(analytic_id: str):
    db = get_db()
    rows = await db.execute_fetchall(
        "SELECT * FROM analytic_records WHERE analytic_id = ? ORDER BY sort_order",
        (analytic_id,),
    )
    return [dict(r) for r in rows]


@router.post("/{analytic_id}/records")
async def create_record(analytic_id: str, body: RecordIn):
    db = get_db()
    rid = str(uuid.uuid4())
    await db.execute(
        "INSERT INTO analytic_records (id, analytic_id, parent_id, sort_order, data_json) VALUES (?,?,?,?,?)",
        (rid, analytic_id, body.parent_id, body.sort_order, json.dumps(body.data_json, ensure_ascii=False)),
    )
    await db.commit()
    await _invalidate_engines_for_analytic(db, analytic_id)
    row = await db.execute_fetchall("SELECT * FROM analytic_records WHERE id = ?", (rid,))
    return dict(row[0])


@router.put("/{analytic_id}/records/{record_id}")
async def update_record(analytic_id: str, record_id: str, body: RecordIn):
    db = get_db()
    await db.execute(
        "UPDATE analytic_records SET parent_id=?, sort_order=?, data_json=? WHERE id=?",
        (body.parent_id, body.sort_order, json.dumps(body.data_json, ensure_ascii=False),
         record_id),
    )
    await db.commit()
    await _invalidate_engines_for_analytic(db, analytic_id)
    row = await db.execute_fetchall("SELECT * FROM analytic_records WHERE id = ?", (record_id,))
    if not row:
        return {"ok": True}
    return dict(row[0])


@router.delete("/{analytic_id}/records/{record_id}")
async def delete_record(analytic_id: str, record_id: str):
    db = get_db()
    await db.execute("DELETE FROM analytic_records WHERE id = ? AND analytic_id = ?", (record_id, analytic_id))
    await db.commit()
    await _invalidate_engines_for_analytic(db, analytic_id)
    return {"ok": True}


@router.post("/{analytic_id}/records/bulk")
async def bulk_upsert_records(analytic_id: str, records: list[RecordIn]):
    db = get_db()
    created = []
    for rec in records:
        rid = str(uuid.uuid4())
        await db.execute(
            "INSERT INTO analytic_records (id, analytic_id, parent_id, sort_order, data_json) VALUES (?,?,?,?,?)",
            (rid, analytic_id, rec.parent_id, rec.sort_order, json.dumps(rec.data_json, ensure_ascii=False)),
        )
        created.append(rid)
    await db.commit()
    await _invalidate_engines_for_analytic(db, analytic_id)
    return {"created": created}


# --- Period generation ---

@router.post("/{analytic_id}/generate-periods")
async def generate_periods(analytic_id: str):
    db = get_db()
    rows = await db.execute_fetchall("SELECT * FROM analytics WHERE id = ?", (analytic_id,))
    if not rows:
        return {"error": "not found"}
    a = dict(rows[0])
    if not a["is_periods"]:
        return {"error": "not a period analytic"}

    period_types = json.loads(a["period_types"]) if a["period_types"] else []
    if not period_types or not a["period_start"] or not a["period_end"]:
        return {"error": "period config incomplete"}

    start = date.fromisoformat(a["period_start"])
    end = date.fromisoformat(a["period_end"])

    # Clear existing records
    await db.execute("DELETE FROM analytic_records WHERE analytic_id = ?", (analytic_id,))

    has_year = "year" in period_types
    has_quarter = "quarter" in period_types
    has_month = "month" in period_types

    sort = 0
    year = start.year
    while year <= end.year:
        year_start = max(start, date(year, 1, 1))
        year_end = min(end, date(year, 12, 31))

        year_id = None
        if has_year:
            year_id = str(uuid.uuid4())
            await db.execute(
                "INSERT INTO analytic_records (id, analytic_id, parent_id, sort_order, data_json) VALUES (?,?,?,?,?)",
                (year_id, analytic_id, None, sort, json.dumps(
                    {"name": str(year), "start": str(year_start), "end": str(year_end)},
                    ensure_ascii=False)),
            )
            sort += 1

        for q in range(4):
            q_start_month = q * 3 + 1
            q_end_month = q * 3 + 3
            q_start = date(year, q_start_month, 1)
            q_end = date(year, q_end_month, monthrange(year, q_end_month)[1])
            if q_start > end or q_end < start:
                continue
            q_start = max(q_start, start)
            q_end = min(q_end, end)

            quarter_id = None
            if has_quarter:
                quarter_id = str(uuid.uuid4())
                await db.execute(
                    "INSERT INTO analytic_records (id, analytic_id, parent_id, sort_order, data_json) VALUES (?,?,?,?,?)",
                    (quarter_id, analytic_id, year_id, sort, json.dumps(
                        {"name": f"{QUARTER_NAMES_RU[q]} {year}", "start": str(q_start), "end": str(q_end)},
                        ensure_ascii=False)),
                )
                sort += 1

            if has_month:
                for m in range(q_start_month, q_end_month + 1):
                    m_start = date(year, m, 1)
                    m_end = date(year, m, monthrange(year, m)[1])
                    if m_start > end or m_end < start:
                        continue
                    m_start = max(m_start, start)
                    m_end = min(m_end, end)
                    mid = str(uuid.uuid4())
                    parent = quarter_id if quarter_id else year_id
                    await db.execute(
                        "INSERT INTO analytic_records (id, analytic_id, parent_id, sort_order, data_json) VALUES (?,?,?,?,?)",
                        (mid, analytic_id, parent, sort, json.dumps(
                            {"name": f"{MONTH_NAMES_RU[m - 1]} {year}", "start": str(m_start), "end": str(m_end)},
                            ensure_ascii=False)),
                    )
                    sort += 1

        year += 1

    await db.commit()
    records = await db.execute_fetchall(
        "SELECT * FROM analytic_records WHERE analytic_id = ? ORDER BY sort_order",
        (analytic_id,),
    )
    return [dict(r) for r in records]


async def _ensure_period_fields(db, analytic_id: str):
    """Ensure the three predefined fields exist for a period analytic."""
    existing = await db.execute_fetchall(
        "SELECT code FROM analytic_fields WHERE analytic_id = ?", (analytic_id,)
    )
    codes = {r["code"] for r in existing}
    predefined = [
        ("Наименование", "name", "string", 0),
        ("Начало", "start", "date", 1),
        ("Окончание", "end", "date", 2),
    ]
    for name, code, dtype, order in predefined:
        if code not in codes:
            fid = str(uuid.uuid4())
            await db.execute(
                "INSERT INTO analytic_fields (id, analytic_id, name, code, data_type, sort_order) VALUES (?,?,?,?,?,?)",
                (fid, analytic_id, name, code, dtype, order),
            )
