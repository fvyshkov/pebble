import uuid
import json
import io
from fastapi import APIRouter, UploadFile, File
from fastapi.responses import StreamingResponse
from openpyxl import Workbook, load_workbook
from backend.db import get_db

router = APIRouter(prefix="/api/excel", tags=["excel"])

INDENT = 4  # spaces per hierarchy level


@router.get("/analytics/{analytic_id}/export")
async def export_analytic_records(analytic_id: str):
    db = get_db()
    fields = await db.execute_fetchall(
        "SELECT * FROM analytic_fields WHERE analytic_id = ? ORDER BY sort_order",
        (analytic_id,),
    )
    records = await db.execute_fetchall(
        "SELECT * FROM analytic_records WHERE analytic_id = ? ORDER BY sort_order",
        (analytic_id,),
    )
    analytic = await db.execute_fetchall("SELECT name FROM analytics WHERE id = ?", (analytic_id,))
    name = analytic[0]["name"] if analytic else "export"

    fields = [dict(f) for f in fields]
    records = [dict(r) for r in records]

    # Build tree
    by_id = {r["id"]: r for r in records}
    children = {}
    roots = []
    for r in records:
        pid = r["parent_id"]
        if pid:
            children.setdefault(pid, []).append(r)
        else:
            roots.append(r)

    wb = Workbook()
    ws = wb.active
    ws.title = name[:31]

    # Header
    headers = [f["name"] for f in fields]
    ws.append(headers)

    def write_rows(nodes, level):
        for node in nodes:
            data = json.loads(node["data_json"]) if isinstance(node["data_json"], str) else node["data_json"]
            row = []
            for i, f in enumerate(fields):
                val = data.get(f["code"], "")
                if i == 0 and level > 0:
                    val = " " * (level * INDENT) + str(val)
                row.append(val)
            ws.append(row)
            if node["id"] in children:
                write_rows(children[node["id"]], level + 1)

    write_rows(roots, 0)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)

    return StreamingResponse(
        buf,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{name}.xlsx"'},
    )


@router.post("/analytics/{analytic_id}/import")
async def import_analytic_records(analytic_id: str, file: UploadFile = File(...)):
    db = get_db()
    fields = await db.execute_fetchall(
        "SELECT * FROM analytic_fields WHERE analytic_id = ? ORDER BY sort_order",
        (analytic_id,),
    )
    fields = [dict(f) for f in fields]

    content = await file.read()
    wb = load_workbook(io.BytesIO(content))
    ws = wb.active

    rows = list(ws.iter_rows(min_row=2, values_only=True))  # skip header

    # Clear existing records
    await db.execute("DELETE FROM analytic_records WHERE analytic_id = ?", (analytic_id,))

    parent_stack: list[tuple[int, str]] = []  # (level, record_id)
    sort = 0

    for row_vals in rows:
        if not row_vals or all(v is None for v in row_vals):
            continue

        first_val = str(row_vals[0]) if row_vals[0] is not None else ""
        stripped = first_val.lstrip(" ")
        spaces = len(first_val) - len(stripped)
        level = spaces // max(INDENT, 1) if spaces > 0 else 0

        data = {}
        for i, f in enumerate(fields):
            val = row_vals[i] if i < len(row_vals) else None
            if i == 0 and val is not None:
                val = str(val).lstrip(" ")
            if val is not None:
                data[f["code"]] = val

        # Determine parent
        while parent_stack and parent_stack[-1][0] >= level:
            parent_stack.pop()
        parent_id = parent_stack[-1][1] if parent_stack else None

        rid = str(uuid.uuid4())
        await db.execute(
            "INSERT INTO analytic_records (id, analytic_id, parent_id, sort_order, data_json) VALUES (?,?,?,?,?)",
            (rid, analytic_id, parent_id, sort, json.dumps(data, ensure_ascii=False)),
        )
        parent_stack.append((level, rid))
        sort += 1

    await db.commit()
    records = await db.execute_fetchall(
        "SELECT * FROM analytic_records WHERE analytic_id = ? ORDER BY sort_order",
        (analytic_id,),
    )
    return [dict(r) for r in records]
