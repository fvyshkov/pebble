"""Import an Excel workbook as a Pebble model.

Detects periods from date header rows, builds indicator hierarchies from row
labels, and classifies cells as manual (input, theme=7 fill) or computed.
"""

import uuid
import json
import io
from datetime import datetime

from fastapi import APIRouter, UploadFile, File, Form
from openpyxl import load_workbook
from backend.db import get_db

router = APIRouter(prefix="/api/import", tags=["import"])

MONTH_NAMES_RU = [
    "", "Январь", "Февраль", "Март", "Апрель", "Май", "Июнь",
    "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь",
]


def _is_input_cell(cell) -> bool:
    """Check if cell has beige/yellow input background (theme=7)."""
    fill = cell.fill
    if fill and fill.fgColor and fill.fgColor.theme == 7:
        return True
    return False


def _detect_periods(ws, max_col: int) -> list[dict]:
    """Detect period columns from date headers in the first 6 rows."""
    periods = []
    date_row = None
    label_row = None
    data_start_col = None

    # Find the row with dates and the row with labels (m1, m2, ...)
    for r in range(1, 7):
        for c in range(1, min(max_col + 1, 50)):
            v = ws.cell(r, c).value
            if isinstance(v, datetime):
                date_row = r
                if data_start_col is None:
                    data_start_col = c
                break
            if isinstance(v, str) and v.startswith("m") and v[1:].isdigit():
                label_row = r
                if data_start_col is None:
                    data_start_col = c
                break

    if data_start_col is None:
        return []

    # Read all period columns
    for c in range(data_start_col, max_col + 1):
        date_val = ws.cell(date_row, c).value if date_row else None
        label_val = ws.cell(label_row, c).value if label_row else None
        if date_val is None and label_val is None:
            continue
        name = ""
        if isinstance(date_val, datetime):
            name = f"{MONTH_NAMES_RU[date_val.month]} {date_val.year}"
        elif label_val:
            name = str(label_val)
        periods.append({"col": c, "name": name, "date": date_val})

    return periods


def _detect_data_start_row(ws) -> int:
    """Find the first data row (after header rows with dates/labels)."""
    for r in range(1, 10):
        v = ws.cell(r, 1).value
        if v and not isinstance(v, datetime):
            # Check if it looks like a header row
            val = str(v).strip()
            if val and val not in ("", "ЕИ", "Отв.исп."):
                return r
    return 7


def _detect_indicators(ws, data_start_row: int, label_col: int, max_row: int) -> list[dict]:
    """Build hierarchical indicator list from row labels.

    Returns list of {name, unit, row, children: [...], is_group}.
    Groups are detected by: row with a name but no data in period columns,
    or rows that act as section headers.
    """
    indicators = []
    current_group = None

    for r in range(data_start_row, max_row + 1):
        name = ws.cell(r, label_col).value
        if name is None or str(name).strip() == "":
            # Empty row = end of current group context
            if current_group:
                indicators.append(current_group)
                current_group = None
            continue

        name = str(name).strip()
        unit_col = label_col + 1
        unit = ws.cell(r, unit_col).value
        unit = str(unit).strip() if unit else ""

        # Check if this looks like a group header (has "ЕИ" or "Отв.исп." pattern, or is bold)
        next_unit = unit
        is_header = unit in ("ЕИ", "") and ws.cell(r, unit_col + 1).value in ("Отв.исп.", None, "")

        # Check if there's actual data in period columns
        has_data = False
        for c_offset in range(3, 8):
            test_col = label_col + c_offset
            if test_col <= ws.max_column and ws.cell(r, test_col).value is not None:
                has_data = True
                break

        if is_header and not has_data:
            # This is a group/section header
            if current_group:
                indicators.append(current_group)
            current_group = {"name": name, "unit": "", "row": r, "children": [], "is_group": True}
        elif current_group:
            current_group["children"].append({"name": name, "unit": unit, "row": r, "children": [], "is_group": False})
        else:
            indicators.append({"name": name, "unit": unit, "row": r, "children": [], "is_group": False})

    if current_group:
        indicators.append(current_group)

    return indicators


@router.post("/excel")
async def import_excel(file: UploadFile = File(...), model_name: str = Form("Imported Model")):
    db = get_db()

    # Ensure unique model name
    existing = await db.execute_fetchall("SELECT id FROM models WHERE name = ?", (model_name,))
    if existing:
        model_name = f"{model_name} ({datetime.now().strftime('%Y-%m-%d %H:%M')})"

    # Create model
    model_id = str(uuid.uuid4())
    await db.execute(
        "INSERT INTO models (id, name, description) VALUES (?, ?, ?)",
        (model_id, model_name, "Imported from Excel"),
    )

    content = await file.read()
    # Load twice: with formulas for detection, data_only for values
    wb_formulas = load_workbook(io.BytesIO(content))
    wb_data = load_workbook(io.BytesIO(content), data_only=True)

    # === Detect periods from first sheet ===
    first_ws = wb_formulas[wb_formulas.sheetnames[0]]
    periods = _detect_periods(first_ws, min(first_ws.max_column, 200))

    if not periods:
        await db.commit()
        return {"model_id": model_id, "model_name": model_name, "sheets": 0, "error": "No periods detected"}

    data_start_col = periods[0]["col"]

    # Create period analytic
    period_analytic_id = str(uuid.uuid4())
    await db.execute(
        "INSERT INTO analytics (id, model_id, name, code, icon, is_periods, data_type, period_types) VALUES (?,?,?,?,?,?,?,?)",
        (period_analytic_id, model_id, "Периоды", "periods", "CalendarMonthOutlined", 1, "sum", '["month"]'),
    )

    # Create period fields
    for fname, fcode, ftype in [("name", "name", "string"), ("start", "start", "date"), ("end", "end", "date")]:
        fid = str(uuid.uuid4())
        await db.execute(
            "INSERT INTO analytic_fields (id, analytic_id, name, code, data_type, sort_order) VALUES (?,?,?,?,?,?)",
            (fid, period_analytic_id, fname, fcode, ftype, 0),
        )

    # Create period records
    period_record_ids = {}  # col -> record_id
    for i, p in enumerate(periods):
        rid = str(uuid.uuid4())
        data = {"name": p["name"]}
        if p.get("date"):
            d = p["date"]
            data["start"] = d.strftime("%Y-%m-%d") if isinstance(d, datetime) else str(d)[:10]
        await db.execute(
            "INSERT INTO analytic_records (id, analytic_id, parent_id, sort_order, data_json) VALUES (?,?,?,?,?)",
            (rid, period_analytic_id, None, i, json.dumps(data, ensure_ascii=False)),
        )
        period_record_ids[p["col"]] = rid

    # === Process each Excel sheet ===
    created_sheets = []

    for sheet_name in wb_formulas.sheetnames:
        ws_f = wb_formulas[sheet_name]
        ws_d = wb_data[sheet_name]

        # Detect periods for this sheet (may differ in start col)
        sheet_periods = _detect_periods(ws_f, min(ws_f.max_column, 200))
        if not sheet_periods:
            continue
        sheet_data_col = sheet_periods[0]["col"]

        # Determine label column(s)
        # Most sheets: col A = indicator name, col B = unit
        # OPEX+CAPEX: col A = product, col B = expense type, col C = line item
        label_col = 1
        if sheet_name == "BS" or sheet_name == "PL":
            label_col = 2  # BS/PL have empty col A, labels in col B

        # Detect data start row
        data_start_row = _detect_data_start_row(ws_f)
        # Skip header rows (dates, labels, empty)
        for r in range(1, ws_f.max_row + 1):
            v = ws_f.cell(r, label_col).value
            if v is not None and str(v).strip() not in ("", "(тыс. сом)", "(тыс сом)"):
                # Check if this row is below the period headers
                if r > 3:
                    data_start_row = r
                    break

        # Build indicator hierarchy
        indicators = _detect_indicators(ws_f, data_start_row, label_col, ws_f.max_row)

        if not indicators:
            continue

        # Create indicator analytic for this sheet
        indicator_analytic_id = str(uuid.uuid4())
        sheet_title = ws_f.cell(1, 1).value or ws_f.cell(1, 2).value or sheet_name
        sheet_title = str(sheet_title).strip()

        analytic_name = f"Показатели ({sheet_name})"
        await db.execute(
            "INSERT INTO analytics (id, model_id, name, code, icon, data_type) VALUES (?,?,?,?,?,?)",
            (indicator_analytic_id, model_id, analytic_name, f"indicators_{sheet_name.lower().replace('.','_')}",
             "ListAltOutlined", "sum"),
        )

        # Create indicator fields
        for fname, fcode, ftype, sort in [("name", "name", "string", 0), ("unit", "unit", "string", 1)]:
            fid = str(uuid.uuid4())
            await db.execute(
                "INSERT INTO analytic_fields (id, analytic_id, name, code, data_type, sort_order) VALUES (?,?,?,?,?,?)",
                (fid, indicator_analytic_id, fname, fcode, ftype, sort),
            )

        # Create indicator records
        indicator_record_ids = {}  # row -> record_id
        sort_idx = 0

        def create_indicator_records(items, parent_id=None):
            nonlocal sort_idx
            for item in items:
                rid = str(uuid.uuid4())
                data = {"name": item["name"]}
                if item.get("unit"):
                    data["unit"] = item["unit"]
                db_coro = db.execute(
                    "INSERT INTO analytic_records (id, analytic_id, parent_id, sort_order, data_json) VALUES (?,?,?,?,?)",
                    (rid, indicator_analytic_id, parent_id, sort_idx, json.dumps(data, ensure_ascii=False)),
                )
                indicator_record_ids[item["row"]] = rid
                sort_idx += 1
                item["_rid"] = rid
                item["_coro"] = db_coro
                if item["children"]:
                    for child in item["children"]:
                        child["_parent_rid"] = rid

        # First pass: assign IDs
        for item in indicators:
            rid = str(uuid.uuid4())
            indicator_record_ids[item["row"]] = rid
            item["_rid"] = rid
            for child in item.get("children", []):
                child_rid = str(uuid.uuid4())
                indicator_record_ids[child["row"]] = child_rid
                child["_rid"] = child_rid

        # Second pass: insert records
        sort_idx = 0
        for item in indicators:
            data = {"name": item["name"]}
            if item.get("unit"):
                data["unit"] = item["unit"]
            await db.execute(
                "INSERT INTO analytic_records (id, analytic_id, parent_id, sort_order, data_json) VALUES (?,?,?,?,?)",
                (item["_rid"], indicator_analytic_id, None, sort_idx, json.dumps(data, ensure_ascii=False)),
            )
            sort_idx += 1
            for child in item.get("children", []):
                data_c = {"name": child["name"]}
                if child.get("unit"):
                    data_c["unit"] = child["unit"]
                await db.execute(
                    "INSERT INTO analytic_records (id, analytic_id, parent_id, sort_order, data_json) VALUES (?,?,?,?,?)",
                    (child["_rid"], indicator_analytic_id, item["_rid"], sort_idx, json.dumps(data_c, ensure_ascii=False)),
                )
                sort_idx += 1

        # Create Pebble sheet
        pebble_sheet_id = str(uuid.uuid4())
        await db.execute(
            "INSERT INTO sheets (id, model_id, name) VALUES (?,?,?)",
            (pebble_sheet_id, model_id, sheet_title),
        )

        # Bind analytics: periods first (columns), then indicators (rows)
        for bind_idx, aid in enumerate([period_analytic_id, indicator_analytic_id]):
            sa_id = str(uuid.uuid4())
            await db.execute(
                "INSERT INTO sheet_analytics (id, sheet_id, analytic_id, sort_order) VALUES (?,?,?,?)",
                (sa_id, pebble_sheet_id, aid, bind_idx),
            )

        # Grant permissions to all users
        users = await db.execute_fetchall("SELECT id FROM users")
        for u in users:
            pid = str(uuid.uuid4())
            try:
                await db.execute(
                    "INSERT INTO sheet_permissions (id, sheet_id, user_id, can_view, can_edit) VALUES (?,?,?,1,1)",
                    (pid, pebble_sheet_id, u["id"]),
                )
            except Exception:
                pass

        # === Import cell data ===
        # Map sheet period columns to shared period record IDs
        # Build col -> period_record_id mapping for this sheet
        sheet_period_map = {}
        for sp in sheet_periods:
            # Match by column name (m1, m2...) or date
            for gp_col, gp_rid in period_record_ids.items():
                # Match by position index
                sp_idx = sheet_periods.index(sp)
                gp_keys = sorted(period_record_ids.keys())
                if sp_idx < len(gp_keys):
                    sheet_period_map[sp["col"]] = period_record_ids[gp_keys[sp_idx]]
                break

        for row_num, indicator_rid in indicator_record_ids.items():
            for col_num, period_rid in sheet_period_map.items():
                cell_data = ws_d.cell(row_num, col_num)
                cell_formula = ws_f.cell(row_num, col_num)

                val = cell_data.value
                if val is None:
                    continue

                # Determine rule
                is_input = _is_input_cell(cell_formula)
                is_formula = str(cell_formula.value).startswith("=") if cell_formula.value else False

                rule = "manual" if is_input else "formula" if is_formula else "manual"
                coord_key = f"{period_rid}|{indicator_rid}"
                value_str = str(val) if val is not None else ""

                cid = str(uuid.uuid4())
                try:
                    await db.execute(
                        "INSERT INTO cell_data (id, sheet_id, coord_key, value, data_type, rule, formula) VALUES (?,?,?,?,?,?,?)",
                        (cid, pebble_sheet_id, coord_key, value_str, "sum", rule, ""),
                    )
                except Exception:
                    pass  # Skip duplicates

        created_sheets.append({"name": sheet_title, "id": pebble_sheet_id})

    await db.commit()

    return {
        "model_id": model_id,
        "model_name": model_name,
        "sheets": len(created_sheets),
        "sheet_list": created_sheets,
        "periods": len(periods),
    }
