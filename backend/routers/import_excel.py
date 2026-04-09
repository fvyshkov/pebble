"""Import an Excel workbook as a Pebble model using Claude API for intelligent
structure analysis.

Flow:
1. Extract text representation of each Excel sheet (headers, row labels, data presence)
2. Send to Claude API → returns JSON describing periods, indicator hierarchies
3. Create model + period analytic with proper year/quarter/month hierarchy
4. Create indicator analytics per sheet with hierarchical records
5. Import cell values (manual vs formula detection by theme color)
"""

import uuid
import json
import io
import os
import logging
from datetime import datetime, date
from calendar import monthrange

from fastapi import APIRouter, UploadFile, File, Form
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from backend.db import get_db

router = APIRouter(prefix="/api/import", tags=["import"])
log = logging.getLogger(__name__)

MONTH_NAMES_RU = [
    "Январь", "Февраль", "Март", "Апрель", "Май", "Июнь",
    "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь",
]
QUARTER_NAMES_RU = ["1-й квартал", "2-й квартал", "3-й квартал", "4-й квартал"]

# ── Claude API prompts ─────────────────────────────────────────────────────

PEBBLE_SYSTEM_PROMPT = """\
You are an assistant that analyzes Excel financial models for import into Pebble.

Pebble object model:
- Model: top-level container (e.g. "BaaS Financial Model")
- Analytic: a dimension/axis (e.g. "Периоды", "Показатели"). Has fields and hierarchical records.
  - Fields define the schema (e.g. "Наименование", "Единица измерения")
  - Records are hierarchical (parent_id tree): groups contain children
- Sheet: a named view binding 2+ analytics. First analytic = columns (periods), rest = rows (indicators)
- Cell: intersection of records from each bound analytic. Has value, rule (manual/formula)

Period analytic has year > quarter > month hierarchy (e.g. "2026" → "1-й квартал" → "Январь 2026").
Indicator analytics contain financial line items organized into groups.

Your task: analyze Excel sheets and return their indicator hierarchy as JSON.
"""

SHEET_ANALYSIS_PROMPT = """\
Analyze ONE sheet from an Excel financial model. The text shows:
- Header rows (dates, labels)
- Row labels with row numbers, formatting (BOLD, indent), data presence (HAS_DATA, INPUT, FORMULA)

Return a JSON describing THIS sheet's indicator hierarchy.

RULES:
1. Build a hierarchical tree. Group headers (BOLD rows without HAS_DATA, or section titles) become parent nodes.
2. Items with HAS_DATA beneath a group become its children.
3. Sections like "Курсы валют", "Процентные ставки": individual items (USD, EUR, rates) MUST be children UNDER the group header, not siblings.
4. Indented rows (indent=N) are children of the nearest preceding non-indented or less-indented row.
5. Rows like "Показатель", "ЕИ", "Отв.исп." are column headers — skip them.
6. display_name: use A1 or B1 title if it's descriptive, otherwise use the sheet tab name.
7. data_start_col: the column number where period data begins (where dates are in headers).
8. Groups with HAS_DATA: a row can be both a group AND have data (e.g. "Итого" row with summed values and child detail rows below). Mark is_group=true if it has children.

Return ONLY valid JSON, no markdown:
{"excel_name":"Tab","display_name":"Title","data_start_col":4,"indicators":[{"name":"Group","unit":"","row":5,"is_group":true,"children":[{"name":"Item","unit":"%","row":6,"is_group":false,"children":[]}]}]}

Sheet content:
"""


# ── Excel text extraction ──────────────────────────────────────────────────

def _extract_sheet_text(ws, sheet_name: str, max_rows: int = 500) -> str:
    """Extract a compact text representation of a sheet for Claude analysis."""
    lines = [f"=== Sheet: {sheet_name} ==="]
    max_col = min(ws.max_column or 1, 200)
    max_row = min(ws.max_row or 1, max_rows)

    # Header rows (first 6)
    lines.append("--- Header rows ---")
    for r in range(1, min(7, max_row + 1)):
        row_vals = []
        for c in range(1, min(max_col + 1, 50)):
            v = ws.cell(r, c).value
            if v is not None:
                if isinstance(v, datetime):
                    row_vals.append(f"{get_column_letter(c)}:{v.strftime('%Y-%m-%d')}")
                else:
                    row_vals.append(f"{get_column_letter(c)}:{str(v)[:60]}")
        if row_vals:
            lines.append(f"  Row {r}: {' | '.join(row_vals)}")

    # Row labels with formatting and data presence
    lines.append("--- Row labels ---")
    for r in range(1, max_row + 1):
        # Collect labels from first few columns
        labels = []
        for c in range(1, min(6, max_col + 1)):
            v = ws.cell(r, c).value
            if v is not None and not isinstance(v, datetime):
                s = str(v).strip()
                if s and len(s) < 100:
                    labels.append(f"{get_column_letter(c)}='{s}'")

        if not labels:
            continue

        # Check formatting
        cell_a = ws.cell(r, 1)
        is_bold = cell_a.font and cell_a.font.bold
        indent = cell_a.alignment.indent if cell_a.alignment and cell_a.alignment.indent else 0

        # Check data presence in period columns (sample cols 4-10)
        has_data = False
        has_formula = False
        for c in range(4, min(15, max_col + 1)):
            cv = ws.cell(r, c).value
            if cv is not None:
                has_data = True
                if isinstance(cv, str) and cv.startswith("="):
                    has_formula = True
                break

        # Check if input cell (theme=7)
        is_input = False
        for c in range(4, min(15, max_col + 1)):
            fill = ws.cell(r, c).fill
            if fill and fill.fgColor and fill.fgColor.theme == 7:
                is_input = True
                break

        flags = []
        if is_bold:
            flags.append("BOLD")
        if indent:
            flags.append(f"indent={int(indent)}")
        if has_data:
            flags.append("HAS_DATA")
        if has_formula:
            flags.append("FORMULA")
        if is_input:
            flags.append("INPUT")

        flag_str = f" [{', '.join(flags)}]" if flags else ""
        lines.append(f"  Row {r}: {' | '.join(labels)}{flag_str}")

    return "\n".join(lines)


def _is_input_cell(cell) -> bool:
    """Check if cell has beige/yellow input background (theme=7)."""
    fill = cell.fill
    if fill and fill.fgColor and fill.fgColor.theme == 7:
        return True
    return False


# ── Claude API call ────────────────────────────────────────────────────────

def _get_claude_client():
    """Create Anthropic client from environment."""
    import anthropic

    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        raise RuntimeError("ANTHROPIC_API_KEY environment variable not set")

    base_url = os.environ.get("ANTHROPIC_BASE_URL")
    kwargs = {"api_key": api_key}
    if base_url:
        kwargs["base_url"] = base_url

    return anthropic.Anthropic(**kwargs)


def _parse_claude_json(response_text: str) -> dict:
    """Parse JSON from Claude response, handling markdown fences."""
    text = response_text.strip()
    if text.startswith("```"):
        text = text.split("\n", 1)[1]
        if "```" in text:
            text = text[: text.rfind("```")]
        text = text.strip()
    return json.loads(text)


async def _analyze_sheet_with_claude(client, sheet_text: str, retries: int = 3) -> dict:
    """Analyze one sheet with Claude API. Returns sheet config dict."""
    import time
    for attempt in range(retries):
        try:
            message = client.messages.create(
                model="claude-sonnet-4-20250514",
                max_tokens=8192,
                system=PEBBLE_SYSTEM_PROMPT,
                messages=[{"role": "user", "content": SHEET_ANALYSIS_PROMPT + sheet_text}],
            )
            return _parse_claude_json(message.content[0].text)
        except Exception as e:
            if attempt < retries - 1 and ("overloaded" in str(e).lower() or "529" in str(e)):
                time.sleep(2 ** attempt)
                continue
            raise


async def _analyze_workbook_with_claude(sheet_texts: dict[str, str], period_dates: list) -> dict:
    """Analyze entire workbook: detect periods from dates, analyze each sheet with Claude."""
    client = _get_claude_client()

    # Period config from detected dates (deterministic, no AI needed)
    if period_dates:
        min_d = min(period_dates)
        max_d = max(period_dates)
        start = f"{min_d.year}-{min_d.month:02d}-01"
        end_m = max_d.month
        end_y = max_d.year
        end_day = monthrange(end_y, end_m)[1]
        end = f"{end_y}-{end_m:02d}-{end_day:02d}"
    else:
        start, end = "2026-01-01", "2028-12-31"

    period_config = {"period_types": ["year", "quarter", "month"], "start": start, "end": end}

    # Analyze each sheet with Claude (sequentially to avoid rate limits)
    sheets = []
    for sheet_name, text in sheet_texts.items():
        try:
            sheet_cfg = await _analyze_sheet_with_claude(client, text)
            sheet_cfg["excel_name"] = sheet_name  # Ensure correct name
            sheets.append(sheet_cfg)
            log.info("Sheet '%s' analyzed: %d indicators", sheet_name,
                     len(sheet_cfg.get("indicators", [])))
        except Exception as e:
            log.warning("Claude analysis failed for sheet '%s': %s", sheet_name, e)

    return {"period_config": period_config, "sheets": sheets}


# ── Detect periods from date headers (fallback) ───────────────────────────

def _detect_periods_from_headers(ws, max_col: int) -> list[dict]:
    """Detect period columns from date headers in the first 6 rows.

    Normalizes all dates to the 1st of the month — Excel formulas like
    =C4+31 accumulate drift (e.g. March 4 instead of March 1), but the
    intent is always month-start boundaries.
    """
    periods = []
    date_row = None

    for r in range(1, 7):
        for c in range(1, min(max_col + 1, 50)):
            v = ws.cell(r, c).value
            if isinstance(v, datetime):
                date_row = r
                break
        if date_row:
            break

    if date_row is None:
        return []

    seen_months = set()
    for c in range(1, max_col + 1):
        v = ws.cell(date_row, c).value
        if isinstance(v, datetime):
            # Normalize to 1st of month (fixes drift from =C4+31 formulas)
            normalized = datetime(v.year, v.month, 1)
            month_key = f"{v.year}-{v.month:02d}"
            if month_key in seen_months:
                continue  # Skip duplicate months
            seen_months.add(month_key)
            periods.append({
                "col": c,
                "name": f"{MONTH_NAMES_RU[v.month - 1]} {v.year}",
                "date": normalized,
            })

    return periods


# ── Period hierarchy creation (mirrors generate_periods from analytics.py) ─

async def _create_period_hierarchy(db, analytic_id: str, period_types: list[str],
                                   start_date: date, end_date: date) -> dict:
    """Create year > quarter > month hierarchy. Returns {col_month_key: record_id} mapping."""
    has_year = "year" in period_types
    has_quarter = "quarter" in period_types
    has_month = "month" in period_types

    # Map: "YYYY-MM" -> record_id (for matching to Excel columns later)
    month_record_ids = {}
    sort = 0
    year = start_date.year

    while year <= end_date.year:
        year_start = max(start_date, date(year, 1, 1))
        year_end = min(end_date, date(year, 12, 31))

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
            if q_start > end_date or q_end < start_date:
                continue
            q_start = max(q_start, start_date)
            q_end = min(q_end, end_date)

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
                    if m_start > end_date or m_end < start_date:
                        continue
                    m_start = max(m_start, start_date)
                    m_end = min(m_end, end_date)
                    mid = str(uuid.uuid4())
                    parent = quarter_id if quarter_id else year_id
                    await db.execute(
                        "INSERT INTO analytic_records (id, analytic_id, parent_id, sort_order, data_json) VALUES (?,?,?,?,?)",
                        (mid, analytic_id, parent, sort, json.dumps(
                            {"name": f"{MONTH_NAMES_RU[m - 1]} {year}", "start": str(m_start), "end": str(m_end)},
                            ensure_ascii=False)),
                    )
                    month_record_ids[f"{year}-{m:02d}"] = mid
                    sort += 1

        year += 1

    return month_record_ids


# ── Indicator records creation (recursive) ─────────────────────────────────

async def _create_indicator_records(db, analytic_id: str, indicators: list[dict]) -> dict:
    """Create hierarchical indicator records. Returns {excel_row: record_id}."""
    row_to_rid = {}
    sort_idx = 0

    async def insert_items(items: list[dict], parent_id: str | None):
        nonlocal sort_idx
        for item in items:
            rid = str(uuid.uuid4())
            data = {"name": item["name"]}
            if item.get("unit"):
                data["unit"] = item["unit"]

            await db.execute(
                "INSERT INTO analytic_records (id, analytic_id, parent_id, sort_order, data_json) VALUES (?,?,?,?,?)",
                (rid, analytic_id, parent_id, sort_idx, json.dumps(data, ensure_ascii=False)),
            )
            row_to_rid[item["row"]] = rid
            sort_idx += 1

            if item.get("children"):
                await insert_items(item["children"], rid)

    await insert_items(indicators, None)
    return row_to_rid


# ── Main import endpoint ───────────────────────────────────────────────────

@router.post("/excel")
async def import_excel(file: UploadFile = File(...), model_name: str = Form("Imported Model")):
    db = get_db()

    # Ensure unique model name
    existing = await db.execute_fetchall("SELECT id FROM models WHERE name = ?", (model_name,))
    if existing:
        model_name = f"{model_name} ({datetime.now().strftime('%Y-%m-%d %H:%M')})"

    content = await file.read()
    wb_formulas = load_workbook(io.BytesIO(content))
    wb_data = load_workbook(io.BytesIO(content), data_only=True)

    # ── Step 1: Extract text and detect dates ──
    sheet_texts = {}
    all_dates = []
    for sn in wb_formulas.sheetnames:
        ws = wb_formulas[sn]
        sheet_texts[sn] = _extract_sheet_text(ws, sn)
        # Collect dates for period detection (normalize to 1st of month)
        for r in range(1, 7):
            for c in range(1, min((ws.max_column or 1) + 1, 200)):
                v = ws.cell(r, c).value
                if isinstance(v, datetime):
                    all_dates.append(datetime(v.year, v.month, 1))

    # ── Step 2: Analyze with Claude API (per sheet) ──
    try:
        analysis = await _analyze_workbook_with_claude(sheet_texts, all_dates)
    except Exception as e:
        log.warning("Claude API analysis failed, falling back to heuristics: %s", e)
        analysis = _fallback_heuristic_analysis(wb_formulas)

    period_config = analysis["period_config"]
    sheets_config = analysis["sheets"]

    # ── Step 3: Create model ──
    model_id = str(uuid.uuid4())
    await db.execute(
        "INSERT INTO models (id, name, description) VALUES (?, ?, ?)",
        (model_id, model_name, "Импортировано из Excel"),
    )

    # ── Step 4: Create period analytic with proper hierarchy ──
    period_analytic_id = str(uuid.uuid4())
    period_types = period_config.get("period_types", ["year", "quarter", "month"])
    period_start = period_config.get("start", "2026-01-01")
    period_end = period_config.get("end", "2028-12-31")

    await db.execute(
        """INSERT INTO analytics (id, model_id, name, code, icon, is_periods, data_type,
           period_types, period_start, period_end, sort_order)
           VALUES (?,?,?,?,?,?,?,?,?,?,?)""",
        (period_analytic_id, model_id, "Периоды", "periods", "CalendarMonthOutlined",
         1, "sum", json.dumps(period_types), period_start, period_end, 0),
    )

    # Create period fields (proper Russian names, matching manual creation)
    for sort_i, (fname, fcode, ftype) in enumerate([
        ("Наименование", "name", "string"),
        ("Начало", "start", "date"),
        ("Окончание", "end", "date"),
    ]):
        fid = str(uuid.uuid4())
        await db.execute(
            "INSERT INTO analytic_fields (id, analytic_id, name, code, data_type, sort_order) VALUES (?,?,?,?,?,?)",
            (fid, period_analytic_id, fname, fcode, ftype, sort_i),
        )

    # Create period records with year > quarter > month hierarchy
    start_d = date.fromisoformat(period_start)
    end_d = date.fromisoformat(period_end)
    month_record_ids = await _create_period_hierarchy(db, period_analytic_id, period_types, start_d, end_d)

    # ── Step 5: Process each sheet ──
    created_sheets = []
    analytic_sort = 1  # 0 is periods

    for sheet_cfg in sheets_config:
        excel_name = sheet_cfg["excel_name"]
        display_name = sheet_cfg.get("display_name", excel_name)
        indicators = sheet_cfg.get("indicators", [])
        data_start_col = sheet_cfg.get("data_start_col", 4)

        if excel_name not in wb_formulas.sheetnames:
            continue
        if not indicators:
            continue

        ws_f = wb_formulas[excel_name]
        ws_d = wb_data[excel_name]

        # Create indicator analytic
        indicator_analytic_id = str(uuid.uuid4())
        analytic_name = f"Показатели ({excel_name})"
        await db.execute(
            """INSERT INTO analytics (id, model_id, name, code, icon, data_type, sort_order)
               VALUES (?,?,?,?,?,?,?)""",
            (indicator_analytic_id, model_id, analytic_name,
             f"indicators_{excel_name.lower().replace('.', '_').replace('+', '_')}",
             "ListAltOutlined", "sum", analytic_sort),
        )
        analytic_sort += 1

        # Create indicator fields (proper Russian names)
        for sort_i, (fname, fcode, ftype) in enumerate([
            ("Наименование", "name", "string"),
            ("Единица измерения", "unit", "string"),
        ]):
            fid = str(uuid.uuid4())
            await db.execute(
                "INSERT INTO analytic_fields (id, analytic_id, name, code, data_type, sort_order) VALUES (?,?,?,?,?,?)",
                (fid, indicator_analytic_id, fname, fcode, ftype, sort_i),
            )

        # Create hierarchical indicator records
        row_to_rid = await _create_indicator_records(db, indicator_analytic_id, indicators)

        # Create Pebble sheet
        pebble_sheet_id = str(uuid.uuid4())
        await db.execute(
            "INSERT INTO sheets (id, model_id, name) VALUES (?,?,?)",
            (pebble_sheet_id, model_id, display_name),
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

        # ── Import cell data ──
        # Build col -> month_record_id mapping from date headers
        # Use ws_d (data_only) because formula-derived dates (=C4+31) are only resolved there
        sheet_periods = _detect_periods_from_headers(ws_d, min(ws_d.max_column or 1, 200))
        col_to_period_rid = {}
        for sp in sheet_periods:
            d = sp["date"]
            if isinstance(d, datetime):
                key = f"{d.year}-{d.month:02d}"
                if key in month_record_ids:
                    col_to_period_rid[sp["col"]] = month_record_ids[key]

        cell_count = 0
        for row_num, indicator_rid in row_to_rid.items():
            for col_num, period_rid in col_to_period_rid.items():
                cell_val = ws_d.cell(row_num, col_num)
                cell_fmt = ws_f.cell(row_num, col_num)

                val = cell_val.value
                if val is None:
                    continue

                is_input = _is_input_cell(cell_fmt)
                is_formula = str(cell_fmt.value).startswith("=") if cell_fmt.value else False
                rule = "manual" if is_input else "formula" if is_formula else "manual"

                coord_key = f"{period_rid}|{indicator_rid}"
                value_str = str(val)

                cid = str(uuid.uuid4())
                try:
                    await db.execute(
                        "INSERT INTO cell_data (id, sheet_id, coord_key, value, data_type, rule, formula) VALUES (?,?,?,?,?,?,?)",
                        (cid, pebble_sheet_id, coord_key, value_str, "sum", rule, ""),
                    )
                    cell_count += 1
                except Exception:
                    pass  # Skip duplicates

        created_sheets.append({"name": display_name, "id": pebble_sheet_id, "cells": cell_count})

    await db.commit()

    return {
        "model_id": model_id,
        "model_name": model_name,
        "sheets": len(created_sheets),
        "sheet_list": created_sheets,
        "periods": len(month_record_ids),
        "period_hierarchy": period_types,
    }


# ── Fallback heuristic analysis (when Claude API unavailable) ──────────────

def _fallback_heuristic_analysis(wb) -> dict:
    """Basic heuristic analysis as fallback when Claude API is not available."""
    # Detect period range from first sheet
    first_ws = wb[wb.sheetnames[0]]
    dates = []
    for r in range(1, 7):
        for c in range(1, min((first_ws.max_column or 1) + 1, 200)):
            v = first_ws.cell(r, c).value
            if isinstance(v, datetime):
                dates.append(v)

    if dates:
        min_d = min(dates)
        max_d = max(dates)
        start = f"{min_d.year}-{min_d.month:02d}-01"
        end_month = max_d.month
        end_year = max_d.year
        end_day = monthrange(end_year, end_month)[1]
        end = f"{end_year}-{end_month:02d}-{end_day:02d}"
    else:
        start, end = "2026-01-01", "2028-12-31"

    period_config = {"period_types": ["year", "quarter", "month"], "start": start, "end": end}

    sheets = []
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        max_row = min(ws.max_row or 1, 500)
        max_col = min(ws.max_column or 1, 200)

        # Find data start col
        data_start_col = 4
        for r in range(1, 7):
            for c in range(1, min(max_col + 1, 50)):
                if isinstance(ws.cell(r, c).value, datetime):
                    data_start_col = c
                    break

        # Detect display name
        title = ws.cell(1, 1).value or ws.cell(1, 2).value or sheet_name
        title = str(title).strip()

        # Detect label column
        label_col = 1
        if sheet_name in ("BS", "PL"):
            label_col = 2

        # Find first data row
        data_start_row = 7
        for r in range(1, max_row + 1):
            v = ws.cell(r, label_col).value
            if v is not None and str(v).strip() not in ("", "(тыс. сом)", "(тыс сом)"):
                if r > 3:
                    data_start_row = r
                    break

        # Build indicator hierarchy with improved grouping
        indicators = []
        current_group = None

        for r in range(data_start_row, max_row + 1):
            name = ws.cell(r, label_col).value
            if name is None or str(name).strip() == "":
                if current_group:
                    indicators.append(current_group)
                    current_group = None
                continue

            name = str(name).strip()
            unit_col = label_col + 1
            unit = ws.cell(r, unit_col).value
            unit = str(unit).strip() if unit else ""

            # Check for data in period columns
            has_data = False
            for c in range(data_start_col, min(data_start_col + 5, max_col + 1)):
                if ws.cell(r, c).value is not None:
                    has_data = True
                    break

            # Check if bold
            is_bold = ws.cell(r, label_col).font and ws.cell(r, label_col).font.bold

            # Group detection: bold or no data in period columns
            is_group_header = (is_bold and not has_data) or (
                not has_data and unit in ("ЕИ", "")
                and ws.cell(r, unit_col + 1).value in ("Отв.исп.", None, "")
            )

            item = {"name": name, "unit": unit, "row": r, "is_group": False, "children": []}

            if is_group_header:
                if current_group:
                    indicators.append(current_group)
                current_group = {"name": name, "unit": "", "row": r, "is_group": True, "children": []}
            elif current_group:
                current_group["children"].append(item)
            else:
                indicators.append(item)

        if current_group:
            indicators.append(current_group)

        if indicators:
            sheets.append({
                "excel_name": sheet_name,
                "display_name": title,
                "data_start_col": data_start_col,
                "indicators": indicators,
            })

    return {"period_config": period_config, "sheets": sheets}
