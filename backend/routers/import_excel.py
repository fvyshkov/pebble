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
from fastapi.responses import StreamingResponse
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
  - Records are hierarchical (parent_id tree): groups contain children
- Sheet: a named view binding 2+ analytics. First = columns (periods), rest = rows
- Cell: value + rule (manual = user input, formula = computed)

Pebble formula syntax (references by indicator NAME, not cell):
- [indicator_name] — value of another indicator, same period, same group
- [indicator_name](периоды="предыдущий") — value from PREVIOUS period
- [SheetName::indicator_name] — cross-sheet reference (use :: separator, sheet display name)
- Standard math: +, -, *, /, parentheses, numbers
- SUM([a], [b], [c]) — sum of multiple indicators
- For first-period special cases (e.g. no previous period), use "0" as value

Examples:
  Excel: =D13*D14 (where R13=количество партнеров, R14=ср. количество выдач)
  Pebble: [количество партнеров] * [среднее количество выдач на 1 партнера]

  Excel: =D20/E18 (prev period's портфель / current ср.срок)
  Pebble: [кредитный портфель](периоды="предыдущий") / [cр. срок портфеля]

  Excel: =(C20+D20)/2*D24/12 (avg of prev+current портфель * rate / 12)
  Pebble: ([кредитный портфель](периоды="предыдущий") + [кредитный портфель]) / 2 * [ср. % ставка портфеля] / 12

  Excel: =D27-C27 (delta from previous period)
  Pebble: [РППУ на конец месяца] - [РППУ на конец месяца](периоды="предыдущий")

  Excel: =SUM(D28:D31) (sum of rows 28-31)
  Pebble: SUM([% доход], [трансфертный расход], [комиссия партнеру], [расходы на провизии])
"""

SHEET_ANALYSIS_PROMPT = """\
Analyze ONE sheet from an Excel financial model. The text shows:
- Header rows (dates, labels)
- Row labels with row numbers, formatting (BOLD, indent), data presence flags
- F1= formula for first period, F2= formula for second period (if different)
- INPUT = manual input cell, FORMULA = computed cell

Return a JSON describing the sheet's indicator hierarchy WITH formulas.

RULES:
1. Build a hierarchical tree. BOLD rows with "ЕИ" = group headers (product types, sections).
2. Items beneath a group are its children indicators.
3. For each indicator: name, unit, row, rule (manual/formula), and for formulas: the Pebble formula.
4. Convert Excel cell references to Pebble [indicator_name] references using the ROW MAPPING.
5. If F1 and F2 differ (e.g. first period is 0, subsequent use @prev), provide TWO formulas: "formula_first" and "formula".
6. Cross-sheet references like ='0'!D10 → [SheetDisplayName::indicator_name] using :: separator and the display_name of the referenced sheet.
7. display_name: from A1/B1 title. data_start_col: first period column number.
8. Skip header rows (Показатель, ЕИ, Отв.исп.) — they are not indicators.
9. CRITICAL: Every [indicator_name] in a formula must EXACTLY match the "name" field of SOME indicator in the JSON output. If the same indicator name appears in multiple groups, append a disambiguating suffix in parentheses to BOTH the name and all formula references. Example: "портфель" in KGS group → name "портфель (KGS)", in RUB group → "портфель (RUB)".
10. For "Итого" / summary rows that SUM across groups: use the EXACT disambiguated names. E.g. formula: "[портфель (KGS)] + [портфель (RUB)] + [портфель (USD)]". NEVER write [портфель] + [портфель] + [портфель] — that's a self-reference!
11. Rows like "ВСЕГО АКТИВЫ", "ИТОГО ОБЯЗАТЕЛЬСТВА" are NOT separate indicators — they are the group header itself. The group header row IS the aggregation row.

Return ONLY valid JSON, no markdown:
{"excel_name":"Tab","display_name":"Title","data_start_col":4,"indicators":[
  {"name":"Тип продукта","unit":"","row":12,"is_group":true,"children":[
    {"name":"показатель","unit":"тыс сом","row":13,"is_group":false,"children":[],
     "rule":"formula","formula":"[a] * [b]"},
    {"name":"ввод","unit":"%","row":14,"is_group":false,"children":[],
     "rule":"manual"}
  ]}
]}

For formulas with first-period exception:
{"rule":"formula","formula":"[портфель](периоды=\"предыдущий\") / [срок]","formula_first":"0"}

Sheet content:
"""

FORMULA_FIX_PROMPT = """\
You have a Pebble model with these sheets and indicator names.
For each formula cell, the formula must reference EXACT indicator names from this list.
Cross-sheet references use :: separator: [SheetDisplayName::indicator_name]

AVAILABLE INDICATORS PER SHEET:
{indicators_context}

FORMULA CELLS TO FIX (current formula → expected behavior based on Excel formula):
{formula_fixes}

Return ONLY a JSON array of fixes:
[{{"coord_key": "...", "formula": "corrected formula", "formula_first": "optional first-period formula"}}]
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

        # Extract first two period formulas for formula cells
        formula_m1 = ""
        formula_m2 = ""
        if has_formula:
            for c in range(4, min(max_col + 1, 50)):
                cv = ws.cell(r, c).value
                if cv and isinstance(cv, str) and cv.startswith("="):
                    if not formula_m1:
                        formula_m1 = str(cv)[:80]
                    elif not formula_m2:
                        formula_m2 = str(cv)[:80]
                        break
                elif cv is not None and not formula_m1:
                    # First period is a constant (e.g. 0 for погашения)
                    formula_m1 = str(cv)[:20]
                    continue

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
        formula_str = ""
        if formula_m1:
            formula_str = f" F1={formula_m1}"
            if formula_m2 and formula_m2 != formula_m1:
                formula_str += f" F2={formula_m2}"
        lines.append(f"  Row {r}: {' | '.join(labels)}{flag_str}{formula_str}")

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
    """Parse JSON from Claude response, handling markdown fences and common issues."""
    import re
    text = response_text.strip()
    if text.startswith("```"):
        text = text.split("\n", 1)[1]
        if "```" in text:
            text = text[: text.rfind("```")]
        text = text.strip()
    # Fix trailing commas before } or ]
    text = re.sub(r',\s*([}\]])', r'\1', text)
    return json.loads(text)


_LLM_CACHE_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), ".llm_cache")

def _llm_cache_get(key: str):
    import hashlib
    h = hashlib.sha256(key.encode()).hexdigest()[:16]
    path = os.path.join(_LLM_CACHE_DIR, f"{h}.json")
    if os.path.exists(path):
        with open(path) as f:
            return json.load(f)
    return None

def _llm_cache_set(key: str, value):
    import hashlib
    os.makedirs(_LLM_CACHE_DIR, exist_ok=True)
    h = hashlib.sha256(key.encode()).hexdigest()[:16]
    path = os.path.join(_LLM_CACHE_DIR, f"{h}.json")
    with open(path, "w") as f:
        json.dump(value, f, ensure_ascii=False)


async def _analyze_sheet_with_claude(client, sheet_text: str, retries: int = 3) -> dict:
    """Analyze one sheet with Claude API. Returns sheet config dict (cached)."""
    import time, asyncio

    # Check cache first
    cached = _llm_cache_get(sheet_text)
    if cached:
        log.info("Cache hit for sheet analysis")
        return cached

    loop = asyncio.get_event_loop()
    for attempt in range(retries):
        try:
            message = await loop.run_in_executor(None, lambda: client.messages.create(
                model="claude-sonnet-4-20250514",
                max_tokens=16384,
                system=PEBBLE_SYSTEM_PROMPT + "\n\nIMPORTANT: Return ONLY valid JSON. No markdown fences, no comments, no trailing commas.",
                messages=[
                    {"role": "user", "content": SHEET_ANALYSIS_PROMPT + sheet_text},
                    {"role": "assistant", "content": "{"},  # prefill to force JSON
                ],
            ))
            result = _parse_claude_json("{" + message.content[0].text)
            _llm_cache_set(sheet_text, result)  # cache for next time
            return result
        except Exception as e:
            if attempt < retries - 1 and ("overloaded" in str(e).lower() or "529" in str(e)):
                await asyncio.sleep(2 ** attempt)
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

async def _create_indicator_records(db, analytic_id: str, indicators: list[dict]) -> tuple[dict, dict]:
    """Create hierarchical indicator records.
    Returns (row_to_rid: {excel_row: record_id}, rid_to_formula: {record_id: {rule, formula, formula_first}})
    """
    row_to_rid = {}
    rid_to_formula = {}
    sort_idx = 0

    async def insert_items(items: list[dict], parent_id: str | None):
        nonlocal sort_idx
        for item in items:
            rid = str(uuid.uuid4())
            data = {"name": item["name"]}
            if item.get("unit"):
                data["unit"] = item["unit"]

            excel_row = item.get("row")
            await db.execute(
                "INSERT INTO analytic_records (id, analytic_id, parent_id, sort_order, data_json, excel_row) VALUES (?,?,?,?,?,?)",
                (rid, analytic_id, parent_id, sort_idx, json.dumps(data, ensure_ascii=False), excel_row),
            )
            row_to_rid[item["row"]] = rid

            # Store formula info
            rule = item.get("rule", "manual")
            if rule == "formula":
                rid_to_formula[rid] = {
                    "rule": "formula",
                    "formula": item.get("formula", ""),
                    "formula_first": item.get("formula_first", ""),
                }

            sort_idx += 1

            if item.get("children"):
                await insert_items(item["children"], rid)

    await insert_items(indicators, None)
    return row_to_rid, rid_to_formula


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
    sheet_sort = 0

    for sheet_cfg in sheets_config:
        excel_name = sheet_cfg["excel_name"]
        display_name = sheet_cfg.get("display_name", excel_name)
        # Sheet name: "ExcelTab. Title" (e.g. "BS. Баланс BaaS")
        sheet_display = display_name if display_name != excel_name else excel_name
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
        row_to_rid, rid_to_formula = await _create_indicator_records(db, indicator_analytic_id, indicators)

        # Create Pebble sheet
        pebble_sheet_id = str(uuid.uuid4())
        await db.execute(
            "INSERT INTO sheets (id, model_id, name, sort_order, excel_code) VALUES (?,?,?,?,?)",
            (pebble_sheet_id, model_id, sheet_display, sheet_sort, excel_name),
        )
        sheet_sort += 1

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

        # Determine which period is "first" (for formula_first)
        sorted_period_cols = sorted(col_to_period_rid.keys())
        first_period_rid = col_to_period_rid[sorted_period_cols[0]] if sorted_period_cols else None

        cell_count = 0
        for row_num, indicator_rid in row_to_rid.items():
            formula_info = rid_to_formula.get(indicator_rid)

            for col_num, period_rid in col_to_period_rid.items():
                cell_val = ws_d.cell(row_num, col_num)
                val = cell_val.value
                if val is None:
                    continue

                # Skip non-numeric strings (month codes like m1/m2, labels, etc.)
                if isinstance(val, str):
                    try:
                        val = float(val.replace(",", ".").replace(" ", ""))
                    except (ValueError, AttributeError):
                        continue

                # Determine rule and formula from Claude analysis
                if formula_info:
                    rule = "formula"
                    is_first = (period_rid == first_period_rid)
                    if is_first and formula_info.get("formula_first"):
                        formula_text = formula_info["formula_first"]
                    else:
                        formula_text = formula_info.get("formula", "")
                else:
                    rule = "manual"
                    formula_text = ""

                coord_key = f"{period_rid}|{indicator_rid}"
                value_str = str(val)

                cid = str(uuid.uuid4())
                try:
                    await db.execute(
                        "INSERT INTO cell_data (id, sheet_id, coord_key, value, data_type, rule, formula) VALUES (?,?,?,?,?,?,?)",
                        (cid, pebble_sheet_id, coord_key, value_str, "sum", rule, formula_text),
                    )
                    cell_count += 1
                except Exception:
                    pass  # Skip duplicates

        created_sheets.append({"name": sheet_display, "id": pebble_sheet_id, "cells": cell_count})

    await db.commit()

    return {
        "model_id": model_id,
        "model_name": model_name,
        "sheets": len(created_sheets),
        "sheet_list": created_sheets,
        "periods": len(month_record_ids),
        "period_hierarchy": period_types,
    }


# ── Streaming import endpoint (SSE) ───────────────────────────────────────

@router.post("/excel-stream")
async def import_excel_stream(file: UploadFile = File(...), model_name: str = Form("Imported Model")):
    import asyncio
    content = await file.read()

    async def generate():
        import time as _time
        _t0 = _time.time()

        def event(msg: str, data: dict | None = None):
            elapsed = _time.time() - _t0
            ts = f"[{int(elapsed)}с]"
            payload = json.dumps({"message": f"{ts} {msg}", **(data or {})}, ensure_ascii=False)
            return f"data: {payload}\n\n"

        db = get_db()
        loop = asyncio.get_event_loop()

        # Unique name
        existing = await db.execute_fetchall("SELECT id FROM models WHERE name = ?", (model_name,))
        if existing:
            model_name_final = f"{model_name} ({datetime.now().strftime('%Y-%m-%d %H:%M')})"
        else:
            model_name_final = model_name

        yield event(f"Загрузка файла Excel...")

        # Run blocking openpyxl in executor to not block event loop
        wb_formulas = await loop.run_in_executor(None, lambda: load_workbook(io.BytesIO(content)))
        wb_data = await loop.run_in_executor(None, lambda: load_workbook(io.BytesIO(content), data_only=True))
        sheet_names = wb_formulas.sheetnames

        yield event(f"Найдено {len(sheet_names)} листов: {', '.join(sheet_names)}")

        # Extract text and dates
        sheet_texts = {}
        all_dates = []
        for sn in sheet_names:
            ws = wb_formulas[sn]
            sheet_texts[sn] = _extract_sheet_text(ws, sn)
            for r in range(1, 7):
                for c in range(1, min((ws.max_column or 1) + 1, 200)):
                    v = ws.cell(r, c).value
                    if isinstance(v, datetime):
                        all_dates.append(datetime(v.year, v.month, 1))

        yield event("Анализ структуры с помощью Claude AI...")

        # Analyze with Claude (per-sheet with progress)
        try:
            client = _get_claude_client()
            if all_dates:
                min_d = min(all_dates)
                max_d = max(all_dates)
                p_start = f"{min_d.year}-{min_d.month:02d}-01"
                end_m = max_d.month
                end_y = max_d.year
                end_day = monthrange(end_y, end_m)[1]
                p_end = f"{end_y}-{end_m:02d}-{end_day:02d}"
            else:
                p_start, p_end = "2026-01-01", "2028-12-31"

            period_config = {"period_types": ["year", "quarter", "month"], "start": p_start, "end": p_end}
            sheets_config = []

            # Launch ALL sheets in parallel for speed
            yield event(f"🔍 Анализ {len(sheet_names)} листов параллельно...")

            async def analyze_one(sn):
                """Analyze one sheet: Claude (with cache+retry) → heuristic fallback."""
                for attempt in range(2):
                    try:
                        cfg = await _analyze_sheet_with_claude(client, sheet_texts[sn])
                        cfg["excel_name"] = sn
                        if len(cfg.get("indicators", [])) > 0:
                            return cfg
                    except Exception:
                        pass
                # Fallback — and cache the result
                fb = await loop.run_in_executor(None, lambda: _fallback_heuristic_analysis(wb_formulas))
                for fb_sheet in fb.get("sheets", []):
                    if fb_sheet["excel_name"] == sn:
                        _llm_cache_set(sheet_texts[sn], fb_sheet)
                        return fb_sheet
                return None

            tasks = [analyze_one(sn) for sn in sheet_names]
            results = await asyncio.gather(*tasks, return_exceptions=True)

            for i, (sn, result) in enumerate(zip(sheet_names, results)):
                if isinstance(result, Exception):
                    yield event(f"   ✗ «{sn}»: ошибка — {result}")
                elif result and len(result.get("indicators", [])) > 0:
                    ind_count = len(result.get("indicators", []))
                    ch_count = sum(len(x.get("children", [])) for x in result.get("indicators", []))
                    yield event(f"   ✓ «{sn}» ({i+1}/{len(sheet_names)}): {ind_count} групп, {ch_count} показателей")
                    sheets_config.append(result)
                else:
                    yield event(f"   ✗ «{sn}»: не удалось разобрать")

            analysis = {"period_config": period_config, "sheets": sheets_config}
        except Exception as e:
            yield event(f"⚠ Claude API недоступен, используем эвристики: {e}")
            analysis = _fallback_heuristic_analysis(wb_formulas)

        period_config = analysis["period_config"]
        sheets_config = analysis["sheets"]

        # Create model
        model_id = str(uuid.uuid4())
        await db.execute(
            "INSERT INTO models (id, name, description) VALUES (?, ?, ?)",
            (model_id, model_name_final, "Импортировано из Excel"),
        )
        yield event(f"Создана модель «{model_name_final}»")

        # Period analytic
        period_types = period_config.get("period_types", ["year", "quarter", "month"])
        period_start = period_config.get("start", "2026-01-01")
        period_end = period_config.get("end", "2028-12-31")

        period_analytic_id = str(uuid.uuid4())
        await db.execute(
            """INSERT INTO analytics (id, model_id, name, code, icon, is_periods, data_type,
               period_types, period_start, period_end, sort_order)
               VALUES (?,?,?,?,?,?,?,?,?,?,?)""",
            (period_analytic_id, model_id, "Периоды", "periods", "CalendarMonthOutlined",
             1, "sum", json.dumps(period_types), period_start, period_end, 0),
        )
        for sort_i, (fname, fcode, ftype) in enumerate([
            ("Наименование", "name", "string"),
            ("Начало", "start", "date"),
            ("Окончание", "end", "date"),
        ]):
            await db.execute(
                "INSERT INTO analytic_fields (id, analytic_id, name, code, data_type, sort_order) VALUES (?,?,?,?,?,?)",
                (str(uuid.uuid4()), period_analytic_id, fname, fcode, ftype, sort_i),
            )

        start_d = date.fromisoformat(period_start)
        end_d = date.fromisoformat(period_end)
        month_record_ids = await _create_period_hierarchy(db, period_analytic_id, period_types, start_d, end_d)
        yield event(f"Создана иерархия периодов: {period_start} — {period_end} ({len(month_record_ids)} месяцев)")

        # Count total indicators across all sheets for progress
        def _count_indicators(items):
            return sum(1 + _count_indicators(it.get("children", [])) for it in items)
        total_indicators = sum(_count_indicators(sc.get("indicators", [])) for sc in sheets_config)
        done_indicators = 0
        yield event(f"📊 Всего {total_indicators} показателей в {len(sheets_config)} листах")

        # Process sheets
        created_sheets = []
        analytic_sort = 1
        sheet_sort = 0
        total_cells = 0

        for sheet_cfg in sheets_config:
            excel_name = sheet_cfg["excel_name"]
            display_name = sheet_cfg.get("display_name", excel_name)
            sheet_display = display_name if display_name != excel_name else excel_name
            indicators = sheet_cfg.get("indicators", [])

            if excel_name not in wb_formulas.sheetnames or not indicators:
                continue

            ws_f = wb_formulas[excel_name]
            ws_d = wb_data[excel_name]

            sheet_indicators = _count_indicators(indicators)
            yield event(f"📋 Создаю лист «{sheet_display}» ({done_indicators}/{total_indicators})...")

            indicator_analytic_id = str(uuid.uuid4())
            await db.execute(
                """INSERT INTO analytics (id, model_id, name, code, icon, data_type, sort_order)
                   VALUES (?,?,?,?,?,?,?)""",
                (indicator_analytic_id, model_id, f"Показатели ({excel_name})",
                 f"indicators_{excel_name.lower().replace('.', '_').replace('+', '_')}",
                 "ListAltOutlined", "sum", analytic_sort),
            )
            analytic_sort += 1

            for sort_i, (fname, fcode, ftype) in enumerate([
                ("Наименование", "name", "string"),
                ("Единица измерения", "unit", "string"),
            ]):
                await db.execute(
                    "INSERT INTO analytic_fields (id, analytic_id, name, code, data_type, sort_order) VALUES (?,?,?,?,?,?)",
                    (str(uuid.uuid4()), indicator_analytic_id, fname, fcode, ftype, sort_i),
                )

            row_to_rid, rid_to_formula = await _create_indicator_records(db, indicator_analytic_id, indicators)

            pebble_sheet_id = str(uuid.uuid4())
            await db.execute("INSERT INTO sheets (id, model_id, name, sort_order, excel_code) VALUES (?,?,?,?,?)",
                             (pebble_sheet_id, model_id, sheet_display, sheet_sort, excel_name))
            sheet_sort += 1

            for bind_idx, aid in enumerate([period_analytic_id, indicator_analytic_id]):
                await db.execute(
                    "INSERT INTO sheet_analytics (id, sheet_id, analytic_id, sort_order) VALUES (?,?,?,?)",
                    (str(uuid.uuid4()), pebble_sheet_id, aid, bind_idx),
                )

            users = await db.execute_fetchall("SELECT id FROM users")
            for u in users:
                try:
                    await db.execute(
                        "INSERT INTO sheet_permissions (id, sheet_id, user_id, can_view, can_edit) VALUES (?,?,?,1,1)",
                        (str(uuid.uuid4()), pebble_sheet_id, u["id"]),
                    )
                except Exception:
                    pass

            # Import cells
            sheet_periods = _detect_periods_from_headers(ws_d, min(ws_d.max_column or 1, 200))
            col_to_period_rid = {}
            for sp in sheet_periods:
                d = sp["date"]
                if isinstance(d, datetime):
                    key = f"{d.year}-{d.month:02d}"
                    if key in month_record_ids:
                        col_to_period_rid[sp["col"]] = month_record_ids[key]

            sorted_period_cols = sorted(col_to_period_rid.keys())
            first_period_rid = col_to_period_rid[sorted_period_cols[0]] if sorted_period_cols else None

            cell_count = 0
            for row_num, indicator_rid in row_to_rid.items():
                formula_info = rid_to_formula.get(indicator_rid)
                for col_num, period_rid in col_to_period_rid.items():
                    val = ws_d.cell(row_num, col_num).value
                    if val is None:
                        continue
                    # Skip non-numeric strings (month codes like m1/m2, labels, etc.)
                    if isinstance(val, str):
                        try:
                            val = float(val.replace(",", ".").replace(" ", ""))
                        except (ValueError, AttributeError):
                            continue
                    if formula_info:
                        rule = "formula"
                        is_first = (period_rid == first_period_rid)
                        if is_first and formula_info.get("formula_first"):
                            formula_text = formula_info["formula_first"]
                        else:
                            formula_text = formula_info.get("formula", "")
                    else:
                        rule = "manual"
                        formula_text = ""
                    try:
                        await db.execute(
                            "INSERT INTO cell_data (id, sheet_id, coord_key, value, data_type, rule, formula) VALUES (?,?,?,?,?,?,?)",
                            (str(uuid.uuid4()), pebble_sheet_id, f"{period_rid}|{indicator_rid}", str(val), "sum", rule, formula_text),
                        )
                        cell_count += 1
                    except Exception:
                        pass

            total_cells += cell_count
            done_indicators += sheet_indicators
            created_sheets.append({"name": sheet_display, "id": pebble_sheet_id, "cells": cell_count})
            yield event(f"   ✓ «{sheet_display}»: {len(row_to_rid)} показателей, {cell_count} ячеек ({done_indicators}/{total_indicators})")

        await db.commit()

        # ── Post-import validation ──
        expected_sheets = len(sheet_names)
        actual_sheets = len(created_sheets)
        if actual_sheets < expected_sheets:
            missing = set(sheet_names) - {sc["excel_name"] for sc in sheets_config}
            yield event(f"⚠ Импортировано {actual_sheets}/{expected_sheets} листов. Пропущены: {', '.join(missing)}")

        # Validate cell counts against Excel
        for cs in created_sheets:
            excel_name = None
            for sc in sheets_config:
                if sc.get("display_name", sc["excel_name"]) == cs["name"].split(". ", 1)[-1] or sc["excel_name"] in cs["name"]:
                    excel_name = sc["excel_name"]
                    break
            if excel_name and excel_name in wb_data.sheetnames:
                ws_check = wb_data[excel_name]
                dsc = next((sc.get("data_start_col", 4) for sc in sheets_config if sc["excel_name"] == excel_name), 4)
                # Count non-empty data cells in Excel
                excel_cells = 0
                for r in range(1, min(ws_check.max_row or 1, 500) + 1):
                    for c in range(dsc, min((ws_check.max_column or 1) + 1, 50)):
                        if ws_check.cell(r, c).value is not None:
                            excel_cells += 1
                ratio = cs["cells"] / excel_cells * 100 if excel_cells > 0 else 0
                if ratio < 80:
                    yield event(f"⚠ «{cs['name']}»: импортировано {cs['cells']}/{excel_cells} ячеек ({ratio:.0f}%)")

        yield event(f"✅ Импорт завершён! {len(created_sheets)} листов, {total_cells} ячеек",
                     {"done": True, "model_id": model_id, "model_name": model_name_final})

    return StreamingResponse(generate(), media_type="text/event-stream")


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
