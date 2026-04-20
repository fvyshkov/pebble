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
- INPUT(bg=color) = cell has non-default background color (yellow, beige, green, etc.)
  This usually means manual user input — numbers typed by hand, not computed.
  If a row has INPUT(bg=...) AND has numbers but NO Excel formula, it's almost certainly manual.
  If a row has INPUT(bg=...) AND ALSO has a FORMULA flag, the color overrides: treat as manual (keep the number, ignore the formula).
- FORMULA = cell contains an Excel formula (computed)

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
12. GROUPING BY NAME PATTERN — MANDATORY, IGNORE MISSING FORMULA: A row is a group header whenever its label matches ANY grouping pattern below, even if the Excel cell is empty or contains no formula. NEVER mark such a row as "manual". Grouping patterns (case-insensitive):
    - ends with "в т.ч.:" / "в т.ч." / "в том числе:" / "в том числе" / "включая:" / "включая"
      EXAMPLE: "общее количество партнеров, в т.ч.:" → is_group=true, rule="sum_children", children=[all indented rows below it]
    - starts with "Итого" / "Всего" / "Всего по " / "Общее " / "Общий " / "Общая "
    - "Суммарн" / "Сумма " prefixes
    For ALL matching rows: is_group=true, rule="sum_children", NO formula. Rows below at greater indent are its children.
13. INDENTATION RULE: A row with indent=N whose immediately following rows have indent>N is a parent group header. Attach those deeper-indent rows as children, even if the header row has no bold/formula. This is a HARD rule — do not put indented rows as siblings of their parent.
14. Sheet 0 / «параметры» caveat: most rows there ARE manual input, BUT a row matching rule 12 above is still a grouping header with `sum_children`, not manual. Do not blanket-mark every row on sheet 0 as manual.

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

        # Check cell background color (non-default = likely manual input)
        is_input = False
        cell_color = None
        for c in range(4, min(15, max_col + 1)):
            cv = ws.cell(r, c).value
            if cv is not None:
                cell_color = _get_cell_bg_color(ws.cell(r, c))
                if cell_color:
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
        if is_input and cell_color:
            flags.append(f"INPUT(bg={cell_color})")

        flag_str = f" [{', '.join(flags)}]" if flags else ""
        formula_str = ""
        if formula_m1:
            formula_str = f" F1={formula_m1}"
            if formula_m2 and formula_m2 != formula_m1:
                formula_str += f" F2={formula_m2}"
        lines.append(f"  Row {r}: {' | '.join(labels)}{flag_str}{formula_str}")

    return "\n".join(lines)


def _get_cell_bg_color(cell) -> str | None:
    """Return human-readable background color description, or None if default/white."""
    fill = cell.fill
    if not fill or fill.patternType != "solid":
        return None
    fg = fill.fgColor
    if not fg:
        return None
    # Theme-based colors
    if fg.type == "theme":
        try:
            theme = int(fg.theme)
            theme_names = {
                0: None,  # white / default
                1: None,  # black text (not a bg color)
                2: None,  # light gray (often default)
                3: "dark gray", 4: "blue", 5: "orange",
                6: "green", 7: "yellow", 8: "teal", 9: "purple",
            }
            return theme_names.get(theme)
        except (TypeError, ValueError):
            pass
    # RGB-based
    if fg.type == "rgb" and isinstance(fg.rgb, str):
        raw = fg.rgb
        # Strip alpha prefix if present (e.g. "00FFFFFF" → "FFFFFF")
        hex6 = raw[-6:] if len(raw) >= 6 else raw
        if hex6.upper() in ("FFFFFF", "000000"):
            return None  # white or black (default)
        r = int(hex6[0:2], 16)
        g = int(hex6[2:4], 16)
        b = int(hex6[4:6], 16)
        if r > 240 and g > 240 and b > 240:
            return None  # near-white
        # Describe the color
        if r > 200 and g > 200 and b < 100:
            return "yellow"
        if r > 200 and g > 180 and b < 150 and b < r - 60:
            return "beige/light yellow"
        if r > 200 and g < 150 and b < 150:
            return "red/pink"
        if r < 150 and g > 200 and b < 150:
            return "green"
        if r < 150 and g < 150 and b > 200:
            return "blue"
        if r > 200 and g > 150 and b < 100:
            return "orange"
        return f"#{hex6}"
    return None


def _is_input_cell(cell) -> bool:
    """Check if cell has a non-default background color (likely manual input)."""
    return _get_cell_bg_color(cell) is not None


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
    """Legacy per-sheet-text cache (pre-shared-cache). Keeps prior warmed
    entries usable. New calls also go through backend.llm_cache via
    cached_messages_create in chat.py."""
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
    # Build lowercase month-name → month-number lookup for text-based headers
    _MONTH_LOWER = {name.lower(): i + 1 for i, name in enumerate(MONTH_NAMES_RU)}

    periods = []
    date_row = None
    text_date_row = None  # fallback: row with text like "январь 2026"

    for r in range(1, 7):
        for c in range(1, min(max_col + 1, 50)):
            v = ws.cell(r, c).value
            if isinstance(v, datetime):
                date_row = r
                break
            # Check for text-based month headers: "январь 2026", "Февраль 2025", etc.
            if isinstance(v, str) and not text_date_row:
                parts = v.strip().split()
                if len(parts) == 2 and parts[0].lower() in _MONTH_LOWER:
                    try:
                        int(parts[1])
                        text_date_row = r
                    except ValueError:
                        pass
        if date_row:
            break

    if date_row is None and text_date_row is None:
        return []

    seen_months = set()
    scan_row = date_row or text_date_row
    for c in range(1, max_col + 1):
        v = ws.cell(scan_row, c).value
        year, month = None, None
        if isinstance(v, datetime):
            year, month = v.year, v.month
        elif isinstance(v, str):
            parts = v.strip().split()
            if len(parts) == 2 and parts[0].lower() in _MONTH_LOWER:
                try:
                    year = int(parts[1])
                    month = _MONTH_LOWER[parts[0].lower()]
                except ValueError:
                    pass
        if year and month:
            normalized = datetime(year, month, 1)
            month_key = f"{year}-{month:02d}"
            if month_key in seen_months:
                continue  # Skip duplicate months
            seen_months.add(month_key)
            periods.append({
                "col": c,
                "name": f"{MONTH_NAMES_RU[month - 1]} {year}",
                "date": normalized,
            })

    return periods


# ── Total column detection (year / quarter totals) ────────────────────────

import re as _re_total

_YEAR_RE = _re_total.compile(r'^\s*(\d{4})\s*$')
_QUARTER_RE = _re_total.compile(
    r'(?:\d[- ]?й?\s*)?(?:кв(?:артал)?|Q)\s*\d?\s*\d{4}|'
    r'\d{4}\s*(?:кв(?:артал)?|Q)\s*\d|'
    r'(?:итого|всего)\s+за\s+(?:год|квартал)',
    _re_total.IGNORECASE,
)
_TOTAL_KEYWORDS = _re_total.compile(
    r'(?:итого|всего|total|год|year)',
    _re_total.IGNORECASE,
)
# Simple SUM pattern: =SUM(...) where interior is a single range or comma-separated refs
_SUM_ONLY_RE = _re_total.compile(
    r'^=?\s*SUM\s*\([^)]+\)\s*$',
    _re_total.IGNORECASE,
)
# Simple addition: =A1+B1+C1+... (all same-row refs, no division/multiplication)
_SIMPLE_ADD_RE = _re_total.compile(
    r'^=?\s*\$?[A-Z]{1,3}\$?\d+(?:\s*\+\s*\$?[A-Z]{1,3}\$?\d+)+\s*$',
    _re_total.IGNORECASE,
)


def _detect_total_columns(
    ws_data, ws_formulas, col_to_period_rid: dict[int, str], max_col: int,
) -> list[dict]:
    """Detect year/quarter total columns by scanning headers.

    Returns list of {col, type: 'year'|'quarter', label}.
    Skips columns that are already in col_to_period_rid (monthly data columns).
    """
    totals = []
    period_cols = set(col_to_period_rid.keys())

    for c in range(1, min(max_col + 1, 100)):
        if c in period_cols:
            continue
        # Scan first 6 header rows for year/quarter/total markers
        for r in range(1, 7):
            val = ws_data.cell(r, c).value
            if val is None:
                continue
            sval = str(val).strip()
            # Year total: header is just a 4-digit year
            ym = _YEAR_RE.match(sval)
            if ym:
                year = int(ym.group(1))
                if 1990 <= year <= 2100:
                    totals.append({"col": c, "type": "year", "label": sval})
                    break
            # Quarter total
            if _QUARTER_RE.search(sval):
                totals.append({"col": c, "type": "quarter", "label": sval})
                break
            # Generic total keywords
            if _TOTAL_KEYWORDS.search(sval):
                totals.append({"col": c, "type": "year", "label": sval})
                break

    return totals


def _is_sum_formula(excel_formula: str) -> bool:
    """Check if an Excel formula is a plain SUM or simple addition."""
    if not excel_formula or not isinstance(excel_formula, str):
        return False
    f = excel_formula.strip()
    if not f.startswith("="):
        return False
    if _SUM_ONLY_RE.match(f):
        return True
    if _SIMPLE_ADD_RE.match(f):
        return True
    return False


# Pattern for AVERAGE(range)
_AVERAGE_RE = _re_total.compile(r'^=?\s*AVERAGE\s*\([^)]+\)\s*$', _re_total.IGNORECASE)
# Single cell ref: =O13 or =$O$13
_SINGLE_REF_RE = _re_total.compile(r'^=?\s*\$?([A-Z]{1,3})\$?(\d+)\s*$', _re_total.IGNORECASE)
# Same-column formula: refs only to cells in the same column as the total
# E.g. =AO5/AO2 where AO is the total column
_SAME_COL_FORMULA_RE = _re_total.compile(r'\$?([A-Z]{1,3})\$?(\d+)')


def _classify_consolidation_formula(
    excel_formula: str,
    target_col: int,
    row_num: int,
    row_to_name: dict[int, str],
    period_cols: set[int],
) -> str | None:
    """Classify a year/quarter total column formula and return a Pebble
    consolidation formula, or None to skip (SUM = default).

    Returns:
      - None: SUM/default, don't store
      - "AVERAGE": average consolidation
      - "LAST": take last child value (stock/balance indicator)
      - formula string: Pebble formula like "[indicator_a] / [indicator_b]"
    """
    from openpyxl.utils import column_index_from_string

    if not excel_formula or not isinstance(excel_formula, str):
        return None
    f = excel_formula.strip().lstrip("=").strip()
    if not f:
        return None

    full = excel_formula.strip()

    # 1. SUM → skip
    if _is_sum_formula(full):
        return None

    # 2. AVERAGE → skip (let LLM determine the correct ratio formula).
    # AVERAGE is almost always wrong for business metrics — weighted averages
    # like "average loans per partner" should be [total_loans]/[partners],
    # not arithmetic mean of monthly values.
    if _AVERAGE_RE.match(full):
        return None

    # 3. Single cell reference (e.g. =O13) — points to a monthly column in same row
    sm = _SINGLE_REF_RE.match(full)
    if sm:
        ref_col = column_index_from_string(sm.group(1).replace("$", ""))
        ref_row = int(sm.group(2).replace("$", ""))
        if ref_row == row_num and ref_col in period_cols:
            return "LAST"
        # Single ref to different row in same column → just a copy, skip
        if ref_col == target_col:
            return None
        return None

    # 4. Formula with refs only in the same total column → ratio formula
    # E.g. =AO5/AO2 where AO is col 41
    refs = _SAME_COL_FORMULA_RE.findall(f)
    if refs:
        all_same_col = all(
            column_index_from_string(col_str.replace("$", "")) == target_col
            for col_str, _ in refs
        )
        if all_same_col:
            # Translate: replace each cell ref with [indicator_name]
            result = f
            for col_str, row_str in refs:
                ref_row = int(row_str.replace("$", ""))
                name = row_to_name.get(ref_row)
                if not name:
                    return None  # can't resolve → skip
                result = result.replace(f"{col_str}{row_str}", f"[{name}]", 1)
                # Also handle $ variants
                result = result.replace(f"${col_str}${row_str}", f"[{name}]")
                result = result.replace(f"${col_str}{row_str}", f"[{name}]")
                result = result.replace(f"{col_str}${row_str}", f"[{name}]")
            return result

    return None


async def _extract_and_store_consolidation_rules(
    db,
    ws_formulas,
    total_cols: list[dict],
    row_to_rid: dict[int, str],
    row_to_name: dict[int, str],
    pebble_sheet_id: str,
    period_cols: set[int] | None = None,
    sheet_row_maps: dict[str, dict[int, str]] | None = None,
    sheet_display_names: dict[str, str] | None = None,
) -> int:
    """Extract consolidation formulas from total columns and store non-SUM ones
    as indicator_formula_rules(kind='consolidation').

    Returns the number of rules written.
    """
    if not total_cols:
        return 0

    # Prefer year-level total columns over quarter-level
    year_cols = [tc for tc in total_cols if tc["type"] == "year"]
    target_cols = year_cols if year_cols else total_cols

    # Use the first suitable total column
    target_col = target_cols[0]["col"]
    pcols = period_cols or set()

    written = 0
    seen_indicators: set[str] = set()

    for row_num, indicator_rid in row_to_rid.items():
        if indicator_rid in seen_indicators:
            continue

        cell_val = ws_formulas.cell(row_num, target_col).value
        if not cell_val or not isinstance(cell_val, str) or not cell_val.startswith("="):
            continue

        pebble_formula = _classify_consolidation_formula(
            cell_val, target_col, row_num, row_to_name, pcols,
        )
        if pebble_formula is None:
            continue  # SUM or unrecognized → default

        seen_indicators.add(indicator_rid)
        await db.execute(
            "INSERT OR IGNORE INTO indicator_formula_rules "
            "(id, sheet_id, indicator_id, kind, scope_json, priority, formula) "
            "VALUES (?, ?, ?, 'consolidation', '{}', 0, ?)",
            (str(uuid.uuid4()), pebble_sheet_id, indicator_rid, pebble_formula),
        )
        written += 1

    return written


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


# ── Post-process: fix missing parent-child grouping ─────────────────────────

import re as _re

_GROUP_PATTERN = _re.compile(
    r'(?:в\s+т\.?\s*ч\.?\s*:?|в\s+том\s+числе\s*:?|включая\s*:?)\s*$',
    _re.IGNORECASE,
)
_GROUP_PREFIX = _re.compile(
    r'^(?:итого|всего|общее\s|общий\s|общая\s|суммарн|сумма\s)',
    _re.IGNORECASE,
)


def _fix_indicator_hierarchy(indicators: list[dict]) -> list[dict]:
    """Post-process indicator list: ensure rows matching grouping name patterns
    (e.g. "в т.ч.:", "Итого") have subsequent rows nested as children.
    This is a safety net after Claude/heuristic analysis — it only promotes
    flat siblings into children when the grouping wasn't already detected."""
    result: list[dict] = []
    i = 0
    while i < len(indicators):
        item = indicators[i]
        name = (item.get("name") or "").strip()
        already_group = item.get("is_group", False) and len(item.get("children", [])) > 0

        # Recursively fix children of already-detected groups
        if item.get("children"):
            item["children"] = _fix_indicator_hierarchy(item["children"])

        # If already a group with children, keep it
        if already_group:
            result.append(item)
            i += 1
            continue

        # Check if this row matches a grouping pattern ("в т.ч.:", "Итого", etc.)
        is_group_by_name = bool(_GROUP_PATTERN.search(name)) or bool(_GROUP_PREFIX.match(name))

        if is_group_by_name and not already_group:
            # Collect subsequent flat siblings as children
            item["is_group"] = True
            if item.get("rule") in (None, "manual", ""):
                item["rule"] = "sum_children"
            existing_children = item.get("children", [])
            j = i + 1
            while j < len(indicators):
                next_item = indicators[j]
                next_name = (next_item.get("name") or "").strip()
                # Stop collecting if we hit another group header or empty
                if (bool(_GROUP_PATTERN.search(next_name)) or
                    bool(_GROUP_PREFIX.match(next_name)) or
                    (next_item.get("is_group", False) and len(next_item.get("children", [])) > 0)):
                    break
                existing_children.append(next_item)
                j += 1
            item["children"] = existing_children
            result.append(item)
            i = j
        else:
            result.append(item)
            i += 1

    return result


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

    # Build row_to_name mapping for formula translator
    row_to_name: dict[int, str] = {}
    def collect_names(items: list[dict]):
        for item in items:
            if item.get("row") and item.get("name"):
                row_to_name[item["row"]] = item["name"]
            if item.get("children"):
                collect_names(item["children"])
    collect_names(indicators)

    return row_to_rid, rid_to_formula, row_to_name


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
        # Scan BOTH formulas and data_only workbooks (formula-derived dates like =B1+31
        # only resolve in data_only mode)
        ws_d_scan = wb_data[sn] if sn in wb_data.sheetnames else None
        for scan_ws in ([ws, ws_d_scan] if ws_d_scan else [ws]):
            for r in range(1, 7):
                for c in range(1, min((scan_ws.max_column or 1) + 1, 200)):
                    v = scan_ws.cell(r, c).value
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

        # Fix grouping hierarchy (safety net after Claude/heuristic analysis)
        indicators = _fix_indicator_hierarchy(indicators)

        # Create hierarchical indicator records
        row_to_rid, rid_to_formula, row_to_name = await _create_indicator_records(db, indicator_analytic_id, indicators)

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
            is_main = 1 if aid == indicator_analytic_id else 0
            await db.execute(
                "INSERT INTO sheet_analytics (id, sheet_id, analytic_id, sort_order, is_main) VALUES (?,?,?,?,?)",
                (sa_id, pebble_sheet_id, aid, bind_idx, is_main),
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

                # Cell color is authoritative: yellow/beige fill (theme=7) means
                # "user input" in the source spreadsheet — treat as manual even
                # if a formula exists. We keep the numeric value Excel computed.
                is_yellow = _is_input_cell(ws_f.cell(row_num, col_num))

                # Determine rule and formula from Claude analysis
                if is_yellow:
                    rule = "manual"
                    formula_text = ""
                elif formula_info:
                    rule = "formula"
                    is_first = (period_rid == first_period_rid)
                    if is_first and formula_info.get("formula_first"):
                        formula_text = formula_info["formula_first"]
                    else:
                        formula_text = formula_info.get("formula", "")
                else:
                    # Fallback: check Excel formula workbook directly
                    excel_formula = ws_f.cell(row_num, col_num).value
                    if isinstance(excel_formula, str) and excel_formula.startswith("="):
                        rule = "formula"
                        formula_text = excel_formula  # raw Excel formula as reference
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

        # Extract consolidation formulas from total columns (year/quarter totals)
        total_cols = _detect_total_columns(ws_d, ws_f, col_to_period_rid, min(ws_d.max_column or 1, 200))
        if total_cols:
            n_consol = await _extract_and_store_consolidation_rules(
                db, ws_f, total_cols, row_to_rid, row_to_name, pebble_sheet_id,
                period_cols=set(col_to_period_rid.keys()),
            )
            if n_consol:
                log.info(f"[import] Extracted {n_consol} consolidation formulas from Excel totals for «{sheet_display}»")

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
            # Scan BOTH formulas and data_only workbooks for dates
            # (formula-derived dates like =B1+31 only resolve in data_only)
            ws_d_scan = wb_data[sn] if sn in wb_data.sheetnames else None
            for scan_ws in ([ws, ws_d_scan] if ws_d_scan else [ws]):
                for r in range(1, 7):
                    for c in range(1, min((scan_ws.max_column or 1) + 1, 200)):
                        v = scan_ws.cell(r, c).value
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

            # Fix grouping hierarchy (safety net after Claude/heuristic analysis)
            indicators = _fix_indicator_hierarchy(indicators)

            row_to_rid, rid_to_formula, row_to_name = await _create_indicator_records(db, indicator_analytic_id, indicators)

            pebble_sheet_id = str(uuid.uuid4())
            await db.execute("INSERT INTO sheets (id, model_id, name, sort_order, excel_code) VALUES (?,?,?,?,?)",
                             (pebble_sheet_id, model_id, sheet_display, sheet_sort, excel_name))
            sheet_sort += 1

            for bind_idx, aid in enumerate([period_analytic_id, indicator_analytic_id]):
                is_main = 1 if aid == indicator_analytic_id else 0
                await db.execute(
                    "INSERT INTO sheet_analytics (id, sheet_id, analytic_id, sort_order, is_main) VALUES (?,?,?,?,?)",
                    (str(uuid.uuid4()), pebble_sheet_id, aid, bind_idx, is_main),
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
                    # Cell color is authoritative: yellow fill = user input,
                    # even if a formula exists. Keep the value Excel computed.
                    is_yellow = _is_input_cell(ws_f.cell(row_num, col_num))
                    if is_yellow:
                        rule = "manual"
                        formula_text = ""
                    elif formula_info:
                        rule = "formula"
                        is_first = (period_rid == first_period_rid)
                        if is_first and formula_info.get("formula_first"):
                            formula_text = formula_info["formula_first"]
                        else:
                            formula_text = formula_info.get("formula", "")
                    else:
                        # Fallback: check Excel formula workbook directly
                        excel_formula = ws_f.cell(row_num, col_num).value
                        if isinstance(excel_formula, str) and excel_formula.startswith("="):
                            rule = "formula"
                            formula_text = excel_formula
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

            # Extract consolidation formulas from total columns
            total_cols = _detect_total_columns(ws_d, ws_f, col_to_period_rid, min(ws_d.max_column or 1, 200))
            n_consol = 0
            if total_cols:
                n_consol = await _extract_and_store_consolidation_rules(
                    db, ws_f, total_cols, row_to_rid, row_to_name, pebble_sheet_id,
                    period_cols=set(col_to_period_rid.keys()),
                )

            total_cells += cell_count
            done_indicators += sheet_indicators
            created_sheets.append({"name": sheet_display, "id": pebble_sheet_id, "cells": cell_count})
            consol_msg = f", {n_consol} формул консолидации из Excel" if n_consol else ""
            yield event(f"   ✓ «{sheet_display}»: {len(row_to_rid)} показателей, {cell_count} ячеек{consol_msg} ({done_indicators}/{total_indicators})")

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

        # ── Post-import: detect period-consolidation formulas for ratios/averages ──
        # Ask Claude which indicators should NOT be summed across periods
        # (e.g. interest rates, averages). Tolerant of API failures.
        if os.environ.get("ANTHROPIC_API_KEY"):
            from backend.formula_suggester import suggest_consolidations_for_sheet
            # Get the periods analytic name (used as context in the prompt).
            period_analytic_name = "Периоды"
            try:
                pa = await db.execute_fetchall(
                    "SELECT name FROM analytics WHERE id = ? LIMIT 1",
                    (period_analytic_id,),
                )
                if pa:
                    period_analytic_name = pa[0]["name"]
            except Exception:
                pass
            total_rules = 0
            async def _suggest_one(cs):
                try:
                    return await suggest_consolidations_for_sheet(db, cs["id"], period_analytic_name)
                except Exception as e:
                    print(f"[import] suggest_consolidations failed for {cs['id']}: {e}")
                    return 0
            suggest_results = await asyncio.gather(*[_suggest_one(cs) for cs in created_sheets])
            total_rules = sum(suggest_results)
            # Propagate formulas across sheets for same-named indicators
            if len(created_sheets) > 1:
                from backend.formula_suggester import propagate_consolidations_across_sheets
                try:
                    total_rules += await propagate_consolidations_across_sheets(db, model_id)
                except Exception as e:
                    print(f"[import] propagate_consolidations failed: {e}")
            if total_rules:
                yield event(f"   ✓ Claude подобрал {total_rules} формул консолидации по периодам")

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

            # Check indentation
            cell_label = ws.cell(r, label_col)
            indent = cell_label.alignment.indent if cell_label.alignment and cell_label.alignment.indent else 0

            # Group detection: name pattern, bold, or no data in period columns
            is_group_by_name = bool(_GROUP_PATTERN.search(name)) or bool(_GROUP_PREFIX.match(name))
            is_group_header = is_group_by_name or (is_bold and not has_data) or (
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
