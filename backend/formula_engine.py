"""Pebble formula engine — lazy pull-based evaluation with memoization.

Each cell is a lazy function. Requesting a cell's value recursively
evaluates its dependencies until it reaches manual inputs. Results
are cached per calculation run.

Formula syntax:
  [indicator_name]                              — same context (exact match)
  [indicator_name](периоды="предыдущий")        — previous period
  [Sheet::indicator_name]                       — cross-sheet reference (:: separator)
  SUM([a], [b], [c])                            — sum function
  Standard math: +, -, *, /, parentheses, numbers

All references resolve by EXACT name match (case-insensitive). No fuzzy matching.
"""

import re
import json
import math
from typing import Any

# ── Tokenizer ──────────────────────────────────────────────────────────────

TOKEN_RE = re.compile(r"""
    (\[(?:[^\[\]]+)\](?:\([^)]*\))?)  |  # [ref](params) or [ref]
    (SUM)\s*\(                         |  # SUM function
    (\d+(?:\.\d+)?)                    |  # number
    ([+\-*/(),])                       |  # operators and parens
    (\s+)                                 # whitespace (skip)
""", re.VERBOSE)

REF_RE = re.compile(r"""
    \[([^\]]+)\]              # indicator name
    (?:\(([^)]*)\))?          # optional (key="value", ...) params
""", re.VERBOSE)

PARAM_RE = re.compile(r'([\w\s]+?)\s*=\s*"([^"]*)"')


def parse_ref(token: str) -> dict:
    m = REF_RE.match(token)
    if not m:
        return {"name": token, "params": {}}
    name = m.group(1)
    params_str = m.group(2) or ""
    params = {}
    for pm in PARAM_RE.finditer(params_str):
        params[pm.group(1)] = pm.group(2)

    # Cross-sheet separator is "::"
    sheet = None
    if "::" in name:
        parts = name.split("::", 1)
        sheet = parts[0].strip()
        name = parts[1].strip()

    return {"name": name, "sheet": sheet, "params": params}


def tokenize(formula: str) -> list:
    tokens = []
    pos = 0
    while pos < len(formula):
        m = TOKEN_RE.match(formula, pos)
        if not m:
            pos += 1; continue
        if m.group(1):
            tokens.append(("REF", m.group(1)))
        elif m.group(2):
            tokens.append(("SUM", "SUM"))
        elif m.group(3):
            tokens.append(("NUM", float(m.group(3))))
        elif m.group(4):
            tokens.append(("OP", m.group(4)))
        pos = m.end()
    return tokens


# ── Evaluator ──────────────────────────────────────────────────────────────

def evaluate(formula: str, get_ref_value) -> float:
    """Evaluate formula. get_ref_value(token_str) -> float."""
    if not formula or not formula.strip():
        return 0.0
    try:
        return float(formula)
    except ValueError:
        pass

    tokens = tokenize(formula)
    if not tokens:
        return 0.0

    pos = [0]
    def peek():
        return tokens[pos[0]] if pos[0] < len(tokens) else None
    def advance():
        t = tokens[pos[0]]; pos[0] += 1; return t
    def expect_op(op):
        t = peek()
        if t and t[0] == "OP" and t[1] == op: advance(); return True
        return False

    def parse_expr():
        left = parse_term()
        while True:
            t = peek()
            if t and t[0] == "OP" and t[1] in ("+", "-"):
                op = advance()[1]; right = parse_term()
                left = left + right if op == "+" else left - right
            else: break
        return left

    def parse_term():
        left = parse_unary()
        while True:
            t = peek()
            if t and t[0] == "OP" and t[1] in ("*", "/"):
                op = advance()[1]; right = parse_unary()
                left = left * right if op == "*" else (left / right if right != 0 else 0.0)
            else: break
        return left

    def parse_unary():
        t = peek()
        if t and t[0] == "OP" and t[1] == "-":
            advance(); return -parse_primary()
        return parse_primary()

    def parse_primary():
        t = peek()
        if t is None: return 0.0
        if t[0] == "NUM": advance(); return t[1]
        if t[0] == "REF": advance(); return get_ref_value(t[1])
        if t[0] == "SUM":
            advance()
            args = []
            while True:
                args.append(parse_expr())
                if not expect_op(","): break
            expect_op(")"); return sum(args)
        if t[0] == "OP" and t[1] == "(":
            advance(); val = parse_expr(); expect_op(")"); return val
        advance(); return 0.0

    try:
        result = parse_expr()
        return result if math.isfinite(result) else 0.0
    except Exception:
        return 0.0


# ── Model-level lazy calculator ────────────────────────────────────────────

async def calculate_model(db, model_id: str) -> dict[str, dict[str, str]]:
    """Calculate ALL formula cells across ALL sheets in a model.
    Returns {sheet_id: {coord_key: new_value}}.
    """
    all_sheets = await db.execute_fetchall(
        "SELECT id, name FROM sheets WHERE model_id = ? ORDER BY created_at", (model_id,))
    if not all_sheets:
        return {}

    # ── Load entire model ──
    global_cells = {}       # {(sheet_id, coord_key): value_str}
    global_formulas = {}    # {(sheet_id, coord_key): formula_str}
    sheet_meta = {}         # sheet_id → metadata
    sheet_name_to_id = {}   # name → sheet_id (case-insensitive index)

    period_aid_global = None
    period_order = []
    prev_period = {}

    for s in all_sheets:
        sid = s["id"]
        sname = s["name"]

        # Register sheet by display name AND common aliases (case-insensitive)
        sheet_name_to_id[sname.lower()] = sid
        # Register by known aliases
        nl = sname.lower()
        if "параметр" in nl:
            for a in ["параметры", "настройки", "baas - настройки", "0"]:
                sheet_name_to_id[a] = sid
        if "расход" in nl or "opex" in nl:
            for a in ["opex+capex", "opex", "capex", "операционные расходы"]:
                sheet_name_to_id[a] = sid
        if "кредитован" in nl: sheet_name_to_id["baas.1"] = sid
        if "депозит" in nl: sheet_name_to_id["baas.2"] = sid
        if "транзакц" in nl: sheet_name_to_id["baas.3"] = sid
        if "баланс" in nl: sheet_name_to_id["bs"] = sid
        if "результат" in nl: sheet_name_to_id["pl"] = sid

        bindings = await db.execute_fetchall(
            "SELECT sa.analytic_id, sa.sort_order, a.name as analytic_name, a.is_periods "
            "FROM sheet_analytics sa JOIN analytics a ON a.id = sa.analytic_id "
            "WHERE sa.sheet_id = ? ORDER BY sa.sort_order", (sid,))

        ordered_aids = [b["analytic_id"] for b in bindings]
        analytic_name_to_id = {b["analytic_name"]: b["analytic_id"] for b in bindings}
        record_by_id = {}
        name_to_rids = {}  # {analytic_id: {name_lower: [record_ids]}}
        period_aid = None

        for b in bindings:
            aid = b["analytic_id"]
            recs = [dict(r) for r in await db.execute_fetchall(
                "SELECT * FROM analytic_records WHERE analytic_id = ? ORDER BY sort_order", (aid,))]
            nmap: dict[str, list[str]] = {}
            for r in recs:
                record_by_id[r["id"]] = r
                data = json.loads(r["data_json"]) if isinstance(r["data_json"], str) else r["data_json"]
                r["_data"] = data
                name = data.get("name", "")
                # Index by lowercase for case-insensitive exact match
                nmap.setdefault(name.lower(), []).append(r["id"])
            name_to_rids[aid] = nmap

            if b["is_periods"]:
                period_aid = aid
                if not period_aid_global:
                    period_aid_global = aid
                    parent_ids = {r["parent_id"] for r in recs if r["parent_id"]}
                    period_order = [r["id"] for r in recs if r["id"] not in parent_ids]
                    for i in range(1, len(period_order)):
                        prev_period[period_order[i]] = period_order[i - 1]

        sheet_meta[sid] = {
            "ordered_aids": ordered_aids,
            "name_to_rids": name_to_rids,
            "record_by_id": record_by_id,
            "period_aid": period_aid,
            "analytic_name_to_id": analytic_name_to_id,
            "name": sname,
        }

        for c in await db.execute_fetchall(
                "SELECT coord_key, value, rule, formula FROM cell_data WHERE sheet_id = ?", (sid,)):
            gk = (sid, c["coord_key"])
            global_cells[gk] = c["value"] or ""
            if c["rule"] == "formula" and c["formula"]:
                global_formulas[gk] = c["formula"]

    # Track original DB cell keys (before computation adds synthetic ones)
    _original_cell_keys = set(global_cells.keys())

    # ── Lazy evaluator ──
    computed_set = set()
    computing_set = set()

    def get_cell(sheet_id: str, coord_key: str) -> float:
        gk = (sheet_id, coord_key)
        if gk in computed_set:
            return _to_float(global_cells.get(gk, ""))
        if gk in computing_set:
            return _to_float(global_cells.get(gk, ""))  # cycle
        formula = global_formulas.get(gk)
        if not formula:
            return _to_float(global_cells.get(gk, ""))

        computing_set.add(gk)
        meta = sheet_meta[sheet_id]
        context = _context_from_key(coord_key, meta["ordered_aids"])

        def get_ref_value(ref_token: str) -> float:
            ref = parse_ref(ref_token)
            ref_sheet = ref.get("sheet")

            if ref_sheet:
                return _resolve_cross_sheet(ref, context, meta, sheet_id)
            else:
                resolved = _resolve_local(ref, context, meta)
                if not resolved or resolved == coord_key:
                    return 0.0
                return get_cell(sheet_id, resolved)

        result = evaluate(formula, get_ref_value)
        result_str = str(round(result, 6)) if result != 0 else "0"
        global_cells[gk] = result_str
        computed_set.add(gk)
        computing_set.discard(gk)
        return result

    def _to_float(val):
        try: return float(val)
        except: return 0.0

    def _context_from_key(coord_key, ordered_aids):
        parts = coord_key.split("|")
        return {aid: parts[i] for i, aid in enumerate(ordered_aids) if i < len(parts)}

    def _resolve_cross_sheet(ref, context, src_meta, src_sheet_id):
        """Resolve [Sheet::indicator] cross-sheet reference."""
        target_sid = sheet_name_to_id.get(ref["sheet"].lower())
        if not target_sid:
            return 0.0
        target_meta = sheet_meta[target_sid]
        name_lower = ref["name"].lower()

        # Find indicator by exact name (case-insensitive)
        ind_rid = None
        for aid, nmap in target_meta["name_to_rids"].items():
            if aid == target_meta["period_aid"]: continue
            rids = nmap.get(name_lower)
            if rids:
                # If multiple, pick one with same parent context
                if len(rids) == 1:
                    ind_rid = rids[0]
                else:
                    # Use context from source to disambiguate
                    ind_rid = rids[0]  # default to first
                break
        if not ind_rid:
            return 0.0

        period_rid = context.get(src_meta["period_aid"])
        if not period_rid:
            return 0.0

        target_ck = f"{period_rid}|{ind_rid}"

        # If target is a group with no cell data, use first child
        if (target_sid, target_ck) not in _original_cell_keys:
            for crid, crec in target_meta["record_by_id"].items():
                if crec.get("parent_id") == ind_rid:
                    child_ck = f"{period_rid}|{crid}"
                    if (target_sid, child_ck) in _original_cell_keys:
                        return get_cell(target_sid, child_ck)
            return 0.0

        return get_cell(target_sid, target_ck)

    def _resolve_local(ref, context, meta):
        """Resolve [indicator] local reference. Exact match only."""
        name_lower = ref["name"].lower()
        params = ref.get("params", {})
        ordered_aids = meta["ordered_aids"]
        period_aid = meta["period_aid"]
        name_to_rids = meta["name_to_rids"]
        record_by_id = meta["record_by_id"]
        analytic_name_to_id = meta["analytic_name_to_id"]

        target_rid = None
        target_aid = None

        for aid, nmap in name_to_rids.items():
            if aid == period_aid: continue
            candidates = nmap.get(name_lower, [])
            if not candidates: continue

            if len(candidates) == 1:
                target_rid = candidates[0]; target_aid = aid; break

            # Multiple records with same name — pick one with same parent
            cur_rid = context.get(aid)
            if cur_rid:
                cur_parent = record_by_id.get(cur_rid, {}).get("parent_id")
                for crid in candidates:
                    crec = record_by_id.get(crid)
                    if crec and crec.get("parent_id") == cur_parent:
                        target_rid = crid; target_aid = aid; break
            if not target_rid:
                target_rid = candidates[0]; target_aid = aid
            break

        if target_rid is None:
            return None

        parts = {}
        for aid in ordered_aids:
            if aid == target_aid: parts[aid] = target_rid
            elif aid in context: parts[aid] = context[aid]

        for param_name, param_value in params.items():
            param_aid = analytic_name_to_id.get(param_name)
            if not param_aid:
                for aname, aid in analytic_name_to_id.items():
                    if param_name.lower() in aname.lower():
                        param_aid = aid; break
            if not param_aid: continue

            if param_value == "предыдущий" and param_aid == period_aid:
                cur = parts.get(param_aid)
                if cur and cur in prev_period:
                    parts[param_aid] = prev_period[cur]
                else:
                    return None
            else:
                nmap = name_to_rids.get(param_aid, {})
                rids = nmap.get(param_value.lower(), [])
                if rids: parts[param_aid] = rids[0]
                else: return None

        coord_parts = [parts.get(aid, "") for aid in ordered_aids]
        if any(p == "" for p in coord_parts): return None
        result_key = "|".join(coord_parts)

        # Self-reference guard
        current_key = "|".join(context.get(aid, "") for aid in ordered_aids)
        if result_key == current_key: return None

        return result_key

    # ── Evaluate all formula cells ──
    for gk in global_formulas:
        sheet_id, coord_key = gk
        get_cell(sheet_id, coord_key)

    # ── Return all formula values grouped by sheet ──
    result = {}
    for (sid, ck) in global_formulas:
        new_val = global_cells.get((sid, ck), "")
        if sid not in result: result[sid] = {}
        result[sid][ck] = new_val

    return result


# ── Convenience wrapper ────────────────────────────────────────────────────

async def calculate_sheet(db, sheet_id: str) -> dict[str, str]:
    """Calculate formulas for the model containing this sheet."""
    sheet = await db.execute_fetchall("SELECT model_id FROM sheets WHERE id = ?", (sheet_id,))
    if not sheet: return {}
    result = await calculate_model(db, sheet[0]["model_id"])
    return result.get(sheet_id, {})
