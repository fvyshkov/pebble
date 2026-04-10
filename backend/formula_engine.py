"""Pebble formula engine — parses and evaluates cell formulas.

Formula syntax:
  [indicator_name]                              — same context (period, product, etc.)
  [indicator_name](периоды="предыдущий")        — previous period
  [indicator_name](периоды="Январь 2026")       — specific period value
  [indicator_name](продукты="Потребительский кредит")  — specific analytic value
  [Sheet.indicator_name]                        — cross-sheet reference
  SUM([a], [b], [c])                            — sum of references

Supports: +, -, *, /, parentheses, numbers, unary minus.
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
    \[([^\]]+)\]              # indicator name (may include Sheet. prefix)
    (?:\(([^)]*)\))?          # optional (key="value", ...) params
""", re.VERBOSE)

PARAM_RE = re.compile(r'([\w\s]+?)\s*=\s*"([^"]*)"')


def parse_ref(token: str) -> dict:
    """Parse a reference token like [name](key="val") into structured form."""
    m = REF_RE.match(token)
    if not m:
        return {"name": token, "params": {}}
    name = m.group(1)
    params_str = m.group(2) or ""
    params = {}
    for pm in PARAM_RE.finditer(params_str):
        params[pm.group(1)] = pm.group(2)

    # Handle Sheet.indicator cross-sheet reference
    # Split on "." only if: no space before the dot AND char after dot is not a space
    # This avoids splitting "ср. сумма" or "ср. % ставка" which are abbreviations
    sheet = None
    if "." in name:
        dot_pos = name.index(".")
        left = name[:dot_pos].strip()
        right = name[dot_pos + 1:]
        # Abbreviations always have space after dot; sheet refs don't
        if left and right and not right.startswith(" "):
            sheet = left
            name = right.strip()

    return {"name": name, "sheet": sheet, "params": params}


def tokenize(formula: str) -> list:
    """Tokenize formula into list of (type, value) tuples."""
    tokens = []
    pos = 0
    while pos < len(formula):
        m = TOKEN_RE.match(formula, pos)
        if not m:
            pos += 1
            continue
        if m.group(1):  # reference
            tokens.append(("REF", m.group(1)))
        elif m.group(2):  # SUM
            tokens.append(("SUM", "SUM"))
        elif m.group(3):  # number
            tokens.append(("NUM", float(m.group(3))))
        elif m.group(4):  # operator/paren
            tokens.append(("OP", m.group(4)))
        # skip whitespace
        pos = m.end()
    return tokens


# ── Evaluator ──────────────────────────────────────────────────────────────

class FormulaContext:
    """Provides cell value resolution for formula evaluation."""

    def __init__(self, cells: dict[str, str], resolve_ref: callable):
        """
        cells: {coord_key: value_str}
        resolve_ref: (ref_dict, current_coord_context) -> coord_key or None
        """
        self.cells = cells
        self.resolve_ref = resolve_ref
        self.current_context = {}  # set before each evaluation

    def get_value(self, ref_token: str) -> float:
        ref = parse_ref(ref_token)
        coord_key = self.resolve_ref(ref, self.current_context)
        if coord_key is None:
            return 0.0
        val = self.cells.get(coord_key, "")
        try:
            return float(val)
        except (ValueError, TypeError):
            return 0.0


def evaluate(formula: str, ctx: FormulaContext) -> float:
    """Evaluate a Pebble formula string. Returns numeric result."""
    if not formula or not formula.strip():
        return 0.0

    # Simple constant check
    try:
        return float(formula)
    except ValueError:
        pass

    tokens = tokenize(formula)
    if not tokens:
        return 0.0

    # Recursive descent parser
    pos = [0]

    def peek():
        if pos[0] < len(tokens):
            return tokens[pos[0]]
        return None

    def advance():
        t = tokens[pos[0]]
        pos[0] += 1
        return t

    def expect_op(op):
        t = peek()
        if t and t[0] == "OP" and t[1] == op:
            advance()
            return True
        return False

    def parse_expr():
        left = parse_term()
        while True:
            t = peek()
            if t and t[0] == "OP" and t[1] in ("+", "-"):
                op = advance()[1]
                right = parse_term()
                left = left + right if op == "+" else left - right
            else:
                break
        return left

    def parse_term():
        left = parse_unary()
        while True:
            t = peek()
            if t and t[0] == "OP" and t[1] in ("*", "/"):
                op = advance()[1]
                right = parse_unary()
                if op == "*":
                    left = left * right
                else:
                    left = left / right if right != 0 else 0.0
            else:
                break
        return left

    def parse_unary():
        t = peek()
        if t and t[0] == "OP" and t[1] == "-":
            advance()
            return -parse_primary()
        return parse_primary()

    def parse_primary():
        t = peek()
        if t is None:
            return 0.0
        if t[0] == "NUM":
            advance()
            return t[1]
        if t[0] == "REF":
            advance()
            return ctx.get_value(t[1])
        if t[0] == "SUM":
            advance()
            # Collect arguments until closing paren
            args = []
            while True:
                args.append(parse_expr())
                if not expect_op(","):
                    break
            expect_op(")")
            return sum(args)
        if t[0] == "OP" and t[1] == "(":
            advance()
            val = parse_expr()
            expect_op(")")
            return val
        advance()  # skip unknown
        return 0.0

    try:
        result = parse_expr()
        return result if math.isfinite(result) else 0.0
    except Exception:
        return 0.0


# ── Sheet calculator ───────────────────────────────────────────────────────

async def calculate_sheet(db, sheet_id: str) -> dict[str, str]:
    """Calculate all formula cells in a sheet. Returns {coord_key: computed_value}.

    1. Load sheet structure (analytics, records, cells)
    2. Build reference resolver
    3. Topological sort by dependencies
    4. Evaluate in order
    """
    # Load bindings
    bindings = await db.execute_fetchall(
        "SELECT sa.analytic_id, sa.sort_order, a.name as analytic_name, a.is_periods "
        "FROM sheet_analytics sa JOIN analytics a ON a.id = sa.analytic_id "
        "WHERE sa.sheet_id = ? ORDER BY sa.sort_order",
        (sheet_id,),
    )
    if not bindings:
        return {}

    # Load records per analytic: {analytic_id: [{id, parent_id, data_json, sort_order}]}
    analytic_records = {}
    record_by_id = {}
    name_to_rids = {}  # {analytic_id: {name: [record_id, ...]}}
    period_analytic_id = None
    period_order = []  # ordered list of leaf period record IDs

    for b in bindings:
        aid = b["analytic_id"]
        recs = await db.execute_fetchall(
            "SELECT * FROM analytic_records WHERE analytic_id = ? ORDER BY sort_order",
            (aid,),
        )
        recs = [dict(r) for r in recs]
        analytic_records[aid] = recs
        name_map: dict[str, list[str]] = {}
        for r in recs:
            record_by_id[r["id"]] = r
            data = json.loads(r["data_json"]) if isinstance(r["data_json"], str) else r["data_json"]
            r["_data"] = data
            name = data.get("name", "")
            name_map.setdefault(name, []).append(r["id"])

        name_to_rids[aid] = name_map

        if b["is_periods"]:
            period_analytic_id = aid
            # Get ordered leaf periods (no children)
            parent_ids = {r["parent_id"] for r in recs if r["parent_id"]}
            period_order = [r["id"] for r in recs if r["id"] not in parent_ids]

    # Analytic name → analytic_id
    analytic_name_to_id = {b["analytic_name"]: b["analytic_id"] for b in bindings}

    # Build ordered analytics list (for coord_key construction)
    ordered_analytic_ids = [b["analytic_id"] for b in bindings]

    # Load cells
    cells_raw = await db.execute_fetchall(
        "SELECT coord_key, value, rule, formula FROM cell_data WHERE sheet_id = ?",
        (sheet_id,),
    )
    cells = {}
    formula_cells = {}
    for c in cells_raw:
        cells[c["coord_key"]] = c["value"] or ""
        if c["rule"] == "formula" and c["formula"]:
            formula_cells[c["coord_key"]] = c["formula"]

    # Previous period map
    prev_period = {}
    for i in range(1, len(period_order)):
        prev_period[period_order[i]] = period_order[i - 1]

    # ── Load cross-sheet data for [Sheet.indicator] references ──
    # Get model_id for this sheet
    sheet_row = await db.execute_fetchall("SELECT model_id FROM sheets WHERE id = ?", (sheet_id,))
    model_id = sheet_row[0]["model_id"] if sheet_row else None

    # Cross-sheet lookup: {sheet_display_name: {indicator_name: {period_rid: value}}}
    xsheet = {}
    if model_id:
        other_sheets = await db.execute_fetchall(
            "SELECT id, name FROM sheets WHERE model_id = ? AND id != ?", (model_id, sheet_id))
        for os in other_sheets:
            os_bindings = await db.execute_fetchall(
                "SELECT sa.analytic_id, a.is_periods FROM sheet_analytics sa "
                "JOIN analytics a ON a.id = sa.analytic_id WHERE sa.sheet_id = ? ORDER BY sa.sort_order",
                (os["id"],))
            os_period_aid = None
            os_indicator_aid = None
            for ob in os_bindings:
                if ob["is_periods"]:
                    os_period_aid = ob["analytic_id"]
                else:
                    os_indicator_aid = ob["analytic_id"]
            if not os_period_aid or not os_indicator_aid:
                continue
            # Build indicator name → record_id map
            os_recs = await db.execute_fetchall(
                "SELECT id, data_json FROM analytic_records WHERE analytic_id = ?", (os_indicator_aid,))
            os_name_to_rid = {}
            for r in os_recs:
                d = json.loads(r["data_json"]) if isinstance(r["data_json"], str) else r["data_json"]
                os_name_to_rid[d.get("name", "")] = r["id"]
            # Load cells
            os_cells = await db.execute_fetchall(
                "SELECT coord_key, value FROM cell_data WHERE sheet_id = ?", (os["id"],))
            indicator_data = {}
            for c in os_cells:
                parts = c["coord_key"].split("|")
                if len(parts) == 2:
                    p_rid, i_rid = parts
                    indicator_data[(p_rid, i_rid)] = c["value"] or ""
            xsheet[os["name"]] = {"name_to_rid": os_name_to_rid, "cells": indicator_data}

    # ── Reference resolver ──
    def resolve_ref(ref: dict, context: dict) -> str | None:
        """Resolve a parsed reference to a coord_key."""
        name = ref["name"]
        sheet_name = ref.get("sheet")
        params = ref.get("params", {})

        # ── Cross-sheet reference: [Sheet.indicator] ──
        if sheet_name:
            # Exact match first, then fuzzy (partial name match)
            xs = xsheet.get(sheet_name)
            if not xs:
                for sn, sd in xsheet.items():
                    if sheet_name.lower() in sn.lower() or sn.lower() in sheet_name.lower():
                        xs = sd
                        break
            if not xs:
                return None
            # Find indicator: exact, then fuzzy
            ind_rid = xs["name_to_rid"].get(name)
            if not ind_rid:
                for iname, irid in xs["name_to_rid"].items():
                    if name.lower() in iname.lower() or iname.lower() in name.lower():
                        ind_rid = irid
                        break
            if not ind_rid:
                return None
            # Get the period from current context
            period_rid = context.get(period_analytic_id)
            if not period_rid:
                return None
            val = xs["cells"].get((period_rid, ind_rid))
            if val is not None:
                # Store directly in cells dict with a synthetic key so evaluator picks it up
                synth_key = f"__xsheet__{sheet_name}__{name}__{period_rid}"
                cells[synth_key] = val
                return synth_key
            return None

        # Find which analytic this indicator belongs to.
        # Key: prefer a record that shares the same parent as the current context record
        # (e.g. within the same product group)
        target_rid = None
        target_analytic_id = None

        for aid, nmap in name_to_rids.items():
            if aid == period_analytic_id:
                continue

            candidates = nmap.get(name, [])

            # If no exact match, try prefix match: "количество партнеров" matches
            # "количество партнеров (рассрочка)" — but only within the same parent group
            if not candidates:
                current_rid = context.get(aid)
                current_parent = record_by_id.get(current_rid, {}).get("parent_id") if current_rid else None
                for rname, rids in nmap.items():
                    if rname.startswith(name) and rname != name:
                        # Check if any candidate shares the current parent
                        for crid in rids:
                            crec = record_by_id.get(crid)
                            if crec and crec.get("parent_id") == current_parent:
                                candidates = [crid]
                                break
                    if candidates:
                        break

            if not candidates:
                continue

            if len(candidates) == 1:
                target_rid = candidates[0]
                target_analytic_id = aid
                break

            # Multiple matches — pick the one with the same parent as current context
            current_rid = context.get(aid)
            if current_rid:
                current_rec = record_by_id.get(current_rid)
                current_parent = current_rec["parent_id"] if current_rec else None
                for crid in candidates:
                    crec = record_by_id.get(crid)
                    if crec and crec["parent_id"] == current_parent:
                        target_rid = crid
                        target_analytic_id = aid
                        break

            if not target_rid:
                target_rid = candidates[0]
                target_analytic_id = aid
            break

        if target_rid is None:
            return None

        # Build coord_key parts
        parts = {}
        for aid in ordered_analytic_ids:
            if aid == target_analytic_id:
                parts[aid] = target_rid
            elif aid in context:
                parts[aid] = context[aid]

        # Apply params overrides
        for param_name, param_value in params.items():
            # Find analytic by name
            param_aid = analytic_name_to_id.get(param_name)
            if not param_aid:
                # Try fuzzy match
                for aname, aid in analytic_name_to_id.items():
                    if param_name.lower() in aname.lower():
                        param_aid = aid
                        break
            if not param_aid:
                continue

            if param_value == "предыдущий" and param_aid == period_analytic_id:
                current_period = parts.get(param_aid)
                if current_period and current_period in prev_period:
                    parts[param_aid] = prev_period[current_period]
                else:
                    return None  # no previous period
            else:
                # Look up by record name
                nmap = name_to_rid.get(param_aid, {})
                if param_value in nmap:
                    parts[param_aid] = nmap[param_value]
                else:
                    return None

        # Build coord_key
        coord_parts = [parts.get(aid, "") for aid in ordered_analytic_ids]
        if any(p == "" for p in coord_parts):
            return None
        return "|".join(coord_parts)

    # ── Evaluate all formula cells ──
    ctx = FormulaContext(cells, resolve_ref)
    computed = {}

    # Iterative evaluation — multiple passes until stable
    for pass_num in range(10):
        changed = False
        for coord_key, formula in formula_cells.items():
            parts = coord_key.split("|")
            context = {}
            for i, aid in enumerate(ordered_analytic_ids):
                if i < len(parts):
                    context[aid] = parts[i]
            ctx.current_context = context

            old_val = cells.get(coord_key, "")
            new_val = evaluate(formula, ctx)
            new_str = str(round(new_val, 6)) if new_val != 0 else "0"

            # First pass: always store (initial values may be stale)
            if pass_num == 0 or new_str != old_val:
                cells[coord_key] = new_str
                computed[coord_key] = new_str
                if new_str != old_val:
                    changed = True

        if pass_num > 0 and not changed:
            break

    return computed
