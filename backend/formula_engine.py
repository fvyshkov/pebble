"""Pebble formula engine — lazy pull-based evaluation with memoization.

Each cell is a lazy function. Requesting a cell's value recursively
evaluates its dependencies until it reaches manual inputs. Results
are cached per calculation run.

Formula syntax:
  [indicator_name]                              — same context
  [indicator_name](периоды="предыдущий")        — previous period
  [indicator_name](периоды="Январь 2026")       — specific period
  [Sheet.indicator_name]                        — cross-sheet reference
  SUM([a], [b], [c])                            — sum function
  Standard math: +, -, *, /, parentheses, numbers
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

ABBREVS = {"ср", "тыс", "мін", "мин", "макс", "кол", "шт", "руб", "дол", "коэф", "кред"}


def parse_ref(token: str) -> dict:
    m = REF_RE.match(token)
    if not m:
        return {"name": token, "params": {}}
    name = m.group(1)
    params_str = m.group(2) or ""
    params = {}
    for pm in PARAM_RE.finditer(params_str):
        params[pm.group(1)] = pm.group(2)

    sheet = None
    if "." in name:
        for i in range(len(name) - 1, -1, -1):
            if name[i] == ".":
                left = name[:i].strip()
                right = name[i + 1:].strip()
                if left.lower() in ABBREVS or len(left) <= 2:
                    continue
                if right:
                    sheet = left; name = right; break

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


# ── Evaluator (stateless — uses callback for value resolution) ─────────────

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
        if t and t[0] == "OP" and t[1] == op:
            advance(); return True
        return False

    def parse_expr():
        left = parse_term()
        while True:
            t = peek()
            if t and t[0] == "OP" and t[1] in ("+", "-"):
                op = advance()[1]; right = parse_term()
                left = left + right if op == "+" else left - right
            else:
                break
        return left

    def parse_term():
        left = parse_unary()
        while True:
            t = peek()
            if t and t[0] == "OP" and t[1] in ("*", "/"):
                op = advance()[1]; right = parse_unary()
                left = left * right if op == "*" else (left / right if right != 0 else 0.0)
            else:
                break
        return left

    def parse_unary():
        t = peek()
        if t and t[0] == "OP" and t[1] == "-":
            advance(); return -parse_primary()
        return parse_primary()

    def parse_primary():
        t = peek()
        if t is None:
            return 0.0
        if t[0] == "NUM":
            advance(); return t[1]
        if t[0] == "REF":
            advance(); return get_ref_value(t[1])
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


# ── Lazy calculator ────────────────────────────────────────────────────────

class SheetCalculator:
    """Lazy pull-based calculator for a sheet.

    Each cell is a lazy function: requesting its value triggers recursive
    evaluation of dependencies. Results are memoized per run.
    """

    def __init__(self, cells: dict, formula_cells: dict,
                 resolve_ref, period_order: list, prev_period: dict):
        """
        cells: {coord_key: value_str} — all cell data (manual + formula)
        formula_cells: {coord_key: formula_str} — only formula cells
        resolve_ref: (ref_dict, context_dict) -> coord_key | None
        """
        self.cells = dict(cells)  # working copy
        self.formula_cells = formula_cells
        self.resolve_ref = resolve_ref
        self.period_order = period_order
        self.prev_period = prev_period
        self._computed = set()    # already computed this run
        self._computing = set()   # currently in stack (cycle detection)

    def get_cell_value(self, coord_key: str, context: dict) -> float:
        """Get cell value, computing lazily if it's a formula."""
        # Already computed this run — return cached
        if coord_key in self._computed:
            return self._to_float(self.cells.get(coord_key, ""))

        # Cycle detection
        if coord_key in self._computing:
            return self._to_float(self.cells.get(coord_key, ""))

        formula = self.formula_cells.get(coord_key)
        if not formula:
            # Manual cell — just return value
            return self._to_float(self.cells.get(coord_key, ""))

        # Mark as computing
        self._computing.add(coord_key)

        # Evaluate — this will recursively pull dependencies
        def get_ref_value(ref_token: str):
            ref = parse_ref(ref_token)
            resolved = self.resolve_ref(ref, context)
            if resolved is None:
                return 0.0
            if resolved == coord_key:
                return 0.0  # self-reference guard
            # Build context for the resolved cell
            ref_context = self._context_from_key(resolved)
            return self.get_cell_value(resolved, ref_context)

        result = evaluate(formula, get_ref_value)
        result_str = str(round(result, 6)) if result != 0 else "0"

        # Cache
        self.cells[coord_key] = result_str
        self._computed.add(coord_key)
        self._computing.discard(coord_key)

        return result

    def _to_float(self, val: str) -> float:
        try:
            return float(val)
        except (ValueError, TypeError):
            return 0.0

    def _context_from_key(self, coord_key: str) -> dict:
        """Extract context (analytic_id → record_id) from coord_key."""
        parts = coord_key.split("|")
        return {aid: parts[i] for i, aid in enumerate(self._ordered_aids) if i < len(parts)}

    def calculate_all(self, ordered_analytic_ids: list) -> dict:
        """Calculate all formula cells. Returns {coord_key: new_value}."""
        self._ordered_aids = ordered_analytic_ids
        original = {k: self.cells.get(k, "") for k in self.formula_cells}

        for coord_key in self.formula_cells:
            context = self._context_from_key(coord_key)
            self.get_cell_value(coord_key, context)

        # Return only changed
        return {k: self.cells[k] for k in self.formula_cells
                if self.cells.get(k, "") != original.get(k, "")}


# ── Model-level calculator ──────────────────────────────────────────────────

async def calculate_model(db, model_id: str) -> dict[str, dict[str, str]]:
    """Calculate ALL formula cells across ALL sheets in a model.

    Returns {sheet_id: {coord_key: new_value}}.
    Uses lazy pull evaluation — any cell can pull dependencies from any sheet.
    """
    all_sheets = await db.execute_fetchall(
        "SELECT id, name FROM sheets WHERE model_id = ? ORDER BY created_at", (model_id,))
    if not all_sheets:
        return {}

    # ── Load entire model data ──
    # Global cells: {(sheet_id, coord_key): value}
    global_cells = {}
    # Formula cells: {(sheet_id, coord_key): formula}
    global_formulas = {}
    # Per-sheet metadata
    sheet_meta = {}  # sheet_id → {ordered_aids, name_to_rids, record_by_id, period_aid, prev_period, analytic_name_to_id}
    # Sheet name → sheet_id
    sheet_name_to_id = {}

    period_aid_global = None
    period_order = []
    prev_period = {}

    for s in all_sheets:
        sid = s["id"]
        sname = s["name"]
        sheet_name_to_id[sname] = sid
        # Aliases
        if "параметр" in sname.lower():
            for alias in ["Параметры", "Настройки", "BaaS - Настройки"]:
                sheet_name_to_id[alias] = sid

        bindings = await db.execute_fetchall(
            "SELECT sa.analytic_id, sa.sort_order, a.name as analytic_name, a.is_periods "
            "FROM sheet_analytics sa JOIN analytics a ON a.id = sa.analytic_id "
            "WHERE sa.sheet_id = ? ORDER BY sa.sort_order", (sid,))

        ordered_aids = [b["analytic_id"] for b in bindings]
        analytic_name_to_id = {b["analytic_name"]: b["analytic_id"] for b in bindings}
        record_by_id = {}
        name_to_rids = {}
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
                nmap.setdefault(name, []).append(r["id"])
                # Add compound names for cross-sheet lookup
                if r["parent_id"] and r["parent_id"] in record_by_id:
                    pdata = record_by_id[r["parent_id"]].get("_data", {})
                    pname = pdata.get("name", "")
                    if pname:
                        nmap.setdefault(f"{name} ({pname})", []).append(r["id"])
                        nmap.setdefault(f"{name} {pname}", []).append(r["id"])
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

        # Load cells
        for c in await db.execute_fetchall(
                "SELECT coord_key, value, rule, formula FROM cell_data WHERE sheet_id = ?", (sid,)):
            gk = (sid, c["coord_key"])
            global_cells[gk] = c["value"] or ""
            if c["rule"] == "formula" and c["formula"]:
                global_formulas[gk] = c["formula"]

    # ── Lazy evaluator ──
    computed_set = set()
    computing_set = set()

    def get_cell(sheet_id: str, coord_key: str) -> float:
        gk = (sheet_id, coord_key)

        if gk in computed_set:
            return _to_float(global_cells.get(gk, ""))

        if gk in computing_set:
            return _to_float(global_cells.get(gk, ""))  # cycle — use current

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
                # Cross-sheet reference
                target_sid = _find_sheet(ref_sheet)
                if not target_sid:
                    return 0.0
                target_meta = sheet_meta[target_sid]
                ind_rid = _find_indicator(ref["name"], target_meta)
                if not ind_rid:
                    return 0.0
                period_rid = context.get(meta["period_aid"])
                if not period_rid:
                    return 0.0
                target_aids = target_meta["ordered_aids"]
                # Build target coord_key: period_rid | ind_rid
                target_parts = []
                for aid in target_aids:
                    if aid == target_meta["period_aid"]:
                        target_parts.append(period_rid)
                    elif any(ind_rid in rids for rids in target_meta["name_to_rids"].get(aid, {}).values()):
                        target_parts.append(ind_rid)
                    else:
                        target_parts.append("")
                if any(p == "" for p in target_parts):
                    # Simpler: just period|indicator
                    target_ck = f"{period_rid}|{ind_rid}"
                else:
                    target_ck = "|".join(target_parts)
                return get_cell(target_sid, target_ck)
            else:
                # Local reference
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

    def _find_sheet(name):
        sid = sheet_name_to_id.get(name)
        if sid: return sid
        nl = name.lower()
        for sn, si in sheet_name_to_id.items():
            sl = sn.lower()
            if nl in sl or sl in nl:
                return si
        # Word overlap
        nw = set(re.sub(r'[^а-яa-z0-9\s]', '', nl).split())
        best = 0; best_sid = None
        for sn, si in sheet_name_to_id.items():
            sw = set(re.sub(r'[^а-яa-z0-9\s]', '', sn.lower()).split())
            o = len(nw & sw)
            if o > best and o >= 1: best = o; best_sid = si
        return best_sid

    def _find_indicator(name, meta):
        """Find indicator record_id by name in a sheet's metadata."""
        for aid, nmap in meta["name_to_rids"].items():
            if aid == meta["period_aid"]: continue
            rids = nmap.get(name)
            if rids: return rids[0]
        # Fuzzy: substring
        nl = name.lower()
        for aid, nmap in meta["name_to_rids"].items():
            if aid == meta["period_aid"]: continue
            for iname, rids in nmap.items():
                il = iname.lower()
                if nl in il or il in nl:
                    return rids[0]
        # Word overlap with stemming
        def norm(s):
            words = set(re.sub(r'[()]', ' ', s.lower()).split())
            stemmed = set()
            for w in words:
                stemmed.add(w)
                if len(w) > 4: stemmed.add(w[:len(w)-2])
            return stemmed
        nw = norm(name)
        best = 0; best_rid = None
        for aid, nmap in meta["name_to_rids"].items():
            if aid == meta["period_aid"]: continue
            for iname, rids in nmap.items():
                iw = norm(iname)
                o = len(nw & iw)
                if o > best and o >= max(2, len(nw) * 0.4):
                    best = o; best_rid = rids[0]
        return best_rid

    def _resolve_local(ref, context, meta):
        """Resolve a local [indicator] reference within a sheet."""
        name = ref["name"]
        params = ref.get("params", {})
        ordered_aids = meta["ordered_aids"]
        period_aid = meta["period_aid"]
        name_to_rids = meta["name_to_rids"]
        record_by_id = meta["record_by_id"]
        analytic_name_to_id = meta["analytic_name_to_id"]

        target_rid = None; target_aid = None
        for aid, nmap in name_to_rids.items():
            if aid == period_aid: continue
            candidates = nmap.get(name, [])
            if not candidates:
                cur_rid = context.get(aid)
                cur_parent = record_by_id.get(cur_rid, {}).get("parent_id") if cur_rid else None
                for rname, rids in nmap.items():
                    if rname.startswith(name) and rname != name:
                        for crid in rids:
                            crec = record_by_id.get(crid)
                            if crec and crec.get("parent_id") == cur_parent:
                                candidates = [crid]; break
                    if candidates: break
            if not candidates: continue
            if len(candidates) == 1:
                target_rid = candidates[0]; target_aid = aid; break
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

        if target_rid is None: return None

        parts = {}
        for aid in ordered_aids:
            if aid == target_aid: parts[aid] = target_rid
            elif aid in context: parts[aid] = context[aid]

        for pname, pval in params.items():
            paid = analytic_name_to_id.get(pname)
            if not paid:
                for aname, aid in analytic_name_to_id.items():
                    if pname.lower() in aname.lower(): paid = aid; break
            if not paid: continue
            if pval == "предыдущий" and paid == period_aid:
                cur = parts.get(paid)
                if cur and cur in prev_period: parts[paid] = prev_period[cur]
                else: return None
            else:
                nmap = name_to_rids.get(paid, {})
                rids = nmap.get(pval, [])
                if rids: parts[paid] = rids[0]
                else: return None

        coord_parts = [parts.get(aid, "") for aid in ordered_aids]
        if any(p == "" for p in coord_parts): return None
        return "|".join(coord_parts)

    # ── Run: evaluate all formula cells ──
    for gk in global_formulas:
        sheet_id, coord_key = gk
        meta = sheet_meta[sheet_id]
        context = _context_from_key(coord_key, meta["ordered_aids"])
        get_cell(sheet_id, coord_key)

    # Return all computed formula values grouped by sheet
    result = {}
    for (sid, ck) in global_formulas:
        new_val = global_cells.get((sid, ck), "")
        if sid not in result: result[sid] = {}
        result[sid][ck] = new_val

    return result


# ── Sheet calculator (convenience wrapper) ─────────────────────────────────

async def calculate_sheet(db, sheet_id: str) -> dict[str, str]:
    """Calculate all formula cells in a sheet using lazy evaluation."""

    # Load bindings
    bindings = await db.execute_fetchall(
        "SELECT sa.analytic_id, sa.sort_order, a.name as analytic_name, a.is_periods "
        "FROM sheet_analytics sa JOIN analytics a ON a.id = sa.analytic_id "
        "WHERE sa.sheet_id = ? ORDER BY sa.sort_order",
        (sheet_id,),
    )
    if not bindings:
        return {}

    # Load records per analytic
    record_by_id = {}
    name_to_rids = {}  # {analytic_id: {name: [record_id, ...]}}
    period_analytic_id = None
    period_order = []

    for b in bindings:
        aid = b["analytic_id"]
        recs = [dict(r) for r in await db.execute_fetchall(
            "SELECT * FROM analytic_records WHERE analytic_id = ? ORDER BY sort_order", (aid,))]
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
            parent_ids = {r["parent_id"] for r in recs if r["parent_id"]}
            period_order = [r["id"] for r in recs if r["id"] not in parent_ids]

    analytic_name_to_id = {b["analytic_name"]: b["analytic_id"] for b in bindings}
    ordered_analytic_ids = [b["analytic_id"] for b in bindings]

    # Load cells
    cells_raw = await db.execute_fetchall(
        "SELECT coord_key, value, rule, formula FROM cell_data WHERE sheet_id = ?", (sheet_id,))
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

    # ── Cross-sheet data ──
    sheet_row = await db.execute_fetchall("SELECT model_id FROM sheets WHERE id = ?", (sheet_id,))
    model_id = sheet_row[0]["model_id"] if sheet_row else None

    xsheet = {}
    if model_id:
        other_sheets = await db.execute_fetchall(
            "SELECT id, name FROM sheets WHERE model_id = ? AND id != ?", (model_id, sheet_id))
        for os in other_sheets:
            os_bindings = await db.execute_fetchall(
                "SELECT sa.analytic_id, a.is_periods FROM sheet_analytics sa "
                "JOIN analytics a ON a.id = sa.analytic_id WHERE sa.sheet_id = ? ORDER BY sa.sort_order",
                (os["id"],))
            os_period_aid = os_indicator_aid = None
            for ob in os_bindings:
                if ob["is_periods"]:
                    os_period_aid = ob["analytic_id"]
                else:
                    os_indicator_aid = ob["analytic_id"]
            if not os_period_aid or not os_indicator_aid:
                continue
            os_recs = [dict(r) for r in await db.execute_fetchall(
                "SELECT id, parent_id, data_json FROM analytic_records WHERE analytic_id = ? ORDER BY sort_order",
                (os_indicator_aid,))]
            os_name_to_rid = {}
            os_by_id = {}
            for r in os_recs:
                d = json.loads(r["data_json"]) if isinstance(r["data_json"], str) else r["data_json"]
                r["_name"] = d.get("name", "")
                os_by_id[r["id"]] = r
                os_name_to_rid[r["_name"]] = r["id"]
            for r in os_recs:
                if r["parent_id"] and r["parent_id"] in os_by_id:
                    pname = os_by_id[r["parent_id"]]["_name"]
                    cname = r["_name"]
                    os_name_to_rid[f"{cname} ({pname})"] = r["id"]
                    os_name_to_rid[f"{cname} {pname}"] = r["id"]
                    os_name_to_rid[f"{cname} ({pname.lower()})"] = r["id"]

            os_cells = {(c["coord_key"].split("|")[0], c["coord_key"].split("|")[1]): c["value"] or ""
                        for c in await db.execute_fetchall(
                            "SELECT coord_key, value FROM cell_data WHERE sheet_id = ?", (os["id"],))
                        if "|" in c["coord_key"]}

            entry = {"name_to_rid": os_name_to_rid, "cells": os_cells}
            xsheet[os["name"]] = entry
            # Aliases
            name_lower = os["name"].lower()
            if "параметр" in name_lower:
                for alias in ["Параметры", "Настройки", "BaaS - Настройки"]:
                    xsheet[alias] = entry

    # ── Reference resolver ──
    def resolve_ref(ref: dict, context: dict) -> str | None:
        name = ref["name"]
        sheet_name = ref.get("sheet")
        params = ref.get("params", {})

        # Cross-sheet
        if sheet_name:
            xs = xsheet.get(sheet_name)
            if not xs:
                sn_lower = sheet_name.lower()
                for sn, sd in xsheet.items():
                    sl = sn.lower()
                    if sn_lower in sl or sl in sn_lower:
                        xs = sd; break
            if not xs:
                sn_words = set(re.sub(r'[^а-яa-z0-9\s]', '', sheet_name.lower()).split())
                best = 0
                for sn, sd in xsheet.items():
                    sw = set(re.sub(r'[^а-яa-z0-9\s]', '', sn.lower()).split())
                    o = len(sn_words & sw)
                    if o > best and o >= 1:
                        best = o; xs = sd
            if not xs:
                return None

            ind_rid = xs["name_to_rid"].get(name)
            if not ind_rid:
                nl = name.lower()
                for iname, irid in xs["name_to_rid"].items():
                    il = iname.lower()
                    if nl in il or il in nl:
                        ind_rid = irid; break
            if not ind_rid:
                def norm(s):
                    words = set(re.sub(r'[()]', ' ', s.lower()).split())
                    stemmed = set()
                    for w in words:
                        stemmed.add(w)
                        if len(w) > 4: stemmed.add(w[:len(w)-2])
                    return stemmed
                nw = norm(name)
                best = 0
                for iname, irid in xs["name_to_rid"].items():
                    iw = norm(iname)
                    o = len(nw & iw)
                    if o > best and o >= max(2, len(nw) * 0.4):
                        best = o; ind_rid = irid
            if not ind_rid:
                return None

            period_rid = context.get(period_analytic_id)
            if not period_rid:
                return None
            val = xs["cells"].get((period_rid, ind_rid))
            if val is not None:
                synth = f"__xs__{sheet_name}__{name}__{period_rid}"
                cells[synth] = val
                return synth
            return None

        # Local reference
        target_rid = None
        target_aid = None
        for aid, nmap in name_to_rids.items():
            if aid == period_analytic_id:
                continue
            candidates = nmap.get(name, [])
            if not candidates:
                current_rid = context.get(aid)
                current_parent = record_by_id.get(current_rid, {}).get("parent_id") if current_rid else None
                for rname, rids in nmap.items():
                    if rname.startswith(name) and rname != name:
                        for crid in rids:
                            crec = record_by_id.get(crid)
                            if crec and crec.get("parent_id") == current_parent:
                                candidates = [crid]; break
                    if candidates: break
            if not candidates:
                continue
            if len(candidates) == 1:
                target_rid = candidates[0]; target_aid = aid; break
            current_rid = context.get(aid)
            if current_rid:
                current_parent = record_by_id.get(current_rid, {}).get("parent_id")
                for crid in candidates:
                    crec = record_by_id.get(crid)
                    if crec and crec.get("parent_id") == current_parent:
                        target_rid = crid; target_aid = aid; break
            if not target_rid:
                target_rid = candidates[0]; target_aid = aid
            break

        if target_rid is None:
            return None

        parts = {}
        for aid in ordered_analytic_ids:
            if aid == target_aid:
                parts[aid] = target_rid
            elif aid in context:
                parts[aid] = context[aid]

        for param_name, param_value in params.items():
            param_aid = analytic_name_to_id.get(param_name)
            if not param_aid:
                for aname, aid in analytic_name_to_id.items():
                    if param_name.lower() in aname.lower():
                        param_aid = aid; break
            if not param_aid:
                continue
            if param_value == "предыдущий" and param_aid == period_analytic_id:
                cur = parts.get(param_aid)
                if cur and cur in prev_period:
                    parts[param_aid] = prev_period[cur]
                else:
                    return None
            else:
                nmap = name_to_rids.get(param_aid, {})
                rids_list = nmap.get(param_value, [])
                if rids_list:
                    parts[param_aid] = rids_list[0]
                else:
                    return None

        coord_parts = [parts.get(aid, "") for aid in ordered_analytic_ids]
        if any(p == "" for p in coord_parts):
            return None
        result_key = "|".join(coord_parts)

        current_key = "|".join(context.get(aid, "") for aid in ordered_analytic_ids)
        if result_key == current_key:
            return None

        return result_key

    # ── Run lazy calculation ──
    calc = SheetCalculator(cells, formula_cells, resolve_ref, period_order, prev_period)
    return calc.calculate_all(ordered_analytic_ids)
