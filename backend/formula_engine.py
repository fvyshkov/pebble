"""Pebble formula engine — lazy pull-based evaluation with memoization.

Each cell is a lazy function. Requesting a cell's value recursively
evaluates its dependencies until it reaches manual inputs. Results
are cached per calculation run.

Formula syntax:
  [indicator_name]                                    — same context (exact match)
  [indicator_name](периоды=предыдущий)                — previous period (unquoted)
  [indicator_name](периоды="предыдущий")              — previous period (quoted, legacy)
  [indicator_name](периоды=период.назад(2))           — two periods back
  [indicator_name](периоды=Январь, подразделения=Москва) — explicit axis values
  [Sheet::indicator_name]                             — cross-sheet reference (:: separator)
  SUM([a], [b], [c])                                  — sum function
  Standard math: +, -, *, /, parentheses, numbers

Parameters support both quoted ("value") and unquoted (value) syntax.
Multiple params are comma-separated.
All references resolve by EXACT name match (case-insensitive). No fuzzy matching.
"""

import re
import json
import math
import itertools
from typing import Any


# ── Tokenizer ──────────────────────────────────────────────────────────────

TOKEN_RE = re.compile(r"""
    (\[(?:[^\[\]]+)\](?:\((?:[^()]*|\([^()]*\))*\))?)  |  # [ref](params) — one nesting level
    (SUM|AVERAGE|IF|MIN|MAX|ABS)\s*\(   |  # functions
    (\d+(?:\.\d+)?)                    |  # number
    ([+\-*/(),<>=!])                   |  # operators, parens, comparison
    (\s+)                                 # whitespace (skip)
""", re.VERBOSE)

REF_RE = re.compile(r"""
    \[([^\]]+)\]                          # indicator name
    (?:\(((?:[^()]*|\([^()]*\))*)\))?     # optional params — one nesting level
""", re.VERBOSE)

# For matching key.назад(N) function calls in param values
_PERIOD_BACK_RE = re.compile(r'\w+\.назад\((\d+)\)')


def parse_ref(token: str) -> dict:
    """Parse a formula reference token like [name] or [name](key=value, ...).

    Supports:
    - Unquoted values:  [ind](периоды=Январь, подразделения=Москва)
    - Quoted values:    [ind](периоды="предыдущий")  (legacy)
    - Period back-ref:  [ind](периоды=период.назад(2))
    - Period identity:  [ind](период=период)  → no-op, ignored
    - Cross-sheet:      [Sheet::indicator]
    """
    m = REF_RE.match(token)
    if not m:
        return {"name": token, "params": {}}
    name = m.group(1)
    params_str = (m.group(2) or "").strip()
    params = {}

    if params_str:
        for raw_pair in params_str.split(","):
            raw_pair = raw_pair.strip()
            if "=" not in raw_pair:
                continue
            key, _, val = raw_pair.partition("=")
            key = key.strip()
            val = val.strip()
            if not key:
                continue

            # Period back-reference: word.назад(N) or key=key.назад(N)
            back_m = _PERIOD_BACK_RE.fullmatch(val)
            if back_m:
                params[key] = f"назад({back_m.group(1)})"
                continue

            # Identity: key=key (same value, no-op)
            if val.lower() == key.lower():
                continue

            # Strip surrounding quotes (legacy quoted syntax)
            if val.startswith('"') and val.endswith('"'):
                val = val[1:-1]

            # "предыдущий" shorthand (keep as-is for _resolve_local)
            params[key] = val

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
            tokens.append(("FUNC", m.group(2).upper()))
        elif m.group(3):
            tokens.append(("NUM", float(m.group(3))))
        elif m.group(4):
            tokens.append(("OP", m.group(4)))
        pos = m.end()
    return tokens


# ── Evaluator ──────────────────────────────────────────────────────────────

def evaluate(formula: str, get_ref_value) -> float:
    """Evaluate formula. get_ref_value(token_str) -> float | None.

    get_ref_value may return None to signal "cell does not exist".
    AVERAGE skips None args (like Excel). All other ops treat None as 0.
    """
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

    def _n(v):
        """Coerce None (missing cell) to 0.0 for arithmetic."""
        return 0.0 if v is None else v

    def parse_comparison():
        """Parse comparison: expr < expr, expr > expr, etc."""
        left = parse_additive()
        while True:
            t = peek()
            if t and t[0] == "OP" and t[1] in ("<", ">", "=", "!"):
                op1 = advance()[1]
                t2 = peek()
                if t2 and t2[0] == "OP" and t2[1] == "=":
                    op1 += advance()[1]
                right = _n(parse_additive())
                left = _n(left)
                if op1 == "<": left = 1.0 if left < right else 0.0
                elif op1 == ">": left = 1.0 if left > right else 0.0
                elif op1 == "<=": left = 1.0 if left <= right else 0.0
                elif op1 == ">=": left = 1.0 if left >= right else 0.0
                elif op1 == "=" or op1 == "==": left = 1.0 if abs(left - right) < 1e-12 else 0.0
                elif op1 == "!=": left = 1.0 if abs(left - right) >= 1e-12 else 0.0
            else:
                break
        return left

    def parse_expr():
        return parse_comparison()

    def parse_additive():
        left = parse_term()
        while True:
            t = peek()
            if t and t[0] == "OP" and t[1] in ("+", "-"):
                op = advance()[1]; right = _n(parse_term())
                left = _n(left) + right if op == "+" else _n(left) - right
            else: break
        return left  # preserves None if no operators

    def parse_term():
        left = parse_unary()
        while True:
            t = peek()
            if t and t[0] == "OP" and t[1] in ("*", "/"):
                op = advance()[1]; right = _n(parse_unary())
                left = _n(left) * right if op == "*" else (_n(left) / right if right != 0 else float('nan'))
            else: break
        return left  # preserves None if no operators

    def parse_unary():
        t = peek()
        if t and t[0] == "OP" and t[1] == "-":
            advance(); val = parse_unary(); return -_n(val)
        return parse_primary()

    def parse_primary():
        t = peek()
        if t is None: return 0.0
        if t[0] == "NUM": advance(); return t[1]
        if t[0] == "REF": advance(); return get_ref_value(t[1])  # may return None
        if t[0] == "FUNC":
            func_name = advance()[1]
            args = []
            while True:
                args.append(parse_expr())
                if not expect_op(","): break
            expect_op(")")
            if func_name == "SUM":
                return sum(_n(a) for a in args)
            elif func_name == "AVERAGE":
                # Excel AVERAGE skips empty/missing cells
                present = [_n(a) for a in args if a is not None]
                return sum(present) / len(present) if present else 0.0
            elif func_name == "IF":
                cond = _n(args[0]) if len(args) > 0 else 0.0
                true_val = _n(args[1]) if len(args) > 1 else 0.0
                false_val = _n(args[2]) if len(args) > 2 else 0.0
                return true_val if cond != 0.0 else false_val
            elif func_name == "MIN":
                return min(_n(a) for a in args) if args else 0.0
            elif func_name == "MAX":
                return max(_n(a) for a in args) if args else 0.0
            elif func_name == "ABS":
                return abs(_n(args[0])) if args else 0.0
            return sum(_n(a) for a in args)  # fallback
        if t[0] == "OP" and t[1] == "(":
            advance(); val = parse_expr(); expect_op(")"); return val
        advance(); return 0.0

    try:
        result = parse_expr()
        if result is None:
            return None  # propagate "missing" to caller
        return result if not math.isinf(result) else 0.0
    except Exception:
        return 0.0


# ── Model-level lazy calculator ────────────────────────────────────────────

async def calculate_model(db, model_id: str) -> dict[str, dict[str, str]]:
    """Calculate ALL formula cells across ALL sheets in a model.
    Returns {sheet_id: {coord_key: new_value}}.
    """
    all_sheets = await db.execute_fetchall(
        "SELECT id, name, excel_code FROM sheets WHERE model_id = ? ORDER BY created_at", (model_id,))
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
        # Register by excel_code (tab name) if available
        excel_code = s["excel_code"] if "excel_code" in s.keys() else None
        if excel_code:
            sheet_name_to_id[excel_code.lower()] = sid
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
            "SELECT sa.analytic_id, sa.sort_order, a.name as analytic_name, a.is_periods, sa.is_main "
            "FROM sheet_analytics sa JOIN analytics a ON a.id = sa.analytic_id "
            "WHERE sa.sheet_id = ? ORDER BY sa.sort_order", (sid,))

        ordered_aids = [b["analytic_id"] for b in bindings]
        analytic_name_to_id = {b["analytic_name"]: b["analytic_id"] for b in bindings}
        record_by_id = {}
        children_by_rid: dict[str, list[str]] = {}
        name_to_rids = {}  # {analytic_id: {name_lower: [record_ids]}}
        period_aid = None
        main_aid = next((b["analytic_id"] for b in bindings if b["is_main"]), None)

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
                # Build parent→children index (scoped to this sheet's records)
                pid = r.get("parent_id")
                if pid:
                    children_by_rid.setdefault(pid, []).append(r["id"])
            name_to_rids[aid] = nmap

            if b["is_periods"]:
                period_aid = aid
                if not period_aid_global:
                    period_aid_global = aid
                    parent_ids = {r["parent_id"] for r in recs if r["parent_id"]}
                    period_order = [r["id"] for r in recs if r["id"] not in parent_ids]
                    for i in range(1, len(period_order)):
                        prev_period[period_order[i]] = period_order[i - 1]

        # Indicator formula rules for this sheet, indexed by indicator_id.
        rules_by_indicator: dict[str, list[dict]] = {}
        for r in await db.execute_fetchall(
                "SELECT id, indicator_id, kind, scope_json, priority, formula "
                "FROM indicator_formula_rules WHERE sheet_id = ?", (sid,)):
            try:
                scope = json.loads(r["scope_json"]) if r["scope_json"] else {}
            except Exception:
                scope = {}
            rules_by_indicator.setdefault(r["indicator_id"], []).append({
                "id": r["id"],
                "kind": r["kind"],
                "scope": scope,
                "priority": r["priority"] or 0,
                "formula": r["formula"] or "",
            })

        sheet_meta[sid] = {
            "ordered_aids": ordered_aids,
            "name_to_rids": name_to_rids,
            "record_by_id": record_by_id,
            "children_by_rid": children_by_rid,
            "period_aid": period_aid,
            "main_aid": main_aid,
            "analytic_name_to_id": analytic_name_to_id,
            "rules_by_indicator": rules_by_indicator,
            "name": sname,
        }

        for c in await db.execute_fetchall(
                "SELECT coord_key, value, rule, formula FROM cell_data WHERE sheet_id = ?", (sid,)):
            gk = (sid, c["coord_key"])
            global_cells[gk] = c["value"] or ""
            if c["rule"] == "formula" and c["formula"]:
                # Skip raw Excel formulas (starting with =) — they can't be evaluated
                if not c["formula"].startswith("="):
                    global_formulas[gk] = c["formula"]

    # Track manual cells — they must not be overwritten by consolidation
    _manual_cells: set[tuple] = set()
    for gk in global_cells:
        if gk not in global_formulas:
            _manual_cells.add(gk)

    # Track original DB cell keys (before computation adds synthetic ones)
    _original_cell_keys = set(global_cells.keys())
    # Snapshot original values so we can detect what actually changed.
    _original_values: dict[tuple, str] = dict(global_cells)

    # ── Pre-filter: remove formulas with unresolvable cross-sheet refs ──
    # If a formula references [Sheet::indicator] and that indicator doesn't exist
    # on the target sheet, the formula would evaluate to 0 and destroy imported values.
    # Better to skip evaluation entirely and keep the imported value.
    _CROSS_SHEET_REF_RE = re.compile(r'\[([^\]]*::([^\]]+))\]')
    _skipped_formulas: set[tuple] = set()
    for gk, formula in list(global_formulas.items()):
        refs = _CROSS_SHEET_REF_RE.findall(formula)
        if not refs:
            continue
        has_bad_ref = False
        for full_ref_name, _ in refs:
            parsed = parse_ref(f"[{full_ref_name}]")
            ref_sheet = parsed.get("sheet")
            if not ref_sheet:
                continue
            target_sid = sheet_name_to_id.get(ref_sheet.lower())
            if not target_sid:
                has_bad_ref = True
                break
            target_meta = sheet_meta.get(target_sid)
            if not target_meta:
                has_bad_ref = True
                break
            name_lower = parsed["name"].lower()
            found = False
            for aid, nmap in target_meta["name_to_rids"].items():
                if aid == target_meta["period_aid"]:
                    continue
                if nmap.get(name_lower):
                    found = True
                    break
            if not found:
                has_bad_ref = True
                break
        if has_bad_ref:
            _skipped_formulas.add(gk)
    for gk in _skipped_formulas:
        del global_formulas[gk]
    if _skipped_formulas:
        print(f"[formula_engine] Skipped {len(_skipped_formulas)} formulas with unresolvable cross-sheet refs")
    # ── Lazy evaluator ──
    # Pre-seed computed_set with skipped formulas so get_cell() returns their stored value
    # without trying indicator rules or consolidation logic.
    computed_set = set(_skipped_formulas)
    computing_set = set()
    # Track which branch produced a cell's value — for resolve-formulas API.
    _computed_sources: dict[tuple, str] = {}  # gk → 'cell' | 'rule:<id>' | 'default-sum'
    _computed_formulas: dict[tuple, str] = {}  # gk → formula text used

    def _is_consolidating(context: dict, meta: dict) -> bool:
        """True if ANY axis (including main) coord points to a record with children."""
        children = meta.get("children_by_rid", {})
        for aid, rid in context.items():
            if children.get(rid):
                return True
        return False

    def _expand_children_one_level(coord_key: str, context: dict, meta: dict) -> list[str]:
        """Expand children along exactly ONE consolidating axis.

        When multiple axes have children (e.g. parent period + parent indicator),
        we must NOT take the Cartesian product — each axis is consolidated
        independently by the outer consolidation loops.  Priority:
        1. Period axis (quarter/year → months)
        2. Main (indicator) axis (section total → leaf indicators)
        3. Any other axis
        """
        children = meta.get("children_by_rid", {})
        ordered_aids = meta["ordered_aids"]
        period_aid = meta.get("period_aid")
        main_aid = meta.get("main_aid")

        # Find all axes that have children
        axes = []
        for aid in ordered_aids:
            rid = context.get(aid)
            if rid and children.get(rid):
                axes.append((aid, children[rid]))
        if not axes:
            return []

        # Pick ONE axis to expand — period first, then main, then other
        chosen = None
        for aid, ch in axes:
            if aid == period_aid:
                chosen = (aid, ch)
                break
        if not chosen:
            for aid, ch in axes:
                if aid == main_aid:
                    chosen = (aid, ch)
                    break
        if not chosen:
            chosen = axes[0]

        expand_aid, expand_children = chosen
        combos = []
        for crid in expand_children:
            parts = []
            for aid in ordered_aids:
                if aid == expand_aid:
                    parts.append(crid)
                else:
                    parts.append(context.get(aid, ""))
            combos.append("|".join(parts))
        return combos

    def _resolve_indicator_formula(sheet_id: str, context: dict, meta: dict):
        """Return (formula_text, source_label) or None."""
        main = meta.get("main_aid")
        if not main:
            return None
        indicator_rid = context.get(main)
        if not indicator_rid:
            return None
        rules = meta.get("rules_by_indicator", {}).get(indicator_rid, [])
        if not rules:
            return None

        # 3a. Scoped rules — pick the best match (subset of non-main coord).
        non_main = {a: r for a, r in context.items() if a != main}
        scoped_hits = []
        for rule in rules:
            if rule["kind"] != "scoped":
                continue
            scope = rule.get("scope") or {}
            if not scope:
                continue
            # scope value may be comma-separated (multi-select periods)
            if all(non_main.get(a) in (r or "").split(",") for a, r in scope.items() if r):
                scoped_hits.append(rule)
        if scoped_hits:
            best = sorted(
                scoped_hits,
                key=lambda r: (-(r.get("priority") or 0), -len(r.get("scope") or {}), r["id"]),
            )[0]
            if best.get("formula"):
                return best["formula"], f"rule:{best['id']}"

        # 3b. Base leaf/consolidation.
        is_consol = _is_consolidating(context, meta)
        base_kind = "consolidation" if is_consol else "leaf"
        for rule in rules:
            if rule["kind"] == base_kind and rule.get("formula"):
                return rule["formula"], f"rule:{rule['id']}"
        return None

    # Track cells with unresolvable references — don't overwrite their DB values.
    _unresolved: set[tuple] = set()

    def get_cell(sheet_id: str, coord_key: str) -> float:
        gk = (sheet_id, coord_key)
        if gk in computed_set:
            return _to_float(global_cells.get(gk, ""))
        if gk in computing_set:
            return _to_float(global_cells.get(gk, ""))  # cycle

        meta = sheet_meta.get(sheet_id)
        if not meta:
            return _to_float(global_cells.get(gk, ""))
        context = _context_from_key(coord_key, meta["ordered_aids"])

        # ── 1. Explicit per-cell formula (cell_data.rule='formula')
        formula = global_formulas.get(gk)
        formula_source = "cell" if formula else None

        # ── 2. Indicator rule (scoped → consolidation/leaf base)
        #    But NEVER override manual cells — user-entered values are sacrosanct.
        if not formula and gk not in _manual_cells:
            resolved = _resolve_indicator_formula(sheet_id, context, meta)
            if resolved:
                formula, formula_source = resolved

        if formula:
            # Special consolidation keywords: AVERAGE, LAST
            if formula == "AVERAGE" and _is_consolidating(context, meta):
                computing_set.add(gk)
                children = list(_expand_children_one_level(coord_key, context, meta))
                total = sum(get_cell(sheet_id, ck) for ck in children)
                result = total / len(children) if children else 0.0
                result_str = str(round(result, 6)) if result != 0 else "0"
                global_cells[gk] = result_str
                computed_set.add(gk)
                computing_set.discard(gk)
                _computed_sources[gk] = formula_source or "rule"
                _computed_formulas[gk] = formula
                return result

            if formula == "LAST" and _is_consolidating(context, meta):
                computing_set.add(gk)
                children = list(_expand_children_one_level(coord_key, context, meta))
                result = get_cell(sheet_id, children[-1]) if children else 0.0
                result_str = str(round(result, 6)) if result != 0 else "0"
                global_cells[gk] = result_str
                computed_set.add(gk)
                computing_set.discard(gk)
                _computed_sources[gk] = formula_source or "rule"
                _computed_formulas[gk] = formula
                return result

            computing_set.add(gk)
            _has_unresolved_ref = False

            def get_ref_value(ref_token: str) -> float:
                nonlocal _has_unresolved_ref
                ref = parse_ref(ref_token)
                ref_sheet = ref.get("sheet")
                if ref_sheet:
                    val = _resolve_cross_sheet(ref, context, meta, sheet_id)
                    if val == 0.0:
                        # Check if the reference actually resolved to a real cell
                        target_sid = sheet_name_to_id.get(ref["sheet"].lower())
                        if not target_sid:
                            _has_unresolved_ref = True
                        else:
                            # Check if indicator was found
                            target_meta = sheet_meta.get(target_sid)
                            if target_meta:
                                name_lower = ref["name"].lower()
                                found = False
                                for aid, nmap in target_meta["name_to_rids"].items():
                                    if aid == target_meta["period_aid"]:
                                        continue
                                    if nmap.get(name_lower):
                                        found = True
                                        break
                                if not found:
                                    _has_unresolved_ref = True
                    return val
                rs = _resolve_local(ref, context, meta)
                if rs is None:
                    return None  # ref couldn't resolve (e.g. prev period doesn't exist)
                if rs == coord_key:
                    return 0.0
                target_gk = (sheet_id, rs)
                # Check if cell exists (has value, formula, or is manual)
                if target_gk not in global_cells and target_gk not in global_formulas and target_gk not in _manual_cells:
                    return None  # cell doesn't exist → missing (AVERAGE will skip)
                val = get_cell(sheet_id, rs)
                # Propagate unresolved status from dependencies
                if target_gk in _unresolved:
                    _has_unresolved_ref = True
                return val

            result = evaluate(formula, get_ref_value)
            if result is None:
                result = 0.0
            # Division by zero in consolidation formula → fall through to SUM
            if not math.isfinite(result) and _is_consolidating(context, meta):
                computing_set.discard(gk)
                pass  # fall through to step 3 (default SUM)
            else:
                if not math.isfinite(result):
                    result = 0.0
                result_str = str(round(result, 6)) if result != 0 else "0"
                global_cells[gk] = result_str
                computed_set.add(gk)
                computing_set.discard(gk)
                _computed_sources[gk] = formula_source or "cell"
                _computed_formulas[gk] = formula
                if _has_unresolved_ref:
                    _unresolved.add(gk)
                return result

        # ── 3. Consolidating coord with no formula → default SUM over children
        # But skip consolidation for manual cells — they have user-entered values
        # that must not be overwritten (e.g. yearly-only sheets like Product/Funnel)
        # Also skip indicator-axis consolidation for cells that don't exist in the
        # original data — creating phantom parent indicator cells causes cascading
        # errors when formulas reference the parent.
        _skip_default_sum = False
        if _is_consolidating(context, meta) and gk not in _original_cell_keys:
            # Check if the consolidation axis is the main (indicator) axis
            main = meta.get("main_aid")
            indicator_rid = context.get(main) if main else None
            if indicator_rid and meta["children_by_rid"].get(indicator_rid):
                _skip_default_sum = True
        if _is_consolidating(context, meta) and gk not in _manual_cells and not _skip_default_sum:
            computing_set.add(gk)
            children_cks = list(_expand_children_one_level(coord_key, context, meta))
            total = 0.0
            for child_ck in children_cks:
                total += get_cell(sheet_id, child_ck)
            total_str = str(round(total, 6)) if total != 0 else "0"
            global_cells[gk] = total_str
            computed_set.add(gk)
            computing_set.discard(gk)
            _computed_sources[gk] = "default-sum"
            # Don't propagate unresolved to consolidation cells — SUM is valid
            # even if some children have unresolvable refs (they contribute 0).
            return total

        # ── 4. Leaf manual value (stored)
        return _to_float(global_cells.get(gk, ""))

    def _to_float(val):
        try: return float(val)
        except: return 0.0

    def _context_from_key(coord_key, ordered_aids):
        parts = coord_key.split("|")
        return {aid: parts[i] for i, aid in enumerate(ordered_aids) if i < len(parts)}

    def _resolve_cross_sheet(ref, context, src_meta, src_sheet_id):
        """Resolve [Sheet::indicator] cross-sheet reference.

        Supports period modifiers like (периоды="предыдущий") and
        (периоды=назад(N)).
        """
        target_sid = sheet_name_to_id.get(ref["sheet"].lower())
        if not target_sid:
            return 0.0
        target_meta = sheet_meta[target_sid]
        raw_name = ref["name"]
        name_lower = raw_name.lower()

        # Handle "name#rowN" format
        row_hint = None
        if "#row" in name_lower:
            nm, row_str = name_lower.rsplit("#row", 1)
            name_lower = nm
            try:
                row_hint = int(row_str)
            except ValueError:
                pass

        # Handle "parent/child" — only if exact name not found
        parent_hint = None
        found_exact = False
        for aid, nmap in target_meta["name_to_rids"].items():
            if aid == target_meta["period_aid"]: continue
            if nmap.get(name_lower):
                found_exact = True; break
        if not found_exact and "/" in raw_name:
            parts_split = raw_name.split("/", 1)
            parent_hint = parts_split[0].strip().lower()
            name_lower = parts_split[1].strip().lower()

        # Find indicator by exact name (case-insensitive)
        ind_rid = None
        for aid, nmap in target_meta["name_to_rids"].items():
            if aid == target_meta["period_aid"]: continue
            rids = nmap.get(name_lower)
            if rids:
                # Row hint — match by excel_row
                if row_hint is not None:
                    target_rbi = target_meta.get("record_by_id", {})
                    for crid in rids:
                        rec = target_rbi.get(crid, {})
                        if rec.get("excel_row") == row_hint:
                            ind_rid = crid; break
                    if ind_rid:
                        break

                # Parent hint — filter by parent name
                if parent_hint and len(rids) > 1:
                    target_rbi = target_meta.get("record_by_id", {})
                    for crid in rids:
                        prec = target_rbi.get(crid, {})
                        pid = prec.get("parent_id")
                        if pid:
                            parent_rec = target_rbi.get(pid, {})
                            pdata = parent_rec.get("_data") or {}
                            pname = pdata.get("name", "").lower()
                            if pname == parent_hint:
                                ind_rid = crid; break
                    if ind_rid:
                        break

                # If multiple, pick one with same parent context
                if len(rids) == 1:
                    ind_rid = rids[0]
                else:
                    # Disambiguate: use source indicator's name to find
                    # a matching section in the target sheet.
                    src_main = src_meta.get("main_aid")
                    src_ind_rid = context.get(src_main) if src_main else None
                    src_name = ""
                    if src_ind_rid:
                        src_rec = src_meta.get("record_by_id", {}).get(src_ind_rid, {})
                        src_data = src_rec.get("_data") or {}
                        src_name = src_data.get("name", "").lower()
                    # Try to match section prefix from source name
                    best_rid = rids[0]
                    if src_name:
                        target_rbi = target_meta.get("record_by_id", {})
                        for crid in rids:
                            # Walk up to section header in target
                            node = target_rbi.get(crid, {}).get("parent_id")
                            while node:
                                prec = target_rbi.get(node, {})
                                pdata = prec.get("_data") or {}
                                pname = pdata.get("name", "").lower()
                                if pname and pname in src_name:
                                    best_rid = crid
                                    break
                                node = prec.get("parent_id")
                            if best_rid != rids[0]:
                                break
                    ind_rid = best_rid
                break
        if not ind_rid:
            return 0.0

        period_rid = context.get(src_meta["period_aid"])
        if not period_rid:
            return 0.0

        # Apply period modifiers from params (e.g. периоды="предыдущий")
        params = ref.get("params", {})
        for param_name, param_value in params.items():
            # Check if this param refers to the period analytic
            param_aid = src_meta["analytic_name_to_id"].get(param_name)
            if not param_aid:
                for aname, aid in src_meta["analytic_name_to_id"].items():
                    if param_name.lower() in aname.lower():
                        param_aid = aid
                        break
            if param_aid and param_aid == src_meta["period_aid"]:
                if param_value == "предыдущий":
                    period_rid = prev_period.get(period_rid)
                    if not period_rid:
                        return 0.0
                elif param_value.startswith("назад("):
                    try:
                        back_n = int(param_value[6:-1])
                    except ValueError:
                        return 0.0
                    for _ in range(back_n):
                        period_rid = prev_period.get(period_rid)
                        if not period_rid:
                            return 0.0

        # Build coord key in target sheet's analytic order
        target_ordered = target_meta["ordered_aids"]
        target_period_aid = target_meta["period_aid"]
        ck_parts = []
        for aid in target_ordered:
            if aid == target_period_aid:
                ck_parts.append(period_rid)
            else:
                ck_parts.append(ind_rid)
        target_ck = "|".join(ck_parts)
        return get_cell(target_sid, target_ck)

    def _resolve_local(ref, context, meta):
        """Resolve [indicator] local reference. Exact match only."""
        raw_name = ref["name"]
        name_lower = raw_name.lower()
        params = ref.get("params", {})
        ordered_aids = meta["ordered_aids"]
        period_aid = meta["period_aid"]
        name_to_rids = meta["name_to_rids"]
        record_by_id = meta["record_by_id"]
        analytic_name_to_id = meta["analytic_name_to_id"]

        # Handle "name#rowN" format — match by excel_row
        row_hint = None
        if "#row" in name_lower:
            nm, row_str = name_lower.rsplit("#row", 1)
            name_lower = nm
            try:
                row_hint = int(row_str)
            except ValueError:
                pass

        # Handle "parent/child" disambiguation format.
        # Only if: first try with full name fails, AND name has exactly
        # one "/" without spaces around it (translator format: "Parent/Child")
        parent_hint = None

        target_rid = None
        target_aid = None

        # First try exact name match (handles names with "/" like "QR прием / оплата")
        found_exact = False
        for aid, nmap in name_to_rids.items():
            if aid == period_aid: continue
            candidates = nmap.get(name_lower, [])
            if candidates:
                found_exact = True
                break

        if not found_exact and "/" in raw_name:
            # Try parent/child split only when exact match fails
            parts_split = raw_name.split("/", 1)
            parent_hint = parts_split[0].strip().lower()
            name_lower = parts_split[1].strip().lower()

        for aid, nmap in name_to_rids.items():
            if aid == period_aid: continue
            candidates = nmap.get(name_lower, [])
            if not candidates: continue

            # If we have a row hint, filter by excel_row
            if row_hint is not None:
                for crid in candidates:
                    rec = record_by_id.get(crid, {})
                    if rec.get("excel_row") == row_hint:
                        target_rid = crid; target_aid = aid; break
                if target_rid:
                    break
                # Row hint didn't match, fall through to normal resolution

            # If we have a parent hint, filter candidates by parent name
            if parent_hint and len(candidates) > 1:
                filtered = []
                for crid in candidates:
                    prec = record_by_id.get(crid, {})
                    pid = prec.get("parent_id")
                    if pid:
                        parent_rec = record_by_id.get(pid, {})
                        pdata = parent_rec.get("_data") or {}
                        pname = pdata.get("name", "").lower()
                        if pname == parent_hint:
                            filtered.append(crid)
                if filtered:
                    candidates = filtered

            if len(candidates) == 1:
                target_rid = candidates[0]; target_aid = aid; break

            # Multiple records with same name — pick nearest by sort_order.
            # Direct parent match first; if no match, walk up ancestors
            # to find common section; finally fall back to closest sort_order.
            cur_rid = context.get(aid)
            if cur_rid:
                cur_rec = record_by_id.get(cur_rid, {})
                cur_parent = cur_rec.get("parent_id")
                # 1. Direct parent match
                for crid in candidates:
                    crec = record_by_id.get(crid)
                    if crec and crec.get("parent_id") == cur_parent:
                        target_rid = crid; target_aid = aid; break
                # 2. Common ancestor — pick candidate sharing deepest ancestor
                if not target_rid:
                    # Build ancestor list with depth (closer = higher depth)
                    cur_ancestors = {}  # rid → depth
                    node = cur_parent
                    depth = 0
                    while node:
                        cur_ancestors[node] = depth
                        node = record_by_id.get(node, {}).get("parent_id")
                        depth += 1
                    best_crid = None
                    best_depth = float('inf')
                    for crid in candidates:
                        node = record_by_id.get(crid, {}).get("parent_id")
                        while node:
                            if node in cur_ancestors:
                                d = cur_ancestors[node]
                                if d < best_depth:
                                    best_depth = d
                                    best_crid = crid
                                break
                            node = record_by_id.get(node, {}).get("parent_id")
                    if best_crid and best_depth < 10:  # reasonable depth
                        target_rid = best_crid; target_aid = aid
                # 3. Closest by sort_order
                if not target_rid:
                    cur_sort = cur_rec.get("sort_order", 0)
                    best = min(candidates,
                               key=lambda c: abs(record_by_id.get(c, {}).get("sort_order", 0) - cur_sort))
                    target_rid = best; target_aid = aid
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

            # Handle period back-references
            is_period_back = False
            back_n = 0
            if param_aid == period_aid:
                if param_value == "предыдущий":
                    is_period_back = True
                    back_n = 1
                elif param_value.startswith("назад("):
                    is_period_back = True
                    try:
                        back_n = int(param_value[6:-1])
                    except ValueError:
                        return None

            if is_period_back:
                cur = parts.get(param_aid)
                for _ in range(back_n):
                    if cur and cur in prev_period:
                        cur = prev_period[cur]
                    else:
                        return None
                parts[param_aid] = cur
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

    # ── Also evaluate cells where an indicator rule applies (no explicit formula).
    #    These are existing cell_data rows (e.g. HEAD-level parents) whose value
    #    should be computed via the indicator's consolidation / scoped rule.
    rule_driven: set[tuple] = set()
    for gk in list(_original_cell_keys):
        if gk in global_formulas or gk in _skipped_formulas:
            continue
        sid, ck = gk
        meta = sheet_meta.get(sid)
        if not meta or not meta.get("rules_by_indicator"):
            continue
        context = _context_from_key(ck, meta["ordered_aids"])
        if _resolve_indicator_formula(sid, context, meta) is not None:
            get_cell(sid, ck)
            rule_driven.add(gk)

    # ── Evaluate indicator rules for ALL leaf combos (not just existing cells).
    #    When a new analytic dimension is added, cells for non-first leaves
    #    don't exist in cell_data yet, but indicator rules should still apply.
    for sid, meta in sheet_meta.items():
        if not meta.get("rules_by_indicator"):
            continue
        ordered_aids = meta["ordered_aids"]
        main_aid = meta.get("main_aid")
        period_aid = meta.get("period_aid")
        if not main_aid or not period_aid:
            continue
        children_by = meta["children_by_rid"]
        name_to_rids = meta.get("name_to_rids", {})

        # Indicators that have rules
        indicators_with_rules = set(meta["rules_by_indicator"].keys())
        if not indicators_with_rules:
            continue

        # Collect leaf records for each axis
        leaf_rids_by_aid: dict[str, list[str]] = {}
        for aid in ordered_aids:
            if aid == main_aid:
                continue
            all_rids = []
            for _, rids_list in name_to_rids.get(aid, {}).items():
                all_rids.extend(rids_list)
            # Leaf = no children
            leaves = [r for r in all_rids if not children_by.get(r)]
            leaf_rids_by_aid[aid] = leaves if leaves else all_rids

        # Generate all leaf-level combos for indicators with rules
        axes_order = [aid for aid in ordered_aids if aid != main_aid]
        if not axes_order:
            continue
        axes_rids = [leaf_rids_by_aid.get(aid, []) for aid in axes_order]
        if not all(axes_rids):
            continue

        for ind_rid in indicators_with_rules:
            for combo in itertools.product(*axes_rids):
                parts = []
                ci = 0
                for aid in ordered_aids:
                    if aid == main_aid:
                        parts.append(ind_rid)
                    else:
                        parts.append(combo[ci])
                        ci += 1
                ck = "|".join(parts)
                gk = (sid, ck)
                if gk in computed_set:
                    continue
                context = _context_from_key(ck, ordered_aids)
                if _resolve_indicator_formula(sid, context, meta) is not None:
                    get_cell(sid, ck)
                    if gk in computed_set:
                        rule_driven.add(gk)

    # ── Compute consolidation cells for parent periods (year/quarter totals).
    #    These cells may not exist in cell_data but need to be computed by
    #    aggregating children (months → quarter → year).
    consol_computed: set[tuple] = set()
    for sid, meta in sheet_meta.items():
        ordered_aids = meta["ordered_aids"]
        children_by = meta["children_by_rid"]
        record_by = meta.get("record_by_id", {})
        name_to_rids = meta.get("name_to_rids", {})
        main_aid = meta.get("main_aid")
        period_aid = meta.get("period_aid")
        if not main_aid or not period_aid:
            continue

        # Collect all indicator record IDs (main analytic)
        ind_rids: set[str] = set()
        for _, rids_list in name_to_rids.get(main_aid, {}).items():
            ind_rids.update(rids_list)
        if not ind_rids:
            continue

        # Collect parent period records (those with children in the period analytic)
        parent_period_rids: list[str] = []
        for rid, ch in children_by.items():
            if ch:
                rec = record_by.get(rid)
                if rec and rec.get("analytic_id") == period_aid:
                    parent_period_rids.append(rid)

        # Collect all record IDs for non-main, non-period analytics
        other_axes_rids: list[list[str]] = []
        other_axes_aids: list[str] = []
        for aid in ordered_aids:
            if aid == main_aid or aid == period_aid:
                continue
            rids_for_axis = []
            for _, rids_list in name_to_rids.get(aid, {}).items():
                rids_for_axis.extend(rids_list)
            if rids_for_axis:
                other_axes_rids.append(rids_for_axis)
                other_axes_aids.append(aid)

        # Build all combos of other axes (usually empty for 2-analytic sheets)
        if other_axes_rids:
            other_combos = list(itertools.product(*other_axes_rids))
        else:
            other_combos = [()]

        for prec_id in parent_period_rids:
            for ind_id in ind_rids:
                for other_vals in other_combos:
                    parts = []
                    oi = 0
                    for aid in ordered_aids:
                        if aid == period_aid:
                            parts.append(prec_id)
                        elif aid == main_aid:
                            parts.append(ind_id)
                        else:
                            parts.append(other_vals[oi])
                            oi += 1
                    ck = "|".join(parts)
                    gk = (sid, ck)
                    if gk not in computed_set:
                        get_cell(sid, ck)
                    # Always mark as consol_computed so it gets saved
                    # (may have been computed recursively by a parent's get_cell)
                    if gk in computed_set:
                        consol_computed.add(gk)

    # ── Compute consolidation along non-period, non-main analytics ──
    #    E.g. HEAD = SUM(F1, F2) for every indicator × period combination.
    #    The period consolidation above only covers parent-period records;
    #    this covers leaf periods AND parent periods for the extra analytic axes.
    for sid, meta in sheet_meta.items():
        ordered_aids = meta["ordered_aids"]
        children_by = meta["children_by_rid"]
        record_by = meta.get("record_by_id", {})
        name_to_rids = meta.get("name_to_rids", {})
        main_aid = meta.get("main_aid")
        period_aid = meta.get("period_aid")
        if not main_aid or not period_aid:
            continue

        # Find ALL analytics (including main) that have parent records needing consolidation.
        # Main-axis parents (e.g. "Итого" group indicators) also need SUM over children.
        consol_axes: list[tuple[str, list[str]]] = []  # (aid, [parent_rids])
        for aid in ordered_aids:
            if aid == period_aid:
                continue  # period consolidation handled above
            parent_rids = []
            for rid, ch in children_by.items():
                if ch:
                    rec = record_by.get(rid)
                    if rec and rec.get("analytic_id") == aid:
                        parent_rids.append(rid)
            if parent_rids:
                consol_axes.append((aid, parent_rids))
        if not consol_axes:
            continue

        # Collect ALL period rids (leaf + parent)
        all_period_rids: list[str] = []
        for _, rids_list in name_to_rids.get(period_aid, {}).items():
            all_period_rids.extend(rids_list)

        # Collect all indicator rids
        ind_rids2: set[str] = set()
        for _, rids_list in name_to_rids.get(main_aid, {}).items():
            ind_rids2.update(rids_list)

        # For each consolidating axis, iterate parent records
        for consol_aid, parent_rids in consol_axes:
            # Other axes: all record IDs for remaining non-main, non-period, non-consol axes
            other_axes_rids2: list[list[str]] = []
            other_axes_aids2: list[str] = []
            for aid in ordered_aids:
                if aid in (main_aid, period_aid, consol_aid):
                    continue
                rids_for_axis = []
                for _, rids_list in name_to_rids.get(aid, {}).items():
                    rids_for_axis.extend(rids_list)
                if rids_for_axis:
                    other_axes_rids2.append(rids_for_axis)
                    other_axes_aids2.append(aid)
            other_combos2 = list(itertools.product(*other_axes_rids2)) if other_axes_rids2 else [()]

            if consol_aid == main_aid:
                # Main-axis consolidation: parent indicator records
                # (e.g. "Итого" group) are the parents, not a separate axis.
                for p_rid in all_period_rids:
                    for parent_rid in parent_rids:
                        for other_vals in other_combos2:
                            parts = []
                            oi = 0
                            for aid in ordered_aids:
                                if aid == period_aid:
                                    parts.append(p_rid)
                                elif aid == main_aid:
                                    parts.append(parent_rid)
                                else:
                                    parts.append(other_vals[oi])
                                    oi += 1
                            ck = "|".join(parts)
                            gk = (sid, ck)
                            # Skip if neither the parent cell nor any child
                            # cell exists in the original data for this
                            # period.  This prevents phantom consolidation
                            # rows when Excel has None for the parent.
                            if gk not in _original_cell_keys and gk not in computed_set:
                                child_rids_for_parent = children_by.get(parent_rid, [])
                                has_child = False
                                for crid in child_rids_for_parent:
                                    cparts = list(parts)
                                    for ci2, a2 in enumerate(ordered_aids):
                                        if a2 == consol_aid:
                                            cparts[ci2] = crid
                                    cgk = (sid, "|".join(cparts))
                                    if cgk in _original_cell_keys or cgk in computed_set:
                                        has_child = True
                                        break
                                if not has_child:
                                    continue
                            if gk not in computed_set:
                                get_cell(sid, ck)
                            if gk in computed_set:
                                consol_computed.add(gk)
            else:
                for p_rid in all_period_rids:
                    for ind_id in ind_rids2:
                        for parent_rid in parent_rids:
                            for other_vals in other_combos2:
                                parts = []
                                oi = 0
                                for aid in ordered_aids:
                                    if aid == period_aid:
                                        parts.append(p_rid)
                                    elif aid == main_aid:
                                        parts.append(ind_id)
                                    elif aid == consol_aid:
                                        parts.append(parent_rid)
                                    else:
                                        parts.append(other_vals[oi])
                                        oi += 1
                                ck = "|".join(parts)
                                gk = (sid, ck)
                                if gk not in computed_set:
                                    get_cell(sid, ck)
                                if gk in computed_set:
                                    consol_computed.add(gk)

    # ── Return only cells whose value actually changed ──
    def _vals_equal(a: str, b: str) -> bool:
        if a == b:
            return True
        try:
            fa, fb = float(a), float(b)
            if fa == fb:
                return True
            if fa == 0 and fb == 0:
                return True
            # Relative tolerance
            if abs(fa) > 1e-9:
                return abs(fa - fb) / abs(fa) < 1e-6
        except (ValueError, TypeError):
            pass
        return False

    if _unresolved:
        print(f"[formula_engine] {len(_unresolved)} cells have unresolvable refs — skipping them")
    result: dict[str, dict[str, str]] = {}
    for gk in list(global_formulas) + list(rule_driven) + list(consol_computed):
        if gk in _unresolved:
            continue  # Don't overwrite cells with unresolvable cross-sheet refs
        sid, ck = gk
        new_val = global_cells.get(gk, "")
        old_val = _original_values.get(gk, "")
        # Skip if value didn't change.
        if _vals_equal(old_val, new_val):
            continue
        result.setdefault(sid, {})[ck] = new_val

    return result


# ── Standalone resolver (no recalc) ─────────────────────────────────────────
# Used by the /resolved-formulas API to tell the UI which formula *would* be
# applied to a given cell, and why. Mirrors the precedence used in get_cell.

async def resolve_formula_for_display(db, sheet_id: str, coord_key: str) -> dict:
    """Return {formula: str, source: 'cell'|'rule:<id>'|'default-sum'|'manual', kind: str|None}.

    Does NOT evaluate — just tells the UI which formula applies.
    """
    # 1. Per-cell explicit formula
    cell_rows = await db.execute_fetchall(
        "SELECT rule, formula FROM cell_data WHERE sheet_id = ? AND coord_key = ?",
        (sheet_id, coord_key),
    )
    cell = dict(cell_rows[0]) if cell_rows else None
    if cell and cell.get("rule") == "formula" and cell.get("formula"):
        return {"formula": cell["formula"], "source": "cell", "kind": None}

    # Load sheet metadata needed for rule resolution.
    bindings = await db.execute_fetchall(
        "SELECT sa.analytic_id, sa.sort_order, sa.is_main, a.is_periods "
        "FROM sheet_analytics sa JOIN analytics a ON a.id = sa.analytic_id "
        "WHERE sa.sheet_id = ? ORDER BY sa.sort_order",
        (sheet_id,),
    )
    ordered_aids = [b["analytic_id"] for b in bindings]
    main_aid = next((b["analytic_id"] for b in bindings if b["is_main"]), None)
    parts = coord_key.split("|")
    context = {aid: parts[i] for i, aid in enumerate(ordered_aids) if i < len(parts)}
    if not main_aid:
        return {"formula": "", "source": "manual", "kind": None}
    indicator_rid = context.get(main_aid)
    if not indicator_rid:
        return {"formula": "", "source": "manual", "kind": None}

    # children index for is_consolidating
    child_rows = await db.execute_fetchall(
        "SELECT id, parent_id FROM analytic_records WHERE analytic_id IN ({})".format(
            ",".join("?" * len(ordered_aids))
        ),
        tuple(ordered_aids),
    )
    children_by_rid: dict[str, list[str]] = {}
    for r in child_rows:
        if r["parent_id"]:
            children_by_rid.setdefault(r["parent_id"], []).append(r["id"])

    # 2. Scoped rules for this indicator on this sheet.
    rule_rows = await db.execute_fetchall(
        "SELECT id, kind, scope_json, priority, formula FROM indicator_formula_rules "
        "WHERE sheet_id = ? AND indicator_id = ?",
        (sheet_id, indicator_rid),
    )
    rules = []
    for r in rule_rows:
        try:
            scope = json.loads(r["scope_json"]) if r["scope_json"] else {}
        except Exception:
            scope = {}
        rules.append({
            "id": r["id"],
            "kind": r["kind"],
            "scope": scope,
            "priority": r["priority"] or 0,
            "formula": r["formula"] or "",
        })

    non_main = {a: rid for a, rid in context.items() if a != main_aid}
    scoped_hits = [
        r for r in rules
        if r["kind"] == "scoped" and r["scope"]
        and all(non_main.get(a) in (v or "").split(",") for a, v in r["scope"].items() if v)
    ]
    if scoped_hits:
        best = sorted(
            scoped_hits,
            key=lambda r: (-(r["priority"]), -len(r["scope"]), r["id"]),
        )[0]
        if best["formula"]:
            return {"formula": best["formula"], "source": f"rule:{best['id']}", "kind": "scoped"}

    # 3. Base consolidation / leaf.
    is_consol = any(
        children_by_rid.get(rid)
        for aid, rid in context.items()
    )
    base_kind = "consolidation" if is_consol else "leaf"
    for r in rules:
        if r["kind"] == base_kind and r["formula"]:
            return {"formula": r["formula"], "source": f"rule:{r['id']}", "kind": base_kind}

    # 4. No rule → default sum for consolidating, else manual.
    if is_consol:
        return {"formula": "SUM", "source": "default-sum", "kind": "consolidation"}
    return {"formula": "", "source": "manual", "kind": None}


# ── Convenience wrapper ────────────────────────────────────────────────────

async def calculate_sheet(db, sheet_id: str) -> dict[str, str]:
    """Calculate formulas for the model containing this sheet."""
    sheet = await db.execute_fetchall("SELECT model_id FROM sheets WHERE id = ?", (sheet_id,))
    if not sheet: return {}
    result = await calculate_model(db, sheet[0]["model_id"])
    return result.get(sheet_id, {})
