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
import time
import zlib
from typing import Any

# zlib magic byte (0x78 = zlib stream); raw bincode starts with random bytes.
# We prepend this 1-byte tag so old uncompressed blobs still load.
_DAG_COMPRESS_MAGIC = b"\x01"   # tag byte for "zlib-compressed payload follows"

# ── Rust V4 engine (only engine — no fallbacks) ──────────────────────────
import pebble_calc as _rust_engine  # mandatory — crash if not installed

# ── V4: Stateful DAG engine cache ────────────────────────────────────────
_engine_cache: dict[str, Any] = {}  # model_id → pebble_calc.CalcEngine


async def _try_load_dag_from_db(db, model_id: str):
    """Try to load a cached DAG blob from the database. Returns CalcEngine or None."""
    rows = await db.execute_fetchall(
        "SELECT dag_blob FROM dag_cache WHERE model_id = ?", (model_id,),
    )
    if not rows or not rows[0]["dag_blob"]:
        return None
    try:
        blob = bytes(rows[0]["dag_blob"])
        if blob[:1] == _DAG_COMPRESS_MAGIC:
            t0 = time.perf_counter()
            blob = zlib.decompress(blob[1:])
            print(f"[formula_engine] V4 DAG decompressed: {time.perf_counter()-t0:.3f}s, {len(blob)/1024**2:.0f} MB")
        engine = _rust_engine.CalcEngine()
        engine.load(blob)
        print(f"[formula_engine] V4 engine loaded from DB cache for model {model_id}")
        return engine
    except Exception as e:
        print(f"[formula_engine] V4 DAG cache load failed: {e}")
        return None


async def _save_dag_to_db(db, model_id: str, engine):
    """Save the built DAG to the database for persistence (zlib-compressed)."""
    try:
        raw = engine.serialize()
        t0 = time.perf_counter()
        blob = _DAG_COMPRESS_MAGIC + zlib.compress(raw, 1)  # level 1 — fast, ~30% ratio
        dt = time.perf_counter() - t0
        await db.execute(
            """INSERT INTO dag_cache (model_id, dag_blob, created_at)
               VALUES (?, ?, datetime('now'))
               ON CONFLICT(model_id) DO UPDATE SET dag_blob = excluded.dag_blob, created_at = excluded.created_at""",
            (model_id, blob),
        )
        await db.commit()
        print(f"[formula_engine] V4 DAG saved to DB cache for model {model_id} "
              f"({len(raw)/1024**2:.0f} MB → {len(blob)/1024**2:.0f} MB, compress={dt:.2f}s)")
    except Exception as e:
        print(f"[formula_engine] V4 DAG cache save failed: {e}")


async def _get_or_build_engine(db, model_id: str, model_json: str):
    """Get cached CalcEngine for model. Tries: memory → DB → full build.
    Returns (engine, result). result is None if engine was loaded from cache."""
    # 1. Check in-memory cache
    engine = _engine_cache.get(model_id)
    if engine is not None and engine.is_built():
        return engine, None

    # 2. Try loading from DB
    engine = await _try_load_dag_from_db(db, model_id)
    if engine is not None:
        _engine_cache[model_id] = engine
        return engine, None

    # 3. Full build
    engine = _rust_engine.CalcEngine()
    result = engine.build(model_json)
    _engine_cache[model_id] = engine

    # Save to DB in background (non-blocking for the response)
    await _save_dag_to_db(db, model_id, engine)

    return engine, result


async def invalidate_engine(db, model_id: str):
    """Called when model structure changes — forces full rebuild on next calc.
    Also marks the model as needing generation."""
    engine = _engine_cache.pop(model_id, None)
    if engine is not None:
        engine.drop_state()
    # Remove from DB cache
    try:
        await db.execute("DELETE FROM dag_cache WHERE model_id = ?", (model_id,))
        await db.commit()
    except Exception:
        pass
    # Mark model as needing generation
    try:
        await db.execute(
            "UPDATE models SET calc_status = 'needs_generation' WHERE id = ?",
            (model_id,),
        )
        await db.commit()
    except Exception:
        pass
    print(f"[formula_engine] V4 engine invalidated for model {model_id}")


async def calculate_model_incremental(
    db, model_id: str,
    changed_cells: list[tuple[str, str, str]],  # [(sheet_id, coord_key, value)]
) -> dict[str, dict[str, str]]:
    """Update specific cell values using cached DAG.
    Falls back to full recalc if engine not cached."""
    # Try incremental path: use cached engine + update_values
    engine = _engine_cache.get(model_id)
    if engine is None or not engine.is_built():
        # No cached engine — try loading from DB
        engine = await _try_load_dag_from_db(db, model_id)
        if engine is not None:
            _engine_cache[model_id] = engine

    if engine is not None and engine.is_built():
        t0 = time.perf_counter()
        changes_json = json.dumps(changed_cells)
        result = engine.update_values(changes_json)
        t1 = time.perf_counter()
        total_cells = sum(len(v) for v in result.values()) if isinstance(result, dict) else 0
        print(f"[formula_engine] V4 incremental: {t1-t0:.3f}s cells={total_cells}")
        return result

    # No cached engine — skip formula recalc (user must click Generate)
    print("[formula_engine] V4 incremental: no cached engine, skipping formula recalc")
    return {}


async def get_dirty_cells(
    db, model_id: str,
    changed_cells: list[tuple[str, str]],  # [(sheet_id, coord_key)]
) -> list[tuple[str, str]]:
    """Return list of (sheet_id, coord_key) affected by the changes.
    Returns empty list if engine not cached."""
    engine = _engine_cache.get(model_id)
    if engine is None or not engine.is_built():
        return []

    changes_json = json.dumps(changed_cells)
    return engine.mark_dirty(changes_json)


# ── Tokenizer ──────────────────────────────────────────────────────────────

TOKEN_RE = re.compile(r"""
    (\[(?:[^\[\]]*(?:\[[^\[\]]*\][^\[\]]*)*)\](?:\((?:[^()]*|\([^()]*\))*\))?)  |  # [ref](params) — supports nested [] in name
    (SUM|AVERAGE|IF|MIN|MAX|ABS|INT)\s*\(   |  # functions
    (\d+(?:\.\d+)?)                    |  # number
    ([+\-*/(),<>=!])                   |  # operators, parens, comparison
    (\s+)                                 # whitespace (skip)
""", re.VERBOSE)

REF_RE = re.compile(r"""
    \[((?:[^\[\]]*(?:\[[^\[\]]*\][^\[\]]*)*)+)\]  # indicator name (supports nested [])
    (?:\(((?:[^()]*|\([^()]*\))*)\))?             # optional params — one nesting level
""", re.VERBOSE)

# For matching key.назад(N) or key.вперед(N) function calls in param values
_PERIOD_BACK_RE = re.compile(r'\w+\.назад\((\d+)\)')
_PERIOD_FWD_RE = re.compile(r'\w+\.вперед\((\d+)\)')


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

            # Period forward-reference: word.вперед(N)
            fwd_m = _PERIOD_FWD_RE.fullmatch(val)
            if fwd_m:
                params[key] = f"вперед({fwd_m.group(1)})"
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
            elif func_name == "INT":
                v = _n(args[0]) if args else 0.0
                # Excel INT() snaps values within 5e-14 of the next integer upward
                # to avoid off-by-one errors from floating-point representation.
                # E.g. 110 * (6/11) in f64 = 59.9999999999999929... => should be 60.
                ceil_v = math.ceil(v)
                if 0 < ceil_v - v < 5e-14:
                    v = ceil_v
                return math.floor(v)
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


# ── Serialization bridge for Rust engine ───────────────────────────────────

def _serialize_for_rust(all_sheets, sheet_meta, global_cells, global_formulas,
                        manual_cells, phantom_cells, sheet_name_to_id,
                        prev_period, period_order) -> str:
    """Serialize loaded model data to JSON for the Rust engine."""
    sheets = []
    for s in all_sheets:
        sid = s["id"]
        meta = sheet_meta[sid]
        cells = {}
        for gk, val in global_cells.items():
            if gk[0] != sid:
                continue
            ck = gk[1]
            formula = global_formulas.get(gk, "")
            if gk in manual_cells:
                rule = "manual"
            elif gk in phantom_cells:
                rule = "phantom"
            elif formula:
                rule = "formula"
            else:
                # No formula text, not manual, not phantom — this is a
                # consolidation cell (rule=formula in DB but formula="").
                # Mark as "phantom" so Rust doesn't treat it as manual
                # and applies default SUM consolidation.
                rule = "phantom"
            cells[ck] = {"value": val, "rule": rule, "formula": formula}

        records = {}
        for rid, rec in meta["record_by_id"].items():
            data = rec.get("_data") or {}
            records[rid] = {
                "id": rid,
                "analytic_id": rec.get("analytic_id", ""),
                "parent_id": rec.get("parent_id"),
                "sort_order": rec.get("sort_order", 0),
                "name": data.get("name", ""),
                "period_key": data.get("period_key"),
                "excel_row": rec.get("excel_row"),
            }

        # Convert rules: ensure scope values are strings
        rules_by_indicator = {}
        for ind_rid, rules in meta.get("rules_by_indicator", {}).items():
            rules_by_indicator[ind_rid] = [
                {
                    "id": r["id"],
                    "kind": r["kind"],
                    "scope": {str(k): str(v) if v else "" for k, v in (r.get("scope") or {}).items()},
                    "priority": r.get("priority", 0),
                    "formula": r.get("formula", ""),
                }
                for r in rules
            ]

        # Convert name_to_rids: ensure all values are lists
        name_to_rids = {}
        for aid, nmap in meta.get("name_to_rids", {}).items():
            name_to_rids[aid] = {name: list(rids) for name, rids in nmap.items()}

        sheets.append({
            "id": sid,
            "name": meta["name"],
            "ordered_aids": meta["ordered_aids"],
            "period_aid": meta.get("period_aid"),
            "main_aid": meta.get("main_aid"),
            "analytic_name_to_id": meta.get("analytic_name_to_id", {}),
            "name_to_rids": name_to_rids,
            "records": records,
            "children_by_rid": {k: list(v) for k, v in meta.get("children_by_rid", {}).items()},
            "rules_by_indicator": rules_by_indicator,
            "cells": cells,
            "rid_to_period_key": meta.get("rid_to_period_key", {}),
            "period_key_to_rid": meta.get("period_key_to_rid", {}),
        })

    return json.dumps({
        "sheets": sheets,
        "sheet_name_to_id": sheet_name_to_id,
        "prev_period": prev_period,
        "period_order": period_order,
    })


# ── Model-level lazy calculator ────────────────────────────────────────────

async def calculate_model(db, model_id: str) -> dict[str, dict[str, str]]:
    """Calculate ALL formula cells across ALL sheets in a model.
    Returns {sheet_id: {coord_key: new_value}}.
    """
    # Fast path: engine state already loaded → skip DB reload + Python↔Rust marshaling.
    # Manual edits go through calculate_model_incremental which keeps the engine in sync,
    # and structural changes call invalidate_engine() which evicts the cache.
    cached = _engine_cache.get(model_id)
    if cached is None or not cached.is_built():
        cached = await _try_load_dag_from_db(db, model_id)
        if cached is not None:
            _engine_cache[model_id] = cached
    if cached is not None and cached.is_built():
        t0 = time.perf_counter()
        rust_result = cached.collect_all_changes()
        dt = time.perf_counter() - t0
        total = sum(len(v) for v in rust_result.values())
        print(f"[formula_engine] V4 cached: compute={dt:.3f}s cells={total}")
        return rust_result

    all_sheets = await db.execute_fetchall(
        "SELECT id, name, excel_code FROM sheets WHERE model_id = ? ORDER BY created_at", (model_id,))
    if not all_sheets:
        return {}

    # ── Load entire model ──
    global_cells = {}       # {(sheet_id, coord_key): value_str}
    global_formulas = {}    # {(sheet_id, coord_key): formula_str}
    _phantom_cells: set[tuple] = set()  # rule=formula, empty formula, val=0
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
                # Build prev_period chain for this period analytic.
                # Include ALL period records (not just leaves) so that navigation
                # works at any granularity level (M, Q, H, Y).
                parent_ids = {r["parent_id"] for r in recs if r["parent_id"]}
                _leaf_order = [r["id"] for r in recs if r["id"] not in parent_ids]
                for i in range(1, len(_leaf_order)):
                    if _leaf_order[i] not in prev_period:
                        prev_period[_leaf_order[i]] = _leaf_order[i - 1]
                # Also build prev_period for non-leaf periods grouped by level.
                # Group by period_key pattern: M (YYYY-MM), Q (YYYY-QN), H (YYYY-HN), Y (YYYY-Y)
                import re as _re
                _level_groups: dict[str, list] = {}  # level → [(sort_order, rid)]
                for r in recs:
                    _d = r.get("_data") or {}
                    pk = _d.get("period_key", "")
                    if _re.match(r'\d{4}-\d{2}$', pk):
                        lvl = "M"
                    elif "-Q" in pk:
                        lvl = "Q"
                    elif "-H" in pk:
                        lvl = "H"
                    elif pk.endswith("-Y"):
                        lvl = "Y"
                    else:
                        continue
                    _level_groups.setdefault(lvl, []).append((r["sort_order"], r["id"]))
                for lvl, items in _level_groups.items():
                    if lvl == "M":
                        continue  # leaves already handled above
                    items.sort(key=lambda x: x[0])
                    for j in range(1, len(items)):
                        rid_cur = items[j][1]
                        rid_prev = items[j - 1][1]
                        if rid_cur not in prev_period:
                            prev_period[rid_cur] = rid_prev
                if not period_aid_global:
                    period_aid_global = aid
                    period_order = _leaf_order

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

        # Build period_key ↔ rid mappings for cross-sheet period translation
        _rid_to_pk: dict[str, str] = {}
        _pk_to_rid: dict[str, str] = {}
        if period_aid:
            for rid, rec in record_by_id.items():
                data = rec.get("_data") or {}
                pk = data.get("period_key", "")
                if pk and rec.get("analytic_id") == period_aid:
                    _rid_to_pk[rid] = pk
                    _pk_to_rid[pk] = rid

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
            "rid_to_period_key": _rid_to_pk,
            "period_key_to_rid": _pk_to_rid,
        }

        for c in await db.execute_fetchall(
                """SELECT cd.coord_key, cd.value, cd.rule,
                          COALESCE(NULLIF(cd.formula,''), f.text, '') AS formula
                   FROM cell_data cd LEFT JOIN formulas f ON f.id = cd.formula_id
                   WHERE cd.sheet_id = ?""", (sid,)):
            gk = (sid, c["coord_key"])
            global_cells[gk] = c["value"] or ""
            if c["rule"] == "formula" and c["formula"]:
                # Skip raw Excel formulas (starting with =) — they can't be evaluated
                if not c["formula"].startswith("="):
                    global_formulas[gk] = c["formula"]
            elif c["rule"] == "formula" and not c["formula"]:
                val = c["value"] or ""
                if not val or val == "0":
                    _phantom_cells.add(gk)

    # Track manual cells — they must not be overwritten by consolidation
    _manual_cells: set[tuple] = set()
    # Track cells that have rule=formula but no formula text — these are
    # auto-created consolidation cells or migration artifacts. They should
    # NOT be marked as manual (they need default SUM consolidation in Rust).
    _formula_no_text: set[tuple] = set()
    for c in await db.execute_fetchall(
            "SELECT sheet_id, coord_key, rule FROM cell_data WHERE rule='formula' AND formula_id IS NULL AND (formula IS NULL OR formula='')"):
        _formula_no_text.add((c["sheet_id"], c["coord_key"]))
    for gk in global_cells:
        if gk not in global_formulas and gk not in _formula_no_text:
            _manual_cells.add(gk)
    # _empty_formula_cells are already identified during the cell loading above:
    # they have rule=formula, empty formula, and val=0/"". These are in global_cells
    # but NOT in global_formulas and NOT truly manual (they're phantom import artifacts).
    _empty_formula_cells = _phantom_cells

    # Track original DB cell keys (before computation adds synthetic ones)
    _original_cell_keys = set(global_cells.keys())
    # Snapshot original values so we can detect what actually changed.
    _original_values: dict[tuple, str] = dict(global_cells)

    # ── Rust V4 engine (only engine) ──
    t0 = time.perf_counter()
    model_json = _serialize_for_rust(
        all_sheets, sheet_meta, global_cells, global_formulas,
        _manual_cells, _phantom_cells, sheet_name_to_id,
        prev_period, period_order,
    )
    t1 = time.perf_counter()
    _engine, rust_result = await _get_or_build_engine(db, model_id, model_json)
    if rust_result is None:
        # Engine was loaded from cache — collect all changes
        rust_result = _engine.collect_all_changes()
    t2 = time.perf_counter()
    total_cells = sum(len(v) for v in rust_result.values())
    print(f"[formula_engine] Rust V4 engine: serialize={t1-t0:.3f}s compute={t2-t1:.3f}s cells={total_cells}")
    return rust_result


# ── Standalone resolver (no recalc) ─────────────────────────────────────────
# Used by the /resolved-formulas API to tell the UI which formula *would* be
# applied to a given cell, and why. Mirrors the precedence used in get_cell.

async def resolve_formula_for_display(db, sheet_id: str, coord_key: str) -> dict:
    """Return {formula: str, source: 'cell'|'rule:<id>'|'default-sum'|'manual', kind: str|None}.

    Does NOT evaluate — just tells the UI which formula applies.
    """
    # 1. Per-cell explicit formula
    cell_rows = await db.execute_fetchall(
        """SELECT cd.rule, COALESCE(NULLIF(cd.formula,''), f.text, '') AS formula
           FROM cell_data cd LEFT JOIN formulas f ON f.id = cd.formula_id
           WHERE cd.sheet_id = ? AND cd.coord_key = ?""",
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
