"""Indicator formula rules — per-indicator, per-sheet formulas for leaf,
consolidation, and scoped cases. See plan at
/Users/mac/.claude/plans/zippy-zooming-pelican.md.
"""
import json
import uuid

from fastapi import APIRouter, HTTPException
from pydantic import BaseModel

from backend.coord_key import (
    unpack as _unpack_coord,
    _load_all as _coord_key_prime,
    normalize as _ck_normalize,
)
from backend.db import get_db
from backend.formula_engine import resolve_formula_for_display

router = APIRouter(prefix="/api/sheets", tags=["indicator_rules"])


class ScopedRuleIn(BaseModel):
    id: str | None = None
    scope: dict[str, str] = {}   # {analytic_id: record_id}
    priority: int = 0
    formula: str = ""


class RulesIn(BaseModel):
    leaf: str = ""
    consolidation: str = ""
    scoped: list[ScopedRuleIn] = []


class PromoteCellIn(BaseModel):
    coord_key: str
    formula: str
    priority: int = 100


class ResolveIn(BaseModel):
    coord_keys: list[str]


# ── List / replace rules for an indicator ─────────────────────────────────

async def _synthesize_scoped_from_cells(
    db, sheet_id: str, indicator_id: str,
) -> tuple[str, list[dict]]:
    """For an indicator that has per-cell formulas in cell_data, synthesize a
    leaf formula (when uniform) or a list of scoped rules (one per distinct
    formula, with scope = comma-separated record_ids per non-main dimension).

    Returns (leaf, scoped_rules). leaf is set only when ALL formula cells
    share one formula text; otherwise scoped_rules contains one entry per
    distinct formula. Manual / empty cells aren't represented as rules —
    they live in cell_data and are surfaced via /resolved-formulas.
    """
    seq_row = await db.execute_fetchall(
        "SELECT seq_id FROM analytic_records WHERE id = ?", (indicator_id,),
    )
    ind_seq = seq_row[0]["seq_id"] if seq_row else None
    if ind_seq is None:
        return "", []
    cell_rows = await db.execute_fetchall(
        """SELECT cd.coord_key, COALESCE(NULLIF(cd.formula,''), f.text, '') AS formula
           FROM cell_data cd LEFT JOIN formulas f ON f.id = cd.formula_id
           WHERE cd.sheet_id = ?
             AND (cd.coord_key = ? OR cd.coord_key LIKE ? OR cd.coord_key LIKE ? OR cd.coord_key LIKE ?)
             AND ((cd.formula IS NOT NULL AND cd.formula != '') OR cd.formula_id IS NOT NULL)""",
        (sheet_id, str(ind_seq), f"{ind_seq}|%", f"%|{ind_seq}", f"%|{ind_seq}|%"),
    )
    bindings = await db.execute_fetchall(
        "SELECT analytic_id, is_main FROM sheet_analytics WHERE sheet_id = ? ORDER BY sort_order",
        (sheet_id,),
    )
    ordered_aids = [b["analytic_id"] for b in bindings]
    main_aid = next((b["analytic_id"] for b in bindings if b["is_main"]), None)
    if not main_aid:
        return "", []
    await _coord_key_prime(db)
    # formula → {analytic_id → ordered list of unique record_ids}
    groups: dict[str, dict[str, list[str]]] = {}
    counts: dict[str, int] = {}
    for cr in cell_rows:
        f = (cr["formula"] or "").strip()
        if not f:
            continue
        parts = _unpack_coord(cr["coord_key"])
        if len(parts) != len(ordered_aids):
            continue
        scope_vals = groups.setdefault(f, {})
        counts[f] = counts.get(f, 0) + 1
        for aid, part in zip(ordered_aids, parts):
            if aid == main_aid:
                continue
            lst = scope_vals.setdefault(aid, [])
            if part not in lst:
                lst.append(part)
    if not groups:
        return "", []
    if len(groups) == 1:
        return next(iter(groups.keys())), []
    scoped: list[dict] = []
    for formula, scope_vals in groups.items():
        scope = {aid: ",".join(rids) for aid, rids in scope_vals.items() if rids}
        scoped.append({
            "id": None,  # synthesized — persisting via PUT creates a real rule
            "scope": scope,
            "priority": 0,
            "formula": formula,
            "synthesized": True,
            "cell_count": counts.get(formula, 0),
        })
    # Largest groups first so the most-applied formula is on top.
    scoped.sort(key=lambda r: (-r["cell_count"], r["formula"]))
    return "", scoped


@router.get("/{sheet_id}/indicators/{indicator_id}/rules")
async def get_rules(sheet_id: str, indicator_id: str):
    db = get_db()
    rows = await db.execute_fetchall(
        """SELECT id, kind, scope_json, priority, formula
           FROM indicator_formula_rules
           WHERE sheet_id = ? AND indicator_id = ?
           ORDER BY priority DESC, created_at""",
        (sheet_id, indicator_id),
    )
    leaf = ""
    consolidation = ""
    scoped = []
    for r in rows:
        if r["kind"] == "leaf":
            leaf = r["formula"] or ""
        elif r["kind"] == "consolidation":
            consolidation = r["formula"] or ""
        elif r["kind"] == "scoped":
            try:
                scope = json.loads(r["scope_json"]) if r["scope_json"] else {}
            except Exception:
                scope = {}
            scoped.append({
                "id": r["id"],
                "scope": scope,
                "priority": r["priority"] or 0,
                "formula": r["formula"] or "",
            })
    # No persisted leaf — derive from per-cell formulas in cell_data:
    # all cells share one formula → that's the leaf; otherwise synthesize one
    # scoped rule per distinct formula (each rule's scope lists every
    # non-main analytic value seen with that formula).
    if not leaf and not scoped:
        derived_leaf, synth_scoped = await _synthesize_scoped_from_cells(
            db, sheet_id, indicator_id,
        )
        if derived_leaf:
            leaf = derived_leaf
        if synth_scoped:
            scoped = synth_scoped

    return {"leaf": leaf, "consolidation": consolidation, "scoped": scoped}


@router.put("/{sheet_id}/indicators/{indicator_id}/rules")
async def put_rules(sheet_id: str, indicator_id: str, body: RulesIn):
    """Replace-all: wipes existing rules for this (sheet, indicator) and
    writes the new set. Use empty strings to clear a base formula."""
    db = get_db()
    await db.execute(
        "DELETE FROM indicator_formula_rules WHERE sheet_id = ? AND indicator_id = ?",
        (sheet_id, indicator_id),
    )
    if body.leaf:
        await db.execute(
            "INSERT INTO indicator_formula_rules "
            "(id, sheet_id, indicator_id, kind, scope_json, priority, formula) "
            "VALUES (?, ?, ?, 'leaf', '{}', 0, ?)",
            (str(uuid.uuid4()), sheet_id, indicator_id, body.leaf),
        )
    if body.consolidation:
        await db.execute(
            "INSERT INTO indicator_formula_rules "
            "(id, sheet_id, indicator_id, kind, scope_json, priority, formula) "
            "VALUES (?, ?, ?, 'consolidation', '{}', 0, ?)",
            (str(uuid.uuid4()), sheet_id, indicator_id, body.consolidation),
        )
    for r in body.scoped:
        await db.execute(
            "INSERT INTO indicator_formula_rules "
            "(id, sheet_id, indicator_id, kind, scope_json, priority, formula) "
            "VALUES (?, ?, ?, 'scoped', ?, ?, ?)",
            (
                r.id or str(uuid.uuid4()),
                sheet_id,
                indicator_id,
                json.dumps(r.scope, ensure_ascii=False),
                r.priority,
                r.formula,
            ),
        )
    await db.commit()
    # Invalidate V4 engine cache (structural change)
    model_row = await db.execute_fetchall(
        "SELECT model_id FROM sheets WHERE id = ?", (sheet_id,),
    )
    if model_row:
        from backend.formula_engine import invalidate_engine
        await invalidate_engine(db, model_row[0]["model_id"])
    return {"ok": True}


# ── Batch: all indicator rules for a sheet ────────────────────────────────

@router.get("/{sheet_id}/indicator-rules-all")
async def get_all_rules(sheet_id: str):
    """Return {record_id: {leaf, consolidation}} for every indicator record
    on this sheet. First checks indicator_formula_rules, then falls back to
    per-cell formulas from cell_data."""
    db = get_db()
    # 1. Check indicator_formula_rules first
    rows = await db.execute_fetchall(
        """SELECT indicator_id, kind, formula
           FROM indicator_formula_rules
           WHERE sheet_id = ?
           ORDER BY indicator_id, priority DESC""",
        (sheet_id,),
    )
    result: dict[str, dict[str, str]] = {}
    for r in rows:
        iid = r["indicator_id"]
        entry = result.setdefault(iid, {"leaf": "", "consolidation": ""})
        if r["kind"] == "leaf" and not entry["leaf"]:
            entry["leaf"] = r["formula"] or ""
        elif r["kind"] == "consolidation" and not entry["consolidation"]:
            entry["consolidation"] = r["formula"] or ""

    # 2. Also collect per-cell formulas from cell_data and detect whether all
    # cells of an indicator share the same formula. If they do, surface that
    # as the indicator-level "leaf". If they diverge — formulas differ per
    # cell — return the sentinel "__per_cell__" so the UI shows that the
    # indicator has cell-specific formulas instead of misleadingly picking
    # one as representative.
    cell_rows = await db.execute_fetchall(
        """SELECT cd.coord_key, COALESCE(NULLIF(cd.formula,''), f.text, '') AS formula, cd.rule
           FROM cell_data cd LEFT JOIN formulas f ON f.id = cd.formula_id
           WHERE cd.sheet_id = ?
             AND ((cd.formula IS NOT NULL AND cd.formula != '') OR cd.formula_id IS NOT NULL)
           ORDER BY cd.coord_key""",
        (sheet_id,),
    )
    # Determine indicator position in coord_key from sheet_analytics binding order.
    bindings = await db.execute_fetchall(
        "SELECT analytic_id, is_main FROM sheet_analytics WHERE sheet_id = ? ORDER BY sort_order",
        (sheet_id,),
    )
    main_idx = None
    for i, b in enumerate(bindings):
        if b["is_main"]:
            main_idx = i
            break
    if main_idx is None:
        return result

    await _coord_key_prime(db)
    distinct_by_ind: dict[str, set[str]] = {}
    for cr in cell_rows:
        parts = _unpack_coord(cr["coord_key"])
        if len(parts) <= main_idx:
            continue
        rec_id = parts[main_idx]
        if rec_id in result and result[rec_id]["leaf"]:
            continue  # indicator-level rule already wins
        f = (cr["formula"] or "").strip()
        if f:
            distinct_by_ind.setdefault(rec_id, set()).add(f)

    for rec_id, formulas in distinct_by_ind.items():
        if rec_id not in result:
            result[rec_id] = {"leaf": "", "consolidation": ""}
        if result[rec_id]["leaf"]:
            continue
        if len(formulas) == 1:
            result[rec_id]["leaf"] = next(iter(formulas))
        else:
            # Multiple distinct per-cell formulas — UI should show this as
            # "разные формулы у клеток", not a single representative.
            result[rec_id]["leaf"] = "__per_cell__"

    return result


# ── Promote a per-cell formula into a scoped rule ─────────────────────────

@router.post("/{sheet_id}/indicators/{indicator_id}/rules/promote-cell")
async def promote_cell(sheet_id: str, indicator_id: str, body: PromoteCellIn):
    """Convert a per-cell formula into a scoped rule on the indicator.
    The scope is derived from coord_key minus the main analytic's part.
    """
    db = get_db()
    bindings = await db.execute_fetchall(
        """SELECT sa.analytic_id, sa.is_main, sa.sort_order
           FROM sheet_analytics sa WHERE sa.sheet_id = ?
           ORDER BY sa.sort_order""",
        (sheet_id,),
    )
    if not bindings:
        raise HTTPException(400, "sheet has no analytic bindings")
    ordered_aids = [b["analytic_id"] for b in bindings]
    main_aid = next((b["analytic_id"] for b in bindings if b["is_main"]), None)
    if not main_aid:
        raise HTTPException(400, "sheet has no main analytic")
    body.coord_key = await _ck_normalize(db, body.coord_key, read_only=False)
    parts = _unpack_coord(body.coord_key)
    if len(parts) != len(ordered_aids):
        raise HTTPException(400, "coord_key length mismatch")
    scope: dict[str, str] = {}
    for aid, part in zip(ordered_aids, parts):
        if aid == main_aid:
            continue
        scope[aid] = part

    # Pick next priority = max existing + 10 (so promoted rules beat older ones).
    max_row = await db.execute_fetchall(
        """SELECT MAX(priority) AS m FROM indicator_formula_rules
           WHERE sheet_id = ? AND indicator_id = ?""",
        (sheet_id, indicator_id),
    )
    base = (max_row[0]["m"] if max_row and max_row[0]["m"] is not None else 0) or 0
    prio = max(body.priority, base + 10)

    rid = str(uuid.uuid4())
    await db.execute(
        "INSERT INTO indicator_formula_rules "
        "(id, sheet_id, indicator_id, kind, scope_json, priority, formula) "
        "VALUES (?, ?, ?, 'scoped', ?, ?, ?)",
        (rid, sheet_id, indicator_id, json.dumps(scope, ensure_ascii=False), prio, body.formula),
    )
    # Clear the per-cell override so the rule takes effect.
    await db.execute(
        "UPDATE cell_data SET rule = 'manual', formula = '' "
        "WHERE sheet_id = ? AND coord_key = ?",
        (sheet_id, body.coord_key),
    )
    await db.commit()
    # Invalidate V4 engine cache (structural change)
    model_row = await db.execute_fetchall(
        "SELECT model_id FROM sheets WHERE id = ?", (sheet_id,),
    )
    if model_row:
        from backend.formula_engine import invalidate_engine
        await invalidate_engine(db, model_row[0]["model_id"])
    return {"ok": True, "rule_id": rid, "priority": prio, "scope": scope}


# ── Resolve which formula would apply to a batch of coords ─────────────────

@router.post("/{sheet_id}/cells/resolved-formulas")
async def resolved_formulas(sheet_id: str, body: ResolveIn):
    db = get_db()
    out = []
    for ck in body.coord_keys:
        info = await resolve_formula_for_display(db, sheet_id, ck)
        out.append({"coord_key": ck, **info})
    return out
