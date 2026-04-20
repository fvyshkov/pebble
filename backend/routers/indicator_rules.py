"""Indicator formula rules — per-indicator, per-sheet formulas for leaf,
consolidation, and scoped cases. See plan at
/Users/mac/.claude/plans/zippy-zooming-pelican.md.
"""
import json
import uuid

from fastapi import APIRouter, HTTPException
from pydantic import BaseModel

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
    # Fallback: if no rules found, look for per-cell formulas in cell_data
    if not leaf and not consolidation and not scoped:
        bindings = await db.execute_fetchall(
            "SELECT analytic_id, is_main FROM sheet_analytics WHERE sheet_id = ? ORDER BY sort_order",
            (sheet_id,),
        )
        main_idx = next((i for i, b in enumerate(bindings) if b["is_main"]), None)
        if main_idx is not None:
            cell_row = await db.execute_fetchall(
                """SELECT formula FROM cell_data
                   WHERE sheet_id = ? AND formula IS NOT NULL AND formula != ''
                   LIMIT 100""",
                (sheet_id,),
            )
            for cr in cell_row:
                # Not efficient but coord_key is not indexed by indicator_id alone;
                # match cells whose coord_key has indicator_id at main_idx position.
                pass
            # Simpler approach: query by coord_key LIKE pattern
            # coord_key parts are joined with |, indicator_id is at main_idx
            like_patterns = []
            if main_idx == 0:
                like_patterns.append(f"{indicator_id}|%")
            elif main_idx == 1:
                like_patterns.append(f"%|{indicator_id}")
            else:
                like_patterns.append(f"%|{indicator_id}|%")

            for pat in like_patterns:
                cell_row = await db.execute_fetchall(
                    """SELECT formula FROM cell_data
                       WHERE sheet_id = ? AND coord_key LIKE ?
                       AND formula IS NOT NULL AND formula != ''
                       LIMIT 1""",
                    (sheet_id, pat),
                )
                if cell_row:
                    leaf = cell_row[0]["formula"] or ""
                    break

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

    if result:
        return result

    # 2. Fallback: extract per-indicator formulas from cell_data.
    # coord_key = "period_rec_id|indicator_rec_id" — take the indicator part.
    # Pick any one non-empty formula per indicator record (they're typically the same).
    cell_rows = await db.execute_fetchall(
        """SELECT coord_key, formula, rule
           FROM cell_data
           WHERE sheet_id = ? AND formula IS NOT NULL AND formula != ''
           ORDER BY coord_key""",
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

    for cr in cell_rows:
        parts = cr["coord_key"].split("|")
        if len(parts) <= main_idx:
            continue
        rec_id = parts[main_idx]
        if rec_id in result:
            continue  # already have a formula for this record
        result[rec_id] = {"leaf": cr["formula"] or "", "consolidation": ""}

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
    parts = body.coord_key.split("|")
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
