//! Engine V3: Pre-resolved DAG + Level-Parallel Evaluation.
//!
//! Instead of recursive get_cell() with HashMap lookups:
//! 1. Enumerate ALL cells upfront (input + generated from consolidation phases)
//! 2. Pre-resolve all formula references to dense cell_ids
//! 3. Topological sort into dependency levels
//! 4. Evaluate each level in parallel with rayon
//!
//! Values stored in Vec<f64> indexed by cell_id — no HashMap during evaluation.

use std::collections::HashMap;
use rustc_hash::FxHashMap;
use rayon::prelude::*;
use crate::types::ModelInput;
use crate::intern::Interner;
use crate::coord::*;
use crate::compiler::*;
use crate::builder::*;

// ── Constants ──────────────────────────────────────────────────────────

pub(crate) const UNRESOLVED: u32 = u32::MAX;
pub(crate) const SELF_REF_SENTINEL: u32 = u32::MAX - 1;

// ── Data structures ────────────────────────────────────────────────────

#[derive(serde::Serialize, serde::Deserialize)]
pub(crate) enum EvalKind {
    /// Not yet resolved — will be determined during resolve_all_refs.
    Pending,
    /// Return stored value (manual cells, phantoms, leaf defaults).
    Manual,
    /// Evaluate compiled formula. resolved_refs[i] → cell_id.
    Formula {
        formula_id: u32,
        resolved_refs: Vec<u32>,
    },
    /// Sum children cell values.
    ConsolidationSum {
        children: Vec<u32>,
    },
}

#[derive(serde::Serialize, serde::Deserialize)]
pub(crate) struct EvalCell {
    pub(crate) sheet_idx: u16,
    pub(crate) coord: CoordKey,
    pub(crate) original_value: f64,
    pub(crate) flags: u8,
    pub(crate) formula_id_hint: u32, // from CellState or NO_FORMULA
    pub(crate) eval_kind: EvalKind,
}

/// Wrapper to send raw pointer across threads.
struct SendPtr(*mut f64);
unsafe impl Send for SendPtr {}
unsafe impl Sync for SendPtr {}

// ── Main entry point ───────────────────────────────────────────────────

pub fn calculate_model(input: &ModelInput) -> HashMap<String, HashMap<String, String>> {
    let t0 = std::time::Instant::now();
    let model = build_model(input);
    let t1 = std::time::Instant::now();
    eprintln!("[v3] build_model: {:.3}s", (t1 - t0).as_secs_f64());

    let (mut cells, mut values, mut lookup) = enumerate_initial_cells(&model);
    let t2 = std::time::Instant::now();
    eprintln!("[v3] enumerate: {:.3}s, {} cells", (t2 - t1).as_secs_f64(), cells.len());

    resolve_all_refs(&model, &mut cells, &mut values, &mut lookup);
    let t3 = std::time::Instant::now();
    eprintln!("[v3] resolve_refs: {:.3}s, {} cells total", (t3 - t2).as_secs_f64(), cells.len());

    let levels = topological_sort(&cells);
    let t4 = std::time::Instant::now();
    eprintln!("[v3] topo_sort: {:.3}s, {} levels", (t4 - t3).as_secs_f64(), levels.len());

    evaluate_parallel(&cells, &levels, &mut values, &model.compiled_formulas);
    let t5 = std::time::Instant::now();
    eprintln!("[v3] evaluate: {:.3}s", (t5 - t4).as_secs_f64());

    let result = collect_changes(&cells, &values, &model);
    let t6 = std::time::Instant::now();
    eprintln!("[v3] collect: {:.3}s", (t6 - t5).as_secs_f64());
    eprintln!("[v3] TOTAL: {:.3}s", (t6 - t0).as_secs_f64());

    result
}

// ── Phase A: Enumerate all cells ───────────────────────────────────────

fn ensure_cell(
    cells: &mut Vec<EvalCell>,
    values: &mut Vec<f64>,
    lookup: &mut [FxHashMap<CoordKey, u32>],
    sheet_idx: usize,
    coord: CoordKey,
    value: f64,
    original_value: f64,
    flags: u8,
    formula_id: u32,
) -> u32 {
    if let Some(&id) = lookup[sheet_idx].get(&coord) {
        return id;
    }
    let id = cells.len() as u32;
    lookup[sheet_idx].insert(coord, id);
    values.push(value);
    cells.push(EvalCell {
        sheet_idx: sheet_idx as u16,
        coord,
        original_value,
        flags,
        formula_id_hint: formula_id,
        eval_kind: EvalKind::Pending,
    });
    id
}

/// Look up cell for a reference target. Only creates new cell if absolutely needed.
fn ensure_cell_for_ref(
    model: &Model,
    sheet_idx: usize,
    coord: CoordKey,
    cells: &mut Vec<EvalCell>,
    values: &mut Vec<f64>,
    lookup: &mut [FxHashMap<CoordKey, u32>],
) -> u32 {
    if let Some(&id) = lookup[sheet_idx].get(&coord) {
        return id;
    }
    // Cell doesn't exist. Check if it SHOULD exist.
    // This is called during reference resolution — avoid creating too many phantom cells.
    // Only create if the cell has a rule OR is consolidating (matches v2 behavior).
    let has_rule = resolve_indicator_formula_id(model, sheet_idx, &coord).is_some();
    if has_rule {
        return ensure_cell(cells, values, lookup, sheet_idx, coord,
                          0.0, 0.0, FLAG_EMPTY_ORIGINAL, CellState::NO_FORMULA);
    }
    let is_consol = is_consolidating(model, sheet_idx, &coord);
    if is_consol {
        return ensure_cell(cells, values, lookup, sheet_idx, coord,
                          0.0, 0.0, FLAG_EMPTY_ORIGINAL, CellState::NO_FORMULA);
    }
    UNRESOLVED
}

pub(crate) fn enumerate_initial_cells(model: &Model) -> (Vec<EvalCell>, Vec<f64>, Vec<FxHashMap<CoordKey, u32>>) {
    let n_sheets = model.sheets.len();
    // Estimate total cells for pre-allocation
    let total_input: usize = model.sheets.iter().map(|s| s.cells.len()).sum();
    let mut cells = Vec::with_capacity(total_input * 2);
    let mut values = Vec::with_capacity(total_input * 2);
    let mut lookup: Vec<FxHashMap<CoordKey, u32>> = (0..n_sheets)
        .map(|si| FxHashMap::with_capacity_and_hasher(
            model.sheets[si].cells.len() * 2,
            Default::default(),
        ))
        .collect();

    // 1. Input cells
    for (si, sheet) in model.sheets.iter().enumerate() {
        for (&coord, &cs) in &sheet.cells {
            ensure_cell(&mut cells, &mut values, &mut lookup, si, coord,
                       cs.value, cs.original_value, cs.flags, cs.formula_id);
        }
    }

    // 2. Leaf combo cells (Phase 3 equivalent)
    for si in 0..n_sheets {
        enumerate_leaf_combos(model, si, &mut cells, &mut values, &mut lookup);
    }

    // 3. Period consolidation cells (Phase 4 equivalent)
    for si in 0..n_sheets {
        enumerate_period_consolidation(model, si, &mut cells, &mut values, &mut lookup);
    }

    // 4. Non-period consolidation cells (Phase 5 equivalent)
    for si in 0..n_sheets {
        enumerate_non_period_consolidation(model, si, &mut cells, &mut values, &mut lookup);
    }

    (cells, values, lookup)
}

fn enumerate_leaf_combos(
    model: &Model, si: usize,
    cells: &mut Vec<EvalCell>, values: &mut Vec<f64>,
    lookup: &mut [FxHashMap<CoordKey, u32>],
) {
    let sheet = &model.sheets[si];
    let main_axis = match sheet.main_axis { Some(ma) => ma, None => return };
    let period_axis = sheet.period_axis;
    let n_axes = sheet.ordered_aids.len();
    let indicator_rids: Vec<u32> = sheet.rules.keys().copied().collect();

    for &ind_rid in &indicator_rids {
        let mut leaf_axes: Vec<(usize, Vec<u32>)> = Vec::new();
        for (axis_idx, _) in sheet.ordered_aids.iter().enumerate() {
            if axis_idx == main_axis || Some(axis_idx) == period_axis { continue; }
            let target_aid = sheet.ordered_aids[axis_idx];
            let mut leaves: Vec<u32> = sheet.records.iter()
                .filter(|(_, rec)| rec.analytic_id == target_aid && !sheet.children.contains_key(&rec.id))
                .map(|(&rid, _)| rid).collect();
            if leaves.is_empty() {
                leaves = sheet.records.iter()
                    .filter(|(_, rec)| rec.analytic_id == target_aid)
                    .map(|(&rid, _)| rid).collect();
            }
            leaves.sort_by_key(|rid| sheet.records.get(rid).map_or(0, |r| r.sort_order));
            leaf_axes.push((axis_idx, leaves));
        }

        let period_leaves: Vec<u32> = if let Some(pa) = period_axis {
            let target_aid = sheet.ordered_aids[pa];
            let mut leaves: Vec<u32> = sheet.records.iter()
                .filter(|(_, rec)| rec.analytic_id == target_aid && !sheet.children.contains_key(&rec.id))
                .map(|(&rid, _)| rid).collect();
            leaves.sort_by_key(|rid| sheet.records.get(rid).map_or(0, |r| r.sort_order));
            leaves
        } else {
            Vec::new()
        };
        let period_iter: Vec<u32> = if period_leaves.is_empty() { vec![0] } else { period_leaves };

        for &p_rid in &period_iter {
            let mut base = vec![0u32; n_axes];
            base[main_axis] = ind_rid;
            if let Some(pa) = period_axis { base[pa] = p_rid; }

            let mut combo_idx = vec![0usize; leaf_axes.len()];
            loop {
                for (i, &(axis_idx, ref leaves)) in leaf_axes.iter().enumerate() {
                    if combo_idx[i] < leaves.len() {
                        base[axis_idx] = leaves[combo_idx[i]];
                    }
                }
                let coord = CoordKey::new(&base[..n_axes]);
                // Add all leaf combos — resolve_all_refs will determine if they have rules
                ensure_cell(cells, values, lookup, si, coord,
                           0.0, 0.0, FLAG_EMPTY_ORIGINAL, CellState::NO_FORMULA);

                if leaf_axes.is_empty() { break; }
                let mut carry = true;
                for i in (0..leaf_axes.len()).rev() {
                    if carry {
                        combo_idx[i] += 1;
                        if combo_idx[i] >= leaf_axes[i].1.len() {
                            combo_idx[i] = 0;
                        } else {
                            carry = false;
                        }
                    }
                }
                if carry { break; }
            }
        }
    }
}

fn enumerate_period_consolidation(
    model: &Model, si: usize,
    cells: &mut Vec<EvalCell>, values: &mut Vec<f64>,
    lookup: &mut [FxHashMap<CoordKey, u32>],
) {
    let sheet = &model.sheets[si];
    let period_axis = match sheet.period_axis { Some(pa) => pa, None => return };
    let main_axis = match sheet.main_axis { Some(ma) => ma, None => return };
    let n_axes = sheet.ordered_aids.len();
    let period_aid = sheet.ordered_aids[period_axis];
    let main_aid_val = sheet.ordered_aids[main_axis];

    let parent_period_rids: Vec<u32> = sheet.children.keys()
        .filter(|&&rid| sheet.records.get(&rid).map_or(false, |r| r.analytic_id == period_aid))
        .copied().collect();
    if parent_period_rids.is_empty() { return; }

    let mut indicator_rids = Vec::new();
    for rids in sheet.name_to_rids.get(main_axis).map(|m| m.values()).into_iter().flatten() {
        indicator_rids.extend(rids.iter().copied());
    }
    for (&rid, rec) in &sheet.records {
        if rec.analytic_id == main_aid_val && !indicator_rids.contains(&rid) {
            indicator_rids.push(rid);
        }
    }

    let mut other_axes_rids = Vec::new();
    for (axis_idx, &aid) in sheet.ordered_aids.iter().enumerate() {
        if axis_idx == period_axis || axis_idx == main_axis { continue; }
        let mut rids: Vec<u32> = sheet.records.iter()
            .filter(|(_, rec)| rec.analytic_id == aid)
            .map(|(&rid, _)| rid).collect();
        rids.sort_by_key(|rid| sheet.records.get(rid).map_or(0, |r| r.sort_order));
        if !rids.is_empty() { other_axes_rids.push((axis_idx, rids)); }
    }

    let other_combos = if other_axes_rids.is_empty() {
        vec![vec![]]
    } else {
        cartesian_product_u32(&other_axes_rids.iter().map(|(_, r)| r.clone()).collect::<Vec<_>>())
    };

    for &p_rid in &parent_period_rids {
        for &ind_rid in &indicator_rids {
            for combo in &other_combos {
                let mut parts = vec![0u32; n_axes];
                parts[period_axis] = p_rid;
                parts[main_axis] = ind_rid;
                for (ci, &(axis_idx, _)) in other_axes_rids.iter().enumerate() {
                    parts[axis_idx] = combo[ci];
                }
                let coord = CoordKey::new(&parts);
                ensure_cell(cells, values, lookup, si, coord,
                           0.0, 0.0, FLAG_EMPTY_ORIGINAL, CellState::NO_FORMULA);
            }
        }
    }
}

fn enumerate_non_period_consolidation(
    model: &Model, si: usize,
    cells: &mut Vec<EvalCell>, values: &mut Vec<f64>,
    lookup: &mut [FxHashMap<CoordKey, u32>],
) {
    let sheet = &model.sheets[si];
    let main_axis = match sheet.main_axis { Some(ma) => ma, None => return };
    let period_axis = sheet.period_axis;
    let n_axes = sheet.ordered_aids.len();

    let mut consol_axes: Vec<(usize, Vec<u32>)> = Vec::new();
    for (axis_idx, &aid) in sheet.ordered_aids.iter().enumerate() {
        if Some(axis_idx) == period_axis { continue; }
        let parent_rids: Vec<u32> = sheet.children.keys()
            .filter(|&&rid| sheet.records.get(&rid).map_or(false, |r| r.analytic_id == aid))
            .copied().collect();
        if !parent_rids.is_empty() { consol_axes.push((axis_idx, parent_rids)); }
    }
    if consol_axes.is_empty() { return; }

    let main_aid_val = sheet.ordered_aids[main_axis];
    let mut indicator_rids: Vec<u32> = sheet.records.iter()
        .filter(|(_, rec)| rec.analytic_id == main_aid_val)
        .map(|(&rid, _)| rid).collect();
    indicator_rids.sort_by_key(|rid| sheet.records.get(rid).map_or(0, |r| r.sort_order));

    let all_period_rids: Vec<u32> = if let Some(pa) = period_axis {
        let period_aid = sheet.ordered_aids[pa];
        let mut prids: Vec<u32> = sheet.records.iter()
            .filter(|(_, rec)| rec.analytic_id == period_aid)
            .map(|(&rid, _)| rid).collect();
        prids.sort_by_key(|rid| sheet.records.get(rid).map_or(0, |r| r.sort_order));
        prids
    } else { vec![] };

    let mut all_axes_rids: Vec<(usize, Vec<u32>)> = Vec::new();
    for (axis_idx, &aid) in sheet.ordered_aids.iter().enumerate() {
        if axis_idx == main_axis || Some(axis_idx) == period_axis { continue; }
        let mut rids: Vec<u32> = sheet.records.iter()
            .filter(|(_, rec)| rec.analytic_id == aid)
            .map(|(&rid, _)| rid).collect();
        rids.sort_by_key(|rid| sheet.records.get(rid).map_or(0, |r| r.sort_order));
        if !rids.is_empty() { all_axes_rids.push((axis_idx, rids)); }
    }

    let period_iter: &[u32] = if all_period_rids.is_empty() { &[0] } else { &all_period_rids };

    for (consol_axis_idx, parent_rids) in &consol_axes {
        let is_main_axis = *consol_axis_idx == main_axis;

        let other_axes: Vec<(usize, &Vec<u32>)> = all_axes_rids.iter()
            .filter(|(ai, _)| *ai != *consol_axis_idx)
            .map(|(ai, rids)| (*ai, rids)).collect();

        let other_combos = if other_axes.is_empty() {
            vec![vec![]]
        } else {
            cartesian_product_u32(&other_axes.iter().map(|(_, r)| (*r).clone()).collect::<Vec<_>>())
        };

        for &p_rid in period_iter {
            if is_main_axis {
                for &parent_rid in parent_rids {
                    for combo in &other_combos {
                        let mut parts = vec![0u32; n_axes];
                        parts[main_axis] = parent_rid;
                        if let Some(pa) = period_axis { parts[pa] = p_rid; }
                        for (ci, &(ai, _)) in other_axes.iter().enumerate() {
                            parts[ai] = combo[ci];
                        }
                        let coord = CoordKey::new(&parts);

                        // has_child check (matches v2 lines 1180-1191)
                        if !lookup[si].contains_key(&coord) {
                            if let Some(children) = sheet.children.get(&parent_rid) {
                                let has_child = children.iter().any(|&crid| {
                                    let child_coord = coord.with_axis(main_axis, crid);
                                    lookup[si].contains_key(&child_coord)
                                });
                                if !has_child { continue; }
                            } else {
                                continue;
                            }
                        }
                        ensure_cell(cells, values, lookup, si, coord,
                                   0.0, 0.0, FLAG_EMPTY_ORIGINAL, CellState::NO_FORMULA);
                    }
                }
            } else {
                for &ind_rid in &indicator_rids {
                    for &parent_rid in parent_rids {
                        for combo in &other_combos {
                            let mut parts = vec![0u32; n_axes];
                            parts[main_axis] = ind_rid;
                            parts[*consol_axis_idx] = parent_rid;
                            if let Some(pa) = period_axis { parts[pa] = p_rid; }
                            for (ci, &(ai, _)) in other_axes.iter().enumerate() {
                                parts[ai] = combo[ci];
                            }
                            let coord = CoordKey::new(&parts);
                            ensure_cell(cells, values, lookup, si, coord,
                                       0.0, 0.0, FLAG_EMPTY_ORIGINAL, CellState::NO_FORMULA);
                        }
                    }
                }
            }
        }
    }
}

// ── Phase B: Pre-resolve references ────────────────────────────────────

pub(crate) fn resolve_all_refs(
    model: &Model,
    cells: &mut Vec<EvalCell>,
    values: &mut Vec<f64>,
    lookup: &mut Vec<FxHashMap<CoordKey, u32>>,
) {
    // First pass: quickly mark pure manual cells to avoid expensive resolution
    for idx in 0..cells.len() {
        if cells[idx].flags & FLAG_MANUAL != 0 && cells[idx].formula_id_hint == CellState::NO_FORMULA {
            cells[idx].eval_kind = EvalKind::Manual;
        }
    }

    // Second pass: resolve remaining cells (may create new cells)
    let mut idx = 0;
    while idx < cells.len() {
        if matches!(cells[idx].eval_kind, EvalKind::Pending) {
            resolve_single_cell(model, idx, cells, values, lookup);
        }
        idx += 1;
    }
}

fn resolve_single_cell(
    model: &Model,
    idx: usize,
    cells: &mut Vec<EvalCell>,
    values: &mut Vec<f64>,
    lookup: &mut Vec<FxHashMap<CoordKey, u32>>,
) {
    let si = cells[idx].sheet_idx as usize;
    let coord = cells[idx].coord;
    let cell_flags = cells[idx].flags;
    let formula_id_hint = cells[idx].formula_id_hint;

    // Step 1: Determine formula_id (per-cell formula or indicator rule)
    let formula_id = if cell_flags & FLAG_HAS_FORMULA != 0 && formula_id_hint != CellState::NO_FORMULA {
        Some(formula_id_hint)
    } else if cell_flags & FLAG_MANUAL != 0 {
        None
    } else {
        resolve_indicator_formula_id(model, si, &coord)
    };

    if let Some(fid) = formula_id {
        let cf = &model.compiled_formulas[fid as usize];
        if cf.instrs.is_empty() {
            // Special formula (AVERAGE/LAST) — treat as consolidation if consolidating
            if is_consolidating(model, si, &coord) {
                let child_ids = expand_children_to_ids(model, si, &coord, cells, values, lookup);
                cells[idx].eval_kind = EvalKind::ConsolidationSum { children: child_ids };
            } else {
                cells[idx].eval_kind = EvalKind::Manual;
            }
            return;
        }

        // Pre-resolve all references in the formula
        let mut resolved_refs = Vec::with_capacity(cf.refs.len());
        let mut has_unresolved_cross = false;

        for cref in &cf.refs {
            let cell_id = resolve_ref_to_id(model, si, &coord, cref, cells, values, lookup);
            resolved_refs.push(cell_id);

            // Check for unresolved cross-sheet ref (v2 lines 221-238)
            if cref.sheet_id.is_some() && cell_id == UNRESOLVED {
                if let Some(&target_name) = cref.sheet_id.as_ref() {
                    if let Some(&tsi) = model.sheet_name_to_idx.get(&target_name) {
                        let ts = &model.sheets[tsi];
                        let mut found = false;
                        for (ai, nm) in ts.name_to_rids.iter().enumerate() {
                            if Some(ai) == ts.period_axis { continue; }
                            if nm.contains_key(&cref.name_lower) {
                                found = true;
                                break;
                            }
                        }
                        if !found { has_unresolved_cross = true; }
                    }
                }
            }
        }

        if has_unresolved_cross {
            cells[idx].flags |= FLAG_UNRESOLVED;
        }

        cells[idx].eval_kind = EvalKind::Formula { formula_id: fid, resolved_refs };
        return;
    }

    // No formula — default SUM consolidation for any parent-coord cell that has
    // children with actual data. We deliberately do NOT skip when the coord wasn't
    // in the original input: synthesized parent cells created during fan-out (new
    // department/version added) need to consolidate the leaves on their slice too.
    // The original `is_manual` check below still preserves user-typed values.
    let is_manual = cell_flags & FLAG_MANUAL != 0;
    if !is_manual && is_consolidating(model, si, &coord) {
        let child_ids = expand_children_to_ids(model, si, &coord, cells, values, lookup);
        cells[idx].eval_kind = EvalKind::ConsolidationSum { children: child_ids };
        return;
    }

    cells[idx].eval_kind = EvalKind::Manual;
}

fn expand_children_to_ids(
    model: &Model, sheet_idx: usize, coord: &CoordKey,
    _cells: &mut Vec<EvalCell>, _values: &mut Vec<f64>,
    lookup: &mut [FxHashMap<CoordKey, u32>],
) -> Vec<u32> {
    let sheet = &model.sheets[sheet_idx];

    // Check axes in priority order: period, main, then others
    let n = coord.len as usize;
    let mut axes_buf = [0usize; 8];
    let mut ai = 0;
    if let Some(pa) = sheet.period_axis { axes_buf[ai] = pa; ai += 1; }
    if let Some(ma) = sheet.main_axis {
        if Some(ma) != sheet.period_axis { axes_buf[ai] = ma; ai += 1; }
    }
    for i in 0..n {
        let already = (0..ai).any(|j| axes_buf[j] == i);
        if !already { axes_buf[ai] = i; ai += 1; }
    }

    for &axis_idx in &axes_buf[..ai] {
        let rid = coord.get(axis_idx);
        if let Some(children) = sheet.children.get(&rid) {
            if !children.is_empty() {
                let mut result = Vec::with_capacity(children.len());
                for &child_rid in children {
                    let child_coord = coord.with_axis(axis_idx, child_rid);
                    // Only reference existing cells — missing children contribute 0.0
                    // (their value slot defaults to 0.0 in the values vec)
                    if let Some(&id) = lookup[sheet_idx].get(&child_coord) {
                        result.push(id);
                    }
                    // If child doesn't exist, we skip it — contributes 0 to the sum
                }
                return result;
            }
        }
    }
    Vec::new()
}

// ── Chain disambiguation helpers ────────────────────────────────────────

/// Verify that `rid`'s ancestor chain on `sheet` matches `parent_lower_chain`
/// (top-most ancestor first). Returns true if every parent in the chain
/// matches the corresponding ancestor record's name_lower (intern id).
fn ancestor_chain_matches(
    sheet: &SheetData,
    rid: u32,
    parent_lower_ids: &[u32],
) -> bool {
    if parent_lower_ids.is_empty() {
        return true;
    }
    let mut cur = match sheet.records.get(&rid) {
        Some(r) => r,
        None => return false,
    };
    // Walk parents bottom-up: child's immediate parent must match the LAST
    // entry in parent_lower_ids, then up the chain.
    for expected_id in parent_lower_ids.iter().rev() {
        let pid = match cur.parent_id {
            Some(p) => p,
            None => return false,
        };
        let prec = match sheet.records.get(&pid) {
            Some(r) => r,
            None => return false,
        };
        if prec.name_lower != *expected_id {
            return false;
        }
        cur = prec;
    }
    true
}

/// Resolve a possibly chain-qualified name on a single sheet.
/// Returns (rid, axis_idx) on success.
///
/// Behavior:
///   - If name has no \x1f: look up directly. Single match → ok. Multiple
///     matches WITHOUT a parent qualifier → caller-defined fallback.
///   - If name has \x1f: split into [a,b,...,leaf]. Look up leaf, filter
///     candidates by walking parent_id chain to verify each ancestor
///     name_lower matches. Single survivor → resolved.
fn resolve_chain_on_sheet(
    model: &Model,
    sheet_idx: usize,
    name_id: u32,
) -> Option<(u32, usize)> {
    let sheet = &model.sheets[sheet_idx];
    let period_axis = sheet.period_axis;
    let name_str = model.interner.get_str(name_id);

    if !name_str.contains('\x1f') {
        // Plain name — single-match only; multi-match handled by caller.
        for (axis_idx, name_map) in sheet.name_to_rids.iter().enumerate() {
            if Some(axis_idx) == period_axis { continue; }
            if let Some(rids) = name_map.get(&name_id) {
                if rids.len() == 1 {
                    return Some((rids[0], axis_idx));
                }
                return None; // ambiguous — caller decides
            }
        }
        return None;
    }

    let parts: Vec<&str> = name_str.split('\x1f').collect();
    let leaf_lower = parts.last().unwrap().trim().to_lowercase();
    let leaf_id = model.interner.get_id(&leaf_lower)?;
    let parent_lower_ids: Vec<u32> = parts[..parts.len() - 1]
        .iter()
        .map(|p| {
            let lower = p.trim().to_lowercase();
            model.interner.get_id(&lower).unwrap_or(u32::MAX)
        })
        .collect();
    if parent_lower_ids.iter().any(|&id| id == u32::MAX) {
        return None;
    }

    for (axis_idx, name_map) in sheet.name_to_rids.iter().enumerate() {
        if Some(axis_idx) == period_axis { continue; }
        if let Some(rids) = name_map.get(&leaf_id) {
            let mut hits: Vec<u32> = Vec::new();
            for &rid in rids {
                if ancestor_chain_matches(sheet, rid, &parent_lower_ids) {
                    hits.push(rid);
                }
            }
            if hits.len() == 1 {
                return Some((hits[0], axis_idx));
            }
            return None; // 0 or multiple — caller decides
        }
    }
    None
}

// ── Reference resolution (returns cell_id instead of f64) ──────────────

fn resolve_ref_to_id(
    model: &Model, sheet_idx: usize, coord: &CoordKey, cref: &CompiledRef,
    cells: &mut Vec<EvalCell>, values: &mut Vec<f64>,
    lookup: &mut [FxHashMap<CoordKey, u32>],
) -> u32 {
    if cref.sheet_id.is_some() {
        return resolve_cross_sheet_ref_to_id(model, sheet_idx, coord, cref, cells, values, lookup);
    }
    resolve_local_ref_to_id(model, sheet_idx, coord, cref, cells, values, lookup)
}

fn resolve_local_ref_to_id(
    model: &Model, sheet_idx: usize, coord: &CoordKey, cref: &CompiledRef,
    cells: &mut Vec<EvalCell>, values: &mut Vec<f64>,
    lookup: &mut [FxHashMap<CoordKey, u32>],
) -> u32 {
    let name_id = cref.name_lower;
    let sheet = &model.sheets[sheet_idx];
    let period_axis = sheet.period_axis;

    let mut target_rid: Option<u32> = None;
    let mut target_axis: Option<usize> = None;

    // Try chain-aware resolution first (handles [a][b][c]... by walking parent_id chain).
    if let Some((rid, axis)) = resolve_chain_on_sheet(model, sheet_idx, name_id) {
        target_rid = Some(rid);
        target_axis = Some(axis);
    }

    // Fallback: plain unqualified ambiguous lookup (multiple matches → disambiguate by parent dim).
    if target_rid.is_none() {
        let name_str = model.interner.get_str(name_id).to_string();
        if !name_str.contains('\x1f') {
            for (axis_idx, name_map) in sheet.name_to_rids.iter().enumerate() {
                if Some(axis_idx) == period_axis { continue; }
                if let Some(rids) = name_map.get(&name_id) {
                    if rids.len() >= 1 {
                        target_rid = Some(disambiguate(model, sheet_idx, axis_idx, rids, coord, cref));
                        target_axis = Some(axis_idx);
                    }
                    break;
                }
            }
        }
    }

    let target_rid = match target_rid {
        Some(r) => r,
        None => return UNRESOLVED,
    };
    let target_axis = target_axis.unwrap();

    let mut target_coord = *coord;
    target_coord.rids[target_axis] = target_rid;

    if apply_params_on_sheet(model, sheet_idx, &mut target_coord, &cref.params).is_none() {
        return UNRESOLVED;
    }

    if target_coord == *coord {
        return SELF_REF_SENTINEL;
    }

    ensure_cell_for_ref(model, sheet_idx, target_coord, cells, values, lookup)
}

fn resolve_cross_sheet_ref_to_id(
    model: &Model, src_sheet_idx: usize, coord: &CoordKey, cref: &CompiledRef,
    cells: &mut Vec<EvalCell>, values: &mut Vec<f64>,
    lookup: &mut [FxHashMap<CoordKey, u32>],
) -> u32 {
    let target_sheet_name = match cref.sheet_id { Some(n) => n, None => return UNRESOLVED };
    let &target_sheet_idx = match model.sheet_name_to_idx.get(&target_sheet_name) {
        Some(idx) => idx, None => return UNRESOLVED,
    };

    let name_id = cref.name_lower;
    let target_sheet = &model.sheets[target_sheet_idx];
    let mut target_rid: Option<u32> = None;
    let mut target_axis: Option<usize> = None;

    // Chain-aware resolution (single match after walking parent chain) —
    // works for plain names (single match) AND multi-level [a][b][c] paths.
    if let Some((rid, axis)) = resolve_chain_on_sheet(model, target_sheet_idx, name_id) {
        target_rid = Some(rid);
        target_axis = Some(axis);
    }

    // Fallback: plain unqualified ambiguous lookup → first rid (legacy behavior).
    // With always-full-path import this branch should rarely trigger; remains
    // as a safety net for hand-edited or legacy formulas.
    if target_rid.is_none() {
        let name_str = model.interner.get_str(name_id).to_string();
        if !name_str.contains('\x1f') {
            for (axis_idx, name_map) in target_sheet.name_to_rids.iter().enumerate() {
                if Some(axis_idx) == target_sheet.period_axis { continue; }
                if let Some(rids) = name_map.get(&name_id) {
                    if !rids.is_empty() {
                        target_rid = Some(rids[0]);
                        target_axis = Some(axis_idx);
                    }
                    break;
                }
            }
        }
    }

    let target_rid = match target_rid { Some(r) => r, None => return UNRESOLVED };
    let target_axis = target_axis.unwrap();

    let target_ordered = &model.sheets[target_sheet_idx].ordered_aids;
    let target_n_axes = target_ordered.len();
    let mut target_parts = vec![0u32; target_n_axes];
    target_parts[target_axis] = target_rid;

    // Period translation
    let src_sheet = &model.sheets[src_sheet_idx];
    if let (Some(src_period_axis), Some(tgt_period_axis)) =
        (src_sheet.period_axis, model.sheets[target_sheet_idx].period_axis)
    {
        let src_period_rid = coord.get(src_period_axis);
        if let Some(tpr) = translate_period(model, src_sheet_idx, target_sheet_idx, src_period_rid) {
            target_parts[tgt_period_axis] = tpr;
        }
    }

    // Fill missing axes with first leaf
    let target_sheet = &model.sheets[target_sheet_idx];
    for (axis_idx, _) in target_sheet.ordered_aids.iter().enumerate() {
        if axis_idx == target_axis { continue; }
        if Some(axis_idx) == target_sheet.period_axis { continue; }
        if target_parts[axis_idx] != 0 { continue; }
        if let Some(first) = find_first_leaf_for_axis(model, target_sheet_idx, axis_idx) {
            target_parts[axis_idx] = first;
        }
    }

    let mut tc = CoordKey::new(&target_parts);
    if apply_params_on_sheet(model, target_sheet_idx, &mut tc, &cref.params).is_none() {
        return UNRESOLVED;
    }

    ensure_cell_for_ref(model, target_sheet_idx, tc, cells, values, lookup)
}

// ── Phase C: Topological sort (Kahn's algorithm) ───────────────────────

pub(crate) fn topological_sort(cells: &[EvalCell]) -> Vec<Vec<u32>> {
    let n = cells.len();
    let mut in_degree = vec![0u32; n];

    // First pass: count edges to size the flat dependents array
    let mut total_edges: usize = 0;
    for cell in cells.iter() {
        let deps = match &cell.eval_kind {
            EvalKind::Manual | EvalKind::Pending => continue,
            EvalKind::Formula { resolved_refs, .. } => resolved_refs.as_slice(),
            EvalKind::ConsolidationSum { children } => children.as_slice(),
        };
        for &dep_id in deps {
            if dep_id < n as u32 { total_edges += 1; }
        }
    }

    // Build CSR (compressed sparse row) for reverse edges: dep → [dependents...]
    let mut dep_count = vec![0u32; n]; // how many dependents each cell has
    for (i, cell) in cells.iter().enumerate() {
        let deps = match &cell.eval_kind {
            EvalKind::Manual | EvalKind::Pending => continue,
            EvalKind::Formula { resolved_refs, .. } => resolved_refs.as_slice(),
            EvalKind::ConsolidationSum { children } => children.as_slice(),
        };
        for &dep_id in deps {
            if dep_id < n as u32 {
                in_degree[i] += 1;
                dep_count[dep_id as usize] += 1;
            }
        }
    }

    // Build offsets
    let mut offsets = vec![0u32; n + 1];
    for i in 0..n {
        offsets[i + 1] = offsets[i] + dep_count[i];
    }
    let mut flat_deps = vec![0u32; total_edges];
    let mut write_pos = offsets[..n].to_vec(); // current write position per cell

    for (i, cell) in cells.iter().enumerate() {
        let deps = match &cell.eval_kind {
            EvalKind::Manual | EvalKind::Pending => continue,
            EvalKind::Formula { resolved_refs, .. } => resolved_refs.as_slice(),
            EvalKind::ConsolidationSum { children } => children.as_slice(),
        };
        for &dep_id in deps {
            if dep_id < n as u32 {
                let pos = write_pos[dep_id as usize] as usize;
                flat_deps[pos] = i as u32;
                write_pos[dep_id as usize] += 1;
            }
        }
    }

    // Kahn's algorithm with levels
    let mut levels: Vec<Vec<u32>> = Vec::new();
    let mut queue: Vec<u32> = Vec::with_capacity(n / 2);

    for i in 0..n {
        if in_degree[i] == 0 { queue.push(i as u32); }
    }

    while !queue.is_empty() {
        let current_level = std::mem::take(&mut queue);
        for &cell_id in &current_level {
            let ci = cell_id as usize;
            let start = offsets[ci] as usize;
            let end = offsets[ci + 1] as usize;
            for &dependent in &flat_deps[start..end] {
                let di = dependent as usize;
                in_degree[di] -= 1;
                if in_degree[di] == 0 { queue.push(dependent); }
            }
        }
        levels.push(current_level);
    }

    // Remaining cells are in cycles
    let cycle_cells: Vec<u32> = (0..n as u32)
        .filter(|&i| in_degree[i as usize] > 0)
        .collect();
    if !cycle_cells.is_empty() {
        levels.push(cycle_cells);
    }

    levels
}

/// Like topological_sort but also returns the reverse-edge CSR (offsets, flat)
/// for forward traversal (mark_dirty / incremental eval).
pub(crate) fn topological_sort_with_reverse(cells: &[EvalCell]) -> (Vec<Vec<u32>>, Vec<u32>, Vec<u32>) {
    let n = cells.len();
    let mut in_degree = vec![0u32; n];

    let mut total_edges: usize = 0;
    for cell in cells.iter() {
        let deps = match &cell.eval_kind {
            EvalKind::Manual | EvalKind::Pending => continue,
            EvalKind::Formula { resolved_refs, .. } => resolved_refs.as_slice(),
            EvalKind::ConsolidationSum { children } => children.as_slice(),
        };
        for &dep_id in deps {
            if dep_id < n as u32 { total_edges += 1; }
        }
    }

    // Build CSR for reverse edges: dep → [dependents...]
    let mut dep_count = vec![0u32; n];
    for (i, cell) in cells.iter().enumerate() {
        let deps = match &cell.eval_kind {
            EvalKind::Manual | EvalKind::Pending => continue,
            EvalKind::Formula { resolved_refs, .. } => resolved_refs.as_slice(),
            EvalKind::ConsolidationSum { children } => children.as_slice(),
        };
        for &dep_id in deps {
            if dep_id < n as u32 {
                in_degree[i] += 1;
                dep_count[dep_id as usize] += 1;
            }
        }
    }

    let mut offsets = vec![0u32; n + 1];
    for i in 0..n {
        offsets[i + 1] = offsets[i] + dep_count[i];
    }
    let mut flat_deps = vec![0u32; total_edges];
    let mut write_pos = offsets[..n].to_vec();

    for (i, cell) in cells.iter().enumerate() {
        let deps = match &cell.eval_kind {
            EvalKind::Manual | EvalKind::Pending => continue,
            EvalKind::Formula { resolved_refs, .. } => resolved_refs.as_slice(),
            EvalKind::ConsolidationSum { children } => children.as_slice(),
        };
        for &dep_id in deps {
            if dep_id < n as u32 {
                let pos = write_pos[dep_id as usize] as usize;
                flat_deps[pos] = i as u32;
                write_pos[dep_id as usize] += 1;
            }
        }
    }

    // Kahn's algorithm with levels
    let mut levels: Vec<Vec<u32>> = Vec::new();
    let mut queue: Vec<u32> = Vec::with_capacity(n / 2);

    for i in 0..n {
        if in_degree[i] == 0 { queue.push(i as u32); }
    }

    while !queue.is_empty() {
        let current_level = std::mem::take(&mut queue);
        for &cell_id in &current_level {
            let ci = cell_id as usize;
            let start = offsets[ci] as usize;
            let end = offsets[ci + 1] as usize;
            for &dependent in &flat_deps[start..end] {
                let di = dependent as usize;
                in_degree[di] -= 1;
                if in_degree[di] == 0 { queue.push(dependent); }
            }
        }
        levels.push(current_level);
    }

    let cycle_cells: Vec<u32> = (0..n as u32)
        .filter(|&i| in_degree[i as usize] > 0)
        .collect();
    if !cycle_cells.is_empty() {
        levels.push(cycle_cells);
    }

    (levels, offsets, flat_deps)
}

// ── Phase D: Parallel evaluation ───────────────────────────────────────

pub(crate) fn evaluate_parallel(
    cells: &[EvalCell],
    levels: &[Vec<u32>],
    values: &mut Vec<f64>,
    formulas: &[CompiledFormula],
) {
    // Store as usize to allow safe Send+Sync capture in rayon closures.
    // Safety: we only write to unique indices within each level, and rayon
    // synchronizes between levels.
    let base = values.as_mut_ptr() as usize;

    for level in levels {
        if level.len() < 256 {
            // Small level — sequential is faster (no rayon overhead)
            for &cell_id in level {
                eval_cell_inline(cells, cell_id, base, formulas);
            }
        } else {
            // Large level — parallel with rayon
            level.par_iter().for_each(|&cell_id| {
                eval_cell_inline(cells, cell_id, base, formulas);
            });
        }
    }
}

#[inline]
pub(crate) fn eval_cell_inline(cells: &[EvalCell], cell_id: u32, base: usize, formulas: &[CompiledFormula]) {
    let cell = &cells[cell_id as usize];
    let ptr = base as *mut f64;
    match &cell.eval_kind {
        EvalKind::Manual | EvalKind::Pending => {}
        EvalKind::Formula { formula_id, resolved_refs } => {
            let val = eval_formula_v3(
                &formulas[*formula_id as usize].instrs,
                resolved_refs,
                ptr as *const f64,
            );
            unsafe { *ptr.add(cell_id as usize) = val; }
        }
        EvalKind::ConsolidationSum { children } => {
            let mut sum = 0.0f64;
            let rp = ptr as *const f64;
            for &child_id in children {
                let v = unsafe { *rp.add(child_id as usize) };
                if v.is_finite() { sum += v; }
            }
            let val = if sum.is_finite() { sum } else { 0.0 };
            unsafe { *ptr.add(cell_id as usize) = val; }
        }
    }
}

// ── Formula VM (same as v2 but PushRef is a Vec index) ─────────────────

fn eval_formula_v3(
    instrs: &[Instr],
    resolved_refs: &[u32],
    values_ptr: *const f64,
) -> f64 {
    let mut stack = [0.0f64; 256];
    let mut sp: usize = 0;

    for instr in instrs {
        match *instr {
            Instr::PushNum(n) => {
                stack[sp] = n;
                sp += 1;
            }
            Instr::PushRef(ref_idx) => {
                let cell_id = resolved_refs[ref_idx as usize];
                let val = if cell_id == SELF_REF_SENTINEL {
                    0.0
                } else if cell_id == UNRESOLVED {
                    f64::NAN
                } else {
                    unsafe { *values_ptr.add(cell_id as usize) }
                };
                stack[sp] = val;
                sp += 1;
            }
            Instr::Add => {
                if sp >= 2 {
                    sp -= 1;
                    let a = if stack[sp - 1].is_nan() { 0.0 } else { stack[sp - 1] };
                    let b = if stack[sp].is_nan() { 0.0 } else { stack[sp] };
                    stack[sp - 1] = a + b;
                }
            }
            Instr::Sub => {
                if sp >= 2 {
                    sp -= 1;
                    let a = if stack[sp - 1].is_nan() { 0.0 } else { stack[sp - 1] };
                    let b = if stack[sp].is_nan() { 0.0 } else { stack[sp] };
                    stack[sp - 1] = a - b;
                }
            }
            Instr::Mul => {
                if sp >= 2 {
                    sp -= 1;
                    let a = if stack[sp - 1].is_nan() { 0.0 } else { stack[sp - 1] };
                    let b = if stack[sp].is_nan() { 0.0 } else { stack[sp] };
                    stack[sp - 1] = a * b;
                }
            }
            Instr::Div => {
                if sp >= 2 {
                    sp -= 1;
                    let a = if stack[sp - 1].is_nan() { 0.0 } else { stack[sp - 1] };
                    let b = if stack[sp].is_nan() { 0.0 } else { stack[sp] };
                    stack[sp - 1] = if b != 0.0 { a / b } else { f64::NAN };
                }
            }
            Instr::Neg => {
                if sp >= 1 {
                    let v = if stack[sp - 1].is_nan() { 0.0 } else { stack[sp - 1] };
                    stack[sp - 1] = -v;
                }
            }
            Instr::CmpLt => {
                if sp >= 2 {
                    sp -= 1;
                    let a = if stack[sp - 1].is_nan() { 0.0 } else { stack[sp - 1] };
                    let b = if stack[sp].is_nan() { 0.0 } else { stack[sp] };
                    stack[sp - 1] = if a < b { 1.0 } else { 0.0 };
                }
            }
            Instr::CmpGt => {
                if sp >= 2 {
                    sp -= 1;
                    let a = if stack[sp - 1].is_nan() { 0.0 } else { stack[sp - 1] };
                    let b = if stack[sp].is_nan() { 0.0 } else { stack[sp] };
                    stack[sp - 1] = if a > b { 1.0 } else { 0.0 };
                }
            }
            Instr::CmpLe => {
                if sp >= 2 {
                    sp -= 1;
                    let a = if stack[sp - 1].is_nan() { 0.0 } else { stack[sp - 1] };
                    let b = if stack[sp].is_nan() { 0.0 } else { stack[sp] };
                    stack[sp - 1] = if a <= b { 1.0 } else { 0.0 };
                }
            }
            Instr::CmpGe => {
                if sp >= 2 {
                    sp -= 1;
                    let a = if stack[sp - 1].is_nan() { 0.0 } else { stack[sp - 1] };
                    let b = if stack[sp].is_nan() { 0.0 } else { stack[sp] };
                    stack[sp - 1] = if a >= b { 1.0 } else { 0.0 };
                }
            }
            Instr::CmpEq => {
                if sp >= 2 {
                    sp -= 1;
                    let a = if stack[sp - 1].is_nan() { 0.0 } else { stack[sp - 1] };
                    let b = if stack[sp].is_nan() { 0.0 } else { stack[sp] };
                    stack[sp - 1] = if a == b { 1.0 } else { 0.0 };
                }
            }
            Instr::CmpNe => {
                if sp >= 2 {
                    sp -= 1;
                    let a = if stack[sp - 1].is_nan() { 0.0 } else { stack[sp - 1] };
                    let b = if stack[sp].is_nan() { 0.0 } else { stack[sp] };
                    stack[sp - 1] = if a != b { 1.0 } else { 0.0 };
                }
            }
            Instr::CallIf => {
                if sp >= 3 {
                    let false_val = if stack[sp - 1].is_nan() { 0.0 } else { stack[sp - 1] };
                    let true_val = if stack[sp - 2].is_nan() { 0.0 } else { stack[sp - 2] };
                    let cond = if stack[sp - 3].is_nan() { 0.0 } else { stack[sp - 3] };
                    sp -= 3;
                    stack[sp] = if cond != 0.0 { true_val } else { false_val };
                    sp += 1;
                }
            }
            Instr::CallSum(n) => {
                let n = n as usize;
                if sp >= n {
                    let sum: f64 = stack[sp - n..sp].iter()
                        .map(|&v| if v.is_nan() { 0.0 } else { v })
                        .sum();
                    sp -= n;
                    stack[sp] = sum;
                    sp += 1;
                }
            }
            Instr::CallAverage(n) => {
                let n = n as usize;
                if sp >= n && n > 0 {
                    let mut sum = 0.0f64;
                    let mut count = 0usize;
                    for &v in &stack[sp - n..sp] {
                        if !v.is_nan() {
                            sum += v;
                            count += 1;
                        }
                    }
                    sp -= n;
                    stack[sp] = if count > 0 { sum / count as f64 } else { 0.0 };
                    sp += 1;
                }
            }
            Instr::CallMin(n) => {
                let n = n as usize;
                if sp >= n && n > 0 {
                    let min = stack[sp - n..sp].iter().copied()
                        .map(|v| if v.is_nan() { 0.0 } else { v })
                        .fold(f64::INFINITY, f64::min);
                    sp -= n;
                    stack[sp] = min;
                    sp += 1;
                }
            }
            Instr::CallMax(n) => {
                let n = n as usize;
                if sp >= n && n > 0 {
                    let max = stack[sp - n..sp].iter().copied()
                        .map(|v| if v.is_nan() { 0.0 } else { v })
                        .fold(f64::NEG_INFINITY, f64::max);
                    sp -= n;
                    stack[sp] = max;
                    sp += 1;
                }
            }
            Instr::CallAbs => {
                if sp >= 1 {
                    let v = if stack[sp - 1].is_nan() { 0.0 } else { stack[sp - 1] };
                    stack[sp - 1] = v.abs();
                }
            }
            Instr::CallInt => {
                if sp >= 1 {
                    let v = if stack[sp - 1].is_nan() { 0.0 } else { stack[sp - 1] };
                    // Excel INT() snaps values within 5e-14 of the next integer upward
                    // to avoid off-by-one errors from floating-point representation.
                    // E.g. 110 * (6/11) in f64 = 59.9999999999999929... => should be 60.
                    let ceil_v = v.ceil();
                    let snapped = if ceil_v - v > 0.0 && ceil_v - v < 5e-14 { ceil_v } else { v };
                    stack[sp - 1] = snapped.floor();
                }
            }
            Instr::CallRound(n) => {
                let n = n as usize;
                if n >= 2 && sp >= 2 {
                    // ROUND(value, decimals)
                    let decimals = stack[sp - 1] as i32;
                    let v = if stack[sp - 2].is_nan() { 0.0 } else { stack[sp - 2] };
                    let factor = 10f64.powi(decimals);
                    stack[sp - 2] = (v * factor).round() / factor;
                    sp -= 1;
                } else if sp >= 1 {
                    // ROUND(value) — round to integer
                    let v = if stack[sp - 1].is_nan() { 0.0 } else { stack[sp - 1] };
                    stack[sp - 1] = v.round();
                }
            }
        }
    }

    let result = if sp > 0 {
        let v = stack[sp - 1];
        if v.is_nan() { 0.0 } else { v }
    } else { 0.0 };

    if result.is_finite() { result } else { 0.0 }
}

// ── Phase E: Collect changes ───────────────────────────────────────────

pub(crate) fn collect_changes(
    cells: &[EvalCell],
    values: &[f64],
    model: &Model,
) -> HashMap<String, HashMap<String, String>> {
    // Pre-cache sheet ID strings
    let sheet_id_strs: Vec<String> = model.sheets.iter()
        .map(|s| model.interner.get_str(s.id).to_string())
        .collect();

    // Pre-cache interner strings as a slice for fast coord_key_to_string
    let interner = &model.interner;

    // Parallel collection by sheet
    let n_sheets = model.sheets.len();
    let mut per_sheet: Vec<Vec<usize>> = vec![Vec::new(); n_sheets];

    // Quick filter: only cells that changed — avoids string formatting for non-changes
    for (cell_id, cell) in cells.iter().enumerate() {
        if cell.flags & FLAG_UNRESOLVED != 0 { continue; }

        let is_manual = cell.flags & FLAG_MANUAL != 0;
        let is_computed = !matches!(cell.eval_kind, EvalKind::Manual | EvalKind::Pending) || is_manual;
        if !is_computed && !is_manual { continue; }

        let value = values[cell_id];
        let is_change = if cell.flags & FLAG_EMPTY_ORIGINAL != 0 {
            true
        } else {
            !vals_equal_f64(cell.original_value, value)
        };
        if !is_change { continue; }

        per_sheet[cell.sheet_idx as usize].push(cell_id);
    }

    // Parallel string formatting per sheet
    let sheet_results: Vec<(usize, HashMap<String, String>)> = per_sheet.into_par_iter()
        .enumerate()
        .filter(|(_, ids)| !ids.is_empty())
        .map(|(si, ids)| {
            let mut sheet_map = HashMap::with_capacity(ids.len());
            for cell_id in ids {
                let cell = &cells[cell_id];
                let value = values[cell_id];
                let ck_str = coord_key_to_string(interner, &cell.coord);
                let val_str = format_result(value);
                sheet_map.insert(ck_str, val_str);
            }
            (si, sheet_map)
        })
        .collect();

    let mut result: HashMap<String, HashMap<String, String>> = HashMap::with_capacity(n_sheets);
    for (si, sheet_map) in sheet_results {
        result.insert(sheet_id_strs[si].clone(), sheet_map);
    }
    result
}

// ── Helper functions (ported from v2) ──────────────────────────────────

fn resolve_indicator_formula_id(model: &Model, sheet_idx: usize, coord: &CoordKey) -> Option<u32> {
    let sheet = &model.sheets[sheet_idx];
    let main_axis = sheet.main_axis?;
    let indicator_rid = coord.get(main_axis);
    let rules = sheet.rules.get(&indicator_rid)?;
    if rules.is_empty() { return None; }

    let is_consol = is_consolidating(model, sheet_idx, coord);

    let mut best_scoped: Option<(i64, u32)> = None;
    for r in rules {
        if r.kind != RuleKind::Scoped { continue; }
        let mut matches = true;
        for &(scope_aid, ref scope_rids) in &r.scope {
            if let Some(&axis_idx) = sheet.aid_to_axis.get(&scope_aid) {
                let cell_rid = coord.get(axis_idx);
                if !scope_rids.contains(&cell_rid) {
                    matches = false;
                    break;
                }
            }
        }
        if matches {
            if best_scoped.is_none() || r.priority > best_scoped.unwrap().0 {
                best_scoped = Some((r.priority, r.formula_id));
            }
        }
    }
    if let Some((_, fid)) = best_scoped { return Some(fid); }

    for r in rules {
        if is_consol && r.kind == RuleKind::Consolidation { return Some(r.formula_id); }
        if !is_consol && r.kind == RuleKind::Leaf { return Some(r.formula_id); }
    }
    None
}

fn is_consolidating(model: &Model, sheet_idx: usize, coord: &CoordKey) -> bool {
    let sheet = &model.sheets[sheet_idx];
    let n = coord.len as usize;
    for i in 0..n {
        if sheet.children.contains_key(&coord.rids[i]) { return true; }
    }
    false
}

fn disambiguate(model: &Model, sheet_idx: usize, axis_idx: usize,
                candidates: &[u32], coord: &CoordKey, cref: &CompiledRef) -> u32 {
    let sheet = &model.sheets[sheet_idx];

    if let Some(row) = cref.row_hint {
        for &rid in candidates {
            if let Some(rec) = sheet.records.get(&rid) {
                if rec.excel_row == Some(row) { return rid; }
            }
        }
    }

    if let Some(parent_name) = cref.parent_hint {
        for &rid in candidates {
            if let Some(rec) = sheet.records.get(&rid) {
                if let Some(pid) = rec.parent_id {
                    if let Some(parent_rec) = sheet.records.get(&pid) {
                        if parent_rec.name_lower == parent_name { return rid; }
                    }
                }
            }
        }
    }

    let current_rid = coord.get(axis_idx);
    if let Some(rec) = sheet.records.get(&current_rid) {
        if let Some(pid) = rec.parent_id {
            if candidates.contains(&pid) { return pid; }
            if let Some(prec) = sheet.records.get(&pid) {
                if let Some(ppid) = prec.parent_id {
                    if candidates.contains(&ppid) { return ppid; }
                }
            }
        }
        for &rid in candidates {
            if let Some(crec) = sheet.records.get(&rid) {
                if crec.parent_id == Some(current_rid) { return rid; }
            }
        }
    }

    let mut best = candidates[0];
    let mut best_sort = i64::MAX;
    for &rid in candidates {
        if let Some(rec) = sheet.records.get(&rid) {
            if rec.sort_order < best_sort {
                best_sort = rec.sort_order;
                best = rid;
            }
        }
    }
    best
}

fn apply_params_on_sheet(model: &Model, sheet_idx: usize, coord: &mut CoordKey,
                         params: &[(u32, ParamValue)]) -> Option<()> {
    for &(param_name_id, ref pv) in params {
        let sheet = &model.sheets[sheet_idx];
        let aid = match sheet.analytic_name_to_aid.get(&param_name_id) {
            Some(&aid) => aid,
            None => {
                let param_str = model.interner.get_str(param_name_id);
                let mut found_aid = None;
                for (&aname_id, &aid) in &sheet.analytic_name_to_aid {
                    let aname_str = model.interner.get_str(aname_id);
                    if aname_str.contains(param_str) {
                        found_aid = Some(aid);
                        break;
                    }
                }
                match found_aid {
                    Some(aid) => aid,
                    None => continue,
                }
            }
        };
        let axis_idx = match sheet.aid_to_axis.get(&aid) {
            Some(&idx) => idx,
            None => continue,
        };

        match pv {
            ParamValue::Previous => {
                let cur = coord.get(axis_idx);
                let prev = *model.prev_period.get(&cur)?;
                coord.rids[axis_idx] = prev;
            }
            ParamValue::Back(n) => {
                let mut cur = coord.get(axis_idx);
                for _ in 0..*n {
                    cur = *model.prev_period.get(&cur)?;
                }
                coord.rids[axis_idx] = cur;
            }
            ParamValue::Forward(n) => {
                let mut cur = coord.get(axis_idx);
                for _ in 0..*n {
                    cur = *model.next_period.get(&cur)?;
                }
                coord.rids[axis_idx] = cur;
            }
            ParamValue::RecordName(name_id) => {
                let sheet = &model.sheets[sheet_idx];
                if let Some(rids) = sheet.name_to_rids.get(axis_idx).and_then(|m| m.get(name_id)) {
                    if let Some(&rid) = rids.first() {
                        coord.rids[axis_idx] = rid;
                    }
                } else if sheet.period_axis == Some(axis_idx) {
                    if let Some(&rid) = sheet.period_key_to_rid.get(name_id) {
                        coord.rids[axis_idx] = rid;
                    } else {
                        return None;
                    }
                }
            }
        }
    }
    Some(())
}

fn translate_period(model: &Model, src_sheet_idx: usize, tgt_sheet_idx: usize,
                    src_period_rid: u32) -> Option<u32> {
    let src_sheet = &model.sheets[src_sheet_idx];
    let tgt_sheet = &model.sheets[tgt_sheet_idx];

    let pk = src_sheet.rid_to_period_key.get(&src_period_rid)?;
    if let Some(&tgt_rid) = tgt_sheet.period_key_to_rid.get(pk) {
        return Some(tgt_rid);
    }

    let pk_str = model.interner.get_str(*pk);
    if pk_str.contains('-') {
        if let Some(year_part) = pk_str.split('-').next() {
            if let Some(year_id) = model.interner.get_id(year_part) {
                if let Some(&tgt_rid) = tgt_sheet.period_key_to_rid.get(&year_id) {
                    return Some(tgt_rid);
                }
            }
        }
    }
    None
}

fn find_first_leaf_for_axis(model: &Model, sheet_idx: usize, axis_idx: usize) -> Option<u32> {
    let sheet = &model.sheets[sheet_idx];
    let aid = sheet.ordered_aids.get(axis_idx)?;
    let mut leaf = None;
    for (&rid, rec) in &sheet.records {
        if rec.analytic_id == *aid {
            if !sheet.children.contains_key(&rid) {
                if leaf.is_none() || rec.sort_order < sheet.records.get(leaf.as_ref().unwrap()).map_or(i64::MAX, |r| r.sort_order) {
                    leaf = Some(rid);
                }
            }
        }
    }
    leaf
}

fn cartesian_product_u32(axes: &[Vec<u32>]) -> Vec<Vec<u32>> {
    if axes.is_empty() { return vec![vec![]]; }
    let mut result = vec![vec![]];
    for axis in axes {
        let mut new_result = Vec::with_capacity(result.len() * axis.len());
        for combo in &result {
            for &rid in axis {
                let mut new_combo = combo.clone();
                new_combo.push(rid);
                new_result.push(new_combo);
            }
        }
        result = new_result;
    }
    result
}

pub(crate) fn coord_key_to_string(interner: &Interner, coord: &CoordKey) -> String {
    let n = coord.len as usize;
    let mut parts = Vec::with_capacity(n);
    for i in 0..n {
        parts.push(interner.get_str(coord.rids[i]));
    }
    parts.join("|")
}

pub(crate) fn format_result(value: f64) -> String {
    if value == 0.0 {
        "0".to_string()
    } else {
        let rounded = (value * 1_000_000.0).round() / 1_000_000.0;
        format!("{}", rounded)
    }
}

pub(crate) fn vals_equal_f64(a: f64, b: f64) -> bool {
    if a == b { return true; }
    if a == 0.0 && b == 0.0 { return true; }
    if a.abs() > 1e-9 {
        return ((a - b) / a).abs() < 1e-6;
    }
    false
}

#[inline]
fn round6(v: f64) -> f64 {
    (v * 1_000_000.0).round() / 1_000_000.0
}
