use std::collections::HashMap;
use crate::types::ModelInput;
use crate::intern::Interner;
use crate::coord::*;
use crate::compiler::*;
use crate::builder::*;

#[cfg(feature = "debug_counters")]
mod dbg {
    use std::sync::atomic::{AtomicU64, Ordering};
    pub static DBG_REF_NAME_NOT_FOUND: AtomicU64 = AtomicU64::new(0);
    pub static DBG_REF_SELF_REF: AtomicU64 = AtomicU64::new(0);
    pub static DBG_REF_NO_CELL_NO_RULE: AtomicU64 = AtomicU64::new(0);
    pub static DBG_REF_PARAM_FAIL: AtomicU64 = AtomicU64::new(0);
    pub static DBG_REF_CROSS_SHEET_FAIL: AtomicU64 = AtomicU64::new(0);
    pub static DBG_REF_OK: AtomicU64 = AtomicU64::new(0);
    pub static DBG_BRANCH_FORMULA: AtomicU64 = AtomicU64::new(0);
    pub static DBG_BRANCH_CONSOL: AtomicU64 = AtomicU64::new(0);
    pub static DBG_BRANCH_CONSOL_NEW: AtomicU64 = AtomicU64::new(0);
    pub static DBG_BRANCH_LEAF: AtomicU64 = AtomicU64::new(0);
    pub static DBG_BRANCH_LEAF_NEW: AtomicU64 = AtomicU64::new(0);
    pub static DBG_BRANCH_SKIP_SUM: AtomicU64 = AtomicU64::new(0);
}

#[cfg(not(feature = "debug_counters"))]
macro_rules! dbg_inc { ($name:ident) => {} }
#[cfg(feature = "debug_counters")]
macro_rules! dbg_inc { ($name:ident) => { dbg::$name.fetch_add(1, std::sync::atomic::Ordering::Relaxed); } }

/// Main entry point: parse JSON, build model, calculate, return changes.
pub fn calculate_model(input: &ModelInput) -> HashMap<String, HashMap<String, String>> {
    let mut model = build_model(input);
    compute_all(&mut model);
    collect_changes(&model)
}

// ── Computation ─────────────────────────────────────────────────────────

fn compute_all(model: &mut Model) {
    // Phase 1: Evaluate all cells that have explicit formulas
    for si in 0..model.sheets.len() {
        let keys: Vec<CoordKey> = model.sheets[si].cells.iter()
            .filter(|(_, c)| c.has_formula())
            .map(|(k, _)| *k)
            .collect();
        for coord in keys {
            get_cell(model, si, coord);
        }
    }

    // Phase 2: Rule-driven cells
    for si in 0..model.sheets.len() {
        let keys: Vec<CoordKey> = model.sheets[si].cells.keys().copied().collect();
        for coord in keys {
            let cell = model.sheets[si].cells.get(&coord).copied();
            if let Some(c) = cell {
                if !c.is_computed() && !c.is_manual() && !c.has_formula() {
                    get_cell(model, si, coord);
                }
            }
        }
    }

    // Phase 3: Leaf combo generation
    for si in 0..model.sheets.len() {
        generate_leaf_combos(model, si);
    }

    // Phase 4: Period consolidation
    for si in 0..model.sheets.len() {
        period_consolidation(model, si);
    }

    // Phase 5: Non-period consolidation
    for si in 0..model.sheets.len() {
        non_period_consolidation(model, si);
    }
}

// ── get_cell: per-sheet recursive evaluator ─────────────────────────────

fn get_cell(model: &mut Model, sheet_idx: usize, coord: CoordKey) -> f64 {
    // Fast path: already computed or cycling
    if let Some(cell) = model.sheets[sheet_idx].cells.get(&coord) {
        if cell.is_computed() {
            return cell.value;
        }
        if cell.is_computing() {
            return cell.value;
        }
    }

    // Check for per-cell formula
    let formula_id = model.sheets[sheet_idx].cells.get(&coord).and_then(|c| {
        if c.has_formula() { Some(c.formula_id) } else { None }
    });

    // If no per-cell formula, check indicator rules
    let formula_id = match formula_id {
        Some(fid) => Some(fid),
        None => {
            if model.sheets[sheet_idx].cells.get(&coord).map_or(false, |c| c.is_manual()) {
                None
            } else {
                resolve_indicator_formula_id(model, sheet_idx, &coord)
            }
        }
    };

    if let Some(fid) = formula_id {
        dbg_inc!(DBG_BRANCH_FORMULA);
        let cf = &model.compiled_formulas[fid as usize];
        if cf.instrs.is_empty() {
            return handle_special_formula(model, sheet_idx, coord, fid);
        }

        // Set COMPUTING flag
        if let Some(cell) = model.sheets[sheet_idx].cells.get_mut(&coord) {
            cell.flags |= FLAG_COMPUTING;
        } else {
            model.sheets[sheet_idx].cells.insert(coord, CellState {
                value: 0.0,
                original_value: 0.0,
                formula_id: fid,
                flags: FLAG_COMPUTING | FLAG_HAS_FORMULA | FLAG_EMPTY_ORIGINAL,
            });
        }

        let (result, has_unresolved) = eval_compiled(model, sheet_idx, coord, fid);
        let result = if result.is_finite() { round6(result) } else { 0.0 };

        if let Some(cell) = model.sheets[sheet_idx].cells.get_mut(&coord) {
            cell.value = result;
            cell.flags = (cell.flags & !(FLAG_COMPUTING)) | FLAG_COMPUTED;
            if has_unresolved {
                cell.flags |= FLAG_UNRESOLVED;
            }
        }

        return result;
    }

    // Default SUM consolidation
    let mut skip_default_sum = false;
    let is_original = model.sheets[sheet_idx].original_cell_keys.contains(&coord);
    if !is_original {
        let sheet = &model.sheets[sheet_idx];
        if let Some(ma) = sheet.main_axis {
            let ind_rid = coord.get(ma);
            if sheet.children.contains_key(&ind_rid) {
                skip_default_sum = true;
                dbg_inc!(DBG_BRANCH_SKIP_SUM);
            }
        }
    }

    let is_manual = model.sheets[sheet_idx].cells.get(&coord).map_or(false, |c| c.is_manual());
    if !skip_default_sum && !is_manual && is_consolidating(model, sheet_idx, &coord) {
        dbg_inc!(DBG_BRANCH_CONSOL);
        if let Some(cell) = model.sheets[sheet_idx].cells.get_mut(&coord) {
            cell.flags |= FLAG_COMPUTING;
        } else {
            dbg_inc!(DBG_BRANCH_CONSOL_NEW);
            model.sheets[sheet_idx].cells.insert(coord, CellState {
                value: 0.0,
                original_value: 0.0,
                formula_id: CellState::NO_FORMULA,
                flags: FLAG_COMPUTING | FLAG_EMPTY_ORIGINAL,
            });
        }

        let children = expand_children(model, sheet_idx, &coord);
        let total: f64 = children.iter()
            .map(|&child_coord| get_cell(model, sheet_idx, child_coord))
            .sum();

        let total = if total.is_finite() { round6(total) } else { 0.0 };
        if let Some(cell) = model.sheets[sheet_idx].cells.get_mut(&coord) {
            cell.value = total;
            cell.flags = (cell.flags & !(FLAG_COMPUTING)) | FLAG_COMPUTED;
        }
        return total;
    }

    // Leaf manual value
    dbg_inc!(DBG_BRANCH_LEAF);
    if let Some(cell) = model.sheets[sheet_idx].cells.get_mut(&coord) {
        cell.flags |= FLAG_COMPUTED;
        cell.value
    } else {
        dbg_inc!(DBG_BRANCH_LEAF_NEW);
        model.sheets[sheet_idx].cells.insert(coord, CellState {
            value: 0.0,
            original_value: 0.0,
            formula_id: CellState::NO_FORMULA,
            flags: FLAG_COMPUTED,
        });
        0.0
    }
}

// ── Stack-based formula VM ──────────────────────────────────────────────

fn eval_compiled(model: &mut Model, sheet_idx: usize,
                 coord: CoordKey, formula_id: u32) -> (f64, bool) {
    let formula = model.compiled_formulas[formula_id as usize].clone();
    let mut stack = [0.0f64; 256];
    let mut sp: usize = 0;
    let mut has_unresolved = false;

    for instr in &formula.instrs {
        match *instr {
            Instr::PushNum(n) => {
                stack[sp] = n;
                sp += 1;
            }
            Instr::PushRef(ref_idx) => {
                let cref = &formula.refs[ref_idx as usize];
                let val = resolve_ref(model, sheet_idx, &coord, cref)
                    .unwrap_or(f64::NAN);
                if val == 0.0 && cref.sheet_id.is_some() {
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
                            if !found {
                                has_unresolved = true;
                            }
                        }
                    }
                }
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
                    let ceil_v = v.ceil();
                    let snapped = if ceil_v - v > 0.0 && ceil_v - v < 5e-14 { ceil_v } else { v };
                    stack[sp - 1] = snapped.floor();
                }
            }
            Instr::CallRound(n) => {
                let n = n as usize;
                if n >= 2 && sp >= 2 {
                    let decimals = stack[sp - 1] as i32;
                    let v = if stack[sp - 2].is_nan() { 0.0 } else { stack[sp - 2] };
                    let factor = 10f64.powi(decimals);
                    stack[sp - 2] = (v * factor).round() / factor;
                    sp -= 1;
                } else if sp >= 1 {
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
    (result, has_unresolved)
}

// ── Reference resolution (numeric) ─────────────────────────────────────

fn resolve_ref(model: &mut Model, sheet_idx: usize,
               coord: &CoordKey, cref: &CompiledRef) -> Option<f64> {
    if cref.sheet_id.is_some() {
        return resolve_cross_sheet_ref(model, sheet_idx, coord, cref);
    }
    resolve_local_ref(model, sheet_idx, coord, cref)
}

fn resolve_local_ref(model: &mut Model, sheet_idx: usize,
                     coord: &CoordKey, cref: &CompiledRef) -> Option<f64> {
    let name_id = cref.name_lower;

    let mut target_rid: Option<u32> = None;
    let mut target_axis: Option<usize> = None;

    let sheet = &model.sheets[sheet_idx];
    let period_axis = sheet.period_axis;

    for (axis_idx, name_map) in sheet.name_to_rids.iter().enumerate() {
        if Some(axis_idx) == period_axis {
            continue;
        }
        if let Some(rids) = name_map.get(&name_id) {
            if rids.len() == 1 {
                target_rid = Some(rids[0]);
                target_axis = Some(axis_idx);
            } else if !rids.is_empty() {
                target_rid = Some(disambiguate(model, sheet_idx, axis_idx, rids, coord, cref));
                target_axis = Some(axis_idx);
            }
            break;
        }
    }

    // If name not found and contains '/', try parent/child split
    if target_rid.is_none() {
        let name_str = model.interner.get_str(name_id).to_string();
        if name_str.contains('/') {
            let parts: Vec<&str> = name_str.splitn(2, '/').collect();
            let child_name = model.interner.intern_lower(parts[1].trim());
            let parent_name = model.interner.intern_lower(parts[0].trim());
            let sheet = &model.sheets[sheet_idx];
            for (axis_idx, name_map) in sheet.name_to_rids.iter().enumerate() {
                if Some(axis_idx) == period_axis {
                    continue;
                }
                if let Some(rids) = name_map.get(&child_name) {
                    if rids.len() == 1 {
                        target_rid = Some(rids[0]);
                        target_axis = Some(axis_idx);
                    } else if !rids.is_empty() {
                        let mut filtered: Vec<u32> = Vec::new();
                        for &rid in rids {
                            if let Some(rec) = sheet.records.get(&rid) {
                                if let Some(pid) = rec.parent_id {
                                    if let Some(prec) = sheet.records.get(&pid) {
                                        if prec.name_lower == parent_name {
                                            filtered.push(rid);
                                        }
                                    }
                                }
                            }
                        }
                        if filtered.len() == 1 {
                            target_rid = Some(filtered[0]);
                        } else if !filtered.is_empty() {
                            target_rid = Some(disambiguate(model, sheet_idx, axis_idx, &filtered, coord, cref));
                        } else {
                            target_rid = Some(disambiguate(model, sheet_idx, axis_idx, rids, coord, cref));
                        }
                        target_axis = Some(axis_idx);
                    }
                    break;
                }
            }
        }
    }

    let target_rid = match target_rid {
        Some(r) => r,
        None => {
            dbg_inc!(DBG_REF_NAME_NOT_FOUND);
            return None;
        }
    };
    let target_axis = target_axis.unwrap();

    let mut target_coord = *coord;
    target_coord.rids[target_axis] = target_rid;

    if apply_params(model, sheet_idx, &mut target_coord, &cref.params).is_none() {
        dbg_inc!(DBG_REF_PARAM_FAIL);
        return None;
    }

    if target_coord == *coord {
        dbg_inc!(DBG_REF_SELF_REF);
        return Some(0.0);
    }

    let exists = model.sheets[sheet_idx].cells.contains_key(&target_coord);
    if !exists {
        let has_rule = resolve_indicator_formula_id(model, sheet_idx, &target_coord).is_some();
        let is_consol = is_consolidating(model, sheet_idx, &target_coord);
        if !has_rule && !is_consol {
            dbg_inc!(DBG_REF_NO_CELL_NO_RULE);
            return None;
        }
    }

    dbg_inc!(DBG_REF_OK);
    Some(get_cell(model, sheet_idx, target_coord))
}

fn resolve_cross_sheet_ref(model: &mut Model, src_sheet_idx: usize,
                           coord: &CoordKey, cref: &CompiledRef) -> Option<f64> {
    let target_sheet_name = cref.sheet_id?;
    let &target_sheet_idx = model.sheet_name_to_idx.get(&target_sheet_name)?;

    let name_id = cref.name_lower;

    let target_sheet = &model.sheets[target_sheet_idx];
    let mut target_rid: Option<u32> = None;
    let mut target_axis: Option<usize> = None;

    for (axis_idx, name_map) in target_sheet.name_to_rids.iter().enumerate() {
        if Some(axis_idx) == target_sheet.period_axis {
            continue;
        }
        if let Some(rids) = name_map.get(&name_id) {
            if rids.len() == 1 {
                target_rid = Some(rids[0]);
                target_axis = Some(axis_idx);
            } else if !rids.is_empty() {
                target_rid = Some(rids[0]);
                target_axis = Some(axis_idx);
            }
            break;
        }
    }

    let target_rid = target_rid?;
    let target_axis = target_axis?;

    let target_ordered = &model.sheets[target_sheet_idx].ordered_aids;
    let target_n_axes = target_ordered.len();
    let mut target_parts = vec![0u32; target_n_axes];
    target_parts[target_axis] = target_rid;

    let src_sheet = &model.sheets[src_sheet_idx];
    if let (Some(src_period_axis), Some(tgt_period_axis)) =
        (src_sheet.period_axis, model.sheets[target_sheet_idx].period_axis)
    {
        let src_period_rid = coord.get(src_period_axis);
        let target_period_rid = translate_period(model, src_sheet_idx, target_sheet_idx, src_period_rid);
        if let Some(tpr) = target_period_rid {
            target_parts[tgt_period_axis] = tpr;
        }
    }

    let target_sheet = &model.sheets[target_sheet_idx];
    for (axis_idx, &_aid) in target_sheet.ordered_aids.iter().enumerate() {
        if axis_idx == target_axis { continue; }
        if Some(axis_idx) == target_sheet.period_axis { continue; }
        if target_parts[axis_idx] != 0 { continue; }
        if let Some(first) = find_first_leaf_for_axis(model, target_sheet_idx, axis_idx) {
            target_parts[axis_idx] = first;
        }
    }

    let target_coord = CoordKey::new(&target_parts);

    let mut tc = target_coord;
    apply_params_on_sheet(model, target_sheet_idx, &mut tc, &cref.params)?;

    Some(get_cell(model, target_sheet_idx, tc))
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

fn apply_params(model: &mut Model, sheet_idx: usize, coord: &mut CoordKey,
                params: &[(u32, ParamValue)]) -> Option<()> {
    apply_params_on_sheet(model, sheet_idx, coord, params)
}

fn apply_params_on_sheet(model: &mut Model, sheet_idx: usize, coord: &mut CoordKey,
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

// ── Disambiguation ──────────────────────────────────────────────────────

fn disambiguate(model: &Model, sheet_idx: usize, axis_idx: usize,
                candidates: &[u32], coord: &CoordKey, cref: &CompiledRef) -> u32 {
    let sheet = &model.sheets[sheet_idx];

    if let Some(row) = cref.row_hint {
        for &rid in candidates {
            if let Some(rec) = sheet.records.get(&rid) {
                if rec.excel_row == Some(row) {
                    return rid;
                }
            }
        }
    }

    if let Some(parent_name) = cref.parent_hint {
        for &rid in candidates {
            if let Some(rec) = sheet.records.get(&rid) {
                if let Some(pid) = rec.parent_id {
                    if let Some(parent_rec) = sheet.records.get(&pid) {
                        if parent_rec.name_lower == parent_name {
                            return rid;
                        }
                    }
                }
            }
        }
    }

    let current_rid = coord.get(axis_idx);
    if let Some(rec) = sheet.records.get(&current_rid) {
        if let Some(pid) = rec.parent_id {
            if candidates.contains(&pid) {
                return pid;
            }
            if let Some(prec) = sheet.records.get(&pid) {
                if let Some(ppid) = prec.parent_id {
                    if candidates.contains(&ppid) {
                        return ppid;
                    }
                }
            }
        }
        for &rid in candidates {
            if let Some(crec) = sheet.records.get(&rid) {
                if crec.parent_id == Some(current_rid) {
                    return rid;
                }
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

// ── Indicator formula rules ─────────────────────────────────────────────

fn resolve_indicator_formula_id(model: &Model, sheet_idx: usize, coord: &CoordKey) -> Option<u32> {
    let sheet = &model.sheets[sheet_idx];
    let main_axis = sheet.main_axis?;
    let indicator_rid = coord.get(main_axis);

    let rules = sheet.rules.get(&indicator_rid)?;
    if rules.is_empty() {
        return None;
    }

    let is_consol = is_consolidating(model, sheet_idx, coord);

    let mut best_scoped: Option<(i64, u32)> = None;
    for r in rules {
        if r.kind != RuleKind::Scoped {
            continue;
        }
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
    if let Some((_, fid)) = best_scoped {
        return Some(fid);
    }

    for r in rules {
        if is_consol && r.kind == RuleKind::Consolidation {
            return Some(r.formula_id);
        }
        if !is_consol && r.kind == RuleKind::Leaf {
            return Some(r.formula_id);
        }
    }

    None
}

// ── Consolidation helpers ───────────────────────────────────────────────

fn is_consolidating(model: &Model, sheet_idx: usize, coord: &CoordKey) -> bool {
    let sheet = &model.sheets[sheet_idx];
    for (axis_idx, _) in sheet.ordered_aids.iter().enumerate() {
        let rid = coord.get(axis_idx);
        if sheet.children.contains_key(&rid) {
            return true;
        }
    }
    false
}

fn expand_children(model: &Model, sheet_idx: usize, coord: &CoordKey) -> Vec<CoordKey> {
    let sheet = &model.sheets[sheet_idx];

    let axes_to_check: Vec<usize> = {
        let mut v = Vec::new();
        if let Some(pa) = sheet.period_axis { v.push(pa); }
        if let Some(ma) = sheet.main_axis {
            if Some(ma) != sheet.period_axis { v.push(ma); }
        }
        for i in 0..coord.len as usize {
            if !v.contains(&i) { v.push(i); }
        }
        v
    };

    for axis_idx in axes_to_check {
        let rid = coord.get(axis_idx);
        if let Some(children) = sheet.children.get(&rid) {
            if !children.is_empty() {
                return children.iter()
                    .map(|&child_rid| coord.with_axis(axis_idx, child_rid))
                    .collect();
            }
        }
    }
    Vec::new()
}

fn handle_special_formula(model: &mut Model, sheet_idx: usize,
                          coord: CoordKey, _formula_id: u32) -> f64 {
    if !is_consolidating(model, sheet_idx, &coord) {
        return model.sheets[sheet_idx].cells.get(&coord).map_or(0.0, |c| c.value);
    }

    if let Some(cell) = model.sheets[sheet_idx].cells.get_mut(&coord) {
        cell.flags |= FLAG_COMPUTING;
    }

    let children = expand_children(model, sheet_idx, &coord);
    if children.is_empty() {
        return 0.0;
    }

    let child_vals: Vec<f64> = children.iter()
        .map(|&cc| get_cell(model, sheet_idx, cc))
        .collect();

    let result = child_vals.iter().sum::<f64>();
    let result = if result.is_finite() { round6(result) } else { 0.0 };

    if let Some(cell) = model.sheets[sheet_idx].cells.get_mut(&coord) {
        cell.value = result;
        cell.flags = (cell.flags & !(FLAG_COMPUTING)) | FLAG_COMPUTED;
    }
    result
}

// ── Leaf combo generation ───────────────────────────────────────────────

fn generate_leaf_combos(model: &mut Model, sheet_idx: usize) {
    let main_axis;
    let period_axis;
    let n_axes;
    let indicator_rids: Vec<u32>;
    let mut all_leaf_axes: Vec<Vec<(usize, Vec<u32>)>> = Vec::new();
    let mut all_period_leaves: Vec<Vec<u32>> = Vec::new();

    {
        let sheet = &model.sheets[sheet_idx];
        main_axis = match sheet.main_axis {
            Some(ma) => ma,
            None => return,
        };
        period_axis = sheet.period_axis;
        n_axes = sheet.ordered_aids.len();
        indicator_rids = sheet.rules.keys().copied().collect();

        for _ind_rid in &indicator_rids {
            let mut leaf_axes: Vec<(usize, Vec<u32>)> = Vec::new();
            for (axis_idx, _) in sheet.ordered_aids.iter().enumerate() {
                if axis_idx == main_axis { continue; }
                if Some(axis_idx) == period_axis { continue; }

                let target_aid = sheet.ordered_aids[axis_idx];
                let mut leaves: Vec<u32> = sheet.records.iter()
                    .filter(|(_, rec)| rec.analytic_id == target_aid && !sheet.children.contains_key(&rec.id))
                    .map(|(&rid, _)| rid)
                    .collect();
                if leaves.is_empty() {
                    leaves = sheet.records.iter()
                        .filter(|(_, rec)| rec.analytic_id == target_aid)
                        .map(|(&rid, _)| rid)
                        .collect();
                }
                leaves.sort_by_key(|rid| sheet.records.get(rid).map_or(0, |r| r.sort_order));
                leaf_axes.push((axis_idx, leaves));
            }
            all_leaf_axes.push(leaf_axes);

            let period_leaves = if let Some(pa) = period_axis {
                let target_aid = sheet.ordered_aids[pa];
                let mut leaves: Vec<u32> = sheet.records.iter()
                    .filter(|(_, rec)| rec.analytic_id == target_aid && !sheet.children.contains_key(&rec.id))
                    .map(|(&rid, _)| rid)
                    .collect();
                leaves.sort_by_key(|rid| sheet.records.get(rid).map_or(0, |r| r.sort_order));
                leaves
            } else {
                Vec::new()
            };
            all_period_leaves.push(period_leaves);
        }
    }

    for (idx, ind_rid) in indicator_rids.iter().enumerate() {
        let ref leaf_axes = all_leaf_axes[idx];
        let ref period_leaves = all_period_leaves[idx];

        let mut base = vec![0u32; n_axes];
        base[main_axis] = *ind_rid;

        let period_iter: Vec<u32> = if period_leaves.is_empty() { vec![0] } else { period_leaves.clone() };

        for &p_rid in &period_iter {
            if let Some(pa) = period_axis {
                base[pa] = p_rid;
            }

            let mut combo_idx = vec![0usize; leaf_axes.len()];
            loop {
                for (i, &(axis_idx, ref leaves)) in leaf_axes.iter().enumerate() {
                    if combo_idx[i] < leaves.len() {
                        base[axis_idx] = leaves[combo_idx[i]];
                    }
                }

                let coord = CoordKey::new(&base[..n_axes]);
                let should_compute = if !model.sheets[sheet_idx].cells.contains_key(&coord) {
                    resolve_indicator_formula_id(model, sheet_idx, &coord).is_some()
                } else {
                    model.sheets[sheet_idx].cells.get(&coord).map_or(false, |c| !c.is_computed())
                };
                if should_compute {
                    get_cell(model, sheet_idx, coord);
                }

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

// ── Period consolidation ────────────────────────────────────────────────

fn period_consolidation(model: &mut Model, sheet_idx: usize) {
    let period_axis;
    let main_axis;
    let n_axes;
    let parent_period_rids: Vec<u32>;
    let indicator_rids: Vec<u32>;
    let other_axes_rids: Vec<(usize, Vec<u32>)>;

    {
        let sheet = &model.sheets[sheet_idx];
        period_axis = match sheet.period_axis {
            Some(pa) => pa,
            None => return,
        };
        main_axis = match sheet.main_axis {
            Some(ma) => ma,
            None => return,
        };
        n_axes = sheet.ordered_aids.len();
        let period_aid = sheet.ordered_aids[period_axis];
        let main_aid_val = sheet.ordered_aids[main_axis];

        parent_period_rids = sheet.children.keys()
            .filter(|&&rid| {
                sheet.records.get(&rid).map_or(false, |r| r.analytic_id == period_aid)
            })
            .copied()
            .collect();

        if parent_period_rids.is_empty() {
            return;
        }

        let mut inds = Vec::new();
        for rids in sheet.name_to_rids.get(main_axis).map(|m| m.values()).into_iter().flatten() {
            inds.extend(rids.iter().copied());
        }
        for (&rid, rec) in &sheet.records {
            if rec.analytic_id == main_aid_val && !inds.contains(&rid) {
                inds.push(rid);
            }
        }
        indicator_rids = inds;

        let mut other = Vec::new();
        for (axis_idx, &aid) in sheet.ordered_aids.iter().enumerate() {
            if axis_idx == period_axis || axis_idx == main_axis {
                continue;
            }
            let mut rids: Vec<u32> = sheet.records.iter()
                .filter(|(_, rec)| rec.analytic_id == aid)
                .map(|(&rid, _)| rid)
                .collect();
            rids.sort_by_key(|rid| sheet.records.get(rid).map_or(0, |r| r.sort_order));
            if !rids.is_empty() {
                other.push((axis_idx, rids));
            }
        }
        other_axes_rids = other;
    }

    let other_combos = if other_axes_rids.is_empty() {
        vec![vec![]]
    } else {
        let axes: Vec<Vec<u32>> = other_axes_rids.iter().map(|(_, rids)| rids.clone()).collect();
        cartesian_product_u32(&axes)
    };

    for &p_rid in &parent_period_rids {
        for &ind_rid in &indicator_rids {
            for combo in &other_combos {
                let mut parts = vec![0u32; n_axes];
                parts[period_axis] = p_rid;
                parts[main_axis] = ind_rid;
                let mut ci = 0;
                for &(axis_idx, _) in &other_axes_rids {
                    parts[axis_idx] = combo[ci];
                    ci += 1;
                }
                let coord = CoordKey::new(&parts);
                if !model.sheets[sheet_idx].cells.get(&coord).map_or(false, |c| c.is_computed()) {
                    get_cell(model, sheet_idx, coord);
                }
            }
        }
    }
}

// ── Non-period consolidation ────────────────────────────────────────────

fn non_period_consolidation(model: &mut Model, sheet_idx: usize) {
    let main_axis;
    let period_axis;
    let n_axes;
    let consol_axes: Vec<(usize, Vec<u32>)>;
    let indicator_rids: Vec<u32>;
    let all_period_rids: Vec<u32>;
    let all_axes_rids: Vec<(usize, Vec<u32>)>;

    {
        let sheet = &model.sheets[sheet_idx];
        main_axis = match sheet.main_axis {
            Some(ma) => ma,
            None => return,
        };
        period_axis = sheet.period_axis;
        n_axes = sheet.ordered_aids.len();

        let mut axes = Vec::new();
        for (axis_idx, &aid) in sheet.ordered_aids.iter().enumerate() {
            if Some(axis_idx) == period_axis { continue; }
            let parent_rids: Vec<u32> = sheet.children.keys()
                .filter(|&&rid| sheet.records.get(&rid).map_or(false, |r| r.analytic_id == aid))
                .copied()
                .collect();
            if !parent_rids.is_empty() {
                axes.push((axis_idx, parent_rids));
            }
        }
        consol_axes = axes;

        if consol_axes.is_empty() { return; }

        let main_aid_val = sheet.ordered_aids[main_axis];
        let mut inds: Vec<u32> = sheet.records.iter()
            .filter(|(_, rec)| rec.analytic_id == main_aid_val)
            .map(|(&rid, _)| rid)
            .collect();
        inds.sort_by_key(|rid| sheet.records.get(rid).map_or(0, |r| r.sort_order));
        indicator_rids = inds;

        all_period_rids = if let Some(pa) = period_axis {
            let period_aid = sheet.ordered_aids[pa];
            let mut prids: Vec<u32> = sheet.records.iter()
                .filter(|(_, rec)| rec.analytic_id == period_aid)
                .map(|(&rid, _)| rid)
                .collect();
            prids.sort_by_key(|rid| sheet.records.get(rid).map_or(0, |r| r.sort_order));
            prids
        } else {
            vec![]
        };

        let mut all_ax = Vec::new();
        for (axis_idx, &aid) in sheet.ordered_aids.iter().enumerate() {
            if axis_idx == main_axis { continue; }
            if Some(axis_idx) == period_axis { continue; }
            let mut rids: Vec<u32> = sheet.records.iter()
                .filter(|(_, rec)| rec.analytic_id == aid)
                .map(|(&rid, _)| rid)
                .collect();
            rids.sort_by_key(|rid| sheet.records.get(rid).map_or(0, |r| r.sort_order));
            if !rids.is_empty() {
                all_ax.push((axis_idx, rids));
            }
        }
        all_axes_rids = all_ax;
    }

    for (consol_axis_idx, parent_rids) in &consol_axes {
        let is_main_axis = *consol_axis_idx == main_axis;

        let other_axes: Vec<(usize, &Vec<u32>)> = all_axes_rids.iter()
            .filter(|(ai, _)| *ai != *consol_axis_idx)
            .map(|(ai, rids)| (*ai, rids))
            .collect();

        let other_combos = if other_axes.is_empty() {
            vec![vec![]]
        } else {
            let axes: Vec<Vec<u32>> = other_axes.iter().map(|(_, rids)| (*rids).clone()).collect();
            cartesian_product_u32(&axes)
        };

        let period_iter: &[u32] = if all_period_rids.is_empty() { &[0] } else { &all_period_rids };

        for &p_rid in period_iter {
            if is_main_axis {
                for &parent_rid in parent_rids {
                    for combo in &other_combos {
                        let mut parts = vec![0u32; n_axes];
                        parts[main_axis] = parent_rid;
                        if let Some(pa) = period_axis { parts[pa] = p_rid; }
                        let mut ci = 0;
                        for &(ai, _) in &other_axes {
                            parts[ai] = combo[ci];
                            ci += 1;
                        }
                        let coord = CoordKey::new(&parts);

                        if !model.sheets[sheet_idx].cells.contains_key(&coord) {
                            let sheet = &model.sheets[sheet_idx];
                            if let Some(children) = sheet.children.get(&parent_rid) {
                                let has_child = children.iter().any(|&crid| {
                                    let child_coord = coord.with_axis(main_axis, crid);
                                    model.sheets[sheet_idx].cells.contains_key(&child_coord)
                                });
                                if !has_child { continue; }
                            } else {
                                continue;
                            }
                        }

                        if !model.sheets[sheet_idx].cells.get(&coord).map_or(false, |c| c.is_computed()) {
                            get_cell(model, sheet_idx, coord);
                        }
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
                            let mut ci = 0;
                            for &(ai, _) in &other_axes {
                                parts[ai] = combo[ci];
                                ci += 1;
                            }
                            let coord = CoordKey::new(&parts);
                            if !model.sheets[sheet_idx].cells.get(&coord).map_or(false, |c| c.is_computed()) {
                                get_cell(model, sheet_idx, coord);
                            }
                        }
                    }
                }
            }
        }
    }
}

/// Cartesian product of u32 vectors.
fn cartesian_product_u32(axes: &[Vec<u32>]) -> Vec<Vec<u32>> {
    if axes.is_empty() {
        return vec![vec![]];
    }
    let mut result = vec![vec![]];
    for axis in axes {
        let mut new_result = Vec::new();
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

// ── Output collection ───────────────────────────────────────────────────

fn collect_changes(model: &Model) -> HashMap<String, HashMap<String, String>> {
    let mut result: HashMap<String, HashMap<String, String>> = HashMap::new();

    for sheet in &model.sheets {
        let sid_str = model.interner.get_str(sheet.id).to_string();

        for (&coord, cell) in &sheet.cells {
            if cell.flags & FLAG_UNRESOLVED != 0 {
                continue;
            }
            if !cell.is_computed() && !cell.is_manual() {
                continue;
            }

            let is_change = if cell.flags & FLAG_EMPTY_ORIGINAL != 0 {
                true
            } else {
                !vals_equal_f64(cell.original_value, cell.value)
            };
            if !is_change {
                continue;
            }

            let ck_str = coord_key_to_string(&model.interner, &coord);
            let val_str = format_result(cell.value);

            result.entry(sid_str.clone()).or_default().insert(ck_str, val_str);
        }
    }

    result
}

fn coord_key_to_string(interner: &Interner, coord: &CoordKey) -> String {
    let n = coord.len as usize;
    let mut parts = Vec::with_capacity(n);
    for i in 0..n {
        parts.push(interner.get_str(coord.rids[i]));
    }
    parts.join("|")
}

fn format_result(value: f64) -> String {
    if value == 0.0 {
        "0".to_string()
    } else {
        let rounded = (value * 1_000_000.0).round() / 1_000_000.0;
        format!("{}", rounded)
    }
}

fn vals_equal_f64(a: f64, b: f64) -> bool {
    if a == b {
        return true;
    }
    if a == 0.0 && b == 0.0 {
        return true;
    }
    if a.abs() > 1e-9 {
        return ((a - b) / a).abs() < 1e-6;
    }
    false
}

#[inline]
fn round6(v: f64) -> f64 {
    (v * 1_000_000.0).round() / 1_000_000.0
}
