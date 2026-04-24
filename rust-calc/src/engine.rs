use crate::evaluator::evaluate;
use crate::parser::parse_ref;
use crate::resolver::{resolve_local, resolve_cross_sheet};
use crate::types::{ModelInput, RecordInput, RuleInput};
use std::collections::{HashMap, HashSet};

/// Runtime metadata for a sheet (built from SheetInput).
pub struct SheetMeta {
    pub ordered_aids: Vec<String>,
    pub period_aid: Option<String>,
    pub main_aid: Option<String>,
    pub analytic_name_to_id: HashMap<String, String>,
    pub name_to_rids: HashMap<String, HashMap<String, Vec<String>>>,
    pub records: HashMap<String, RecordInput>,
    pub children_by_rid: HashMap<String, Vec<String>>,
    pub rules_by_indicator: HashMap<String, Vec<RuleInput>>,
    pub name: String,
    pub rid_to_period_key: HashMap<String, String>,
    pub period_key_to_rid: HashMap<String, String>,
}

/// Main calculation entry point. Takes a ModelInput and returns
/// {sheet_id: {coord_key: new_value_str}} — only cells that changed.
pub fn calculate_model(input: &ModelInput) -> HashMap<String, HashMap<String, String>> {
    // Build sheet metadata index
    let mut all_meta: HashMap<String, SheetMeta> = HashMap::new();
    let mut global_cells: HashMap<(String, String), String> = HashMap::new();
    let mut global_formulas: HashMap<(String, String), String> = HashMap::new();
    let mut phantom_cells: HashSet<(String, String)> = HashSet::new();
    let mut manual_cells: HashSet<(String, String)> = HashSet::new();
    let empty_formula_cells: HashSet<(String, String)>;

    for sheet in &input.sheets {
        let meta = SheetMeta {
            ordered_aids: sheet.ordered_aids.clone(),
            period_aid: sheet.period_aid.clone(),
            main_aid: sheet.main_aid.clone(),
            analytic_name_to_id: sheet.analytic_name_to_id.clone(),
            name_to_rids: sheet.name_to_rids.clone(),
            records: sheet.records.clone(),
            children_by_rid: sheet.children_by_rid.clone(),
            rules_by_indicator: sheet.rules_by_indicator.clone(),
            name: sheet.name.clone(),
            rid_to_period_key: sheet.rid_to_period_key.clone(),
            period_key_to_rid: sheet.period_key_to_rid.clone(),
        };

        for (ck, cell) in &sheet.cells {
            let gk = (sheet.id.clone(), ck.clone());
            global_cells.insert(gk.clone(), cell.value.clone());
            if cell.rule == "formula" && !cell.formula.is_empty() {
                // Skip raw Excel formulas (starting with =)
                if !cell.formula.starts_with('=') {
                    global_formulas.insert(gk.clone(), cell.formula.clone());
                }
            } else if cell.rule == "formula" && cell.formula.is_empty() {
                let val = &cell.value;
                if val.is_empty() || val == "0" {
                    phantom_cells.insert(gk.clone());
                }
            }
        }

        all_meta.insert(sheet.id.clone(), meta);
    }

    // Track manual cells
    for gk in global_cells.keys() {
        if !global_formulas.contains_key(gk) {
            manual_cells.insert(gk.clone());
        }
    }
    empty_formula_cells = phantom_cells.clone();

    let original_cell_keys: HashSet<(String, String)> = global_cells.keys().cloned().collect();
    let original_values: HashMap<(String, String), String> = global_cells.clone();

    // Pre-filter: remove formulas with unresolvable cross-sheet refs
    let mut skipped_formulas: HashSet<(String, String)> = HashSet::new();
    for (gk, formula) in &global_formulas {
        if !formula.contains("::") {
            continue;
        }
        // Simple check: find all [xxx::yyy] patterns
        let mut has_bad_ref = false;
        let mut pos = 0;
        let chars: Vec<char> = formula.chars().collect();
        while pos < chars.len() {
            if chars[pos] == '[' {
                // Find matching ]
                let start = pos;
                let mut depth = 1;
                pos += 1;
                while pos < chars.len() && depth > 0 {
                    match chars[pos] {
                        '[' => depth += 1,
                        ']' => depth -= 1,
                        _ => {}
                    }
                    pos += 1;
                }
                let ref_str: String = chars[start..pos].iter().collect();
                if ref_str.contains("::") {
                    let parsed = parse_ref(&ref_str);
                    if let Some(sheet_name) = &parsed.sheet {
                        let target_sid = input.sheet_name_to_id.get(&sheet_name.to_lowercase());
                        if target_sid.is_none() {
                            has_bad_ref = true;
                            break;
                        }
                        let target_sid = target_sid.unwrap();
                        let target_meta = all_meta.get(target_sid);
                        if target_meta.is_none() {
                            has_bad_ref = true;
                            break;
                        }
                        let target_meta = target_meta.unwrap();
                        let name_lower = parsed.name.to_lowercase();
                        let mut found = false;
                        for (aid, nmap) in &target_meta.name_to_rids {
                            if Some(aid) == target_meta.period_aid.as_ref() {
                                continue;
                            }
                            if nmap.get(&name_lower).map_or(false, |v| !v.is_empty()) {
                                found = true;
                                break;
                            }
                        }
                        if !found {
                            has_bad_ref = true;
                            break;
                        }
                    }
                }
                continue;
            }
            pos += 1;
        }
        if has_bad_ref {
            skipped_formulas.insert(gk.clone());
        }
    }
    for gk in &skipped_formulas {
        global_formulas.remove(gk);
    }

    // Computation state
    let mut computed_set: HashSet<(String, String)> = skipped_formulas.clone();
    let mut computing_set: HashSet<(String, String)> = HashSet::new();
    let mut computed_formulas: HashMap<(String, String), String> = HashMap::new();
    let mut unresolved: HashSet<(String, String)> = HashSet::new();

    // We need to use a recursive get_cell approach. Since Rust doesn't allow
    // easy recursive closures with mutable state, we'll use a struct.
    let mut engine = EngineState {
        global_cells: &mut global_cells,
        global_formulas: &global_formulas,
        manual_cells: &manual_cells,
        empty_formula_cells: &empty_formula_cells,
        original_cell_keys: &original_cell_keys,
        computed_set: &mut computed_set,
        computing_set: &mut computing_set,
        computed_formulas: &mut computed_formulas,
        unresolved: &mut unresolved,
        all_meta: &all_meta,
        sheet_name_to_id: &input.sheet_name_to_id,
        prev_period: &input.prev_period,
    };

    // Evaluate all formula cells
    let formula_keys: Vec<(String, String)> = engine.global_formulas.keys().cloned().collect();
    for gk in &formula_keys {
        engine.get_cell(&gk.0, &gk.1);
    }

    // Evaluate cells where indicator rule applies (no explicit formula)
    let mut rule_driven: HashSet<(String, String)> = HashSet::new();
    let original_keys: Vec<(String, String)> = original_cell_keys.iter().cloned().collect();
    for gk in &original_keys {
        if engine.global_formulas.contains_key(gk) || skipped_formulas.contains(gk) {
            continue;
        }
        let meta = match engine.all_meta.get(&gk.0) {
            Some(m) => m,
            None => continue,
        };
        if meta.rules_by_indicator.is_empty() {
            continue;
        }
        let context = context_from_key(&gk.1, &meta.ordered_aids);
        if resolve_indicator_formula(&context, meta).is_some() {
            engine.get_cell(&gk.0, &gk.1);
            rule_driven.insert(gk.clone());
        }
    }

    // Evaluate indicator rules for ALL leaf combos
    let sheet_ids: Vec<String> = engine.all_meta.keys().cloned().collect();
    for sid in &sheet_ids {
        let meta = engine.all_meta.get(sid).unwrap();
        if meta.rules_by_indicator.is_empty() {
            continue;
        }
        let main_aid = match &meta.main_aid {
            Some(a) => a.clone(),
            None => continue,
        };
        if meta.period_aid.is_none() {
            continue;
        }

        let indicators_with_rules: HashSet<String> = meta.rules_by_indicator.keys().cloned().collect();
        if indicators_with_rules.is_empty() {
            continue;
        }

        // Collect leaf records for each non-main axis
        let mut leaf_rids_by_aid: HashMap<String, Vec<String>> = HashMap::new();
        for aid in &meta.ordered_aids {
            if *aid == main_aid {
                continue;
            }
            let mut all_rids = Vec::new();
            if let Some(nmap) = meta.name_to_rids.get(aid) {
                for rids_list in nmap.values() {
                    all_rids.extend(rids_list.iter().cloned());
                }
            }
            let leaves: Vec<String> = all_rids.iter()
                .filter(|r| !meta.children_by_rid.contains_key(*r))
                .cloned()
                .collect();
            leaf_rids_by_aid.insert(aid.clone(), if leaves.is_empty() { all_rids } else { leaves });
        }

        let axes_order: Vec<String> = meta.ordered_aids.iter()
            .filter(|a| **a != main_aid)
            .cloned()
            .collect();
        if axes_order.is_empty() {
            continue;
        }
        let axes_rids: Vec<&Vec<String>> = axes_order.iter()
            .map(|a| leaf_rids_by_aid.get(a).unwrap())
            .collect();
        if axes_rids.iter().any(|v| v.is_empty()) {
            continue;
        }

        // Generate all combos
        let combos = cartesian_product(&axes_rids);

        for ind_rid in &indicators_with_rules {
            for combo in &combos {
                let mut parts = Vec::new();
                let mut ci = 0;
                for aid in &meta.ordered_aids {
                    if *aid == main_aid {
                        parts.push(ind_rid.clone());
                    } else {
                        parts.push(combo[ci].clone());
                        ci += 1;
                    }
                }
                let ck = parts.join("|");
                let gk = (sid.clone(), ck.clone());
                if engine.computed_set.contains(&gk) {
                    continue;
                }
                let context = context_from_key(&ck, &meta.ordered_aids);
                if resolve_indicator_formula(&context, meta).is_some() {
                    engine.get_cell(sid, &ck);
                    if engine.computed_set.contains(&gk) {
                        rule_driven.insert(gk);
                    }
                }
            }
        }
    }

    // Period consolidation
    let mut consol_computed: HashSet<(String, String)> = HashSet::new();
    for sid in &sheet_ids {
        let meta = engine.all_meta.get(sid).unwrap();
        let main_aid = match &meta.main_aid { Some(a) => a, None => continue };
        let period_aid = match &meta.period_aid { Some(a) => a, None => continue };

        // Collect indicator RIDs
        let mut ind_rids: HashSet<String> = HashSet::new();
        if let Some(nmap) = meta.name_to_rids.get(main_aid) {
            for rids_list in nmap.values() {
                ind_rids.extend(rids_list.iter().cloned());
            }
        }
        if ind_rids.is_empty() {
            continue;
        }

        // Parent period records
        let parent_period_rids: Vec<String> = meta.children_by_rid.iter()
            .filter(|(rid, ch)| {
                !ch.is_empty() && meta.records.get(*rid).map_or(false, |r| r.analytic_id == *period_aid)
            })
            .map(|(rid, _)| rid.clone())
            .collect();

        // Other axes
        let mut other_axes_rids: Vec<Vec<String>> = Vec::new();
        for aid in &meta.ordered_aids {
            if aid == main_aid || aid == period_aid {
                continue;
            }
            let mut rids = Vec::new();
            if let Some(nmap) = meta.name_to_rids.get(aid) {
                for rids_list in nmap.values() {
                    rids.extend(rids_list.iter().cloned());
                }
            }
            if !rids.is_empty() {
                other_axes_rids.push(rids);
            }
        }
        let other_combos = if other_axes_rids.is_empty() {
            vec![vec![]]
        } else {
            cartesian_product(&other_axes_rids.iter().collect::<Vec<_>>())
        };

        for prec_id in &parent_period_rids {
            for ind_id in &ind_rids {
                for other_vals in &other_combos {
                    let ck = build_coord_key(&meta.ordered_aids, period_aid, main_aid, prec_id, ind_id, other_vals, &meta.ordered_aids.iter().filter(|a| *a != main_aid && *a != period_aid).cloned().collect::<Vec<_>>());
                    let gk = (sid.clone(), ck.clone());
                    if !engine.computed_set.contains(&gk) {
                        engine.get_cell(sid, &ck);
                    }
                    if engine.computed_set.contains(&gk) {
                        consol_computed.insert(gk);
                    }
                }
            }
        }
    }

    // Non-period consolidation
    for sid in &sheet_ids {
        let meta = engine.all_meta.get(sid).unwrap();
        let main_aid = match &meta.main_aid { Some(a) => a, None => continue };
        let period_aid = match &meta.period_aid { Some(a) => a, None => continue };

        // Find consolidating axes (including main, excluding period)
        let mut consol_axes: Vec<(String, Vec<String>)> = Vec::new();
        for aid in &meta.ordered_aids {
            if aid == period_aid {
                continue;
            }
            let parent_rids: Vec<String> = meta.children_by_rid.iter()
                .filter(|(rid, ch)| {
                    !ch.is_empty() && meta.records.get(*rid).map_or(false, |r| r.analytic_id == *aid)
                })
                .map(|(rid, _)| rid.clone())
                .collect();
            if !parent_rids.is_empty() {
                consol_axes.push((aid.clone(), parent_rids));
            }
        }
        if consol_axes.is_empty() {
            continue;
        }

        // Collect ALL period rids
        let mut all_period_rids: Vec<String> = Vec::new();
        if let Some(nmap) = meta.name_to_rids.get(period_aid) {
            for rids_list in nmap.values() {
                all_period_rids.extend(rids_list.iter().cloned());
            }
        }

        // Collect all indicator rids
        let mut ind_rids2: HashSet<String> = HashSet::new();
        if let Some(nmap) = meta.name_to_rids.get(main_aid) {
            for rids_list in nmap.values() {
                ind_rids2.extend(rids_list.iter().cloned());
            }
        }

        for (consol_aid, parent_rids) in &consol_axes {
            // Other axes: all non-main, non-period, non-consol
            let mut other_axes_rids2: Vec<Vec<String>> = Vec::new();
            let mut other_axes_aids2: Vec<String> = Vec::new();
            for aid in &meta.ordered_aids {
                if aid == main_aid || aid == period_aid || aid == consol_aid {
                    continue;
                }
                let mut rids = Vec::new();
                if let Some(nmap) = meta.name_to_rids.get(aid) {
                    for rids_list in nmap.values() {
                        rids.extend(rids_list.iter().cloned());
                    }
                }
                if !rids.is_empty() {
                    other_axes_rids2.push(rids);
                    other_axes_aids2.push(aid.clone());
                }
            }
            let other_combos2 = if other_axes_rids2.is_empty() {
                vec![vec![]]
            } else {
                cartesian_product(&other_axes_rids2.iter().collect::<Vec<_>>())
            };

            if consol_aid == main_aid {
                // Main-axis consolidation
                for p_rid in &all_period_rids {
                    for parent_rid in parent_rids {
                        for other_vals in &other_combos2 {
                            let ck = build_coord_key_with_consol(
                                &meta.ordered_aids, period_aid, main_aid, consol_aid,
                                p_rid, parent_rid, other_vals,
                                &other_axes_aids2,
                            );
                            let gk = (sid.clone(), ck.clone());
                            // Skip phantom consolidation
                            if !engine.original_cell_keys.contains(&gk) && !engine.computed_set.contains(&gk) {
                                let child_rids = meta.children_by_rid.get(parent_rid);
                                let has_child = child_rids.map_or(false, |crs| {
                                    crs.iter().any(|crid| {
                                        let child_ck = build_coord_key_with_consol(
                                            &meta.ordered_aids, period_aid, main_aid, consol_aid,
                                            p_rid, crid, other_vals,
                                            &other_axes_aids2,
                                        );
                                        let cgk = (sid.clone(), child_ck);
                                        engine.original_cell_keys.contains(&cgk) || engine.computed_set.contains(&cgk)
                                    })
                                });
                                if !has_child {
                                    continue;
                                }
                            }
                            if !engine.computed_set.contains(&gk) {
                                engine.get_cell(sid, &ck);
                            }
                            if engine.computed_set.contains(&gk) {
                                consol_computed.insert(gk);
                            }
                        }
                    }
                }
            } else {
                // Non-main axis consolidation
                for p_rid in &all_period_rids {
                    for ind_id in &ind_rids2 {
                        for parent_rid in parent_rids {
                            for other_vals in &other_combos2 {
                                let ck = build_coord_key_4way(
                                    &meta.ordered_aids, period_aid, main_aid, consol_aid,
                                    p_rid, ind_id, parent_rid, other_vals,
                                    &other_axes_aids2,
                                );
                                let gk = (sid.clone(), ck.clone());
                                if !engine.computed_set.contains(&gk) {
                                    engine.get_cell(sid, &ck);
                                }
                                if engine.computed_set.contains(&gk) {
                                    consol_computed.insert(gk);
                                }
                            }
                        }
                    }
                }
            }
        }
    }

    // Collect changes
    let mut result: HashMap<String, HashMap<String, String>> = HashMap::new();
    let all_computed: HashSet<(String, String)> = global_formulas.keys().cloned()
        .chain(rule_driven.into_iter())
        .chain(consol_computed.into_iter())
        .collect();

    for gk in &all_computed {
        if engine.unresolved.contains(gk) {
            continue;
        }
        let (sid, ck) = gk;
        // Check __empty__
        if engine.computed_formulas.get(gk).map_or(false, |f| f == "__empty__") {
            result.entry(sid.clone()).or_default().insert(ck.clone(), "__empty__".to_string());
            continue;
        }
        let new_val = engine.global_cells.get(gk).cloned().unwrap_or_default();
        let old_val = original_values.get(gk).cloned().unwrap_or_default();
        if vals_equal(&old_val, &new_val) {
            continue;
        }
        result.entry(sid.clone()).or_default().insert(ck.clone(), new_val);
    }

    result
}

// Helper struct to hold mutable computation state
struct EngineState<'a> {
    global_cells: &'a mut HashMap<(String, String), String>,
    global_formulas: &'a HashMap<(String, String), String>,
    manual_cells: &'a HashSet<(String, String)>,
    empty_formula_cells: &'a HashSet<(String, String)>,
    original_cell_keys: &'a HashSet<(String, String)>,
    computed_set: &'a mut HashSet<(String, String)>,
    computing_set: &'a mut HashSet<(String, String)>,
    computed_formulas: &'a mut HashMap<(String, String), String>,
    unresolved: &'a mut HashSet<(String, String)>,
    all_meta: &'a HashMap<String, SheetMeta>,
    sheet_name_to_id: &'a HashMap<String, String>,
    prev_period: &'a HashMap<String, String>,
}

impl<'a> EngineState<'a> {
    fn get_cell(&mut self, sheet_id: &str, coord_key: &str) -> f64 {
        let gk = (sheet_id.to_string(), coord_key.to_string());
        if self.computed_set.contains(&gk) {
            return to_float(self.global_cells.get(&gk).map(|s| s.as_str()).unwrap_or(""));
        }
        if self.computing_set.contains(&gk) {
            return to_float(self.global_cells.get(&gk).map(|s| s.as_str()).unwrap_or(""));
        }

        let meta = match self.all_meta.get(sheet_id) {
            Some(m) => m,
            None => return to_float(self.global_cells.get(&gk).map(|s| s.as_str()).unwrap_or("")),
        };
        let context = context_from_key(coord_key, &meta.ordered_aids);

        // 1. Explicit per-cell formula
        let mut formula: Option<String> = self.global_formulas.get(&gk).cloned();

        // 2. Indicator rule (but NEVER override manual cells)
        if formula.is_none() && !self.manual_cells.contains(&gk) {
            if let Some((f, _)) = resolve_indicator_formula(&context, meta) {
                formula = Some(f);
            }
        }

        // Handle __empty__
        if formula.as_deref() == Some("__empty__") {
            self.global_cells.insert(gk.clone(), String::new());
            self.computed_set.insert(gk.clone());
            self.computed_formulas.insert(gk, "__empty__".to_string());
            return 0.0;
        }

        if let Some(ref formula_str) = formula {
            // Special consolidation keywords
            if formula_str == "AVERAGE" && is_consolidating(&context, meta) {
                self.computing_set.insert(gk.clone());
                let children = expand_children_one_level(coord_key, &context, meta);
                let total: f64 = children.iter().map(|ck| self.get_cell(sheet_id, ck)).sum();
                let result = if !children.is_empty() { total / children.len() as f64 } else { 0.0 };
                let result_str = format_result(result);
                self.global_cells.insert(gk.clone(), result_str);
                self.computed_set.insert(gk.clone());
                self.computing_set.remove(&gk);
                self.computed_formulas.insert(gk, formula_str.clone());
                return result;
            }

            if formula_str == "LAST" && is_consolidating(&context, meta) {
                self.computing_set.insert(gk.clone());
                let children = expand_children_one_level(coord_key, &context, meta);
                let result = if let Some(last) = children.last() {
                    self.get_cell(sheet_id, last)
                } else {
                    0.0
                };
                let result_str = format_result(result);
                self.global_cells.insert(gk.clone(), result_str);
                self.computed_set.insert(gk.clone());
                self.computing_set.remove(&gk);
                self.computed_formulas.insert(gk, formula_str.clone());
                return result;
            }

            self.computing_set.insert(gk.clone());
            let mut has_unresolved_ref = false;

            // Evaluate formula
            let sheet_id_owned = sheet_id.to_string();
            let coord_key_owned = coord_key.to_string();

            // We need a way for the ref evaluator to call back into get_cell.
            // Since Rust doesn't easily allow recursive mutable borrows,
            // we'll use a simple approach: collect ref values before evaluation.
            // Actually, we need true recursion. Let's use unsafe for this.
            let result = {
                let engine_ptr = self as *mut EngineState;
                let mut ref_evaluator = |ref_token: &str| -> Option<f64> {
                    let engine = unsafe { &mut *engine_ptr };
                    let ref_parsed = parse_ref(ref_token);

                    if ref_parsed.sheet.is_some() {
                        let val = resolve_cross_sheet(
                            &ref_parsed, &context, meta,
                            engine.sheet_name_to_id,
                            engine.all_meta,
                            engine.prev_period,
                            &mut |sid, ck| engine.get_cell(sid, ck),
                        );
                        if val == 0.0 {
                            // Check if reference actually resolved
                            if let Some(sname) = &ref_parsed.sheet {
                                let target_sid = engine.sheet_name_to_id.get(&sname.to_lowercase());
                                if target_sid.is_none() {
                                    has_unresolved_ref = true;
                                } else if let Some(target_meta) = target_sid.and_then(|s| engine.all_meta.get(s)) {
                                    let nl = ref_parsed.name.to_lowercase();
                                    let mut found = false;
                                    for (aid, nmap) in &target_meta.name_to_rids {
                                        if Some(aid) == target_meta.period_aid.as_ref() { continue; }
                                        if nmap.get(&nl).map_or(false, |v| !v.is_empty()) {
                                            found = true;
                                            break;
                                        }
                                    }
                                    if !found {
                                        has_unresolved_ref = true;
                                    }
                                }
                            }
                        }
                        return Some(val);
                    }

                    // Local reference
                    let rs = resolve_local(&ref_parsed, &context, meta, engine.prev_period);
                    match rs {
                        None => None,
                        Some(ref rs_key) if rs_key == &coord_key_owned => Some(0.0),
                        Some(rs_key) => {
                            let target_gk = (sheet_id_owned.clone(), rs_key.clone());
                            // Check if cell exists
                            if !engine.global_cells.contains_key(&target_gk)
                                && !engine.global_formulas.contains_key(&target_gk)
                                && !engine.manual_cells.contains(&target_gk) {
                                return None;
                            }
                            // Phantom cells
                            if engine.empty_formula_cells.contains(&target_gk)
                                && !engine.computed_set.contains(&target_gk) {
                                return None;
                            }
                            let val = engine.get_cell(&sheet_id_owned, &rs_key);
                            if engine.unresolved.contains(&target_gk) {
                                has_unresolved_ref = true;
                            }
                            Some(val)
                        }
                    }
                };

                evaluate(formula_str, &mut ref_evaluator)
            };

            let result = result.unwrap_or(0.0);

            // Division by zero in consolidation → fall through to SUM
            if !result.is_finite() && is_consolidating(&context, meta) {
                self.computing_set.remove(&gk);
                // Fall through to step 3
            } else {
                let result = if !result.is_finite() { 0.0 } else { result };
                let result_str = format_result(result);
                self.global_cells.insert(gk.clone(), result_str);
                self.computed_set.insert(gk.clone());
                self.computing_set.remove(&gk);
                self.computed_formulas.insert(gk.clone(), formula_str.clone());
                if has_unresolved_ref {
                    self.unresolved.insert(gk);
                }
                return result;
            }
        }

        // 3. Default SUM consolidation
        let mut skip_default_sum = false;
        if is_consolidating(&context, meta) && !self.original_cell_keys.contains(&gk) {
            if let Some(main) = &meta.main_aid {
                if let Some(ind_rid) = context.get(main) {
                    if meta.children_by_rid.contains_key(ind_rid) {
                        skip_default_sum = true;
                    }
                }
            }
        }

        if is_consolidating(&context, meta) && !self.manual_cells.contains(&gk) && !skip_default_sum {
            self.computing_set.insert(gk.clone());
            let children_cks = expand_children_one_level(coord_key, &context, meta);
            let total: f64 = children_cks.iter().map(|ck| self.get_cell(sheet_id, ck)).sum();
            let total_str = format_result(total);
            self.global_cells.insert(gk.clone(), total_str);
            self.computed_set.insert(gk.clone());
            self.computing_set.remove(&gk);
            return total;
        }

        // 4. Leaf manual value
        to_float(self.global_cells.get(&gk).map(|s| s.as_str()).unwrap_or(""))
    }
}

fn to_float(val: &str) -> f64 {
    val.parse::<f64>().unwrap_or(0.0)
}

fn format_result(result: f64) -> String {
    if result == 0.0 {
        "0".to_string()
    } else {
        let rounded = (result * 1_000_000.0).round() / 1_000_000.0;
        let s = format!("{}", rounded);
        s
    }
}

fn context_from_key(coord_key: &str, ordered_aids: &[String]) -> HashMap<String, String> {
    let parts: Vec<&str> = coord_key.split('|').collect();
    let mut context = HashMap::new();
    for (i, aid) in ordered_aids.iter().enumerate() {
        if i < parts.len() {
            context.insert(aid.clone(), parts[i].to_string());
        }
    }
    context
}

fn is_consolidating(context: &HashMap<String, String>, meta: &SheetMeta) -> bool {
    for (_aid, rid) in context {
        if meta.children_by_rid.contains_key(rid) {
            return true;
        }
    }
    false
}

fn expand_children_one_level(
    coord_key: &str,
    context: &HashMap<String, String>,
    meta: &SheetMeta,
) -> Vec<String> {
    // Find axes with children
    let mut axes: Vec<(String, Vec<String>)> = Vec::new();
    for aid in &meta.ordered_aids {
        if let Some(rid) = context.get(aid) {
            if let Some(children) = meta.children_by_rid.get(rid) {
                if !children.is_empty() {
                    axes.push((aid.clone(), children.clone()));
                }
            }
        }
    }
    if axes.is_empty() {
        return Vec::new();
    }

    // Pick ONE axis: period first, then main, then other
    let chosen = axes.iter()
        .find(|(aid, _)| Some(aid) == meta.period_aid.as_ref())
        .or_else(|| axes.iter().find(|(aid, _)| Some(aid) == meta.main_aid.as_ref()))
        .unwrap_or(&axes[0]);

    let (expand_aid, expand_children) = chosen;
    let _ = coord_key; // used for context only

    expand_children.iter().map(|crid| {
        let parts: Vec<String> = meta.ordered_aids.iter().map(|aid| {
            if aid == expand_aid {
                crid.clone()
            } else {
                context.get(aid).cloned().unwrap_or_default()
            }
        }).collect();
        parts.join("|")
    }).collect()
}

/// Resolve indicator formula rule: scoped → consolidation/leaf.
pub fn resolve_indicator_formula(
    context: &HashMap<String, String>,
    meta: &SheetMeta,
) -> Option<(String, String)> {
    let main = meta.main_aid.as_ref()?;
    let indicator_rid = context.get(main)?;
    let rules = meta.rules_by_indicator.get(indicator_rid)?;
    if rules.is_empty() {
        return None;
    }

    // Scoped rules
    let non_main: HashMap<&String, &String> = context.iter()
        .filter(|(a, _)| *a != main)
        .collect();
    let mut scoped_hits: Vec<&RuleInput> = Vec::new();
    for rule in rules {
        if rule.kind != "scoped" || rule.scope.is_empty() {
            continue;
        }
        let matches = rule.scope.iter().all(|(a, r)| {
            if r.is_empty() {
                return true;
            }
            let vals: Vec<&str> = r.split(',').collect();
            non_main.get(a).map_or(false, |v| vals.contains(&v.as_str()))
        });
        if matches {
            scoped_hits.push(rule);
        }
    }
    if !scoped_hits.is_empty() {
        scoped_hits.sort_by(|a, b| {
            b.priority.cmp(&a.priority)
                .then_with(|| b.scope.len().cmp(&a.scope.len()))
                .then_with(|| a.id.cmp(&b.id))
        });
        let best = scoped_hits[0];
        if !best.formula.is_empty() {
            return Some((best.formula.clone(), format!("rule:{}", best.id)));
        }
    }

    // Base leaf/consolidation
    let is_consol = is_consolidating(context, meta);
    let base_kind = if is_consol { "consolidation" } else { "leaf" };
    for rule in rules {
        if rule.kind == base_kind && !rule.formula.is_empty() {
            return Some((rule.formula.clone(), format!("rule:{}", rule.id)));
        }
    }

    None
}

fn vals_equal(a: &str, b: &str) -> bool {
    if a == b {
        return true;
    }
    let fa = match a.parse::<f64>() {
        Ok(v) => v,
        Err(_) => return false,
    };
    let fb = match b.parse::<f64>() {
        Ok(v) => v,
        Err(_) => return false,
    };
    if fa == fb {
        return true;
    }
    if fa == 0.0 && fb == 0.0 {
        return true;
    }
    if fa.abs() > 1e-9 {
        return ((fa - fb) / fa).abs() < 1e-6;
    }
    false
}

/// Cartesian product of multiple vectors of strings.
fn cartesian_product(axes: &[&Vec<String>]) -> Vec<Vec<String>> {
    if axes.is_empty() {
        return vec![vec![]];
    }
    let mut result = vec![vec![]];
    for axis in axes {
        let mut new_result = Vec::new();
        for existing in &result {
            for item in *axis {
                let mut new = existing.clone();
                new.push(item.clone());
                new_result.push(new);
            }
        }
        result = new_result;
    }
    result
}

/// Build coord_key for period consolidation (period_aid + main_aid + others).
fn build_coord_key(
    ordered_aids: &[String],
    period_aid: &str,
    main_aid: &str,
    period_rid: &str,
    ind_rid: &str,
    other_vals: &[String],
    _other_aids: &[String],
) -> String {
    let mut parts = Vec::new();
    let mut oi = 0;
    for aid in ordered_aids {
        if aid == period_aid {
            parts.push(period_rid.to_string());
        } else if aid == main_aid {
            parts.push(ind_rid.to_string());
        } else {
            if oi < other_vals.len() {
                parts.push(other_vals[oi].clone());
                oi += 1;
            }
        }
    }
    parts.join("|")
}

/// Build coord_key for consolidation with a specific consol axis.
fn build_coord_key_with_consol(
    ordered_aids: &[String],
    period_aid: &str,
    main_aid: &str,
    consol_aid: &str,
    period_rid: &str,
    consol_rid: &str,
    other_vals: &[String],
    _other_aids: &[String],
) -> String {
    let mut parts = Vec::new();
    let mut oi = 0;
    for aid in ordered_aids {
        if aid == period_aid {
            parts.push(period_rid.to_string());
        } else if aid == main_aid || aid == consol_aid {
            parts.push(consol_rid.to_string());
        } else {
            if oi < other_vals.len() {
                parts.push(other_vals[oi].clone());
                oi += 1;
            }
        }
    }
    parts.join("|")
}

/// Build coord_key for 4-way (period + main + consol + others).
fn build_coord_key_4way(
    ordered_aids: &[String],
    period_aid: &str,
    main_aid: &str,
    consol_aid: &str,
    period_rid: &str,
    ind_rid: &str,
    consol_rid: &str,
    other_vals: &[String],
    _other_aids: &[String],
) -> String {
    let mut parts = Vec::new();
    let mut oi = 0;
    for aid in ordered_aids {
        if aid == period_aid {
            parts.push(period_rid.to_string());
        } else if aid == main_aid {
            parts.push(ind_rid.to_string());
        } else if aid == consol_aid {
            parts.push(consol_rid.to_string());
        } else {
            if oi < other_vals.len() {
                parts.push(other_vals[oi].clone());
                oi += 1;
            }
        }
    }
    parts.join("|")
}
