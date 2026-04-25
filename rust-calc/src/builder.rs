use std::collections::{HashMap, HashSet};
use rustc_hash::{FxHashMap, FxHashSet};
use crate::types::ModelInput;
use crate::intern::Interner;
use crate::coord::{CoordKey, CellState, FLAG_EMPTY_ORIGINAL};
use crate::compiler::{self, CompiledFormula};

/// Per-record compact data.
#[derive(Clone, Debug, serde::Serialize, serde::Deserialize)]
pub struct RecordCompact {
    pub id: u32,
    pub analytic_id: u32,
    pub parent_id: Option<u32>,
    pub sort_order: i64,
    pub name_lower: u32,
    pub period_key: Option<u32>,
    pub excel_row: Option<i64>,
}

/// Rule kind.
#[derive(Clone, Debug, PartialEq, serde::Serialize, serde::Deserialize)]
pub enum RuleKind {
    Leaf,
    Consolidation,
    Scoped,
}

/// Compact indicator formula rule.
#[derive(Clone, Debug, serde::Serialize, serde::Deserialize)]
pub struct RuleCompact {
    pub id: u32,
    pub kind: RuleKind,
    pub scope: Vec<(u32, Vec<u32>)>, // (interned_aid, [interned_rid, ...])
    pub priority: i64,
    pub formula_id: u32, // index into compiled_formulas
}

/// Per-sheet metadata with numeric indices.
#[derive(Clone, serde::Serialize, serde::Deserialize)]
pub struct SheetData {
    pub id: u32,                      // interned sheet ID
    pub name_lower: u32,              // interned lowercase sheet name
    pub ordered_aids: Vec<u32>,       // interned analytic IDs in order
    pub period_axis: Option<usize>,   // index into ordered_aids
    pub main_axis: Option<usize>,     // index into ordered_aids
    /// name_to_rids[axis_idx][name_lower_id] -> Vec<rid>
    pub name_to_rids: Vec<HashMap<u32, Vec<u32>>>,
    pub records: HashMap<u32, RecordCompact>,
    pub children: HashMap<u32, Vec<u32>>,
    pub rules: HashMap<u32, Vec<RuleCompact>>, // indicator_rid -> rules
    pub rid_to_period_key: HashMap<u32, u32>,
    pub period_key_to_rid: HashMap<u32, u32>,
    pub analytic_name_to_aid: HashMap<u32, u32>, // name_lower_id -> aid
    pub aid_to_axis: HashMap<u32, usize>,        // aid -> index in ordered_aids
    pub cells: FxHashMap<CoordKey, CellState>,     // per-sheet cell storage
    pub original_cell_keys: FxHashSet<CoordKey>,   // keys present in original input
}

/// The full model with numeric indices.
#[derive(serde::Serialize, serde::Deserialize)]
pub struct Model {
    pub interner: Interner,
    pub sheets: Vec<SheetData>,
    pub sheet_id_to_idx: HashMap<u32, usize>,
    pub sheet_name_to_idx: HashMap<u32, usize>,
    pub prev_period: HashMap<u32, u32>,
    pub next_period: HashMap<u32, u32>,
    pub compiled_formulas: Vec<CompiledFormula>,
}

/// Build a Model from a deserialized ModelInput.
pub fn build_model(input: &ModelInput) -> Model {
    let mut interner = Interner::new();
    let mut sheets = Vec::new();
    let mut sheet_id_to_idx = HashMap::new();
    let mut sheet_name_to_idx = HashMap::new();
    let mut compiled_formulas: Vec<CompiledFormula> = Vec::new();
    let mut formula_cache: HashMap<String, u32> = HashMap::new();

    for (sheet_idx, s) in input.sheets.iter().enumerate() {
        let sid = interner.intern(&s.id);
        let sname_lower = interner.intern_lower(&s.name);
        sheet_id_to_idx.insert(sid, sheet_idx);
        sheet_name_to_idx.insert(sname_lower, sheet_idx);

        // Intern analytic IDs
        let ordered_aids: Vec<u32> = s.ordered_aids.iter()
            .map(|a| interner.intern(a))
            .collect();

        let period_aid = s.period_aid.as_ref().map(|a| interner.intern(a));
        let main_aid = s.main_aid.as_ref().map(|a| interner.intern(a));

        let period_axis = period_aid.and_then(|pa| ordered_aids.iter().position(|&a| a == pa));
        let main_axis = main_aid.and_then(|ma| ordered_aids.iter().position(|&a| a == ma));

        // Build records
        let mut records = HashMap::new();
        let mut children: HashMap<u32, Vec<u32>> = HashMap::new();
        for (rid_str, rec) in &s.records {
            let rid = interner.intern(rid_str);
            let aid = interner.intern(&rec.analytic_id);
            let parent = if rec.parent_id.as_deref() == Some("") { None }
                         else { rec.parent_id.as_ref().map(|p| interner.intern(p)) };
            let pk = rec.period_key.as_ref()
                .filter(|k| !k.is_empty())
                .map(|k| interner.intern(k));

            records.insert(rid, RecordCompact {
                id: rid,
                analytic_id: aid,
                parent_id: parent,
                sort_order: rec.sort_order,
                name_lower: interner.intern_lower(&rec.name),
                period_key: pk,
                excel_row: rec.excel_row,
            });

            if let Some(pid) = parent {
                children.entry(pid).or_default().push(rid);
            }
        }

        // Build name_to_rids per axis
        let mut name_to_rids: Vec<HashMap<u32, Vec<u32>>> = vec![HashMap::new(); ordered_aids.len()];
        for (name_str, nmap) in &s.name_to_rids {
            let aid = interner.intern(name_str);
            if let Some(axis_idx) = ordered_aids.iter().position(|&a| a == aid) {
                for (name, rids) in nmap {
                    let name_lower = interner.intern_lower(name);
                    let interned_rids: Vec<u32> = rids.iter().map(|r| interner.intern(r)).collect();
                    name_to_rids[axis_idx].insert(name_lower, interned_rids);
                }
            }
        }

        // Build analytic_name_to_aid and aid_to_axis
        let mut analytic_name_to_aid = HashMap::new();
        let mut aid_to_axis = HashMap::new();
        // We need analytic names from the input
        for (aname, aid_str) in &s.analytic_name_to_id {
            let aname_lower = interner.intern_lower(aname);
            let aid = interner.intern(aid_str);
            analytic_name_to_aid.insert(aname_lower, aid);
        }
        for (i, &aid) in ordered_aids.iter().enumerate() {
            aid_to_axis.insert(aid, i);
        }

        // Build rules
        let mut rules: HashMap<u32, Vec<RuleCompact>> = HashMap::new();
        for (ind_rid_str, rule_list) in &s.rules_by_indicator {
            let ind_rid = interner.intern(ind_rid_str);
            let mut compact_rules = Vec::new();
            for r in rule_list {
                let kind = match r.kind.as_str() {
                    "leaf" => RuleKind::Leaf,
                    "consolidation" => RuleKind::Consolidation,
                    _ => RuleKind::Scoped,
                };
                // scope is {aid: rid_or_empty}
                let scope: Vec<(u32, Vec<u32>)> = r.scope.iter()
                    .filter(|(_, v)| !v.is_empty())
                    .map(|(k, v)| {
                        let aid = interner.intern(k);
                        let rid = interner.intern(v);
                        (aid, vec![rid])
                    }).collect();

                // Compile the formula
                let formula_id = get_or_compile_formula(
                    &r.formula, &mut formula_cache, &mut compiled_formulas, &mut interner);

                compact_rules.push(RuleCompact {
                    id: interner.intern(&r.id),
                    kind,
                    scope,
                    priority: r.priority,
                    formula_id,
                });
            }
            rules.insert(ind_rid, compact_rules);
        }

        // Build rid_to_period_key / period_key_to_rid
        let mut rid_to_period_key = HashMap::new();
        let mut period_key_to_rid = HashMap::new();
        for (rid_str, pk_str) in &s.rid_to_period_key {
            let rid = interner.intern(rid_str);
            let pk = interner.intern(pk_str);
            rid_to_period_key.insert(rid, pk);
            period_key_to_rid.insert(pk, rid);
        }

        // Build cells (per-sheet)
        let mut sheet_cells: FxHashMap<CoordKey, CellState> = FxHashMap::default();
        for (ck_str, cell) in &s.cells {
            let parts: Vec<u32> = ck_str.split('|')
                .map(|p| interner.intern(p))
                .collect();
            let coord = CoordKey::new(&parts);

            let is_empty_original = cell.value.is_empty()
                || cell.value.parse::<f64>().is_err();
            let value = cell.value.parse::<f64>().unwrap_or(0.0);

            let mut cell_state = match cell.rule.as_str() {
                "manual" => CellState::new_manual(value),
                "phantom" => CellState::new_phantom(value),
                _ => {
                    if cell.formula.is_empty() {
                        CellState::new_manual(value)
                    } else {
                        let fid = get_or_compile_formula(
                            &cell.formula, &mut formula_cache, &mut compiled_formulas, &mut interner);
                        CellState::new_formula(value, fid)
                    }
                }
            };
            if is_empty_original {
                cell_state.flags |= FLAG_EMPTY_ORIGINAL;
            }

            sheet_cells.insert(coord, cell_state);
        }
        let sheet_original_keys: FxHashSet<CoordKey> = sheet_cells.keys().copied().collect();

        // Sort children by sort_order
        for (_, ch) in children.iter_mut() {
            ch.sort_by_key(|rid| records.get(rid).map_or(0, |r| r.sort_order));
        }

        sheets.push(SheetData {
            id: sid,
            name_lower: sname_lower,
            ordered_aids,
            period_axis,
            main_axis,
            name_to_rids,
            records,
            children,
            rules,
            rid_to_period_key,
            period_key_to_rid,
            analytic_name_to_aid,
            aid_to_axis,
            cells: sheet_cells,
            original_cell_keys: sheet_original_keys,
        });
    }

    // Build prev_period / next_period
    let mut prev_period = HashMap::new();
    let mut next_period = HashMap::new();
    for (rid_str, prev_str) in &input.prev_period {
        let rid = interner.intern(rid_str);
        let prev = interner.intern(prev_str);
        prev_period.insert(rid, prev);
        next_period.insert(prev, rid);
    }

    Model {
        interner,
        sheets,
        sheet_id_to_idx,
        sheet_name_to_idx,
        prev_period,
        next_period,
        compiled_formulas,
    }
}

fn get_or_compile_formula(
    formula_text: &str,
    cache: &mut HashMap<String, u32>,
    formulas: &mut Vec<CompiledFormula>,
    interner: &mut Interner,
) -> u32 {
    if let Some(&id) = cache.get(formula_text) {
        return id;
    }
    let compiled = compiler::compile(formula_text, interner);
    let id = formulas.len() as u32;
    formulas.push(compiled);
    cache.insert(formula_text.to_string(), id);
    id
}
