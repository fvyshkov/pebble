//! Engine V4: Stateful DAG — build once, recalculate incrementally.
//!
//! Wraps V3's phases into a persistent struct:
//! - `build()` runs enumerate + resolve + topo_sort + evaluate (same as V3)
//! - `mark_dirty()` BFS-forward via stored reverse edges → returns stale cell coords
//! - `update_values()` sets new values, marks dirty, evaluates only dirty subgraph
//!
//! DAG depends only on structure (formulas, analytics, rules) — not values.
//! Value-only changes reuse the stored DAG → milliseconds instead of seconds.

use std::collections::{HashMap, VecDeque};
use rustc_hash::FxHashMap;
use rayon::prelude::*;
use crate::types::ModelInput;
use crate::coord::*;
use crate::builder::*;
use crate::engine_v3::{
    EvalCell,
    enumerate_initial_cells, resolve_all_refs,
    topological_sort_with_reverse, evaluate_parallel,
    eval_cell_inline, collect_changes,
    coord_key_to_string, format_result,
};

// ── Stateful engine ───────────────────────────────────────────────────

#[derive(serde::Serialize, serde::Deserialize)]
pub struct CalcEngine {
    model: Model,
    cells: Vec<EvalCell>,
    values: Vec<f64>,
    lookup: Vec<FxHashMap<CoordKey, u32>>,
    levels: Vec<Vec<u32>>,
    // Reverse-edge CSR: for cell i, dependents are rev_flat[rev_offsets[i]..rev_offsets[i+1]]
    rev_offsets: Vec<u32>,
    rev_flat: Vec<u32>,
}

impl CalcEngine {
    /// Full build from model JSON. Runs all V3 phases and stores state.
    pub fn build(input: &ModelInput) -> Self {
        let t0 = std::time::Instant::now();
        let model = build_model(input);
        let t1 = std::time::Instant::now();
        eprintln!("[v4] build_model: {:.3}s", (t1 - t0).as_secs_f64());

        let (mut cells, mut values, mut lookup) = enumerate_initial_cells(&model);
        let t2 = std::time::Instant::now();
        eprintln!("[v4] enumerate: {:.3}s, {} cells", (t2 - t1).as_secs_f64(), cells.len());

        resolve_all_refs(&model, &mut cells, &mut values, &mut lookup);
        let t3 = std::time::Instant::now();
        eprintln!("[v4] resolve_refs: {:.3}s, {} cells total", (t3 - t2).as_secs_f64(), cells.len());

        let (levels, rev_offsets, rev_flat) = topological_sort_with_reverse(&cells);
        let t4 = std::time::Instant::now();
        eprintln!("[v4] topo_sort: {:.3}s, {} levels, {} reverse edges",
                 (t4 - t3).as_secs_f64(), levels.len(), rev_flat.len());

        evaluate_parallel(&cells, &levels, &mut values, &model.compiled_formulas);
        let t5 = std::time::Instant::now();
        eprintln!("[v4] evaluate: {:.3}s", (t5 - t4).as_secs_f64());

        eprintln!("[v4] BUILD TOTAL: {:.3}s, {} cells", (t5 - t0).as_secs_f64(), cells.len());

        CalcEngine { model, cells, values, lookup, levels, rev_offsets, rev_flat }
    }

    /// Collect all changes (initial build result). Same as V3 collect_changes.
    pub fn collect_all_changes(&self) -> HashMap<String, HashMap<String, String>> {
        collect_changes(&self.cells, &self.values, &self.model)
    }

    /// BFS forward from changed cells via reverse edges.
    /// Returns a boolean vec where dirty[i] = true means cell i is affected.
    fn mark_dirty_ids(&self, changed_ids: &[u32]) -> Vec<bool> {
        let n = self.cells.len();
        let mut dirty = vec![false; n];
        let mut queue: VecDeque<u32> = VecDeque::with_capacity(changed_ids.len() * 16);

        for &cid in changed_ids {
            if (cid as usize) < n {
                dirty[cid as usize] = true;
                queue.push_back(cid);
            }
        }

        while let Some(cid) = queue.pop_front() {
            let ci = cid as usize;
            let start = self.rev_offsets[ci] as usize;
            let end = self.rev_offsets[ci + 1] as usize;
            for &dep in &self.rev_flat[start..end] {
                let di = dep as usize;
                if !dirty[di] {
                    dirty[di] = true;
                    queue.push_back(dep);
                }
            }
        }

        dirty
    }

    /// Mark dirty — returns list of (sheet_id_str, coord_key_str) for all
    /// cells transitively affected by the given changes.
    pub fn mark_dirty_external(&self, changes: &[(u16, CoordKey)]) -> Vec<(String, String)> {
        let t0 = std::time::Instant::now();

        let mut changed_ids: Vec<u32> = Vec::with_capacity(changes.len());
        for &(sheet_idx, coord) in changes {
            if let Some(&cid) = self.lookup.get(sheet_idx as usize)
                .and_then(|m| m.get(&coord))
            {
                changed_ids.push(cid);
            }
        }

        let dirty = self.mark_dirty_ids(&changed_ids);

        // Convert dirty cell_ids to (sheet_id, coord_key) strings
        let sheet_id_strs: Vec<String> = self.model.sheets.iter()
            .map(|s| self.model.interner.get_str(s.id).to_string())
            .collect();

        let mut result = Vec::new();
        for (i, is_dirty) in dirty.iter().enumerate() {
            if *is_dirty {
                let cell = &self.cells[i];
                let sid = &sheet_id_strs[cell.sheet_idx as usize];
                let ck = coord_key_to_string(&self.model.interner, &cell.coord);
                result.push((sid.clone(), ck));
            }
        }

        let t1 = std::time::Instant::now();
        eprintln!("[v4] mark_dirty: {:.3}s, {} changed → {} dirty",
                 (t1 - t0).as_secs_f64(), changes.len(), result.len());

        result
    }

    /// Update cell values and re-evaluate only the affected subgraph.
    /// Returns changes in the same format as collect_all_changes.
    pub fn update_values(&mut self, changes: &[(u16, CoordKey, f64)])
        -> HashMap<String, HashMap<String, String>>
    {
        let t0 = std::time::Instant::now();

        // 1. Apply new values and collect changed cell_ids
        let mut changed_ids: Vec<u32> = Vec::with_capacity(changes.len());
        for &(sheet_idx, coord, new_val) in changes {
            if let Some(&cid) = self.lookup.get(sheet_idx as usize)
                .and_then(|m| m.get(&coord))
            {
                self.values[cid as usize] = new_val;
                // Also update original_value so collect doesn't consider it "changed"
                // relative to the new manual value
                self.cells[cid as usize].original_value = new_val;
                changed_ids.push(cid);
            }
        }

        if changed_ids.is_empty() {
            eprintln!("[v4] update_values: no matching cells found");
            return HashMap::new();
        }

        // 2. Mark dirty (BFS forward)
        let dirty = self.mark_dirty_ids(&changed_ids);
        let dirty_count: usize = dirty.iter().filter(|&&d| d).count();
        let t1 = std::time::Instant::now();

        // 3. Evaluate only dirty cells, walking levels in order
        let base = self.values.as_mut_ptr() as usize;
        for level in &self.levels {
            let dirty_in_level: Vec<u32> = level.iter()
                .copied()
                .filter(|&cid| dirty[cid as usize])
                .collect();

            if dirty_in_level.is_empty() { continue; }

            if dirty_in_level.len() < 256 {
                for &cid in &dirty_in_level {
                    eval_cell_inline(&self.cells, cid, base, &self.model.compiled_formulas);
                }
            } else {
                dirty_in_level.par_iter().for_each(|&cid| {
                    eval_cell_inline(&self.cells, cid, base, &self.model.compiled_formulas);
                });
            }
        }
        let t2 = std::time::Instant::now();

        // 4. Collect changes for dirty cells only
        let sheet_id_strs: Vec<String> = self.model.sheets.iter()
            .map(|s| self.model.interner.get_str(s.id).to_string())
            .collect();

        let n_sheets = self.model.sheets.len();
        let mut per_sheet: Vec<Vec<usize>> = vec![Vec::new(); n_sheets];

        for (cell_id, is_dirty) in dirty.iter().enumerate() {
            if !is_dirty { continue; }
            let cell = &self.cells[cell_id];
            if cell.flags & FLAG_UNRESOLVED != 0 { continue; }
            per_sheet[cell.sheet_idx as usize].push(cell_id);
        }

        let interner = &self.model.interner;
        let values = &self.values;
        let cells = &self.cells;

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

        let t3 = std::time::Instant::now();
        eprintln!("[v4] update_values: {} changes → {} dirty, mark={:.3}s eval={:.3}s collect={:.3}s total={:.3}s",
                 changes.len(), dirty_count,
                 (t1 - t0).as_secs_f64(), (t2 - t1).as_secs_f64(),
                 (t3 - t2).as_secs_f64(), (t3 - t0).as_secs_f64());

        result
    }

    /// Resolve external string identifiers to internal (sheet_idx, CoordKey).
    /// sheet_id and coord_key parts must match interned strings.
    pub fn resolve_external_coord(&self, sheet_id: &str, coord_key: &str) -> Option<(u16, CoordKey)> {
        let sid = self.model.interner.get_id(sheet_id)?;
        let &sheet_idx = self.model.sheet_id_to_idx.get(&sid)?;

        let parts: Vec<&str> = coord_key.split('|').collect();
        let n_axes = self.model.sheets[sheet_idx].ordered_aids.len();
        if parts.len() != n_axes { return None; }

        let mut rids = [0u32; MAX_AXES];
        for (i, part) in parts.iter().enumerate() {
            rids[i] = self.model.interner.get_id(part)?;
        }

        Some((sheet_idx as u16, CoordKey { rids, len: n_axes as u8 }))
    }

    /// Number of cells in the DAG.
    pub fn cell_count(&self) -> usize {
        self.cells.len()
    }

    /// Number of reverse edges (dependency links).
    pub fn edge_count(&self) -> usize {
        self.rev_flat.len()
    }

    /// Number of topological levels.
    pub fn level_count(&self) -> usize {
        self.levels.len()
    }

    /// Serialize to binary (bincode).
    pub fn to_bytes(&self) -> Vec<u8> {
        let t0 = std::time::Instant::now();
        let bytes = bincode::serialize(self).expect("CalcEngine serialization failed");
        let t1 = std::time::Instant::now();
        eprintln!("[v4] serialize: {:.3}s, {} bytes ({:.1} MB)",
                 (t1 - t0).as_secs_f64(), bytes.len(), bytes.len() as f64 / 1_048_576.0);
        bytes
    }

    /// Deserialize from binary (bincode).
    pub fn from_bytes(data: &[u8]) -> Result<Self, String> {
        let t0 = std::time::Instant::now();
        let engine: CalcEngine = bincode::deserialize(data)
            .map_err(|e| format!("CalcEngine deserialization failed: {}", e))?;
        let t1 = std::time::Instant::now();
        eprintln!("[v4] deserialize: {:.3}s, {} cells from {} bytes ({:.1} MB)",
                 (t1 - t0).as_secs_f64(), engine.cells.len(),
                 data.len(), data.len() as f64 / 1_048_576.0);
        Ok(engine)
    }
}
