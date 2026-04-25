mod types;
mod tokenizer;
mod parser;
mod evaluator;
mod resolver;
pub mod engine;

// V2: optimized numeric engine
mod intern;
mod coord;
mod compiler;
mod builder;
pub mod engine_v2;
pub mod engine_v3;
pub mod engine_v4;

use pyo3::prelude::*;
use pyo3::types::PyBytes;
use std::collections::HashMap;
use std::sync::Mutex;

/// Calculate all formula cells (V1 — string-based engine).
#[pyfunction]
fn calculate(input_json: &str) -> PyResult<HashMap<String, HashMap<String, String>>> {
    let input: types::ModelInput = serde_json::from_str(input_json)
        .map_err(|e| PyErr::new::<pyo3::exceptions::PyValueError, _>(
            format!("Failed to parse model input: {}", e)
        ))?;
    let result = engine::calculate_model(&input);
    Ok(result)
}

/// Calculate all formula cells (V2 — optimized numeric engine).
#[pyfunction]
fn calculate_v2(input_json: &str) -> PyResult<HashMap<String, HashMap<String, String>>> {
    let input: types::ModelInput = serde_json::from_str(input_json)
        .map_err(|e| PyErr::new::<pyo3::exceptions::PyValueError, _>(
            format!("Failed to parse model input: {}", e)
        ))?;
    let result = engine_v2::calculate_model(&input);
    Ok(result)
}

/// Calculate all formula cells (V3 — DAG parallel engine).
#[pyfunction]
fn calculate_v3(input_json: &str) -> PyResult<HashMap<String, HashMap<String, String>>> {
    let input: types::ModelInput = serde_json::from_str(input_json)
        .map_err(|e| PyErr::new::<pyo3::exceptions::PyValueError, _>(
            format!("Failed to parse model input: {}", e)
        ))?;
    let result = engine_v3::calculate_model(&input);
    Ok(result)
}

// ── V4: Stateful DAG engine ──────────────────────────────────────────

#[pyclass]
struct CalcEngine {
    inner: Mutex<Option<engine_v4::CalcEngine>>,
}

#[pymethods]
impl CalcEngine {
    #[new]
    fn new() -> Self {
        CalcEngine { inner: Mutex::new(None) }
    }

    /// Build DAG from full model JSON. Returns all computed changes.
    fn build(&self, input_json: &str) -> PyResult<HashMap<String, HashMap<String, String>>> {
        let input: types::ModelInput = serde_json::from_str(input_json)
            .map_err(|e| PyErr::new::<pyo3::exceptions::PyValueError, _>(
                format!("Failed to parse model input: {}", e)
            ))?;
        let engine = engine_v4::CalcEngine::build(&input);
        let result = engine.collect_all_changes();
        *self.inner.lock().unwrap() = Some(engine);
        Ok(result)
    }

    /// Update cell values incrementally. changes_json = [["sheet_id", "coord_key", "value"], ...]
    fn update_values(&self, changes_json: &str) -> PyResult<HashMap<String, HashMap<String, String>>> {
        let mut guard = self.inner.lock().unwrap();
        let engine = guard.as_mut().ok_or_else(||
            PyErr::new::<pyo3::exceptions::PyRuntimeError, _>("Engine not built — call build() first")
        )?;

        let raw_changes: Vec<(String, String, String)> = serde_json::from_str(changes_json)
            .map_err(|e| PyErr::new::<pyo3::exceptions::PyValueError, _>(
                format!("Failed to parse changes: {}", e)
            ))?;

        let mut resolved: Vec<(u16, crate::coord::CoordKey, f64)> = Vec::with_capacity(raw_changes.len());
        for (sheet_id, coord_key, value_str) in &raw_changes {
            let val: f64 = value_str.parse().unwrap_or(0.0);
            if let Some((si, ck)) = engine.resolve_external_coord(sheet_id, coord_key) {
                resolved.push((si, ck, val));
            }
        }

        Ok(engine.update_values(&resolved))
    }

    /// Mark dirty — returns list of [sheet_id, coord_key] for all affected cells.
    fn mark_dirty(&self, changes_json: &str) -> PyResult<Vec<(String, String)>> {
        let guard = self.inner.lock().unwrap();
        let engine = guard.as_ref().ok_or_else(||
            PyErr::new::<pyo3::exceptions::PyRuntimeError, _>("Engine not built — call build() first")
        )?;

        let raw_changes: Vec<(String, String)> = serde_json::from_str(changes_json)
            .map_err(|e| PyErr::new::<pyo3::exceptions::PyValueError, _>(
                format!("Failed to parse changes: {}", e)
            ))?;

        let mut resolved: Vec<(u16, crate::coord::CoordKey)> = Vec::with_capacity(raw_changes.len());
        for (sheet_id, coord_key) in &raw_changes {
            if let Some((si, ck)) = engine.resolve_external_coord(sheet_id, coord_key) {
                resolved.push((si, ck));
            }
        }

        Ok(engine.mark_dirty_external(&resolved))
    }

    /// Collect all computed changes (for use after loading from cache).
    fn collect_all_changes(&self) -> PyResult<HashMap<String, HashMap<String, String>>> {
        let guard = self.inner.lock().unwrap();
        let engine = guard.as_ref().ok_or_else(||
            PyErr::new::<pyo3::exceptions::PyRuntimeError, _>("Engine not built")
        )?;
        Ok(engine.collect_all_changes())
    }

    /// Check if engine has been built.
    fn is_built(&self) -> bool {
        self.inner.lock().unwrap().is_some()
    }

    /// Drop cached state (free memory).
    fn drop_state(&self) {
        *self.inner.lock().unwrap() = None;
    }

    /// Serialize built engine to bytes (for DB persistence).
    fn serialize<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyBytes>> {
        let guard = self.inner.lock().unwrap();
        let engine = guard.as_ref().ok_or_else(||
            PyErr::new::<pyo3::exceptions::PyRuntimeError, _>("Engine not built — call build() first")
        )?;
        let bytes = engine.to_bytes();
        Ok(PyBytes::new(py, &bytes))
    }

    /// Load engine from serialized bytes (from DB cache).
    fn load(&self, data: &[u8]) -> PyResult<()> {
        let engine = engine_v4::CalcEngine::from_bytes(data)
            .map_err(|e| PyErr::new::<pyo3::exceptions::PyValueError, _>(e))?;
        *self.inner.lock().unwrap() = Some(engine);
        Ok(())
    }

    /// Get engine stats.
    fn stats(&self) -> PyResult<HashMap<String, usize>> {
        let guard = self.inner.lock().unwrap();
        let engine = guard.as_ref().ok_or_else(||
            PyErr::new::<pyo3::exceptions::PyRuntimeError, _>("Engine not built")
        )?;
        let mut m = HashMap::new();
        m.insert("cells".to_string(), engine.cell_count());
        m.insert("edges".to_string(), engine.edge_count());
        m.insert("levels".to_string(), engine.level_count());
        Ok(m)
    }
}

/// Python module definition.
#[pymodule]
fn pebble_calc(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_function(wrap_pyfunction!(calculate, m)?)?;
    m.add_function(wrap_pyfunction!(calculate_v2, m)?)?;
    m.add_function(wrap_pyfunction!(calculate_v3, m)?)?;
    m.add_class::<CalcEngine>()?;
    Ok(())
}
