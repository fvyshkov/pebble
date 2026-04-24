mod types;
mod tokenizer;
mod parser;
mod evaluator;
mod resolver;
pub mod engine;

use pyo3::prelude::*;
use std::collections::HashMap;

/// Calculate all formula cells across all sheets in a model.
/// Takes a JSON string (ModelInput), returns {sheet_id: {coord_key: value_str}}.
#[pyfunction]
fn calculate(input_json: &str) -> PyResult<HashMap<String, HashMap<String, String>>> {
    let input: types::ModelInput = serde_json::from_str(input_json)
        .map_err(|e| PyErr::new::<pyo3::exceptions::PyValueError, _>(
            format!("Failed to parse model input: {}", e)
        ))?;
    let result = engine::calculate_model(&input);
    Ok(result)
}

/// Python module definition.
#[pymodule]
fn pebble_calc(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_function(wrap_pyfunction!(calculate, m)?)?;
    Ok(())
}
