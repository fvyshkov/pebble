use serde::Deserialize;
use std::collections::HashMap;

/// The entire model snapshot needed for computation.
#[derive(Deserialize)]
pub struct ModelInput {
    pub sheets: Vec<SheetInput>,
    pub sheet_name_to_id: HashMap<String, String>,
    pub prev_period: HashMap<String, String>,
    pub period_order: Vec<String>,
}

#[derive(Deserialize)]
pub struct SheetInput {
    pub id: String,
    pub name: String,
    pub ordered_aids: Vec<String>,
    pub period_aid: Option<String>,
    pub main_aid: Option<String>,
    pub analytic_name_to_id: HashMap<String, String>,
    /// aid -> {name_lower -> [record_ids]}
    pub name_to_rids: HashMap<String, HashMap<String, Vec<String>>>,
    /// rid -> record data
    pub records: HashMap<String, RecordInput>,
    /// parent_rid -> [child_rids]
    pub children_by_rid: HashMap<String, Vec<String>>,
    /// indicator_rid -> [rules]
    pub rules_by_indicator: HashMap<String, Vec<RuleInput>>,
    /// coord_key -> cell
    pub cells: HashMap<String, CellInput>,
    pub rid_to_period_key: HashMap<String, String>,
    pub period_key_to_rid: HashMap<String, String>,
}

#[derive(Deserialize, Clone)]
pub struct RecordInput {
    pub id: String,
    pub analytic_id: String,
    pub parent_id: Option<String>,
    pub sort_order: i64,
    pub name: String,
    pub period_key: Option<String>,
    pub excel_row: Option<i64>,
}

#[derive(Deserialize, Clone)]
pub struct CellInput {
    pub value: String,
    pub rule: String,
    pub formula: String,
}

#[derive(Deserialize, Clone)]
pub struct RuleInput {
    pub id: String,
    pub kind: String,
    pub scope: HashMap<String, String>,
    pub priority: i64,
    pub formula: String,
}
