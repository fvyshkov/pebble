use crate::parser::RefParsed;
use crate::engine::SheetMeta;
use std::collections::HashMap;

/// Resolve a local [indicator] reference. Returns new coord_key or None.
pub fn resolve_local(
    ref_parsed: &RefParsed,
    context: &HashMap<String, String>,
    meta: &SheetMeta,
    prev_period: &HashMap<String, String>,
) -> Option<String> {
    let mut name_lower = ref_parsed.name.to_lowercase();

    // Handle "name#rowN" format
    let mut row_hint: Option<i64> = None;
    if let Some(idx) = name_lower.rfind("#row") {
        let row_str = &name_lower[idx + 4..];
        row_hint = row_str.parse::<i64>().ok();
        name_lower = name_lower[..idx].to_string();
    }

    // Handle "parent/child" disambiguation
    let mut parent_hint: Option<String> = None;

    // First try exact name match
    let mut found_exact = false;
    for (aid, nmap) in &meta.name_to_rids {
        if Some(aid) == meta.period_aid.as_ref() {
            continue;
        }
        if nmap.contains_key(&name_lower) {
            found_exact = true;
            break;
        }
    }

    if !found_exact && ref_parsed.name.contains('/') {
        let parts: Vec<&str> = ref_parsed.name.splitn(2, '/').collect();
        parent_hint = Some(parts[0].trim().to_lowercase());
        name_lower = parts[1].trim().to_lowercase();
    }

    let mut target_rid: Option<String> = None;
    let mut target_aid: Option<String> = None;

    for (aid, nmap) in &meta.name_to_rids {
        if Some(aid) == meta.period_aid.as_ref() {
            continue;
        }
        let candidates = match nmap.get(&name_lower) {
            Some(c) if !c.is_empty() => c.clone(),
            _ => continue,
        };

        // Row hint — match by excel_row
        if let Some(rh) = row_hint {
            for crid in &candidates {
                if let Some(rec) = meta.records.get(crid) {
                    if rec.excel_row == Some(rh) {
                        target_rid = Some(crid.clone());
                        target_aid = Some(aid.clone());
                        break;
                    }
                }
            }
            if target_rid.is_some() {
                break;
            }
        }

        // Parent hint — filter by parent name
        let mut filtered_candidates = candidates.clone();
        if parent_hint.is_some() && candidates.len() > 1 {
            let ph = parent_hint.as_ref().unwrap();
            let mut filtered = Vec::new();
            for crid in &candidates {
                if let Some(rec) = meta.records.get(crid) {
                    if let Some(pid) = &rec.parent_id {
                        if let Some(prec) = meta.records.get(pid) {
                            if prec.name.to_lowercase() == *ph {
                                filtered.push(crid.clone());
                            }
                        }
                    }
                }
            }
            if !filtered.is_empty() {
                filtered_candidates = filtered;
            }
        }

        if filtered_candidates.len() == 1 {
            target_rid = Some(filtered_candidates[0].clone());
            target_aid = Some(aid.clone());
            break;
        }

        // Multiple records — disambiguate
        if let Some(cur_rid) = context.get(aid) {
            let cur_rec = meta.records.get(cur_rid);
            let cur_parent = cur_rec.and_then(|r| r.parent_id.clone());

            // 1. Direct parent match
            for crid in &filtered_candidates {
                if let Some(crec) = meta.records.get(crid) {
                    if crec.parent_id == cur_parent {
                        target_rid = Some(crid.clone());
                        target_aid = Some(aid.clone());
                        break;
                    }
                }
            }

            // 2. Common ancestor — pick candidate sharing deepest ancestor
            if target_rid.is_none() {
                let mut cur_ancestors: HashMap<String, usize> = HashMap::new();
                let mut node = cur_parent.clone();
                let mut depth = 0usize;
                while let Some(n) = node {
                    cur_ancestors.insert(n.clone(), depth);
                    node = meta.records.get(&n).and_then(|r| r.parent_id.clone());
                    depth += 1;
                }

                let mut best_crid: Option<String> = None;
                let mut best_depth = usize::MAX;
                for crid in &filtered_candidates {
                    let mut node = meta.records.get(crid).and_then(|r| r.parent_id.clone());
                    while let Some(n) = node {
                        if let Some(&d) = cur_ancestors.get(&n) {
                            if d < best_depth {
                                best_depth = d;
                                best_crid = Some(crid.clone());
                            }
                            break;
                        }
                        node = meta.records.get(&n).and_then(|r| r.parent_id.clone());
                    }
                }
                if let Some(best) = best_crid {
                    if best_depth < 10 {
                        target_rid = Some(best);
                        target_aid = Some(aid.clone());
                    }
                }
            }

            // 3. Closest by sort_order
            if target_rid.is_none() {
                let cur_sort = cur_rec.map(|r| r.sort_order).unwrap_or(0);
                let best = filtered_candidates.iter().min_by_key(|c| {
                    let s = meta.records.get(*c).map(|r| r.sort_order).unwrap_or(0);
                    (s - cur_sort).unsigned_abs()
                });
                if let Some(b) = best {
                    target_rid = Some(b.clone());
                    target_aid = Some(aid.clone());
                }
            }
        }

        if target_rid.is_none() {
            target_rid = Some(filtered_candidates[0].clone());
            target_aid = Some(aid.clone());
        }
        break;
    }

    let target_rid = target_rid?;
    let target_aid = target_aid?;

    // Build parts map
    let mut parts: HashMap<String, String> = HashMap::new();
    for aid in &meta.ordered_aids {
        if *aid == target_aid {
            parts.insert(aid.clone(), target_rid.clone());
        } else if let Some(v) = context.get(aid) {
            parts.insert(aid.clone(), v.clone());
        }
    }

    // Apply param modifiers
    for (param_name, param_value) in &ref_parsed.params {
        let param_aid = find_analytic_by_name(param_name, &meta.analytic_name_to_id);
        let param_aid = match param_aid {
            Some(a) => a,
            None => continue,
        };

        let is_period = meta.period_aid.as_ref() == Some(&param_aid);

        if is_period {
            // Period navigation
            if param_value == "предыдущий" {
                let cur = parts.get(&param_aid)?.clone();
                let next = prev_period.get(&cur)?;
                parts.insert(param_aid, next.clone());
            } else if param_value.starts_with("назад(") {
                let n: usize = param_value[("назад(".len())..param_value.len() - 1].parse().ok()?;
                let mut cur = parts.get(&param_aid)?.clone();
                for _ in 0..n {
                    cur = prev_period.get(&cur)?.clone();
                }
                parts.insert(param_aid, cur);
            } else if param_value.starts_with("вперед(") {
                let n: usize = param_value[("вперед(".len())..param_value.len() - 1].parse().ok()?;
                // Build reverse map
                let next_period: HashMap<&String, &String> = prev_period.iter().map(|(k, v)| (v, k)).collect();
                let mut cur = parts.get(&param_aid)?.clone();
                for _ in 0..n {
                    cur = next_period.get(&cur).copied()?.clone();
                }
                parts.insert(param_aid, cur);
            } else {
                // Absolute period key — scan records
                let mut found = false;
                for (rid, rec) in &meta.records {
                    if let Some(pk) = &rec.period_key {
                        if pk == param_value {
                            parts.insert(param_aid.clone(), rid.clone());
                            found = true;
                            break;
                        }
                    }
                }
                if !found {
                    return None;
                }
            }
        } else {
            // Non-period axis — resolve by name
            let nmap = meta.name_to_rids.get(&param_aid)?;
            let rids = nmap.get(&param_value.to_lowercase())?;
            if rids.is_empty() {
                return None;
            }
            parts.insert(param_aid, rids[0].clone());
        }
    }

    // Build coord_key
    let coord_parts: Vec<String> = meta.ordered_aids.iter()
        .map(|aid| parts.get(aid).cloned().unwrap_or_default())
        .collect();
    if coord_parts.iter().any(|p| p.is_empty()) {
        return None;
    }
    let result_key = coord_parts.join("|");

    // Self-reference guard
    let current_key: String = meta.ordered_aids.iter()
        .map(|aid| context.get(aid).cloned().unwrap_or_default())
        .collect::<Vec<_>>()
        .join("|");
    if result_key == current_key {
        return None;
    }

    Some(result_key)
}

/// Resolve a cross-sheet [Sheet::indicator] reference.
/// Returns the value of the target cell.
pub fn resolve_cross_sheet(
    ref_parsed: &RefParsed,
    context: &HashMap<String, String>,
    src_meta: &SheetMeta,
    sheet_name_to_id: &HashMap<String, String>,
    all_meta: &HashMap<String, SheetMeta>,
    prev_period: &HashMap<String, String>,
    get_cell: &mut dyn FnMut(&str, &str) -> f64,
) -> f64 {
    let sheet_name = match &ref_parsed.sheet {
        Some(s) => s,
        None => return 0.0,
    };

    let target_sid = match sheet_name_to_id.get(&sheet_name.to_lowercase()) {
        Some(s) => s.clone(),
        None => return 0.0,
    };
    let target_meta = match all_meta.get(&target_sid) {
        Some(m) => m,
        None => return 0.0,
    };

    let mut name_lower = ref_parsed.name.to_lowercase();

    // Handle "name#rowN"
    let mut row_hint: Option<i64> = None;
    if let Some(idx) = name_lower.rfind("#row") {
        let row_str = &name_lower[idx + 4..];
        row_hint = row_str.parse::<i64>().ok();
        name_lower = name_lower[..idx].to_string();
    }

    // Handle "parent/child" — only if exact name not found
    let mut parent_hint: Option<String> = None;
    let mut found_exact = false;
    for (aid, nmap) in &target_meta.name_to_rids {
        if Some(aid) == target_meta.period_aid.as_ref() {
            continue;
        }
        if nmap.contains_key(&name_lower) {
            found_exact = true;
            break;
        }
    }
    if !found_exact && ref_parsed.name.contains('/') {
        let parts: Vec<&str> = ref_parsed.name.splitn(2, '/').collect();
        parent_hint = Some(parts[0].trim().to_lowercase());
        name_lower = parts[1].trim().to_lowercase();
    }

    // Find indicator by exact name (case-insensitive)
    let mut ind_rid: Option<String> = None;
    for (aid, nmap) in &target_meta.name_to_rids {
        if Some(aid) == target_meta.period_aid.as_ref() {
            continue;
        }
        let rids = match nmap.get(&name_lower) {
            Some(r) if !r.is_empty() => r,
            _ => continue,
        };

        // Row hint
        if let Some(rh) = row_hint {
            for crid in rids {
                if let Some(rec) = target_meta.records.get(crid) {
                    if rec.excel_row == Some(rh) {
                        ind_rid = Some(crid.clone());
                        break;
                    }
                }
            }
            if ind_rid.is_some() {
                break;
            }
        }

        // Parent hint
        if let Some(ph) = &parent_hint {
            if rids.len() > 1 {
                for crid in rids {
                    if let Some(rec) = target_meta.records.get(crid) {
                        if let Some(pid) = &rec.parent_id {
                            if let Some(prec) = target_meta.records.get(pid) {
                                if prec.name.to_lowercase() == *ph {
                                    ind_rid = Some(crid.clone());
                                    break;
                                }
                            }
                        }
                    }
                }
                if ind_rid.is_some() {
                    break;
                }
            }
        }

        if rids.len() == 1 {
            ind_rid = Some(rids[0].clone());
        } else {
            // Disambiguate: use source indicator's section name
            let src_main = src_meta.main_aid.as_ref();
            let src_ind_rid = src_main.and_then(|m| context.get(m));
            let src_name = src_ind_rid
                .and_then(|rid| src_meta.records.get(rid))
                .map(|r| r.name.to_lowercase())
                .unwrap_or_default();

            let mut best_rid = rids[0].clone();
            if !src_name.is_empty() {
                for crid in rids {
                    let mut node = target_meta.records.get(crid).and_then(|r| r.parent_id.clone());
                    while let Some(n) = node {
                        if let Some(prec) = target_meta.records.get(&n) {
                            let pname = prec.name.to_lowercase();
                            if !pname.is_empty() && src_name.contains(&pname) {
                                best_rid = crid.clone();
                                break;
                            }
                            node = prec.parent_id.clone();
                        } else {
                            break;
                        }
                    }
                    if best_rid != rids[0] {
                        break;
                    }
                }
            }
            ind_rid = Some(best_rid);
        }
        break;
    }

    let ind_rid = match ind_rid {
        Some(r) => r,
        None => return 0.0,
    };

    let src_period_aid = match &src_meta.period_aid {
        Some(a) => a,
        None => return 0.0,
    };
    let mut period_rid = match context.get(src_period_aid) {
        Some(r) => r.clone(),
        None => return 0.0,
    };

    // Apply period modifiers from params
    for (param_name, param_value) in &ref_parsed.params {
        let param_aid = find_analytic_by_name(param_name, &src_meta.analytic_name_to_id);
        let param_aid = match param_aid {
            Some(a) if Some(&a) == src_meta.period_aid.as_ref() => a,
            _ => continue,
        };
        let _ = param_aid; // used for type check

        if param_value == "предыдущий" {
            period_rid = match prev_period.get(&period_rid) {
                Some(r) => r.clone(),
                None => return 0.0,
            };
        } else if param_value.starts_with("назад(") {
            let n: usize = match param_value[6..param_value.len() - 1].parse() {
                Ok(n) => n,
                Err(_) => return 0.0,
            };
            for _ in 0..n {
                period_rid = match prev_period.get(&period_rid) {
                    Some(r) => r.clone(),
                    None => return 0.0,
                };
            }
        } else if param_value.starts_with("вперед(") {
            let n: usize = match param_value[7..param_value.len() - 1].parse() {
                Ok(n) => n,
                Err(_) => return 0.0,
            };
            let next_period: HashMap<&String, &String> = prev_period.iter().map(|(k, v)| (v, k)).collect();
            for _ in 0..n {
                period_rid = match next_period.get(&period_rid).copied() {
                    Some(r) => r.clone(),
                    None => return 0.0,
                };
            }
        } else {
            // Absolute period key
            let pk_to_rid = &src_meta.period_key_to_rid;
            period_rid = match pk_to_rid.get(param_value) {
                Some(r) => r.clone(),
                None => return 0.0,
            };
        }
    }

    // Build coord key in target sheet's analytic order.
    // Translate period if source and target use different period analytics.
    let target_period_aid = target_meta.period_aid.as_ref();

    let mut target_period_rid = period_rid.clone();
    if src_meta.period_aid != target_meta.period_aid {
        // Different period analytics — translate via period_key
        if let Some(src_pk) = src_meta.rid_to_period_key.get(&period_rid) {
            if let Some(tgt_rid) = target_meta.period_key_to_rid.get(src_pk) {
                target_period_rid = tgt_rid.clone();
            } else {
                // Try fallback: monthly → yearly
                target_period_rid = try_period_fallback(src_pk, &target_meta.period_key_to_rid);
                if target_period_rid.is_empty() {
                    return 0.0;
                }
            }
        } else {
            return 0.0;
        }
    }

    let mut ck_parts = Vec::new();
    for aid in &target_meta.ordered_aids {
        if target_period_aid == Some(aid) {
            ck_parts.push(target_period_rid.clone());
        } else {
            ck_parts.push(ind_rid.clone());
        }
    }
    let target_ck = ck_parts.join("|");

    get_cell(&target_sid, &target_ck)
}

/// Try to find a matching analytic by name or partial match.
fn find_analytic_by_name(param_name: &str, analytic_name_to_id: &HashMap<String, String>) -> Option<String> {
    // Exact match first
    if let Some(aid) = analytic_name_to_id.get(param_name) {
        return Some(aid.clone());
    }
    // Partial match (case-insensitive)
    let lower = param_name.to_lowercase();
    for (aname, aid) in analytic_name_to_id {
        if aname.to_lowercase().contains(&lower) {
            return Some(aid.clone());
        }
    }
    None
}

/// Try period fallback: monthly → yearly, quarterly → yearly
fn try_period_fallback(src_pk: &str, target_pk_to_rid: &HashMap<String, String>) -> String {
    // YYYY-MM → YYYY-Y
    if src_pk.len() == 7 && src_pk.as_bytes().get(4) == Some(&b'-') {
        let year = &src_pk[..4];
        let fallback = format!("{}-Y", year);
        if let Some(rid) = target_pk_to_rid.get(&fallback) {
            return rid.clone();
        }
    }
    // YYYY-QN → YYYY-Y
    if src_pk.contains("-Q") {
        let year = &src_pk[..4];
        let fallback = format!("{}-Y", year);
        if let Some(rid) = target_pk_to_rid.get(&fallback) {
            return rid.clone();
        }
    }
    // YYYY-HN → YYYY-Y
    if src_pk.contains("-H") {
        let year = &src_pk[..4];
        let fallback = format!("{}-Y", year);
        if let Some(rid) = target_pk_to_rid.get(&fallback) {
            return rid.clone();
        }
    }
    String::new()
}
