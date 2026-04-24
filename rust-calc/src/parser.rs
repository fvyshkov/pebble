use std::collections::HashMap;

/// Parsed reference: [Sheet::name](key=value, ...)
#[derive(Debug, Clone)]
pub struct RefParsed {
    pub name: String,
    pub sheet: Option<String>,
    pub params: HashMap<String, String>,
}

/// Parse a reference token like "[name]" or "[Sheet::name](key=value, ...)".
pub fn parse_ref(token: &str) -> RefParsed {
    // Strip outer brackets
    let inner = if token.starts_with('[') {
        // Find the matching ']' — track byte offsets for slicing
        let mut depth = 0;
        let mut bracket_end_byte = 0;
        for (byte_offset, ch) in token.char_indices() {
            match ch {
                '[' => depth += 1,
                ']' => {
                    depth -= 1;
                    if depth == 0 {
                        bracket_end_byte = byte_offset;
                        break;
                    }
                }
                _ => {}
            }
        }
        if bracket_end_byte == 0 {
            return RefParsed { name: token.to_string(), sheet: None, params: HashMap::new() };
        }
        // '[' is 1 byte, ']' is 1 byte
        let name_part = &token[1..bracket_end_byte];
        let rest = &token[bracket_end_byte + 1..];
        (name_part, rest)
    } else {
        return RefParsed { name: token.to_string(), sheet: None, params: HashMap::new() };
    };

    let (name_str, params_str) = inner;

    // Parse params from (...) if present
    let params = if params_str.starts_with('(') && params_str.ends_with(')') {
        parse_params(&params_str[1..params_str.len() - 1])
    } else {
        HashMap::new()
    };

    // Check for cross-sheet separator "::"
    let (sheet, name) = if name_str.contains("::") {
        let mut parts = name_str.splitn(2, "::");
        let s = parts.next().unwrap_or("").trim().to_string();
        let n = parts.next().unwrap_or("").trim().to_string();
        (Some(s), n)
    } else {
        (None, name_str.to_string())
    };

    RefParsed { name, sheet, params }
}

fn parse_params(params_str: &str) -> HashMap<String, String> {
    let mut params = HashMap::new();

    for raw_pair in params_str.split(',') {
        let raw_pair = raw_pair.trim();
        if !raw_pair.contains('=') {
            continue;
        }
        let (key, val) = match raw_pair.split_once('=') {
            Some((k, v)) => (k.trim(), v.trim()),
            None => continue,
        };
        if key.is_empty() {
            continue;
        }

        // Period back-reference: word.назад(N)
        if let Some(n) = extract_period_func(val, "назад") {
            params.insert(key.to_string(), format!("назад({})", n));
            continue;
        }

        // Period forward-reference: word.вперед(N)
        if let Some(n) = extract_period_func(val, "вперед") {
            params.insert(key.to_string(), format!("вперед({})", n));
            continue;
        }

        // Identity: key=key (same value, no-op) — skip
        if val.to_lowercase() == key.to_lowercase() {
            continue;
        }

        // Strip surrounding quotes
        let val = if val.starts_with('"') && val.ends_with('"') && val.len() >= 2 {
            val.strip_prefix('"').unwrap_or(val).strip_suffix('"').unwrap_or(val)
        } else {
            val
        };

        params.insert(key.to_string(), val.to_string());
    }

    params
}

/// Extract N from patterns like "word.назад(N)" or "word.вперед(N)"
fn extract_period_func(val: &str, func_name: &str) -> Option<i32> {
    let marker = format!(".{}(", func_name);
    if let Some(idx) = val.find(&marker) {
        let after = &val[idx + marker.len()..];
        if let Some(end) = after.find(')') {
            let n_str = &after[..end];
            return n_str.parse::<i32>().ok();
        }
    }
    None
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_simple_ref() {
        let r = parse_ref("[Revenue]");
        assert_eq!(r.name, "Revenue");
        assert!(r.sheet.is_none());
        assert!(r.params.is_empty());
    }

    #[test]
    fn test_cross_sheet_ref() {
        let r = parse_ref("[Sheet1::Revenue]");
        assert_eq!(r.name, "Revenue");
        assert_eq!(r.sheet.as_deref(), Some("Sheet1"));
    }

    #[test]
    fn test_ref_with_params() {
        let r = parse_ref("[Revenue](периоды=предыдущий)");
        assert_eq!(r.name, "Revenue");
        assert_eq!(r.params.get("периоды").map(|s| s.as_str()), Some("предыдущий"));
    }

    #[test]
    fn test_period_back() {
        let r = parse_ref("[ind](периоды=период.назад(2))");
        assert_eq!(r.params.get("периоды").map(|s| s.as_str()), Some("назад(2)"));
    }

    #[test]
    fn test_period_forward() {
        let r = parse_ref("[ind](периоды=период.вперед(3))");
        assert_eq!(r.params.get("периоды").map(|s| s.as_str()), Some("вперед(3)"));
    }

    #[test]
    fn test_identity_param_skipped() {
        let r = parse_ref("[ind](период=период)");
        assert!(r.params.is_empty());
    }

    #[test]
    fn test_quoted_param() {
        let r = parse_ref("[ind](периоды=\"предыдущий\")");
        assert_eq!(r.params.get("периоды").map(|s| s.as_str()), Some("предыдущий"));
    }

    #[test]
    fn test_multiple_params() {
        let r = parse_ref("[ind](периоды=Январь, подразделения=Москва)");
        assert_eq!(r.params.get("периоды").map(|s| s.as_str()), Some("Январь"));
        assert_eq!(r.params.get("подразделения").map(|s| s.as_str()), Some("Москва"));
    }
}
