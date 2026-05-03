/// Token types produced by the tokenizer.
#[derive(Debug, Clone, PartialEq)]
pub enum Token {
    Ref(String),       // [indicator] or [Sheet::indicator](params)
    Func(String),      // SUM, AVERAGE, IF, MIN, MAX, ABS (followed by '(')
    Num(f64),          // numeric literal
    Op(char),          // + - * / ( ) , < > = !
}

/// Tokenize a formula string into a list of tokens.
pub fn tokenize(formula: &str) -> Vec<Token> {
    let chars: Vec<char> = formula.chars().collect();
    let len = chars.len();
    let mut tokens = Vec::new();
    let mut pos = 0;

    while pos < len {
        let ch = chars[pos];

        // Skip whitespace
        if ch.is_whitespace() {
            pos += 1;
            continue;
        }

        // Reference: [...](...)
        if ch == '[' {
            if let Some((ref_str, end)) = scan_reference(&chars, pos) {
                tokens.push(Token::Ref(ref_str));
                pos = end;
                continue;
            }
        }

        // Function names: SUM, AVERAGE, IF, MIN, MAX, ABS followed by '('
        if ch.is_ascii_alphabetic() {
            let start = pos;
            while pos < len && chars[pos].is_ascii_alphabetic() {
                pos += 1;
            }
            let word: String = chars[start..pos].iter().collect();
            let upper = word.to_uppercase();
            // Skip whitespace between function name and '('
            let mut peek = pos;
            while peek < len && chars[peek].is_whitespace() {
                peek += 1;
            }
            if peek < len && chars[peek] == '(' && matches!(upper.as_str(), "SUM" | "AVERAGE" | "IF" | "MIN" | "MAX" | "ABS" | "INT" | "ROUND") {
                tokens.push(Token::Func(upper));
                pos = peek + 1; // skip the '('
            }
            // else: unknown word, skip
            continue;
        }

        // Numbers
        if ch.is_ascii_digit() || (ch == '.' && pos + 1 < len && chars[pos + 1].is_ascii_digit()) {
            let start = pos;
            while pos < len && chars[pos].is_ascii_digit() {
                pos += 1;
            }
            if pos < len && chars[pos] == '.' {
                pos += 1;
                while pos < len && chars[pos].is_ascii_digit() {
                    pos += 1;
                }
            }
            let num_str: String = chars[start..pos].iter().collect();
            if let Ok(n) = num_str.parse::<f64>() {
                tokens.push(Token::Num(n));
            }
            continue;
        }

        // Operators
        if "+-*/(),<>=!".contains(ch) {
            tokens.push(Token::Op(ch));
            pos += 1;
            continue;
        }

        // Unknown character — skip
        pos += 1;
    }

    tokens
}

/// Scan a reference starting at chars[pos] == '['.
/// Returns (full_ref_string_including_brackets_and_params, end_pos) or None.
fn scan_reference(chars: &[char], start: usize) -> Option<(String, usize)> {
    let len = chars.len();
    if start >= len || chars[start] != '[' {
        return None;
    }

    // Scan for matching ']', handling one level of nesting
    let mut pos = start + 1;
    let mut depth = 1;
    while pos < len && depth > 0 {
        match chars[pos] {
            '[' => depth += 1,
            ']' => depth -= 1,
            _ => {}
        }
        pos += 1;
    }
    if depth != 0 {
        return None;
    }
    // pos is now just after the closing ']'
    let bracket_end = pos;

    // Parent-qualified ref: [a][b][c]... → single token covering every
    // immediately-adjacent bracket pair (no space between ']' and '[').
    // We consume bracket pairs greedily; if any of them is malformed we
    // stop at the last good final_end.
    let mut final_end = bracket_end;
    while pos < len && chars[pos] == '[' {
        let scan_start = pos;
        pos += 1;
        let mut d2 = 1;
        while pos < len && d2 > 0 {
            match chars[pos] {
                '[' => d2 += 1,
                ']' => d2 -= 1,
                _ => {}
            }
            pos += 1;
        }
        if d2 == 0 {
            final_end = pos;
        } else {
            pos = scan_start;
            break;
        }
    }

    // Check for optional params (...) — supports one level of nesting
    pos = final_end;
    if pos < len && chars[pos] == '(' {
        pos += 1;
        let mut pdepth = 1;
        while pos < len && pdepth > 0 {
            match chars[pos] {
                '(' => pdepth += 1,
                ')' => pdepth -= 1,
                _ => {}
            }
            pos += 1;
        }
        // Include params in the ref string
        let full: String = chars[start..pos].iter().collect();
        return Some((full, pos));
    }

    let full: String = chars[start..final_end].iter().collect();
    Some((full, final_end))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_simple_ref() {
        let tokens = tokenize("[Revenue]");
        assert_eq!(tokens, vec![Token::Ref("[Revenue]".to_string())]);
    }

    #[test]
    fn test_arithmetic() {
        let tokens = tokenize("[A] + [B] * 2");
        assert_eq!(tokens, vec![
            Token::Ref("[A]".to_string()),
            Token::Op('+'),
            Token::Ref("[B]".to_string()),
            Token::Op('*'),
            Token::Num(2.0),
        ]);
    }

    #[test]
    fn test_function() {
        let tokens = tokenize("SUM([A], [B])");
        assert_eq!(tokens, vec![
            Token::Func("SUM".to_string()),
            Token::Ref("[A]".to_string()),
            Token::Op(','),
            Token::Ref("[B]".to_string()),
            Token::Op(')'),
        ]);
    }

    #[test]
    fn test_ref_with_params() {
        let tokens = tokenize("[Revenue](периоды=предыдущий)");
        assert_eq!(tokens, vec![Token::Ref("[Revenue](периоды=предыдущий)".to_string())]);
    }

    #[test]
    fn test_cross_sheet() {
        let tokens = tokenize("[Sheet1::Revenue]");
        assert_eq!(tokens, vec![Token::Ref("[Sheet1::Revenue]".to_string())]);
    }

    #[test]
    fn test_comparison() {
        let tokens = tokenize("[A] >= [B]");
        assert_eq!(tokens, vec![
            Token::Ref("[A]".to_string()),
            Token::Op('>'),
            Token::Op('='),
            Token::Ref("[B]".to_string()),
        ]);
    }

    #[test]
    fn test_if_function() {
        let tokens = tokenize("IF([A] > 0, [B], [C])");
        assert_eq!(tokens, vec![
            Token::Func("IF".to_string()),
            Token::Ref("[A]".to_string()),
            Token::Op('>'),
            Token::Num(0.0),
            Token::Op(','),
            Token::Ref("[B]".to_string()),
            Token::Op(','),
            Token::Ref("[C]".to_string()),
            Token::Op(')'),
        ]);
    }

    #[test]
    fn test_nested_brackets() {
        let tokens = tokenize("[name[sub]]");
        assert_eq!(tokens, vec![Token::Ref("[name[sub]]".to_string())]);
    }

    #[test]
    fn test_parent_qualified_ref() {
        // [parent][child] should be a single token
        let tokens = tokenize("[Факторинг][прибыль]");
        assert_eq!(tokens, vec![Token::Ref("[Факторинг][прибыль]".to_string())]);
    }

    #[test]
    fn test_parent_qualified_with_params() {
        let tokens = tokenize("[parent][child](периоды=\"предыдущий\")");
        assert_eq!(tokens, vec![Token::Ref("[parent][child](периоды=\"предыдущий\")".to_string())]);
    }

    #[test]
    fn test_parent_qualified_in_expression() {
        let tokens = tokenize("[A][B] + [C]");
        assert_eq!(tokens, vec![
            Token::Ref("[A][B]".to_string()),
            Token::Op('+'),
            Token::Ref("[C]".to_string()),
        ]);
    }

    #[test]
    fn test_three_level_chain_single_token() {
        let tokens = tokenize("[gp][p][child]");
        assert_eq!(tokens, vec![Token::Ref("[gp][p][child]".to_string())]);
    }

    #[test]
    fn test_three_level_chain_in_sum() {
        let tokens = tokenize("SUM([a][b][c],[a][b][d])");
        assert_eq!(tokens, vec![
            Token::Func("SUM".to_string()),
            Token::Ref("[a][b][c]".to_string()),
            Token::Op(','),
            Token::Ref("[a][b][d]".to_string()),
            Token::Op(')'),
        ]);
    }

    #[test]
    fn test_three_level_with_params() {
        let tokens = tokenize("[a][b][c](периоды=предыдущий)");
        assert_eq!(tokens, vec![
            Token::Ref("[a][b][c](периоды=предыдущий)".to_string()),
        ]);
    }

    #[test]
    fn test_cross_sheet_three_level() {
        let tokens = tokenize("[Sheet::a][b][c]");
        assert_eq!(tokens, vec![
            Token::Ref("[Sheet::a][b][c]".to_string()),
        ]);
    }
}
