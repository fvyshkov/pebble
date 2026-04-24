use crate::tokenizer::{Token, tokenize};

/// Evaluate a formula string.
/// `get_ref_value(ref_token)` resolves a reference token like "[indicator]" to a value.
/// Returns None if all references are missing (propagate "missing" to caller).
pub fn evaluate<F>(formula: &str, get_ref_value: &mut F) -> Option<f64>
where
    F: FnMut(&str) -> Option<f64>,
{
    let trimmed = formula.trim();
    if trimmed.is_empty() {
        return Some(0.0);
    }
    // Try parsing as a plain number first
    if let Ok(n) = trimmed.parse::<f64>() {
        return Some(n);
    }

    let tokens = tokenize(formula);
    if tokens.is_empty() {
        return Some(0.0);
    }

    let mut parser = ExprParser {
        tokens: &tokens,
        pos: 0,
        get_ref_value,
    };

    match std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| parser.parse_expr())) {
        Ok(result) => {
            match result {
                None => None, // propagate "missing"
                Some(v) if v.is_infinite() => Some(0.0),
                Some(v) => Some(v),
            }
        }
        Err(_) => Some(0.0),
    }
}

struct ExprParser<'a, F> {
    tokens: &'a [Token],
    pos: usize,
    get_ref_value: &'a mut F,
}

impl<'a, F> ExprParser<'a, F>
where
    F: FnMut(&str) -> Option<f64>,
{
    fn peek(&self) -> Option<&Token> {
        self.tokens.get(self.pos)
    }

    fn advance(&mut self) -> Option<&Token> {
        if self.pos < self.tokens.len() {
            let t = &self.tokens[self.pos];
            self.pos += 1;
            Some(t)
        } else {
            None
        }
    }

    fn expect_op(&mut self, op: char) -> bool {
        if let Some(Token::Op(c)) = self.peek() {
            if *c == op {
                self.advance();
                return true;
            }
        }
        false
    }

    /// Coerce None (missing cell) to 0.0 for arithmetic.
    fn n(v: Option<f64>) -> f64 {
        v.unwrap_or(0.0)
    }

    fn parse_expr(&mut self) -> Option<f64> {
        self.parse_comparison()
    }

    fn parse_comparison(&mut self) -> Option<f64> {
        let mut left = self.parse_additive();
        loop {
            match self.peek() {
                Some(Token::Op('<')) | Some(Token::Op('>')) | Some(Token::Op('=')) | Some(Token::Op('!')) => {
                    let op1 = if let Some(Token::Op(c)) = self.advance() { *c } else { break; };
                    // Check for compound operators: <=, >=, !=, ==
                    let op = if let Some(Token::Op('=')) = self.peek() {
                        self.advance();
                        format!("{}=", op1)
                    } else {
                        op1.to_string()
                    };
                    let right = Self::n(self.parse_additive());
                    let l = Self::n(left);
                    left = Some(match op.as_str() {
                        "<" => if l < right { 1.0 } else { 0.0 },
                        ">" => if l > right { 1.0 } else { 0.0 },
                        "<=" => if l <= right { 1.0 } else { 0.0 },
                        ">=" => if l >= right { 1.0 } else { 0.0 },
                        "=" | "==" => if (l - right).abs() < 1e-12 { 1.0 } else { 0.0 },
                        "!=" => if (l - right).abs() >= 1e-12 { 1.0 } else { 0.0 },
                        _ => 0.0,
                    });
                }
                _ => break,
            }
        }
        left
    }

    fn parse_additive(&mut self) -> Option<f64> {
        let mut left = self.parse_term();
        loop {
            match self.peek() {
                Some(Token::Op('+')) => {
                    self.advance();
                    let right = Self::n(self.parse_term());
                    left = Some(Self::n(left) + right);
                }
                Some(Token::Op('-')) => {
                    self.advance();
                    let right = Self::n(self.parse_term());
                    left = Some(Self::n(left) - right);
                }
                _ => break,
            }
        }
        left // preserves None if no operators
    }

    fn parse_term(&mut self) -> Option<f64> {
        let mut left = self.parse_unary();
        loop {
            match self.peek() {
                Some(Token::Op('*')) => {
                    self.advance();
                    let right = Self::n(self.parse_unary());
                    left = Some(Self::n(left) * right);
                }
                Some(Token::Op('/')) => {
                    self.advance();
                    let right = Self::n(self.parse_unary());
                    left = Some(if right != 0.0 { Self::n(left) / right } else { f64::NAN });
                }
                _ => break,
            }
        }
        left // preserves None if no operators
    }

    fn parse_unary(&mut self) -> Option<f64> {
        if let Some(Token::Op('-')) = self.peek() {
            self.advance();
            let val = self.parse_unary();
            return Some(-Self::n(val));
        }
        self.parse_primary()
    }

    fn parse_primary(&mut self) -> Option<f64> {
        match self.peek().cloned() {
            None => Some(0.0),
            Some(Token::Num(n)) => {
                self.advance();
                Some(n)
            }
            Some(Token::Ref(ref_str)) => {
                let s = ref_str.clone();
                self.advance();
                (self.get_ref_value)(&s)
            }
            Some(Token::Func(ref name)) => {
                let func_name = name.clone();
                self.advance();
                let mut args = Vec::new();
                loop {
                    args.push(self.parse_expr());
                    if !self.expect_op(',') {
                        break;
                    }
                }
                self.expect_op(')');

                Some(match func_name.as_str() {
                    "SUM" => args.iter().map(|a| Self::n(*a)).sum(),
                    "AVERAGE" => {
                        let present: Vec<f64> = args.iter().filter(|a| a.is_some()).map(|a| Self::n(*a)).collect();
                        if present.is_empty() { 0.0 } else { present.iter().sum::<f64>() / present.len() as f64 }
                    }
                    "IF" => {
                        let cond = Self::n(args.first().copied().flatten());
                        let true_val = Self::n(args.get(1).copied().flatten());
                        let false_val = Self::n(args.get(2).copied().flatten());
                        if cond != 0.0 { true_val } else { false_val }
                    }
                    "MIN" => {
                        if args.is_empty() { 0.0 } else {
                            args.iter().map(|a| Self::n(*a)).fold(f64::INFINITY, f64::min)
                        }
                    }
                    "MAX" => {
                        if args.is_empty() { 0.0 } else {
                            args.iter().map(|a| Self::n(*a)).fold(f64::NEG_INFINITY, f64::max)
                        }
                    }
                    "ABS" => {
                        Self::n(args.first().copied().flatten()).abs()
                    }
                    _ => args.iter().map(|a| Self::n(*a)).sum(), // fallback
                })
            }
            Some(Token::Op('(')) => {
                self.advance();
                let val = self.parse_expr();
                self.expect_op(')');
                val
            }
            _ => {
                self.advance();
                Some(0.0)
            }
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn eval(formula: &str) -> Option<f64> {
        evaluate(formula, &mut |_| Some(0.0))
    }

    fn eval_with<F: FnMut(&str) -> Option<f64>>(formula: &str, f: F) -> Option<f64> {
        evaluate(formula, &mut { f })
    }

    #[test]
    fn test_number() {
        assert_eq!(eval("42"), Some(42.0));
        assert_eq!(eval("3.14"), Some(3.14));
    }

    #[test]
    fn test_arithmetic() {
        assert_eq!(eval("2 + 3"), Some(5.0));
        assert_eq!(eval("10 - 4"), Some(6.0));
        assert_eq!(eval("3 * 4"), Some(12.0));
        assert_eq!(eval("10 / 4"), Some(2.5));
    }

    #[test]
    fn test_precedence() {
        assert_eq!(eval("2 + 3 * 4"), Some(14.0));
        assert_eq!(eval("(2 + 3) * 4"), Some(20.0));
    }

    #[test]
    fn test_unary_minus() {
        assert_eq!(eval("-5"), Some(-5.0));
        assert_eq!(eval("-5 + 3"), Some(-2.0));
    }

    #[test]
    fn test_division_by_zero() {
        let r = eval("1 / 0");
        assert!(r.unwrap().is_nan());
    }

    #[test]
    fn test_ref_values() {
        let r = eval_with("[A] + [B]", |name| {
            match name {
                "[A]" => Some(10.0),
                "[B]" => Some(20.0),
                _ => None,
            }
        });
        assert_eq!(r, Some(30.0));
    }

    #[test]
    fn test_average_skips_none() {
        let r = eval_with("AVERAGE([A], [B], [C])", |name| {
            match name {
                "[A]" => Some(10.0),
                "[B]" => None,
                "[C]" => Some(20.0),
                _ => None,
            }
        });
        assert_eq!(r, Some(15.0));
    }

    #[test]
    fn test_if() {
        let r = eval_with("IF([A] > 0, [B], [C])", |name| {
            match name {
                "[A]" => Some(1.0),
                "[B]" => Some(100.0),
                "[C]" => Some(200.0),
                _ => None,
            }
        });
        assert_eq!(r, Some(100.0));
    }

    #[test]
    fn test_comparison() {
        assert_eq!(eval("3 > 2"), Some(1.0));
        assert_eq!(eval("2 > 3"), Some(0.0));
        assert_eq!(eval("3 >= 3"), Some(1.0));
        assert_eq!(eval("3 = 3"), Some(1.0));
        assert_eq!(eval("3 != 4"), Some(1.0));
    }

    #[test]
    fn test_empty_formula() {
        assert_eq!(eval(""), Some(0.0));
        assert_eq!(eval("  "), Some(0.0));
    }

    #[test]
    fn test_none_propagation() {
        // Single ref returning None → propagate
        let r = eval_with("[X]", |_| None);
        assert_eq!(r, None);
    }

    #[test]
    fn test_none_coerced_in_arithmetic() {
        // None + number → 0 + number
        let r = eval_with("[X] + 5", |_| None);
        assert_eq!(r, Some(5.0));
    }
}
