use crate::intern::Interner;
use crate::tokenizer::{tokenize, Token};
use crate::parser::parse_ref;

/// A single instruction in a compiled formula.
#[derive(Clone, Copy, Debug, serde::Serialize, serde::Deserialize)]
pub enum Instr {
    PushNum(f64),
    PushRef(u32),    // index into CompiledFormula.refs
    Add,
    Sub,
    Mul,
    Div,
    Neg,
    CmpLt,
    CmpGt,
    CmpLe,
    CmpGe,
    CmpEq,
    CmpNe,
    CallIf,          // pops 3: false_val, true_val, cond
    CallSum(u8),     // pops N args, pushes sum
    CallAverage(u8), // pops N args, pushes average
    CallMin(u8),
    CallMax(u8),
    CallAbs,
    CallInt,         // floor(x) — Excel INT()
    CallRound(u8),   // round(x, n_decimals)
}

/// How a parameter modifies an axis value.
#[derive(Clone, Debug, serde::Serialize, serde::Deserialize)]
pub enum ParamValue {
    Previous,           // предыдущий
    Back(u16),          // назад(N)
    Forward(u16),       // вперед(N)
    RecordName(u32),    // interned lowercase record name
}

/// A pre-parsed reference with numeric IDs.
#[derive(Clone, Debug, serde::Serialize, serde::Deserialize)]
pub struct CompiledRef {
    pub name_lower: u32,          // interned lowercase indicator name
    pub sheet_id: Option<u32>,    // interned target sheet ID (if cross-sheet)
    pub row_hint: Option<i64>,    // #rowN hint
    pub parent_hint: Option<u32>, // interned lowercase parent name
    pub params: Vec<(u32, ParamValue)>, // (interned analytic_name_lower, value)
}

/// A compiled formula ready for stack-based evaluation.
#[derive(Clone, Debug, serde::Serialize, serde::Deserialize)]
pub struct CompiledFormula {
    pub instrs: Vec<Instr>,
    pub refs: Vec<CompiledRef>,
}

/// Compile a formula string into bytecode.
pub fn compile(formula: &str, interner: &mut Interner) -> CompiledFormula {
    let tokens = tokenize(formula);
    let mut refs: Vec<CompiledRef> = Vec::new();
    let mut instrs: Vec<Instr> = Vec::new();

    // Check for special single-keyword formulas
    let trimmed = formula.trim();
    if trimmed.eq_ignore_ascii_case("AVERAGE") || trimmed.eq_ignore_ascii_case("LAST") {
        // These are handled as special cases in engine, not as bytecode
        return CompiledFormula { instrs, refs };
    }

    let mut pos = 0;
    compile_expr(&tokens, &mut pos, &mut instrs, &mut refs, interner);

    CompiledFormula { instrs, refs }
}

// ── Recursive descent compiler ──────────────────────────────────────────

fn compile_expr(tokens: &[Token], pos: &mut usize, instrs: &mut Vec<Instr>,
                refs: &mut Vec<CompiledRef>, interner: &mut Interner) {
    compile_comparison(tokens, pos, instrs, refs, interner);
}

fn compile_comparison(tokens: &[Token], pos: &mut usize, instrs: &mut Vec<Instr>,
                      refs: &mut Vec<CompiledRef>, interner: &mut Interner) {
    compile_additive(tokens, pos, instrs, refs, interner);

    while *pos < tokens.len() {
        match &tokens[*pos] {
            Token::Op('<') => {
                if *pos + 1 < tokens.len() && tokens[*pos + 1] == Token::Op('=') {
                    *pos += 2;
                    compile_additive(tokens, pos, instrs, refs, interner);
                    instrs.push(Instr::CmpLe);
                } else {
                    *pos += 1;
                    compile_additive(tokens, pos, instrs, refs, interner);
                    instrs.push(Instr::CmpLt);
                }
            }
            Token::Op('>') => {
                if *pos + 1 < tokens.len() && tokens[*pos + 1] == Token::Op('=') {
                    *pos += 2;
                    compile_additive(tokens, pos, instrs, refs, interner);
                    instrs.push(Instr::CmpGe);
                } else {
                    *pos += 1;
                    compile_additive(tokens, pos, instrs, refs, interner);
                    instrs.push(Instr::CmpGt);
                }
            }
            Token::Op('=') => {
                *pos += 1;
                compile_additive(tokens, pos, instrs, refs, interner);
                instrs.push(Instr::CmpEq);
            }
            Token::Op('!') => {
                if *pos + 1 < tokens.len() && tokens[*pos + 1] == Token::Op('=') {
                    *pos += 2;
                    compile_additive(tokens, pos, instrs, refs, interner);
                    instrs.push(Instr::CmpNe);
                } else {
                    break;
                }
            }
            _ => break,
        }
    }
}

fn compile_additive(tokens: &[Token], pos: &mut usize, instrs: &mut Vec<Instr>,
                    refs: &mut Vec<CompiledRef>, interner: &mut Interner) {
    compile_term(tokens, pos, instrs, refs, interner);

    while *pos < tokens.len() {
        match &tokens[*pos] {
            Token::Op('+') => {
                *pos += 1;
                compile_term(tokens, pos, instrs, refs, interner);
                instrs.push(Instr::Add);
            }
            Token::Op('-') => {
                *pos += 1;
                compile_term(tokens, pos, instrs, refs, interner);
                instrs.push(Instr::Sub);
            }
            _ => break,
        }
    }
}

fn compile_term(tokens: &[Token], pos: &mut usize, instrs: &mut Vec<Instr>,
                refs: &mut Vec<CompiledRef>, interner: &mut Interner) {
    compile_unary(tokens, pos, instrs, refs, interner);

    while *pos < tokens.len() {
        match &tokens[*pos] {
            Token::Op('*') => {
                *pos += 1;
                compile_unary(tokens, pos, instrs, refs, interner);
                instrs.push(Instr::Mul);
            }
            Token::Op('/') => {
                *pos += 1;
                compile_unary(tokens, pos, instrs, refs, interner);
                instrs.push(Instr::Div);
            }
            _ => break,
        }
    }
}

fn compile_unary(tokens: &[Token], pos: &mut usize, instrs: &mut Vec<Instr>,
                 refs: &mut Vec<CompiledRef>, interner: &mut Interner) {
    if *pos < tokens.len() {
        if let Token::Op('-') = &tokens[*pos] {
            *pos += 1;
            compile_unary(tokens, pos, instrs, refs, interner);
            instrs.push(Instr::Neg);
            return;
        }
        if let Token::Op('+') = &tokens[*pos] {
            *pos += 1;
            compile_unary(tokens, pos, instrs, refs, interner);
            return;
        }
    }
    compile_primary(tokens, pos, instrs, refs, interner);
}

fn compile_primary(tokens: &[Token], pos: &mut usize, instrs: &mut Vec<Instr>,
                   refs: &mut Vec<CompiledRef>, interner: &mut Interner) {
    if *pos >= tokens.len() {
        instrs.push(Instr::PushNum(0.0));
        return;
    }

    match &tokens[*pos] {
        Token::Num(n) => {
            instrs.push(Instr::PushNum(*n));
            *pos += 1;
        }
        Token::Ref(ref_str) => {
            let ref_idx = compile_ref(ref_str, refs, interner);
            instrs.push(Instr::PushRef(ref_idx));
            *pos += 1;
        }
        Token::Func(name) => {
            let func_name = name.clone();
            *pos += 1; // skip func name (tokenizer already consumed the '(')
            let mut arg_count: u8 = 0;

            if *pos < tokens.len() && tokens[*pos] == Token::Op(')') {
                *pos += 1; // empty function call
            } else {
                loop {
                    compile_expr(tokens, pos, instrs, refs, interner);
                    arg_count += 1;
                    if *pos < tokens.len() && tokens[*pos] == Token::Op(',') {
                        *pos += 1;
                    } else {
                        break;
                    }
                }
                // Skip closing ')'
                if *pos < tokens.len() && tokens[*pos] == Token::Op(')') {
                    *pos += 1;
                }
            }

            match func_name.as_str() {
                "SUM" => instrs.push(Instr::CallSum(arg_count)),
                "AVERAGE" => instrs.push(Instr::CallAverage(arg_count)),
                "IF" => instrs.push(Instr::CallIf),
                "MIN" => instrs.push(Instr::CallMin(arg_count)),
                "MAX" => instrs.push(Instr::CallMax(arg_count)),
                "ABS" => instrs.push(Instr::CallAbs),
                "INT" => instrs.push(Instr::CallInt),
                "ROUND" => instrs.push(Instr::CallRound(arg_count)),
                _ => instrs.push(Instr::CallSum(arg_count)), // fallback
            }
        }
        Token::Op('(') => {
            *pos += 1;
            compile_expr(tokens, pos, instrs, refs, interner);
            if *pos < tokens.len() && tokens[*pos] == Token::Op(')') {
                *pos += 1;
            }
        }
        _ => {
            *pos += 1;
            instrs.push(Instr::PushNum(0.0));
        }
    }
}

/// Compile a reference token into a CompiledRef, returning its index.
fn compile_ref(ref_str: &str, refs: &mut Vec<CompiledRef>, interner: &mut Interner) -> u32 {
    let parsed = parse_ref(ref_str);
    let name_lower = interner.intern_lower(&parsed.name);

    // Parse row hint: "name#row123" -> row_hint = 123
    let (actual_name_lower, row_hint) = {
        let name_lc = parsed.name.to_lowercase();
        if let Some(idx) = name_lc.find("#row") {
            let base = &name_lc[..idx];
            let row_str = &name_lc[idx + 4..];
            let row = row_str.parse::<i64>().ok();
            (interner.intern(base), row)
        } else {
            (name_lower, None)
        }
    };

    // Don't split on "/" at compile time — the resolver handles the
    // parent/child fallback at lookup time (matching V1 behavior which
    // only splits when the exact name isn't found in name_to_rids).
    let final_name = actual_name_lower;
    let parent_hint: Option<u32> = None;

    // Convert sheet name to interned ID
    let sheet_id = parsed.sheet.as_ref().map(|s| interner.intern_lower(s));

    // Convert params
    let mut params = Vec::new();
    for (key, val) in &parsed.params {
        let key_lower = interner.intern_lower(key);
        let pv = if val == "предыдущий" {
            ParamValue::Previous
        } else if val.starts_with("назад(") && val.ends_with(')') {
            // "назад(" = 5 Cyrillic chars (2 bytes each) + '(' (1 byte) = 11 bytes
            let prefix_len = "назад(".len();
            let n_str = &val[prefix_len..val.len() - 1];
            let n = n_str.parse::<u16>().unwrap_or(1);
            ParamValue::Back(n)
        } else if val.starts_with("вперед(") && val.ends_with(')') {
            // "вперед(" = 6 Cyrillic chars (2 bytes each) + '(' (1 byte) = 13 bytes
            let prefix_len = "вперед(".len();
            let n_str = &val[prefix_len..val.len() - 1];
            let n = n_str.parse::<u16>().unwrap_or(1);
            ParamValue::Forward(n)
        } else {
            let val_lower = interner.intern_lower(val);
            ParamValue::RecordName(val_lower)
        };
        params.push((key_lower, pv));
    }

    let idx = refs.len() as u32;
    refs.push(CompiledRef {
        name_lower: final_name,
        sheet_id,
        row_hint,
        parent_hint,
        params,
    });
    idx
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_compile_simple_add() {
        let mut int = Interner::new();
        let cf = compile("[A] + [B]", &mut int);
        // Should be: PushRef(0), PushRef(1), Add
        assert_eq!(cf.instrs.len(), 3);
        assert!(matches!(cf.instrs[0], Instr::PushRef(0)));
        assert!(matches!(cf.instrs[1], Instr::PushRef(1)));
        assert!(matches!(cf.instrs[2], Instr::Add));
        assert_eq!(cf.refs.len(), 2);
    }

    #[test]
    fn test_compile_num() {
        let mut int = Interner::new();
        let cf = compile("42.5", &mut int);
        assert_eq!(cf.instrs.len(), 1);
        assert!(matches!(cf.instrs[0], Instr::PushNum(n) if (n - 42.5).abs() < 1e-10));
    }

    #[test]
    fn test_compile_sum_function() {
        let mut int = Interner::new();
        let cf = compile("SUM([A], [B], [C])", &mut int);
        // PushRef(0), PushRef(1), PushRef(2), CallSum(3)
        assert_eq!(cf.instrs.len(), 4);
        assert!(matches!(cf.instrs[3], Instr::CallSum(3)));
    }

    #[test]
    fn test_compile_if() {
        let mut int = Interner::new();
        let cf = compile("IF([A] > 0, [B], [C])", &mut int);
        assert!(cf.instrs.iter().any(|i| matches!(i, Instr::CallIf)));
    }

    #[test]
    fn test_compile_precedence() {
        let mut int = Interner::new();
        let cf = compile("[A] + [B] * 2", &mut int);
        // Should be: PushRef(0), PushRef(1), PushNum(2), Mul, Add
        assert_eq!(cf.instrs.len(), 5);
        assert!(matches!(cf.instrs[3], Instr::Mul));
        assert!(matches!(cf.instrs[4], Instr::Add));
    }

    #[test]
    fn test_compile_negation() {
        let mut int = Interner::new();
        let cf = compile("-[A]", &mut int);
        assert!(matches!(cf.instrs[0], Instr::PushRef(0)));
        assert!(matches!(cf.instrs[1], Instr::Neg));
    }

    #[test]
    fn test_compile_ref_with_params() {
        let mut int = Interner::new();
        let cf = compile("[Revenue](периоды=предыдущий)", &mut int);
        assert_eq!(cf.refs.len(), 1);
        assert_eq!(cf.refs[0].params.len(), 1);
        assert!(matches!(cf.refs[0].params[0].1, ParamValue::Previous));
    }

    #[test]
    fn test_compile_cross_sheet() {
        let mut int = Interner::new();
        let cf = compile("[Sheet1::Revenue]", &mut int);
        assert_eq!(cf.refs.len(), 1);
        assert!(cf.refs[0].sheet_id.is_some());
    }

    #[test]
    fn test_compile_special_keywords() {
        let mut int = Interner::new();
        let cf = compile("AVERAGE", &mut int);
        assert!(cf.instrs.is_empty());
        let cf2 = compile("LAST", &mut int);
        assert!(cf2.instrs.is_empty());
    }
}
