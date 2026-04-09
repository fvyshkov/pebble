/**
 * Formula engine for Pebble.
 *
 * Supports:
 * - References: [SheetName].[AnalyticName].[RecordName]
 * - Arithmetic: +, -, *, /, (, )
 * - Comparisons: >, <, >=, <=, ==, !=
 * - если [cond] то [expr] иначе_если [cond] то [expr] иначе [expr] конец_если
 * - Numbers, negative numbers
 */

export interface FormulaContext {
  // Resolves [Sheet].[Analytic].[Record] to a numeric value
  resolveRef: (sheetName: string, analyticName: string, recordName: string) => number | null
}

interface Token {
  type: 'number' | 'ref' | 'op' | 'paren' | 'keyword' | 'cmp'
  value: string
  refParts?: [string, string, string]
}

function tokenize(formula: string): Token[] {
  const tokens: Token[] = []
  let i = 0
  const s = formula.trim()

  while (i < s.length) {
    // Skip whitespace
    if (/\s/.test(s[i])) { i++; continue }

    // Reference: [xxx].[yyy].[zzz]
    if (s[i] === '[') {
      const parts: string[] = []
      while (i < s.length && s[i] === '[') {
        const end = s.indexOf(']', i + 1)
        if (end === -1) break
        parts.push(s.slice(i + 1, end))
        i = end + 1
        if (i < s.length && s[i] === '.') i++ // skip dot
      }
      if (parts.length >= 3) {
        tokens.push({ type: 'ref', value: parts.join('.'), refParts: [parts[0], parts[1], parts[2]] })
      }
      continue
    }

    // Number
    if (/[0-9]/.test(s[i]) || (s[i] === '-' && (tokens.length === 0 || tokens[tokens.length - 1].type === 'op' || tokens[tokens.length - 1].type === 'paren' || tokens[tokens.length - 1].type === 'cmp' || tokens[tokens.length - 1].value === 'то'))) {
      let num = ''
      if (s[i] === '-') { num += '-'; i++ }
      while (i < s.length && /[0-9.]/.test(s[i])) { num += s[i]; i++ }
      tokens.push({ type: 'number', value: num })
      continue
    }

    // Operators
    if ('+-*/'.includes(s[i])) {
      tokens.push({ type: 'op', value: s[i] }); i++; continue
    }

    // Parens
    if ('()'.includes(s[i])) {
      tokens.push({ type: 'paren', value: s[i] }); i++; continue
    }

    // Comparisons
    if (s[i] === '>' || s[i] === '<' || s[i] === '!' || s[i] === '=') {
      let cmp = s[i]; i++
      if (i < s.length && s[i] === '=') { cmp += '='; i++ }
      tokens.push({ type: 'cmp', value: cmp }); continue
    }

    // Keywords (Russian)
    const keywords = ['если', 'то', 'иначе_если', 'иначе', 'конец_если']
    let found = false
    for (const kw of keywords) {
      if (s.slice(i, i + kw.length) === kw && (i + kw.length >= s.length || /[\s\[\]()+\-*/>=<!]/.test(s[i + kw.length]))) {
        tokens.push({ type: 'keyword', value: kw })
        i += kw.length
        found = true
        break
      }
    }
    if (found) continue

    // Skip unknown char
    i++
  }

  return tokens
}

function evalTokens(tokens: Token[], ctx: FormulaContext): number | null {
  let pos = 0

  function peek(): Token | null { return pos < tokens.length ? tokens[pos] : null }
  function advance(): Token { return tokens[pos++] }

  function parseExpr(): number | null {
    // Handle если...то...иначе...конец_если
    const t = peek()
    if (t && t.type === 'keyword' && t.value === 'если') {
      return parseIfElse()
    }
    return parseAddSub()
  }

  function parseIfElse(): number | null {
    advance() // skip 'если'
    const cond = parseComparison()
    // expect 'то'
    const to = peek()
    if (!to || to.value !== 'то') return null
    advance()
    const thenVal = parseExpr()

    const next = peek()
    if (next && next.value === 'иначе_если') {
      advance() // skip 'иначе_если'
      // Recurse as new if
      const condElif = parseComparison()
      const to2 = peek()
      if (!to2 || to2.value !== 'то') return null
      advance()
      const elifVal = parseExpr()

      let elseVal: number | null = 0
      const n2 = peek()
      if (n2 && n2.value === 'иначе') {
        advance()
        elseVal = parseExpr()
      }
      // expect конец_если
      const end = peek()
      if (end && end.value === 'конец_если') advance()

      if (cond !== null && cond !== 0) return thenVal
      if (condElif !== null && condElif !== 0) return elifVal
      return elseVal
    }

    let elseVal: number | null = 0
    const n2 = peek()
    if (n2 && n2.value === 'иначе') {
      advance()
      elseVal = parseExpr()
    }
    // expect конец_если
    const end = peek()
    if (end && end.value === 'конец_если') advance()

    if (cond !== null && cond !== 0) return thenVal
    return elseVal
  }

  function parseComparison(): number | null {
    let left = parseAddSub()
    const t = peek()
    if (t && t.type === 'cmp') {
      advance()
      const right = parseAddSub()
      if (left === null || right === null) return null
      switch (t.value) {
        case '>': return left > right ? 1 : 0
        case '<': return left < right ? 1 : 0
        case '>=': return left >= right ? 1 : 0
        case '<=': return left <= right ? 1 : 0
        case '==': return left === right ? 1 : 0
        case '!=': return left !== right ? 1 : 0
      }
    }
    return left
  }

  function parseAddSub(): number | null {
    let left = parseMulDiv()
    while (peek() && peek()!.type === 'op' && (peek()!.value === '+' || peek()!.value === '-')) {
      const op = advance().value
      const right = parseMulDiv()
      if (left === null || right === null) return null
      left = op === '+' ? left + right : left - right
    }
    return left
  }

  function parseMulDiv(): number | null {
    let left = parseAtom()
    while (peek() && peek()!.type === 'op' && (peek()!.value === '*' || peek()!.value === '/')) {
      const op = advance().value
      const right = parseAtom()
      if (left === null || right === null) return null
      left = op === '*' ? left * right : (right !== 0 ? left / right : null)
    }
    return left
  }

  function parseAtom(): number | null {
    const t = peek()
    if (!t) return null

    if (t.type === 'number') {
      advance()
      return parseFloat(t.value)
    }

    if (t.type === 'ref' && t.refParts) {
      advance()
      return ctx.resolveRef(t.refParts[0], t.refParts[1], t.refParts[2])
    }

    if (t.type === 'paren' && t.value === '(') {
      advance()
      const val = parseExpr()
      if (peek() && peek()!.value === ')') advance()
      return val
    }

    if (t.type === 'keyword' && t.value === 'если') {
      return parseIfElse()
    }

    // Unknown, skip
    advance()
    return null
  }

  return parseExpr()
}

export function evaluateFormula(formula: string, ctx: FormulaContext): number | null {
  if (!formula || !formula.trim()) return null
  try {
    const tokens = tokenize(formula)
    if (tokens.length === 0) return null
    return evalTokens(tokens, ctx)
  } catch {
    return null
  }
}
