import { describe, it, expect } from 'vitest'
import { evaluateFormula, FormulaContext } from './formulaEngine'

const emptyCtx: FormulaContext = { resolveRef: () => null }

describe('formulaEngine', () => {
  it('evaluates simple numbers', () => {
    expect(evaluateFormula('42', emptyCtx)).toBe(42)
    expect(evaluateFormula('-5', emptyCtx)).toBe(-5)
    expect(evaluateFormula('3.14', emptyCtx)).toBeCloseTo(3.14)
  })

  it('evaluates arithmetic', () => {
    expect(evaluateFormula('2 + 3', emptyCtx)).toBe(5)
    expect(evaluateFormula('10 - 4', emptyCtx)).toBe(6)
    expect(evaluateFormula('3 * 7', emptyCtx)).toBe(21)
    expect(evaluateFormula('20 / 4', emptyCtx)).toBe(5)
  })

  it('respects operator precedence', () => {
    expect(evaluateFormula('2 + 3 * 4', emptyCtx)).toBe(14)
    expect(evaluateFormula('(2 + 3) * 4', emptyCtx)).toBe(20)
  })

  it('handles division by zero', () => {
    expect(evaluateFormula('1 / 0', emptyCtx)).toBeNull()
  })

  it('evaluates comparisons inside если', () => {
    // Comparisons only work inside если...то...конец_если context
    expect(evaluateFormula('если 5 > 3 то 1 иначе 0 конец_если', emptyCtx)).toBe(1)
    expect(evaluateFormula('если 2 > 3 то 1 иначе 0 конец_если', emptyCtx)).toBe(0)
    expect(evaluateFormula('если 3 >= 3 то 1 иначе 0 конец_если', emptyCtx)).toBe(1)
    expect(evaluateFormula('если 2 <= 3 то 1 иначе 0 конец_если', emptyCtx)).toBe(1)
    expect(evaluateFormula('если 3 == 3 то 1 иначе 0 конец_если', emptyCtx)).toBe(1)
    expect(evaluateFormula('если 3 != 4 то 1 иначе 0 конец_если', emptyCtx)).toBe(1)
  })

  it('evaluates если...то...иначе...конец_если', () => {
    expect(evaluateFormula('если 1 > 0 то 10 иначе 20 конец_если', emptyCtx)).toBe(10)
    expect(evaluateFormula('если 0 > 1 то 10 иначе 20 конец_если', emptyCtx)).toBe(20)
  })

  it('evaluates если...иначе_если...конец_если', () => {
    const f = 'если 0 > 1 то 10 иначе_если 1 > 0 то 30 иначе 50 конец_если'
    expect(evaluateFormula(f, emptyCtx)).toBe(30)
  })

  it('resolves references', () => {
    const ctx: FormulaContext = {
      resolveRef: (s, a, r) => {
        if (s === 'Бюджет' && a === 'Показатели' && r === 'Выручка') return 1000
        return 0
      },
    }
    expect(evaluateFormula('[Бюджет].[Показатели].[Выручка] * 0.2', ctx)).toBe(200)
  })

  it('handles complex formula with refs and conditions', () => {
    const ctx: FormulaContext = {
      resolveRef: (s, a, r) => {
        if (r === 'Выручка') return 500
        if (r === 'Расходы') return 300
        return 0
      },
    }
    const f = 'если [Б].[П].[Выручка] > [Б].[П].[Расходы] то [Б].[П].[Выручка] - [Б].[П].[Расходы] иначе 0 конец_если'
    expect(evaluateFormula(f, ctx)).toBe(200)
  })

  it('returns null for empty formula', () => {
    expect(evaluateFormula('', emptyCtx)).toBeNull()
    expect(evaluateFormula('   ', emptyCtx)).toBeNull()
  })

  it('handles nested parens', () => {
    expect(evaluateFormula('((2 + 3) * (4 - 1))', emptyCtx)).toBe(15)
  })
})
