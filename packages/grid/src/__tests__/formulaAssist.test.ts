import { describe, expect, it } from 'vitest'
import { applyFormulaSuggestion, resolveFormulaAssistState, resolveNameBoxDisplayValue } from '../formulaAssist.js'

describe('formula assist helpers', () => {
  it('suggests common functions for a typed prefix', () => {
    const state = resolveFormulaAssistState({
      value: '=su',
      caret: 3,
    })

    expect(state.tokenStart).toBe(1)
    expect(state.tokenEnd).toBe(3)
    expect(state.suggestions.some((entry) => entry.kind === 'function' && entry.name === 'SUM')).toBe(true)
  })

  it('tracks the active argument for nested functions', () => {
    const state = resolveFormulaAssistState({
      value: '=IF(A1>0,SUM(B1:B3),XLOOKUP("id",A:A,',
      caret: '=IF(A1>0,SUM(B1:B3),XLOOKUP("id",A:A,'.length,
    })

    expect(state.activeFunction?.entry.name).toBe('XLOOKUP')
    expect(state.activeFunction?.activeArgumentIndex).toBe(2)
    expect(state.activeFunction?.signature).toContain('return_array')
  })

  it('includes defined names in suggestions and the name box display', () => {
    const definedNames = [
      {
        name: 'TaxRate',
        value: {
          kind: 'cell-ref' as const,
          sheetName: 'Sheet1',
          address: 'B2',
        },
      },
    ]

    const state = resolveFormulaAssistState({
      value: '=ta',
      caret: 3,
      definedNames,
    })

    expect(state.suggestions[0]).toMatchObject({
      kind: 'defined-name',
      name: 'TaxRate',
    })
    expect(
      resolveNameBoxDisplayValue({
        sheetName: 'Sheet1',
        address: 'B2',
        definedNames,
      }),
    ).toBe('TaxRate')
  })

  it('replaces a function prefix with a callable suggestion and keeps the caret inside parens', () => {
    const next = applyFormulaSuggestion({
      value: '=su',
      tokenStart: 1,
      tokenEnd: 3,
      suggestion: {
        kind: 'function',
        name: 'SUM',
        category: 'aggregation',
        summary: 'Add numbers, ranges, and spill results.',
        signature: 'SUM(number1, [number2], ...)',
      },
    })

    expect(next).toEqual({
      value: '=SUM()',
      caret: 5,
    })
  })
})
