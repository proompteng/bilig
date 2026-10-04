import { describe, expect, it } from 'vitest'
import { applyFormulaSuggestion, resolveFormulaAssistState, resolveNameBoxDisplayValue } from '../formulaAssist.js'

describe('formula assist helpers', () => {
  it('keeps argument help without reopening suggestions after accepting a function', () => {
    const state = resolveFormulaAssistState({ value: '=SUM()', caret: 5 })
    expect(state.suggestions).toEqual([])
    expect(state.activeFunction?.entry.name).toBe('SUM')
    expect(resolveFormulaAssistState({ value: '=SUM(A1,', caret: 8 }).suggestions).toEqual([])
    expect(resolveFormulaAssistState({ value: '=SUM(SU', caret: 7 }).suggestions.some((entry) => entry.name === 'SUM')).toBe(true)
  })

  it('shows the actual SUBSTITUTE parameters', () => {
    const state = resolveFormulaAssistState({ value: '=SUBSTITUTE(', caret: 12 })
    expect(state.activeFunction?.signature).toBe('SUBSTITUTE(text, old_text, new_text, [instance_num])')
  })

  it('suggests common functions for a typed prefix', () => {
    const state = resolveFormulaAssistState({
      value: '=su',
      caret: 3,
    })

    expect(state.tokenStart).toBe(1)
    expect(state.tokenEnd).toBe(3)
    expect(state.suggestions.some((entry) => entry.kind === 'function' && entry.name === 'SUM')).toBe(true)
  })

  it('deduplicates generated function suggestions by function name', () => {
    const state = resolveFormulaAssistState({
      value: '=sc',
      caret: 3,
    })
    const functionNames = state.suggestions.filter((entry) => entry.kind === 'function').map((entry) => entry.name)

    expect(functionNames).toContain('SCAN')
    expect(new Set(functionNames).size).toBe(functionNames.length)
  })

  it('surfaces Google Sheets SORTN help without generic fallback text', () => {
    const state = resolveFormulaAssistState({
      value: '=sortn(',
      caret: '=sortn('.length,
    })

    expect(state.activeFunction?.entry.name).toBe('SORTN')
    expect(state.activeFunction?.entry.summary).toContain('tie')
    expect(state.activeFunction?.signature).toContain('display_ties_mode')
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

  it('shows a range-ref defined name when the visible selection summary matches it', () => {
    expect(
      resolveNameBoxDisplayValue({
        sheetName: 'Sheet1',
        address: 'B2',
        selectionLabel: 'B2:D5',
        definedNames: [
          {
            name: 'QuarterlyData',
            value: {
              kind: 'range-ref',
              sheetName: 'Sheet1',
              startAddress: 'B2',
              endAddress: 'D5',
            },
          },
        ],
      }),
    ).toBe('QuarterlyData')
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
