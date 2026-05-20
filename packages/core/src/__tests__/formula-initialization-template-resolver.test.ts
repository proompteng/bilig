import { describe, expect, it, vi } from 'vitest'
import { createTemplateBank } from '../formula/template-bank.js'
import {
  tryBuildInitialPrefixSumTemplateKey,
  tryBuildInitialSimpleRowRelativeBinaryTemplate,
  tryBuildInitialSimpleRowRelativeBinaryTemplateKey,
} from '../engine/services/formula-initialization-template-keys.js'
import { createInitialTemplateFormulaResolver } from '../engine/services/formula-initialization-template-resolver.js'

describe('initial template formula resolver', () => {
  it('reuses simple row-relative binary templates without recompiling each row', () => {
    const templateBank = createTemplateBank()
    const compileTemplateFormula = vi.fn((source: string, row: number, col: number) => templateBank.resolve(source, row, col))
    const resolve = createInitialTemplateFormulaResolver(compileTemplateFormula)

    const first = resolve('A1+B1', 0, 4)
    const second = resolve('A2+B2', 1, 4)
    const third = resolve('A3+B3', 2, 4)

    expect(compileTemplateFormula).toHaveBeenCalledTimes(1)
    expect(second.templateId).toBe(first.templateId)
    expect(third.templateId).toBe(first.templateId)
    expect(second.compiled.source).toBe('A2+B2')
    expect(third.compiled.deps).toEqual(['A3', 'B3'])
    expect(third.compiled.parsedSymbolicRefs).toEqual([
      {
        address: 'A3',
        col: 0,
        colAbsolute: false,
        row: 2,
        rowAbsolute: false,
      },
      {
        address: 'B3',
        col: 1,
        colAbsolute: false,
        row: 2,
        rowAbsolute: false,
      },
    ])
  })

  it('reuses row-relative binary templates with row-literal suffixes', () => {
    const templateBank = createTemplateBank()
    const compileTemplateFormula = vi.fn((source: string, row: number, col: number) => templateBank.resolve(source, row, col))
    const resolve = createInitialTemplateFormulaResolver(compileTemplateFormula)

    const first = resolve('E2*2+2', 1, 5)
    const second = resolve('E3*2+3', 2, 5)

    expect(compileTemplateFormula).toHaveBeenCalledTimes(1)
    expect(second.templateId).toBe(first.templateId)
    expect(second.compiled.source).toBe('E3*2+3')
    expect(Array.from(second.compiled.constants)).toEqual([2, 3])
    expect(second.compiled.deps).toEqual(['E3'])
  })

  it('reuses anchored prefix SUM compilation while preserving row-specific ranges', () => {
    const templateBank = createTemplateBank()
    const compileTemplateFormula = vi.fn((source: string, row: number, col: number) => templateBank.resolve(source, row, col))
    const resolve = createInitialTemplateFormulaResolver(compileTemplateFormula)

    const first = resolve('SUM(A1:A1)', 0, 4)
    const second = resolve('SUM(A1:A2)', 1, 4)
    const third = resolve('SUM(A1:A3)', 2, 4)

    expect(compileTemplateFormula).toHaveBeenCalledTimes(1)
    expect(second.templateId).toBe(first.templateId)
    expect(third.templateId).toBe(first.templateId)
    expect(second.compiled.source).toBe('SUM(A1:A2)')
    expect(third.compiled.source).toBe('SUM(A1:A3)')
    expect(second.compiled.parsedSymbolicRanges).toEqual([
      {
        address: 'A1:A2',
        kind: 'range',
        refKind: 'cells',
        startAddress: 'A1',
        endAddress: 'A2',
        startRow: 0,
        endRow: 1,
        startCol: 0,
        endCol: 0,
      },
    ])
    expect(third.compiled.parsedDeps).toEqual(third.compiled.parsedSymbolicRanges)
    expect(third.compiled.directAggregateCandidate).toMatchObject({
      aggregateKind: 'sum',
      callee: 'SUM',
      symbolicRangeIndex: 0,
    })
  })

  it('rejects unsafe row numbers in initial template shortcuts', () => {
    const unsafeRow = String(Number.MAX_SAFE_INTEGER + 1)

    expect(tryBuildInitialSimpleRowRelativeBinaryTemplateKey(`A${unsafeRow}+1`, Number.MAX_SAFE_INTEGER, 0)).toBeUndefined()
    expect(tryBuildInitialPrefixSumTemplateKey(`SUM(A1:A${unsafeRow})`, Number.MAX_SAFE_INTEGER, 0)).toBeUndefined()
  })

  it('builds parsed refs for simple row-relative binary formulas during initial template detection', () => {
    expect(tryBuildInitialSimpleRowRelativeBinaryTemplate('A4+B4', 3, 4)).toEqual({
      key: 'c-4+c-3',
      parsedRefs: {
        parsedDeps: [
          {
            address: 'A4',
            col: 0,
            colAbsolute: false,
            kind: 'cell',
            row: 3,
            rowAbsolute: false,
          },
          {
            address: 'B4',
            col: 1,
            colAbsolute: false,
            kind: 'cell',
            row: 3,
            rowAbsolute: false,
          },
        ],
        parsedSymbolicRefs: [
          {
            address: 'A4',
            col: 0,
            colAbsolute: false,
            row: 3,
            rowAbsolute: false,
          },
          {
            address: 'B4',
            col: 1,
            colAbsolute: false,
            row: 3,
            rowAbsolute: false,
          },
        ],
        symbolicRefs: ['A4', 'B4'],
      },
      usesRowLiteralSuffix: false,
    })
  })
})
