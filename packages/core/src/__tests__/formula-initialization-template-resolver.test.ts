import { describe, expect, it, vi } from 'vitest'
import { createTemplateBank } from '../formula/template-bank.js'
import {
  tryBuildInitialPrefixSumTemplateKey,
  tryBuildInitialSimpleRowRelativeBinaryTemplateKey,
} from '../engine/services/formula-initialization-template-keys.js'
import { createInitialTemplateFormulaResolver } from '../engine/services/formula-initialization-template-resolver.js'

describe('initial template formula resolver', () => {
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
})
