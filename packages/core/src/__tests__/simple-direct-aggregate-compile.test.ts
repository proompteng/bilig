import { describe, expect, it } from 'vitest'
import { tryCompileSimpleDirectAggregateFormula } from '../formula/simple-direct-aggregate-compile.js'

describe('tryCompileSimpleDirectAggregateFormula', () => {
  it('normalizes same-column aggregate ranges on the direct compiler fast path', () => {
    const compiled = tryCompileSimpleDirectAggregateFormula('sum(a1:a32)')

    expect(compiled?.directAggregateCandidate?.aggregateKind).toBe('sum')
    expect(compiled?.symbolicRanges).toEqual(['A1:A32'])
    expect(compiled?.parsedSymbolicRanges?.[0]).toMatchObject({
      startAddress: 'A1',
      endAddress: 'A32',
      startRow: 0,
      endRow: 31,
      startCol: 0,
      endCol: 0,
    })
  })

  it('uses the parser fallback for supported absolute same-column ranges', () => {
    const compiled = tryCompileSimpleDirectAggregateFormula('AVERAGE($A$1:$A$4)')

    expect(compiled?.directAggregateCandidate?.aggregateKind).toBe('average')
    expect(compiled?.symbolicRanges).toEqual(['$A$1:$A$4'])
    expect(compiled?.parsedSymbolicRanges?.[0]).toMatchObject({
      startAddress: 'A1',
      endAddress: 'A4',
      startRow: 0,
      endRow: 3,
      startCol: 0,
      endCol: 0,
    })
  })

  it('leaves unsupported direct aggregate ranges on the fallback path', () => {
    expect(tryCompileSimpleDirectAggregateFormula('SUM(A1:B2)')).toBeUndefined()
    expect(tryCompileSimpleDirectAggregateFormula('SUM(A32:A1)')?.parsedSymbolicRanges?.[0]).toMatchObject({
      startAddress: 'A1',
      endAddress: 'A32',
    })
    expect(tryCompileSimpleDirectAggregateFormula('SUM(Sheet2!A1:A2)')).toBeUndefined()
  })
})
