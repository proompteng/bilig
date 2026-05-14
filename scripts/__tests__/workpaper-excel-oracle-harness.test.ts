import { describe, expect, it } from 'vitest'

import { classifyFormulaComparison, type NormalizedFormulaValue } from '../workpaper-excel-oracle-harness.ts'

const numberValue = (value: number): NormalizedFormulaValue => ({ kind: 'number', value })

describe('WorkPaper Excel oracle harness classifier', () => {
  it('classifies a workbook where cache equals Excel and Bilig matches', () => {
    expect(
      classifyFormulaComparison({
        actualBiligValue: numberValue(3),
        embeddedCacheValue: numberValue(3),
        excelOracleValue: numberValue(3),
        formula: 'A1+B1',
      }),
    ).toBe('bilig_matches_excel')
  })

  it('classifies a workbook where the cache is stale but Bilig matches fresh Excel', () => {
    expect(
      classifyFormulaComparison({
        actualBiligValue: numberValue(3),
        embeddedCacheValue: numberValue(2),
        excelOracleValue: numberValue(3),
        formula: 'A1+B1',
      }),
    ).toBe('cache_stale_bilig_matches_excel')
  })

  it('classifies a workbook where the cache is stale and Bilig mismatches fresh Excel', () => {
    expect(
      classifyFormulaComparison({
        actualBiligValue: numberValue(4),
        embeddedCacheValue: numberValue(2),
        excelOracleValue: numberValue(3),
        formula: 'A1+B1',
      }),
    ).toBe('cache_stale_bilig_mismatches_excel')
  })

  it('skips volatile formulas before comparing cache or Bilig output', () => {
    expect(
      classifyFormulaComparison({
        actualBiligValue: numberValue(45100),
        embeddedCacheValue: numberValue(45000),
        excelOracleValue: numberValue(45100),
        formula: 'TODAY()',
      }),
    ).toBe('volatile_skipped')
  })

  it('marks a workbook without an Excel oracle as missing oracle instead of an accuracy failure', () => {
    expect(
      classifyFormulaComparison({
        actualBiligValue: numberValue(3),
        embeddedCacheValue: numberValue(2),
        formula: 'A1+B1',
      }),
    ).toBe('missing_excel_oracle')
  })
})
