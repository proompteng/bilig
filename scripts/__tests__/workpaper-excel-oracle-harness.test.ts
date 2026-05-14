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

  it('rejects an Excel oracle value when Excel rewrites the formula as an unsupported UDF', () => {
    expect(
      classifyFormulaComparison({
        actualBiligValue: { kind: 'string', value: '2026-04-01' },
        embeddedCacheValue: { kind: 'string', value: '2026-04-01' },
        excelOracleFormula: 'IFERROR(_xludf.XLOOKUP(C14,Bank!$D$2:$D$31,Bank!$B$2:$B$31,"",0),"")',
        excelOracleValue: { kind: 'string', value: '' },
        formula: 'IFERROR(XLOOKUP(C14,Bank!$D$2:$D$31,Bank!$B$2:$B$31,"",0),"")',
      }),
    ).toBe('missing_excel_oracle')
  })

  it('accepts Excel compatibility prefixes when comparing oracle formulas', () => {
    expect(
      classifyFormulaComparison({
        actualBiligValue: numberValue(20),
        embeddedCacheValue: numberValue(20),
        excelOracleFormula: '_xlfn.XLOOKUP(2,A2:A4,B2:B4)',
        excelOracleValue: numberValue(20),
        formula: 'XLOOKUP(2,A2:A4,B2:B4)',
      }),
    ).toBe('bilig_matches_excel')
  })
})
