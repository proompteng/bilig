import { describe, expect, it } from 'vitest'

import { ImportedWorkbookArena } from '../xlsx-large-simple-arena.js'
import { LargeSimpleFormulaRecords, readLargeSimpleFormulaTypeCode } from '../xlsx-large-simple-formula-records.js'

describe('large simple XLSX formula records', () => {
  it('resolves pooled formula text and shared formulas into the import arena', () => {
    const arena = new ImportedWorkbookArena()
    const records = new LargeSimpleFormulaRecords()
    const directCell = arena.addCell({ sheetIndex: 0, row: 0, column: 0, value: 1 })
    const sharedBaseCell = arena.addCell({ sheetIndex: 0, row: 1, column: 0, value: 2 })
    const sharedFollowerCell = arena.addCell({ sheetIndex: 0, row: 2, column: 0, value: 3 })

    records.add(directCell, 0, 0, readLargeSimpleFormulaTypeCode(null), null, 'IF(B1&lt;3,&quot;yes&quot;,&quot;no&quot;)')
    records.add(sharedBaseCell, 1, 0, readLargeSimpleFormulaTypeCode('shared'), 0, 'B1+1')
    records.add(sharedFollowerCell, 2, 0, readLargeSimpleFormulaTypeCode('shared'), 0, '')

    expect(records.resolveIntoArena(arena)).toBe(true)
    expect(arena.materializeSheetCells(0)).toEqual([
      { address: 'A1', value: 1, formula: 'IF(B1<3,"yes","no")' },
      { address: 'A2', value: 2, formula: 'B1+1' },
      { address: 'A3', value: 3, formula: 'B2+1' },
    ])
  })
})
