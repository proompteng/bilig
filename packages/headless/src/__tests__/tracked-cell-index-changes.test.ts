import { describe, expect, it } from 'vitest'
import { SpreadsheetEngine } from '@bilig/core'

import { materializeTrackedIndexChanges } from '../tracked-cell-index-changes.js'

describe('materializeTrackedIndexChanges', () => {
  it('merges explicit and recalculated slices while materializing same-sheet changes', () => {
    const engine = new SpreadsheetEngine({ workbookName: 'tracked-index-merge' })
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 1)
    engine.setCellValue('Sheet1', 'B1', 2)
    engine.setCellValue('Sheet1', 'A2', 3)
    engine.setCellValue('Sheet1', 'B2', 4)

    const a1 = engine.workbook.getCellIndex('Sheet1', 'A1')
    const b1 = engine.workbook.getCellIndex('Sheet1', 'B1')
    const a2 = engine.workbook.getCellIndex('Sheet1', 'A2')
    const b2 = engine.workbook.getCellIndex('Sheet1', 'B2')
    expect(a1).toBeDefined()
    expect(b1).toBeDefined()
    expect(a2).toBeDefined()
    expect(b2).toBeDefined()

    const changes = materializeTrackedIndexChanges(engine, Uint32Array.of(a1, a2, b1, b2), {
      explicitChangedCount: 2,
    })

    expect(changes.map((change) => change.a1)).toEqual(['A1', 'B1', 'A2', 'B2'])
  })
})
