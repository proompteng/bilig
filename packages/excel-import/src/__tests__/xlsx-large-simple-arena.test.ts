import { describe, expect, it } from 'vitest'

import { ImportedWorkbookArena, ImportedWorksheetStyleIndexArena } from '../xlsx-large-simple-arena.js'

describe('large simple XLSX import arena', () => {
  it('releases typed storage and pools after materialization', () => {
    const arena = new ImportedWorkbookArena()
    const styleIndexes = new ImportedWorksheetStyleIndexArena()
    const cell = arena.addCell({ sheetIndex: 0, row: 0, column: 0, value: 'Alpha' })
    arena.setFormula(cell, 'B1')
    styleIndexes.add(0, 0, 1)

    expect(arena.materializeSheetCells(0)).toEqual([{ address: 'A1', value: 'Alpha', formula: 'B1' }])
    expect(styleIndexes.count).toBe(1)

    arena.release()
    styleIndexes.release()

    const snapshot = arena.snapshot()
    expect(snapshot.sheetIndexes).toHaveLength(0)
    expect(snapshot.rows).toHaveLength(0)
    expect(snapshot.strings).toHaveLength(0)
    expect(snapshot.formulas).toHaveLength(0)
    expect(arena.cellCount).toBe(0)
    expect(arena.readPreviewText(0, 0)).toBe('')
    expect(styleIndexes.count).toBe(0)
  })
})
