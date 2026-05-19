import { describe, expect, it } from 'vitest'

import { ImportedWorkbookArena, ImportedWorksheetStyleIndexArena } from '../xlsx-large-simple-arena.js'
import { ImportedWorkbookStringPool } from '../xlsx-large-simple-string-pool.js'

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

  it('canonicalizes repeated strings and formulas through a shared import pool', () => {
    const pool = new ImportedWorkbookStringPool()
    const firstArena = new ImportedWorkbookArena(pool)
    const secondArena = new ImportedWorkbookArena(pool)

    const firstCell = firstArena.addCell({ sheetIndex: 0, row: 0, column: 0, value: 'Repeated label' })
    const secondCell = secondArena.addCell({ sheetIndex: 1, row: 0, column: 0, value: 'Repeated label' })
    firstArena.setFormula(firstCell, 'A1&"!"')
    secondArena.setFormula(secondCell, 'A1&"!"')

    expect(firstArena.materializeSheetCells(0)).toEqual([{ address: 'A1', value: 'Repeated label', formula: 'A1&"!"' }])
    expect(secondArena.materializeSheetCells(1)).toEqual([{ address: 'A1', value: 'Repeated label', formula: 'A1&"!"' }])
    expect(pool.count).toBe(2)
  })
})
