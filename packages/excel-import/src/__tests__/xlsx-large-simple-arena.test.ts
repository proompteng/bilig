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
    expect(arena.readPreviewText(0, 0)).toBe('Alpha')
    expect(styleIndexes.count).toBe(1)

    arena.release()
    styleIndexes.release()

    const snapshot = arena.snapshot()
    expect(snapshot.sheetIndex).toBeNull()
    expect(snapshot.sheetIndexes).toBeUndefined()
    expect(snapshot.rows).toHaveLength(0)
    expect(snapshot.columns).toBeInstanceOf(Uint16Array)
    expect(snapshot.numberValues).toBeUndefined()
    expect(snapshot.booleanValues).toBeUndefined()
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

  it('can retain repeated strings without building de-duplication maps', () => {
    const pool = new ImportedWorkbookStringPool()
    const arena = new ImportedWorkbookArena(pool, { deduplicateStrings: false })

    arena.addCell({ sheetIndex: 0, row: 0, column: 0, value: 'Repeated label' })
    arena.addCell({ sheetIndex: 0, row: 1, column: 0, value: 'Repeated label' })

    expect(arena.snapshot().strings).toEqual(['Repeated label', 'Repeated label'])
    expect(arena.materializeSheetCells(0)).toEqual([
      { address: 'A1', value: 'Repeated label' },
      { address: 'A2', value: 'Repeated label' },
    ])
    expect(pool.count).toBe(0)
  })

  it('bounds repeated string and formula interning for large import arenas', () => {
    const pool = new ImportedWorkbookStringPool()
    const arena = new ImportedWorkbookArena(pool, {
      deduplicateStrings: 'bounded',
      deduplicateFormulas: 'bounded',
      dedupeMaxEntries: 2,
    })

    const firstRepeated = arena.addCell({ sheetIndex: 0, row: 0, column: 0, value: 'Repeated label' })
    const secondRepeated = arena.addCell({ sheetIndex: 0, row: 1, column: 0, value: 'Repeated label' })
    arena.setFormula(firstRepeated, 'A1&"!"')
    arena.setFormula(secondRepeated, 'A1&"!"')

    expect(arena.snapshot().strings).toEqual(['Repeated label'])
    expect(arena.snapshot().formulas).toEqual(['A1&"!"'])

    arena.addCell({ sheetIndex: 0, row: 2, column: 0, value: 'Unique 1' })
    arena.addCell({ sheetIndex: 0, row: 3, column: 0, value: 'Unique 2' })
    arena.addCell({ sheetIndex: 0, row: 4, column: 0, value: 'Repeated label' })

    expect(arena.snapshot().strings).toEqual(['Repeated label', 'Unique 1', 'Unique 2', 'Repeated label'])
  })

  it('releases preview and de-duplication scratch while lazy cells stay readable', () => {
    const pool = new ImportedWorkbookStringPool()
    const arena = new ImportedWorkbookArena(pool)
    const cell = arena.addCell({ sheetIndex: 0, row: 0, column: 0, value: 'Alpha' })
    arena.setFormula(cell, 'B1')
    const lazyCells = arena.createLazySheetCells(0)

    expect(arena.readPreviewText(0, 0)).toBe('Alpha')

    arena.releaseMaterializationScratch()

    expect(arena.readPreviewText(0, 0)).toBe('')
    expect(lazyCells).toHaveLength(1)
    expect(lazyCells[0]).toEqual({ address: 'A1', value: 'Alpha', formula: 'B1' })
    expect([...lazyCells]).toEqual([{ address: 'A1', value: 'Alpha', formula: 'B1' }])
  })

  it('uses scalar sheet ownership for single-sheet arenas and falls back only for mixed ownership', () => {
    const arena = new ImportedWorkbookArena()
    arena.addCell({ sheetIndex: 4, row: 0, column: 0, value: 1 })

    expect(arena.snapshot().sheetIndex).toBe(4)
    expect(arena.snapshot().sheetIndexes).toBeUndefined()
    expect(arena.snapshot().columns).toBeInstanceOf(Uint16Array)
    expect(arena.snapshot().numberValues).toEqual(new Float64Array([1]))
    expect(arena.snapshot().stringIds).toBeUndefined()
    expect(arena.snapshot().booleanValues).toBeUndefined()
    expect(arena.snapshot().formulaIds).toBeUndefined()
    expect(arena.materializeSheetCells(3)).toEqual([])
    expect(arena.materializeSheetCells(4)).toEqual([{ address: 'A1', value: 1 }])

    arena.addCell({ sheetIndex: 5, row: 1, column: 0, value: 2 })

    expect(arena.snapshot().sheetIndexes).toEqual(new Uint32Array([4, 5]))
    expect(arena.materializeSheetCells(4)).toEqual([{ address: 'A1', value: 1 }])
    expect(arena.materializeSheetCells(5)).toEqual([{ address: 'A2', value: 2 }])
  })

  it('allocates number, string, boolean, and formula storage only when cells need those pools', () => {
    const stringOnlyArena = new ImportedWorkbookArena()
    stringOnlyArena.addCell({ sheetIndex: 0, row: 0, column: 0, value: 'Only text' })
    expect(stringOnlyArena.snapshot().numberValues).toBeUndefined()
    expect(stringOnlyArena.materializeSheetCells(0)).toEqual([{ address: 'A1', value: 'Only text' }])

    const arena = new ImportedWorkbookArena()
    const numericCell = arena.addCell({ sheetIndex: 0, row: 0, column: 0, value: 42 })

    expect(arena.snapshot().numberValues).toEqual(new Float64Array([42]))
    expect(arena.snapshot().stringIds).toBeUndefined()
    expect(arena.snapshot().booleanValues).toBeUndefined()
    expect(arena.snapshot().formulaIds).toBeUndefined()

    arena.addCell({ sheetIndex: 0, row: 1, column: 0, value: 'Label' })
    expect(arena.snapshot().stringIds).toEqual(new Uint32Array([0xffffffff, 0]))
    expect(arena.snapshot().booleanValues).toBeUndefined()
    expect(arena.snapshot().formulaIds).toBeUndefined()

    arena.addCell({ sheetIndex: 0, row: 2, column: 0, value: true })
    expect(arena.snapshot().booleanValues).toEqual(new Uint8Array([0, 0, 1]))

    arena.setFormula(numericCell, '1+1')
    expect(arena.snapshot().formulaIds).toEqual(new Uint32Array([0, 0xffffffff, 0xffffffff]))
  })

  it('keeps preview values in fixed slots without overriding values with formulas', () => {
    const arena = new ImportedWorkbookArena()
    const valueFormulaCell = arena.addCell({ sheetIndex: 0, row: 0, column: 0, value: 42 })
    const formulaOnlyCell = arena.addCell({ sheetIndex: 0, row: 0, column: 1, value: undefined })
    arena.addCell({ sheetIndex: 0, row: 8, column: 0, value: 'Outside preview' })

    arena.setFormula(valueFormulaCell, '40+2')
    arena.setFormula(formulaOnlyCell, 'SUM(A1:A1)')

    expect(arena.readPreviewText(0, 0)).toBe('42')
    expect(arena.readPreviewText(0, 1)).toBe('=SUM(A1:A1)')
    expect(arena.readPreviewText(8, 0)).toBe('')
    expect(arena.materializeSheetCells(0)).toEqual([
      { address: 'A1', value: 42, formula: '40+2' },
      { address: 'B1', formula: 'SUM(A1:A1)' },
      { address: 'A9', value: 'Outside preview' },
    ])
  })
})
