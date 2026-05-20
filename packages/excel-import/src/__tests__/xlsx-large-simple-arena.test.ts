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

  it('keeps large arena typed-array capacity close to live cell count', () => {
    const arena = new ImportedWorkbookArena()
    const styleIndexes = new ImportedWorksheetStyleIndexArena()
    const cellCount = 100_000

    for (let row = 0; row < cellCount; row += 1) {
      arena.addCell({ sheetIndex: 0, row, column: 0, value: row })
      styleIndexes.add(row, 0, 1)
    }

    expect(arena.cellCount).toBe(cellCount)
    expect(styleIndexes.count).toBe(cellCount)
    expect(arena.allocatedCellCapacity).toBeLessThanOrEqual(Math.ceil(cellCount * 1.25))
    expect(styleIndexes.allocatedCapacity).toBeLessThanOrEqual(Math.ceil(cellCount * 1.25))
    expect(arena.allocatedCellCapacity).toBeLessThan(131_072)
    expect(styleIndexes.allocatedCapacity).toBeLessThan(131_072)
  })

  it('tracks shared-string references so non-shared sheets skip resolution work', () => {
    const numericOnlyArena = new ImportedWorkbookArena()
    numericOnlyArena.addCell({ sheetIndex: 0, row: 0, column: 0, value: 42 })
    const referencedIndexes = new Set<number>()

    numericOnlyArena.collectSharedStringIndexes(referencedIndexes)

    expect(numericOnlyArena.sharedStringRefCount).toBe(0)
    expect(referencedIndexes.size).toBe(0)
    expect(numericOnlyArena.resolveSharedStrings([])).toEqual([])
    expect(numericOnlyArena.materializeSheetCells(0)).toEqual([{ address: 'A1', value: 42 }])

    const sharedArena = new ImportedWorkbookArena()
    sharedArena.addSharedStringCell({ sheetIndex: 1, row: 0, column: 0, sharedStringIndex: 3 })
    sharedArena.addSharedStringCell({ sheetIndex: 1, row: 0, column: 1, sharedStringIndex: 5 })
    const sharedIndexes = new Set<number>()

    sharedArena.collectSharedStringIndexes(sharedIndexes)

    expect(sharedArena.sharedStringRefCount).toBe(2)
    expect([...sharedIndexes].toSorted((left, right) => left - right)).toEqual([3, 5])
    expect(
      sharedArena.resolveSharedStrings([
        { text: 'unused 0', rich: false },
        { text: 'unused 1', rich: false },
        { text: 'unused 2', rich: false },
        { text: 'Alpha', rich: false },
        { text: 'unused 4', rich: false },
        { text: 'Rich Beta', rich: true, xml: '<si><r><t>Rich Beta</t></r></si>' },
      ]),
    ).toEqual([
      {
        address: 'B1',
        text: 'Rich Beta',
        storage: 'sharedString',
        xml: '<si><r><t>Rich Beta</t></r></si>',
      },
    ])
    expect(sharedArena.sharedStringRefCount).toBe(0)
    expect(sharedArena.materializeSheetCells(1)).toEqual([
      { address: 'A1', value: 'Alpha' },
      { address: 'B1', value: 'Rich Beta' },
    ])
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
