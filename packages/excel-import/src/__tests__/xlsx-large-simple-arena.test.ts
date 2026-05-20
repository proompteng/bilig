import { describe, expect, it } from 'vitest'

import { ImportedWorkbookArena, ImportedWorksheetStyleIndexArena } from '../xlsx-large-simple-arena.js'
import { createLazyWorkbookRichTextCells, mergeWorkbookRichTextCells } from '../xlsx-large-simple-lazy-rich-text-cells.js'
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
    expect(snapshot.integerValues).toBeUndefined()
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

  it('keeps lazy shared-string cells readable without duplicating shared strings into arena pools', () => {
    const arena = new ImportedWorkbookArena()
    const cell = arena.addSharedStringCell({ sheetIndex: 0, row: 0, column: 0, sharedStringIndex: 2 })
    arena.setFormula(cell, 'A1&""')

    const richTextCells = arena.retainSharedStringReferences([
      { text: 'unused', rich: false },
      { text: 'also unused', rich: false },
      { text: 'Shared label', rich: true, xml: '<si><r><t>Shared label</t></r></si>' },
    ])
    const lazyCells = arena.createLazySheetCells(0)

    expect(richTextCells).toEqual([
      {
        address: 'A1',
        text: 'Shared label',
        storage: 'sharedString',
        xml: '<si><r><t>Shared label</t></r></si>',
      },
    ])
    expect(arena.snapshot().strings).toEqual([])
    expect(arena.readPreviewText(0, 0)).toBe('Shared label')

    arena.releaseMaterializationScratch()

    expect(lazyCells).toHaveLength(1)
    expect(lazyCells[0]).toEqual({ address: 'A1', value: 'Shared label', formula: 'A1&""' })
    expect([...lazyCells]).toEqual([{ address: 'A1', value: 'Shared label', formula: 'A1&""' }])
  })

  it('eagerly materializes retained shared strings without duplicating them into arena pools', () => {
    const arena = new ImportedWorkbookArena()
    arena.addSharedStringCell({ sheetIndex: 0, row: 0, column: 0, sharedStringIndex: 1 })
    arena.addSharedStringCell({ sheetIndex: 0, row: 0, column: 1, sharedStringIndex: 2 })

    expect(
      arena.retainSharedStringReferences([
        { text: 'unused', rich: false },
        { text: 'First shared', rich: false },
        { text: 'Second shared', rich: false },
      ]),
    ).toEqual([])
    expect(arena.snapshot().strings).toEqual([])
    expect(arena.materializeSheetCells(0)).toEqual([
      { address: 'A1', value: 'First shared' },
      { address: 'B1', value: 'Second shared' },
    ])

    arena.release()

    expect(arena.snapshot().strings).toEqual([])
    expect(arena.cellCount).toBe(0)
  })

  it('keeps large shared-string rich-text artifacts lazy', () => {
    const arena = new ImportedWorkbookArena()
    arena.addSharedStringCell({ sheetIndex: 0, row: 0, column: 0, sharedStringIndex: 0 })
    arena.addSharedStringCell({ sheetIndex: 0, row: 1, column: 0, sharedStringIndex: 1 })

    const richTextCells = arena.retainSharedStringReferences(
      [
        { text: 'First rich', rich: true, xml: '<si><r><t>First rich</t></r></si>' },
        { text: 'Second rich', rich: true, xml: '<si><r><t>Second rich</t></r></si>' },
      ],
      { lazyRichTextCellThreshold: 1 },
    )

    expect(Array.isArray(richTextCells)).toBe(true)
    expect(richTextCells).toHaveLength(2)
    expect(richTextCells?.[0]).toEqual({
      address: 'A1',
      text: 'First rich',
      storage: 'sharedString',
      xml: '<si><r><t>First rich</t></r></si>',
    })
    expect(richTextCells?.at(-1)).toEqual({
      address: 'A2',
      text: 'Second rich',
      storage: 'sharedString',
      xml: '<si><r><t>Second rich</t></r></si>',
    })
    expect(richTextCells?.map((cell) => cell.address)).toEqual(['A1', 'A2'])
    const toJSON = Reflect.get(richTextCells ?? [], 'toJSON')
    expect(typeof toJSON === 'function' ? toJSON.call(richTextCells) : null).toEqual([...richTextCells!])
  })

  it('merges imported rich-text artifacts without forcing lazy materialization', () => {
    let materialized = 0
    const appended = createLazyWorkbookRichTextCells(2, (index) => {
      materialized += 1
      return {
        address: index === 0 ? 'B1' : 'C1',
        text: index === 0 ? 'Second' : 'Third',
        storage: 'sharedString',
        xml: `<si><t>${index === 0 ? 'Second' : 'Third'}</t></si>`,
      }
    })

    const merged = mergeWorkbookRichTextCells(
      [{ address: 'A1', text: 'First', storage: 'inlineString', xml: '<is><r><t>First</t></r></is>' }],
      appended,
    )

    expect(Array.isArray(merged)).toBe(true)
    expect(merged).toHaveLength(3)
    expect(materialized).toBe(0)
    expect(merged[2]).toEqual({ address: 'C1', text: 'Third', storage: 'sharedString', xml: '<si><t>Third</t></si>' })
    expect(materialized).toBe(1)
  })

  it('packs deferred shared-string indexes into float storage when a mixed sheet needs doubles', () => {
    const arena = new ImportedWorkbookArena()
    arena.addSharedStringCell({ sheetIndex: 0, row: 0, column: 0, sharedStringIndex: 1 })
    arena.addCell({ sheetIndex: 0, row: 0, column: 1, value: 42 })
    arena.addCell({ sheetIndex: 0, row: 0, column: 2, value: 42.5 })

    expect(arena.snapshot().integerValues).toEqual(new Int32Array([0, 42, 0]))
    expect(arena.snapshot().numberValues).toEqual(new Float64Array([1, Number.NaN, 42.5]))

    expect(
      arena.retainSharedStringReferences([
        { text: 'unused', rich: false },
        { text: 'Packed shared label', rich: false },
      ]),
    ).toEqual([])
    expect(arena.materializeSheetCells(0)).toEqual([
      { address: 'A1', value: 'Packed shared label' },
      { address: 'B1', value: 42 },
      { address: 'C1', value: 42.5 },
    ])
  })

  it('keeps earlier shared-string indexes readable when string cells prevent packing', () => {
    const arena = new ImportedWorkbookArena()
    arena.addCell({ sheetIndex: 0, row: 0, column: 0, value: 'Inline' })
    arena.addSharedStringCell({ sheetIndex: 0, row: 0, column: 1, sharedStringIndex: 1 })
    arena.addCell({ sheetIndex: 0, row: 0, column: 2, value: 42 })
    arena.addSharedStringCell({ sheetIndex: 0, row: 0, column: 3, sharedStringIndex: 2 })

    expect(
      arena.retainSharedStringReferences([
        { text: 'unused', rich: false },
        { text: 'First shared', rich: false },
        { text: 'Second shared', rich: false },
      ]),
    ).toEqual([])
    expect(arena.materializeSheetCells(0)).toEqual([
      { address: 'A1', value: 'Inline' },
      { address: 'B1', value: 'First shared' },
      { address: 'C1', value: 42 },
      { address: 'D1', value: 'Second shared' },
    ])
  })

  it('uses scalar sheet ownership for single-sheet arenas and falls back only for mixed ownership', () => {
    const arena = new ImportedWorkbookArena()
    arena.addCell({ sheetIndex: 4, row: 0, column: 0, value: 1 })

    expect(arena.snapshot().sheetIndex).toBe(4)
    expect(arena.snapshot().sheetIndexes).toBeUndefined()
    expect(arena.snapshot().columns).toBeInstanceOf(Uint16Array)
    expect(arena.snapshot().integerValues).toEqual(new Int32Array([1]))
    expect(arena.snapshot().numberValues).toBeUndefined()
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

  it('keeps near-dense row-major coordinates readable after sparse cells break the dense path', () => {
    const arena = new ImportedWorkbookArena()
    arena.reserveDenseRowMajorCellCapacity(0, 4, 3)

    arena.addCell({ sheetIndex: 0, row: 0, column: 0, value: 'A1' })
    arena.addCell({ sheetIndex: 0, row: 0, column: 1, value: 'B1' })
    arena.addCell({ sheetIndex: 0, row: 0, column: 3, value: 'D1' })
    arena.addCell({ sheetIndex: 0, row: 1, column: 0, value: 'A2' })

    const lazyCells = arena.createLazySheetCells(0)

    expect(lazyCells).toHaveLength(4)
    expect(lazyCells.map((cell) => cell.address)).toEqual(['A1', 'B1', 'D1', 'A2'])
    expect(arena.materializeSheetCells(0)).toEqual([
      { address: 'A1', value: 'A1' },
      { address: 'B1', value: 'B1' },
      { address: 'D1', value: 'D1' },
      { address: 'A2', value: 'A2' },
    ])
    expect(arena.sheetCellsAreDenseRowMajor(0, 4, 3)).toBe(false)

    const snapshot = arena.snapshot()
    expect(snapshot.rows).toEqual(new Uint32Array([0, 0, 0, 1]))
    expect(snapshot.columns).toEqual(new Uint16Array([0, 1, 3, 0]))
  })

  it('allocates number, string, boolean, and formula storage only when cells need those pools', () => {
    const stringOnlyArena = new ImportedWorkbookArena()
    stringOnlyArena.addCell({ sheetIndex: 0, row: 0, column: 0, value: 'Only text' })
    expect(stringOnlyArena.snapshot().numberValues).toBeUndefined()
    expect(stringOnlyArena.snapshot().integerValues).toBeUndefined()
    expect(stringOnlyArena.materializeSheetCells(0)).toEqual([{ address: 'A1', value: 'Only text' }])

    const arena = new ImportedWorkbookArena()
    const numericCell = arena.addCell({ sheetIndex: 0, row: 0, column: 0, value: 42 })

    expect(arena.snapshot().integerValues).toEqual(new Int32Array([42]))
    expect(arena.snapshot().numberValues).toBeUndefined()
    expect(arena.snapshot().stringIds).toBeUndefined()
    expect(arena.snapshot().booleanValues).toBeUndefined()
    expect(arena.snapshot().formulaIds).toBeUndefined()

    arena.addCell({ sheetIndex: 0, row: 1, column: 0, value: 42.25 })
    expect(arena.snapshot().numberValues).toEqual(new Float64Array([Number.NaN, 42.25]))

    arena.addCell({ sheetIndex: 0, row: 2, column: 0, value: 'Label' })
    expect(arena.snapshot().stringIds).toEqual(new Uint32Array([0xffffffff, 0xffffffff, 0]))
    expect(arena.snapshot().booleanValues).toBeUndefined()
    expect(arena.snapshot().formulaIds).toBeUndefined()

    arena.addCell({ sheetIndex: 0, row: 3, column: 0, value: true })
    expect(arena.snapshot().booleanValues).toEqual(new Uint8Array([0, 0, 0, 1]))

    arena.setFormula(numericCell, '1+1')
    expect(arena.snapshot().formulaIds).toEqual(new Uint32Array([0, 0xffffffff, 0xffffffff, 0xffffffff]))
  })

  it('keeps non-integer and negative-zero numbers in float storage for exact JS semantics', () => {
    const arena = new ImportedWorkbookArena()
    arena.addCell({ sheetIndex: 0, row: 0, column: 0, value: -0 })
    arena.addCell({ sheetIndex: 0, row: 0, column: 1, value: 2147483648 })
    arena.addCell({ sheetIndex: 0, row: 0, column: 2, value: -2147483649 })
    arena.addCell({ sheetIndex: 0, row: 0, column: 3, value: 2147483647 })

    expect(arena.snapshot().numberValues).toEqual(new Float64Array([-0, 2147483648, -2147483649, Number.NaN]))
    expect(arena.snapshot().integerValues).toEqual(new Int32Array([0, 0, 0, 2147483647]))
    expect(arena.materializeSheetCells(0)).toEqual([
      { address: 'A1', value: -0 },
      { address: 'B1', value: 2147483648 },
      { address: 'C1', value: -2147483649 },
      { address: 'D1', value: 2147483647 },
    ])
    expect(Object.is(arena.materializeSheetCells(0)[0]?.value, -0)).toBe(true)
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
