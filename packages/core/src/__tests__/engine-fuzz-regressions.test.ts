import { describe, expect, it } from 'vitest'
import { ErrorCode, ValueTag } from '@bilig/protocol'
import { SpreadsheetEngine } from '../engine.js'
import { createEngineSeedSnapshot } from './engine-fuzz-helpers.js'

describe('engine fuzz regressions', () => {
  it('does not mutate source cells when moveRange is blocked by protection', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'move-protection-regression' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setRangeValues({ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'C5' }, [
      ['Id', 'Status', 'Amount'],
      [1, 'Draft', 10],
      [2, 'Final', 20],
      [3, 'Review', 30],
      [4, 'Final', 40],
    ])
    engine.setRangeProtection({
      id: 'protect-a1-c4',
      range: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'C4' },
      hideFormulas: true,
    })

    const before = engine.exportSnapshot()
    expect(() =>
      engine.moveRange(
        { sheetName: 'Sheet1', startAddress: 'A5', endAddress: 'A6' },
        { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A2' },
      ),
    ).toThrow(/Failed to execute local transaction/)
    expect(engine.exportSnapshot()).toEqual(before)
  })

  it('restores pivot materialization dimensions after undoing source clears', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'pivot-undo-regression' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.createSheet('Pivot')
    engine.setRangeValues({ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'C5' }, [
      ['Region', 'Quarter', 'Sales'],
      ['East', 'Q1', 10],
      ['West', 'Q1', 6],
      ['East', 'Q2', 12],
      ['West', 'Q2', 9],
    ])
    engine.setTable({
      name: 'QuarterlySales',
      sheetName: 'Sheet1',
      startAddress: 'A1',
      endAddress: 'C5',
      columnNames: ['Region', 'Quarter', 'Sales'],
      headerRow: true,
      totalsRow: false,
    })
    engine.setPivotTable('Pivot', 'B2', {
      name: 'QuarterlyPivot',
      source: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'C5' },
      groupBy: ['Region'],
      values: [{ sourceColumn: 'Sales', summarizeBy: 'sum' }],
    })

    expect(engine.getPivotTable('Pivot', 'B2')?.rows).toBe(3)
    engine.clearRange({ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A2' })
    expect(engine.getPivotTable('Pivot', 'B2')?.rows).toBe(1)
    expect(engine.undo()).toBe(true)
    expect(engine.getPivotTable('Pivot', 'B2')?.rows).toBe(3)
  })

  it('treats structural deletes on blank sheets as history no-ops', async () => {
    const seed = new SpreadsheetEngine({ workbookName: 'blank-structural-noop-seed' })
    await seed.ready()
    seed.createSheet('Sheet1')
    const snapshot = seed.exportSnapshot()

    const engine = new SpreadsheetEngine({ workbookName: 'blank-structural-noop' })
    await engine.ready()
    engine.importSnapshot(snapshot)

    const before = engine.exportSnapshot()
    engine.deleteColumns('Sheet1', 0, 1)
    expect(engine.exportSnapshot()).toEqual(before)
    expect(engine.undo()).toBe(false)
  })

  it('restores sparse style metadata after undoing coalesced structural deletes', async () => {
    const seedSnapshot = await createEngineSeedSnapshot('sparse-format', 'structural-style-undo-regression')
    const engine = new SpreadsheetEngine({ workbookName: seedSnapshot.workbook.name, replicaId: 'structural-style-undo-regression' })
    await engine.ready()
    engine.importSnapshot(structuredClone(seedSnapshot))

    engine.setRangeStyle({ sheetName: 'Sheet1', startAddress: 'C5', endAddress: 'C5' }, { fill: { backgroundColor: '#dbeafe' } })
    engine.deleteRows('Sheet1', 3, 1)
    engine.deleteColumns('Sheet1', 0, 1)

    expect(engine.undo()).toBe(true)
    expect(engine.undo()).toBe(true)
    expect(engine.undo()).toBe(true)
    expect(engine.exportSnapshot()).toEqual(seedSnapshot)
  })

  it('restores sparse style range shapes after undoing a partial clear', async () => {
    const seedSnapshot = await createEngineSeedSnapshot('sparse-format', 'sparse-style-clear-undo-regression')
    const engine = new SpreadsheetEngine({
      workbookName: seedSnapshot.workbook.name,
      replicaId: 'sparse-style-clear-undo-regression',
    })
    await engine.ready()
    engine.importSnapshot(structuredClone(seedSnapshot))

    engine.setRangeStyle({ sheetName: 'Sheet1', startAddress: 'C4', endAddress: 'D4' }, { fill: { backgroundColor: '#dbeafe' } })
    const styledSnapshot = engine.exportSnapshot()

    engine.clearRangeStyle({ sheetName: 'Sheet1', startAddress: 'D3', endAddress: 'D4' })

    expect(engine.undo()).toBe(true)
    expect(engine.exportSnapshot()).toEqual(styledSnapshot)
  })

  it('rebinds structurally rewritten formulas when dependency addresses shift after prior ref errors', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'structural-ref-error-rebind-regression' })
    await engine.ready()
    engine.createSheet('Sheet1')

    engine.setCellFormula('Sheet1', 'A1', 'C3+D4')
    engine.deleteRows('Sheet1', 2, 1)
    engine.insertColumns('Sheet1', 0, 1)

    expect(engine.getCell('Sheet1', 'B1').formula).toBe('#REF!+E3')

    engine.deleteColumns('Sheet1', 4, 1)

    expect(engine.getCell('Sheet1', 'B1').formula).toBe('#REF!+#REF!')
    expect(engine.undo()).toBe(true)
    expect(engine.getCell('Sheet1', 'B1').formula).toBe('#REF!+E3')
  })

  it('restores formula graphs after undoing mixed row inserts and column deletes', async () => {
    const seedSnapshot = await createEngineSeedSnapshot('formula-graph', 'structural-formula-undo-regression')
    const engine = new SpreadsheetEngine({
      workbookName: seedSnapshot.workbook.name,
      replicaId: 'structural-formula-undo-regression',
    })
    await engine.ready()
    engine.importSnapshot(structuredClone(seedSnapshot))

    engine.insertRows('Sheet1', 0, 1)
    engine.deleteColumns('Sheet1', 0, 1)
    engine.insertRows('Sheet1', 0, 1)

    expect(engine.undo()).toBe(true)
    expect(engine.undo()).toBe(true)
    expect(engine.undo()).toBe(true)
    expect(engine.exportSnapshot()).toEqual(seedSnapshot)
  })

  it('clears target formats when moving empty cells over formatted blanks and keeps undo aligned', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'move-empty-format-regression' })
    await engine.ready()
    engine.createSheet('Sheet1')

    engine.setRangeNumberFormat({ sheetName: 'Sheet1', startAddress: 'B1', endAddress: 'B2' }, '0.00')
    engine.fillRange(
      { sheetName: 'Sheet1', startAddress: 'B2', endAddress: 'C3' },
      { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B2' },
    )

    expect(engine.getCell('Sheet1', 'A1').format).toBe('0.00')

    engine.moveRange(
      { sheetName: 'Sheet1', startAddress: 'B4', endAddress: 'C4' },
      { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B1' },
    )

    expect(engine.getCell('Sheet1', 'A1').format).toBeUndefined()
    expect(engine.exportSnapshot().sheets[0]?.cells ?? []).toEqual([])

    expect(engine.undo()).toBe(true)
    expect(engine.getCell('Sheet1', 'A1').format).toBe('0.00')
  })

  it('propagates cycle errors to dependent formulas after direct formula writes', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'cycle-dependent-regression' })
    await engine.ready()
    engine.createSheet('Sheet1')

    engine.setCellFormula('Sheet1', 'A1', 'SUM(B1:B1)')
    engine.setCellFormula('Sheet1', 'C1', 'SUM(B1:C2)')
    engine.setCellFormula('Sheet1', 'A2', 'SUM(B1:C2)')

    expect(engine.getCell('Sheet1', 'C1').value).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Cycle,
    })
    expect(engine.getCell('Sheet1', 'A2').value).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Cycle,
    })
  })

  it('propagates cycle errors through range dependents after direct formula writes', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'cycle-range-dependent-regression' })
    await engine.ready()
    engine.createSheet('Sheet1')

    engine.setCellValue('Sheet1', 'A1', 1)
    engine.setCellValue('Sheet1', 'B1', 'text')
    engine.setCellValue('Sheet1', 'C2', 'text')
    engine.setCellFormula('Sheet1', 'B3', 'SUM(A1:C3)')
    engine.setCellFormula('Sheet1', 'E4', 'SUM(B1:C4)')

    expect(engine.getCell('Sheet1', 'B3').value).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Cycle,
    })
    expect(engine.getCell('Sheet1', 'E4').value).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Cycle,
    })
  })

  it('preserves cycle errors for self-referential range formulas after CSV roundtrip import', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'csv-cycle-roundtrip-regression' })
    await engine.ready()
    engine.createSheet('Sheet1')

    engine.setRangeValues({ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' }, [[false]])
    engine.setCellFormula('Sheet1', 'A2', 'A1+A4')
    engine.setCellFormula('Sheet1', 'A3', 'SUM(A1:A4)')
    engine.setRangeValues({ sheetName: 'Sheet1', startAddress: 'A4', endAddress: 'A4' }, [['text:lB<`x']])
    engine.setCellFormula('Sheet1', 'A5', 'A1+A1')

    expect(engine.getCell('Sheet1', 'A3').value).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Cycle,
    })

    const restored = new SpreadsheetEngine({ workbookName: 'csv-cycle-roundtrip-regression-restored' })
    await restored.ready()
    restored.createSheet('Sheet1')
    restored.importSheetCsv('Sheet1', engine.exportSheetCsv('Sheet1'))

    expect(restored.getCell('Sheet1', 'A3').value).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Cycle,
    })
  })
})
