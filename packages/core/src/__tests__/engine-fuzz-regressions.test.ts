import { describe, expect, it } from 'vitest'
import { ErrorCode, ValueTag, type WorkbookConditionalFormatSnapshot } from '@bilig/protocol'
import { SpreadsheetEngine } from '../engine.js'
import {
  applyActionAndCaptureResult,
  createEngineSeedSnapshot,
  exportReplaySnapshot,
  normalizeSnapshotForSemanticComparison,
  type CoreAction,
} from './engine-fuzz-helpers.js'

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

  it('keeps pivot dimensions aligned after copy, row delete, and undo replay', async () => {
    const seedSnapshot = await createEngineSeedSnapshot('pivot-analytics', 'pivot-copy-delete-undo-regression')
    const engine = new SpreadsheetEngine({
      workbookName: seedSnapshot.workbook.name,
      replicaId: 'pivot-copy-delete-undo-regression',
    })
    await engine.ready()
    engine.importSnapshot(seedSnapshot)

    engine.copyRange(
      { sheetName: 'Sheet1', startAddress: 'D1', endAddress: 'E1' },
      { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B1' },
    )
    const pivotAfterCopy = engine.getPivotTable('Pivot', 'B2')
    expect(pivotAfterCopy).toMatchObject({ rows: 1, cols: 1 })

    engine.deleteRows('Sheet1', 0, 1)
    expect(engine.undo()).toBe(true)

    expect(engine.getPivotTable('Pivot', 'B2')).toMatchObject({
      rows: pivotAfterCopy?.rows,
      cols: pivotAfterCopy?.cols,
    })
  })

  it('keeps pivot source clears aligned across move undo and redo replay', async () => {
    const seedSnapshot = await createEngineSeedSnapshot('pivot-analytics', 'pivot-clear-move-redo-regression')
    const engine = new SpreadsheetEngine({
      workbookName: seedSnapshot.workbook.name,
      replicaId: 'pivot-clear-move-redo-regression',
    })
    await engine.ready()
    engine.importSnapshot(structuredClone(seedSnapshot))

    engine.insertRows('Sheet1', 0, 1)
    engine.moveRange(
      { sheetName: 'Sheet1', startAddress: 'E3', endAddress: 'E4' },
      { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A2' },
    )
    const pivotAfterMove = engine.getPivotTable('Pivot', 'B2')

    expect(pivotAfterMove).toMatchObject({ rows: 1, cols: 1 })
    expect(engine.undo()).toBe(true)
    expect(engine.getPivotTable('Pivot', 'B2')).toMatchObject({ rows: 3, cols: 2 })
    expect(engine.redo()).toBe(true)
    expect(engine.getPivotTable('Pivot', 'B2')).toMatchObject({
      rows: pivotAfterMove?.rows,
      cols: pivotAfterMove?.cols,
    })
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

  it('does not record history for duplicate formula writes', async () => {
    const seedSnapshot = await createEngineSeedSnapshot('blank', 'duplicate-formula-history-regression')
    const engine = new SpreadsheetEngine({
      workbookName: seedSnapshot.workbook.name,
      replicaId: 'duplicate-formula-history-regression',
    })
    await engine.ready()
    engine.importSnapshot(structuredClone(seedSnapshot))

    engine.setCellFormula('Sheet1', 'E3', 'F6/C3')
    const afterFirstWrite = engine.exportSnapshot()

    engine.setCellFormula('Sheet1', 'E3', 'F6/C3')
    expect(engine.exportSnapshot()).toEqual(afterFirstWrite)

    const sheetId = seedSnapshot.sheets[0]?.id
    if (sheetId === undefined) {
      throw new Error('Expected blank seed to include Sheet1 id')
    }
    engine.setCellFormulaAt(sheetId, 2, 4, 'F6/C3')
    expect(engine.exportSnapshot()).toEqual(afterFirstWrite)

    expect(engine.undo()).toBe(true)
    expect(engine.undo()).toBe(false)
  })

  it('rejects coordinate formula writes for unknown sheet ids', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'unknown-sheet-id-write-regression' })
    await engine.ready()

    expect(() => engine.setCellFormulaAt(999_999, 0, 0, 'A1')).toThrow('Unknown sheet id: 999999')
  })

  it('does not record history for duplicate conditional format writes', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'duplicate-conditional-format-history-regression' })
    await engine.ready()
    engine.createSheet('Sheet1')
    const format = {
      id: 'duplicate-cf',
      range: { sheetName: 'Sheet1', startAddress: 'B1', endAddress: 'B2' },
      rule: { kind: 'cellIs', operator: 'greaterThan', values: [10] },
      style: { fill: { backgroundColor: '#ff0000' } },
      priority: 1,
    } satisfies WorkbookConditionalFormatSnapshot

    engine.setConditionalFormat(format)
    const afterFirstWrite = engine.exportSnapshot()

    engine.setConditionalFormat(format)
    expect(engine.exportSnapshot()).toEqual(afterFirstWrite)
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

  it('captures family-deferred formula sources before structural delete undo replay', async () => {
    const seedSnapshot = await createEngineSeedSnapshot('cross-sheet-graph', 'family-deferred-formula-undo-regression')
    const engine = new SpreadsheetEngine({
      workbookName: seedSnapshot.workbook.name,
      replicaId: 'family-deferred-formula-undo-regression',
    })
    await engine.ready()
    engine.importSnapshot(structuredClone(seedSnapshot))

    engine.clearRange({ sheetName: 'Sheet1', startAddress: 'B3', endAddress: 'C3' })
    engine.insertColumns('Sheet1', 0, 1)
    expect(engine.getCell('Sheet1', 'D2').formula).toBe('C2*Summary!B2')

    engine.deleteColumns('Sheet1', 0, 1)
    expect(engine.getCell('Sheet1', 'C2').formula).toBe('B2*Summary!B2')

    expect(engine.undo()).toBe(true)
    expect(engine.getCell('Sheet1', 'D2').formula).toBe('C2*Summary!B2')

    expect(engine.undo()).toBe(true)
    expect(engine.getCell('Sheet1', 'C2').formula).toBe('B2*Summary!B2')

    expect(engine.undo()).toBe(true)
    expect(engine.exportSnapshot()).toEqual(seedSnapshot)
  })

  it('imports snapshots with structurally invalidated formula dependencies as ref errors', async () => {
    const seedSnapshot = await createEngineSeedSnapshot('formula-graph', 'invalid-range-dependency-import-regression')
    const engine = new SpreadsheetEngine({
      workbookName: seedSnapshot.workbook.name,
      replicaId: 'invalid-range-dependency-import-regression',
    })
    await engine.ready()
    engine.importSnapshot(structuredClone(seedSnapshot))

    engine.deleteColumns('Sheet1', 0, 1)
    engine.setCellFormula('Sheet1', 'A1', 'A1+A1')
    engine.deleteRows('Sheet1', 0, 1)
    engine.copyRange(
      { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' },
      { sheetName: 'Sheet1', startAddress: 'B1', endAddress: 'B1' },
    )
    engine.setRangeValues({ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' }, [['north']])
    engine.insertRows('Sheet1', 0, 1)
    engine.copyRange(
      { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' },
      { sheetName: 'Sheet1', startAddress: 'B1', endAddress: 'B1' },
    )

    const snapshot = engine.exportSnapshot()
    const restored = new SpreadsheetEngine({
      workbookName: snapshot.workbook.name,
      replicaId: 'invalid-range-dependency-import-restored',
    })
    await restored.ready()

    expect(() => restored.importSnapshot(structuredClone(snapshot))).not.toThrow()
  })

  it('restores structurally rewritten formula templates after refs collapse to #REF', async () => {
    const seedSnapshot = await createEngineSeedSnapshot('formula-graph', 'structural-template-ref-restore-regression')
    const engine = new SpreadsheetEngine({
      workbookName: seedSnapshot.workbook.name,
      replicaId: 'structural-template-ref-restore-regression',
    })
    await engine.ready()
    engine.importSnapshot(structuredClone(seedSnapshot))

    engine.deleteColumns('Sheet1', 0, 1)
    engine.setCellFormula('Sheet1', 'A1', 'A1+A1')
    engine.deleteColumns('Sheet1', 0, 1)
    engine.deleteRows('Sheet1', 0, 1)
    engine.setRangeNumberFormat({ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' }, '0.00')
    engine.fillRange(
      { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' },
      { sheetName: 'Sheet1', startAddress: 'B1', endAddress: 'B1' },
    )

    const snapshot = engine.exportSnapshot()
    expect(snapshot.sheets[0]?.cells).toContainEqual({ address: 'C2', formula: 'SUM(#REF!)' })

    const restored = new SpreadsheetEngine({
      workbookName: snapshot.workbook.name,
      replicaId: 'structural-template-ref-restore-regression-restored',
    })
    await restored.ready()
    restored.importSnapshot(snapshot)

    expect(restored.exportSnapshot()).toEqual(snapshot)
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

  it('restores explicit formats on deleted formula cells during structural undo', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'delete-formula-format-undo-regression' })
    await engine.ready()
    engine.createSheet('Sheet1')

    engine.setRangeNumberFormat({ sheetName: 'Sheet1', startAddress: 'B4', endAddress: 'C4' }, '0.00')
    engine.fillRange(
      { sheetName: 'Sheet1', startAddress: 'C3', endAddress: 'D4' },
      { sheetName: 'Sheet1', startAddress: 'E4', endAddress: 'F5' },
    )
    engine.deleteRows('Sheet1', 0, 1)
    engine.setCellFormula('Sheet1', 'E4', 'A1+A1')

    expect(engine.getCell('Sheet1', 'E4').format).toBe('0.00')

    engine.deleteRows('Sheet1', 2, 2)

    expect(engine.undo()).toBe(true)
    expect(engine.getCell('Sheet1', 'E4')).toMatchObject({
      formula: 'A1+A1',
      format: '0.00',
    })
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

  it('preserves shifted range sum precision after CSV roundtrip import', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'csv-shifted-sum-precision-regression' })
    await engine.ready()
    engine.createSheet('Sheet1')

    engine.setCellFormula('Sheet1', 'A1', 'SUM(A5:A5)')
    engine.setCellFormula('Sheet1', 'A2', 'A5*A5')
    engine.setCellFormula('Sheet1', 'A3', 'IF(A5>0,"text:yes","text:no")')
    engine.setCellFormula('Sheet1', 'A4', 'IF(A5>0,"text:yes","text:no")')
    engine.setCellValue('Sheet1', 'A5', 1429783918)

    const restored = new SpreadsheetEngine({ workbookName: 'csv-shifted-sum-precision-regression-restored' })
    await restored.ready()
    restored.createSheet('Sheet1')
    restored.importSheetCsv('Sheet1', engine.exportSheetCsv('Sheet1'))

    expect(restored.getCell('Sheet1', 'A1').value).toEqual({
      tag: ValueTag.Number,
      value: 1429783918,
    })
  })

  it('preserves narrow SUM range precision after CSV roundtrip import', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'csv-sum-prefix-precision-regression' })
    await engine.ready()
    engine.createSheet('Sheet1')

    engine.setCellFormula('Sheet1', 'A1', 'SUM(A5:A5)')
    engine.setCellFormula('Sheet1', 'A2', 'A5+A5')
    engine.setCellFormula('Sheet1', 'A3', 'A5+A5')
    engine.setCellFormula('Sheet1', 'A4', 'A5*A5')
    engine.setRangeValues({ sheetName: 'Sheet1', startAddress: 'A5', endAddress: 'A5' }, [[663897248]])

    const restored = new SpreadsheetEngine({ workbookName: 'csv-sum-prefix-precision-regression-restored' })
    await restored.ready()
    restored.createSheet('Sheet1')
    restored.importSheetCsv('Sheet1', engine.exportSheetCsv('Sheet1'))

    expect(restored.getCell('Sheet1', 'A1').value).toEqual({
      tag: ValueTag.Number,
      value: 663897248,
    })
  })

  it('restores snapshots when a formula runtime image contains a stale template id', async () => {
    const seedSnapshot = await createEngineSeedSnapshot('formula-graph', 'stale-template-runtime-image-regression')
    const engine = new SpreadsheetEngine({
      workbookName: seedSnapshot.workbook.name,
      replicaId: 'stale-template-runtime-image-primary',
    })
    await engine.ready()
    engine.importSnapshot(structuredClone(seedSnapshot))

    const actions: CoreAction[] = [
      { kind: 'deleteColumns', start: 0, count: 1 },
      { kind: 'formula', address: 'A1', formula: 'A1+A1' },
      { kind: 'deleteRows', start: 0, count: 1 },
      {
        kind: 'format',
        range: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' },
        format: '0.00',
      },
      {
        kind: 'clear',
        range: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' },
      },
      {
        kind: 'style',
        range: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' },
        patch: { fill: { backgroundColor: '#dbeafe' } },
      },
    ]
    actions.forEach((action) => applyActionAndCaptureResult(engine, action))

    const snapshot = engine.exportSnapshot()
    const restored = new SpreadsheetEngine({
      workbookName: seedSnapshot.workbook.name,
      replicaId: 'stale-template-runtime-image-restored',
    })
    await restored.ready()

    expect(() => restored.importSnapshot(structuredClone(snapshot))).not.toThrow()
    expect(normalizeSnapshotForSemanticComparison(restored.exportSnapshot())).toEqual(normalizeSnapshotForSemanticComparison(snapshot))
  })

  it('normalizes sparse style metadata by covered cells after undo replay', async () => {
    const seedSnapshot = await createEngineSeedSnapshot('sparse-format', 'history-style-run-normalization-regression')
    const engine = new SpreadsheetEngine({
      workbookName: seedSnapshot.workbook.name,
      replicaId: 'history-style-run-normalization',
    })
    await engine.ready()
    engine.importSnapshot(structuredClone(seedSnapshot))

    const applied: CoreAction[] = []
    const applyAccepted = (action: CoreAction) => {
      const result = applyActionAndCaptureResult(engine, action)
      if (result.accepted) {
        applied.push(action)
      }
    }

    applyAccepted({ kind: 'insertRows', start: 0, count: 1 })
    applyAccepted({
      kind: 'style',
      range: { sheetName: 'Sheet1', startAddress: 'D2', endAddress: 'E3' },
      patch: { alignment: { horizontal: 'right', wrap: true } },
    })
    applyAccepted({ kind: 'insertRows', start: 0, count: 1 })
    applyAccepted({ kind: 'deleteColumns', start: 0, count: 1 })

    expect(engine.undo()).toBe(true)
    expect(applied.pop()?.kind).toBe('deleteColumns')

    applyAccepted({
      kind: 'style',
      range: { sheetName: 'Sheet1', startAddress: 'F3', endAddress: 'F3' },
      patch: { alignment: { horizontal: 'right', wrap: true } },
    })

    const expectedSnapshot = await exportReplaySnapshot(seedSnapshot, applied)

    expect(normalizeSnapshotForSemanticComparison(engine.exportSnapshot())).toEqual(
      normalizeSnapshotForSemanticComparison(expectedSnapshot),
    )
  })
})
