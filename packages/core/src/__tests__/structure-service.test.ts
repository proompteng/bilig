import { Effect } from 'effect'
import { describe, expect, it } from 'vitest'
import { ValueTag } from '@bilig/protocol'
import { SpreadsheetEngine } from '../engine.js'
import type { EngineStructureService } from '../engine/services/structure-service.js'

function isEngineStructureService(value: unknown): value is EngineStructureService {
  if (typeof value !== 'object' || value === null) {
    return false
  }
  return (
    typeof Reflect.get(value, 'captureSheetCellState') === 'function' &&
    typeof Reflect.get(value, 'captureRowRangeCellState') === 'function' &&
    typeof Reflect.get(value, 'captureColumnRangeCellState') === 'function' &&
    typeof Reflect.get(value, 'applyStructuralAxisOp') === 'function'
  )
}

function getStructureService(engine: SpreadsheetEngine): EngineStructureService {
  const runtime = Reflect.get(engine, 'runtime')
  if (typeof runtime !== 'object' || runtime === null) {
    throw new TypeError('Expected engine runtime')
  }
  const structure = Reflect.get(runtime, 'structure')
  if (!isEngineStructureService(structure)) {
    throw new TypeError('Expected engine structure service')
  }
  return structure
}

describe('EngineStructureService', () => {
  it('captures sheet cell state in row-major order for undo reconstruction', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'structure-capture' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'B2', 20)
    engine.setCellFormat('Sheet1', 'B2', '0.00')
    engine.setCellFormula('Sheet1', 'A3', 'B2*2')

    const ops = Effect.runSync(getStructureService(engine).captureSheetCellState('Sheet1'))

    expect(ops).toEqual([
      { kind: 'setCellValue', sheetName: 'Sheet1', address: 'B2', value: 20 },
      { kind: 'setCellFormat', sheetName: 'Sheet1', address: 'B2', format: '0.00' },
      { kind: 'setCellFormula', sheetName: 'Sheet1', address: 'A3', formula: 'B2*2' },
    ])
  })

  it('rewrites metadata-backed ranges and formula bindings across row inserts', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'structure-rewrite' })
    await engine.ready()
    engine.createSheet('Data')
    engine.createSheet('Pivot')
    engine.setRangeValues({ sheetName: 'Data', startAddress: 'A1', endAddress: 'B4' }, [
      ['Region', 'Sales'],
      ['East', 10],
      ['West', 7],
      ['East', 5],
    ])
    engine.setDefinedName('SalesRange', '=Data!A1:B4')
    engine.setCellFormula('Pivot', 'E2', 'SUM(Data!B1:B4)')
    engine.setFreezePane('Data', 1, 0)
    engine.setFilter('Data', { sheetName: 'Data', startAddress: 'A1', endAddress: 'B4' })
    engine.setSort('Data', { sheetName: 'Data', startAddress: 'A1', endAddress: 'B4' }, [{ keyAddress: 'B1', direction: 'asc' }])
    engine.setTable({
      name: 'Sales',
      sheetName: 'Data',
      startAddress: 'A1',
      endAddress: 'B4',
      columnNames: ['Region', 'Sales'],
      headerRow: true,
      totalsRow: false,
    })
    engine.setPivotTable('Pivot', 'B2', {
      name: 'SalesPivot',
      source: { sheetName: 'Data', startAddress: 'A1', endAddress: 'B4' },
      groupBy: ['Region'],
      values: [{ sourceColumn: 'Sales', summarizeBy: 'sum' }],
    })

    const result = Effect.runSync(
      getStructureService(engine).applyStructuralAxisOp({
        kind: 'insertRows',
        sheetName: 'Data',
        start: 0,
        count: 1,
      }),
    )

    expect(result.changedCellIndices).toEqual([])
    expect(engine.getDefinedName('SalesRange')).toEqual({
      name: 'SalesRange',
      value: '=Data!A2:B5',
    })
    expect(engine.getCell('Pivot', 'E2').formula).toBe('SUM(Data!B2:B5)')
    expect(engine.getFreezePane('Data')).toEqual({ sheetName: 'Data', rows: 2, cols: 0 })
    expect(engine.getFilters('Data')).toEqual([{ sheetName: 'Data', range: { sheetName: 'Data', startAddress: 'A2', endAddress: 'B5' } }])
    expect(engine.getSorts('Data')).toEqual([
      {
        sheetName: 'Data',
        range: { sheetName: 'Data', startAddress: 'A2', endAddress: 'B5' },
        keys: [{ keyAddress: 'B2', direction: 'asc' }],
      },
    ])
    expect(engine.getTables()).toEqual([
      {
        name: 'Sales',
        sheetName: 'Data',
        startAddress: 'A2',
        endAddress: 'B5',
        columnNames: ['Region', 'Sales'],
        headerRow: true,
        totalsRow: false,
      },
    ])
    expect(engine.getPivotTable('Pivot', 'B2')?.source).toEqual({
      sheetName: 'Data',
      startAddress: 'A2',
      endAddress: 'B5',
    })
  })

  it('rewrites range-backed defined names across column inserts', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'structure-defined-range-rewrite' })
    await engine.ready()
    engine.createSheet('Data')
    engine.setRangeValues({ sheetName: 'Data', startAddress: 'A1', endAddress: 'B3' }, [
      ['Qty', 'Amount'],
      [1, 10],
      [2, 20],
    ])
    engine.setDefinedName('SalesRange', {
      kind: 'range-ref',
      sheetName: 'Data',
      startAddress: 'A1',
      endAddress: 'B3',
    })
    engine.setCellFormula('Data', 'C1', 'SUM(SalesRange)')

    Effect.runSync(
      getStructureService(engine).applyStructuralAxisOp({
        kind: 'insertColumns',
        sheetName: 'Data',
        start: 0,
        count: 1,
      }),
    )

    expect(engine.getDefinedName('SalesRange')).toEqual({
      name: 'SalesRange',
      value: {
        kind: 'range-ref',
        sheetName: 'Data',
        startAddress: 'B1',
        endAddress: 'C3',
      },
    })
    expect(engine.getCell('Data', 'D1').formula).toBe('SUM(SalesRange)')
    expect(engine.getCellValue('Data', 'D1')).toMatchObject({ tag: 1, value: 33 })
  })

  it('rewrites formulas and axis identities across column deletes and moves through the service', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'structure-delete-move-columns' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 2)
    engine.setCellValue('Sheet1', 'B1', 3)
    engine.setCellValue('Sheet1', 'C1', 5)
    engine.setCellFormula('Sheet1', 'E1', 'SUM(A1:B1)')
    engine.updateColumnMetadata('Sheet1', 0, 1, 90, true)

    Effect.runSync(
      getStructureService(engine).applyStructuralAxisOp({
        kind: 'deleteColumns',
        sheetName: 'Sheet1',
        start: 0,
        count: 1,
      }),
    )

    expect(engine.getCell('Sheet1', 'D1').formula).toBe('SUM(A1:A1)')
    expect(engine.getColumnAxisEntries('Sheet1')).toEqual([])

    engine.updateColumnMetadata('Sheet1', 1, 1, 110, false)
    engine.setCellFormula('Sheet1', 'D2', 'B1')

    Effect.runSync(
      getStructureService(engine).applyStructuralAxisOp({
        kind: 'moveColumns',
        sheetName: 'Sheet1',
        start: 1,
        count: 1,
        target: 0,
      }),
    )

    expect(engine.getCell('Sheet1', 'D2').formula).toBe('A1')
    expect(engine.getColumnAxisEntries('Sheet1')).toEqual([{ id: expect.any(String), index: 0, size: 110, hidden: false }])
  })

  it('plans the structural transaction before mutating workbook axis metadata', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'structure-transaction-order' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 1)

    const workbook = engine.workbook
    if (typeof workbook !== 'object' || workbook === null) {
      throw new TypeError('Expected workbook store')
    }

    const events: string[] = []
    const originalPlanStructuralAxisTransform = Reflect.get(workbook, 'planStructuralAxisTransform')
    const originalApplyPlannedStructuralTransaction = Reflect.get(workbook, 'applyPlannedStructuralTransaction')
    const originalInsertRows = Reflect.get(workbook, 'insertRows')
    if (
      typeof originalPlanStructuralAxisTransform !== 'function' ||
      typeof originalApplyPlannedStructuralTransaction !== 'function' ||
      typeof originalInsertRows !== 'function'
    ) {
      throw new TypeError('Expected structural workbook methods')
    }

    Reflect.set(workbook, 'planStructuralAxisTransform', (...args: unknown[]) => {
      events.push('planStructuralAxisTransform')
      return originalPlanStructuralAxisTransform.apply(workbook, args)
    })
    Reflect.set(workbook, 'applyPlannedStructuralTransaction', (...args: unknown[]) => {
      events.push('applyPlannedStructuralTransaction')
      return originalApplyPlannedStructuralTransaction.apply(workbook, args)
    })
    Reflect.set(workbook, 'insertRows', (...args: unknown[]) => {
      events.push('insertRows')
      return originalInsertRows.apply(workbook, args)
    })

    Effect.runSync(
      getStructureService(engine).applyStructuralAxisOp({
        kind: 'insertRows',
        sheetName: 'Sheet1',
        start: 0,
        count: 1,
      }),
    )

    expect(events).toEqual(['planStructuralAxisTransform', 'insertRows', 'applyPlannedStructuralTransaction'])
  })

  it('keeps repeated direct aggregate row inserts off the topology and dirty-formula path', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'structure-preserve-row-aggregates' })
    await engine.ready()
    engine.createSheet('Sheet1')
    for (let row = 1; row <= 4; row += 1) {
      engine.setCellValue('Sheet1', `A${row}`, row)
      engine.setCellFormula('Sheet1', `B${row}`, `SUM(A1:A${row})`)
    }

    const result = Effect.runSync(
      getStructureService(engine).applyStructuralAxisOp({
        kind: 'insertRows',
        sheetName: 'Sheet1',
        start: 1,
        count: 1,
      }),
    )

    expect(result.topologyChanged).toBe(false)
    expect(result.formulaCellIndices).toEqual([])
    expect(engine.getCell('Sheet1', 'B1').formula).toBe('SUM(A1:A1)')
    expect(engine.getCell('Sheet1', 'B3').formula).toBe('SUM(A1:A3)')
    expect(engine.getCell('Sheet1', 'B5').formula).toBe('SUM(A1:A5)')
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Sheet1', 'B3')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(engine.getCellValue('Sheet1', 'B5')).toEqual({ tag: ValueTag.Number, value: 10 })
  })

  it('avoids full graph refresh when a structural delete only removes preserved aggregate nodes', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'structure-delete-row-aggregates' })
    await engine.ready()
    engine.createSheet('Sheet1')
    for (let row = 1; row <= 4; row += 1) {
      engine.setCellValue('Sheet1', `A${row}`, row)
      engine.setCellFormula('Sheet1', `B${row}`, `SUM(A1:A${row})`)
    }

    engine.resetPerformanceCounters()
    const result = Effect.runSync(
      getStructureService(engine).applyStructuralAxisOp({
        kind: 'deleteRows',
        sheetName: 'Sheet1',
        start: 1,
        count: 1,
      }),
    )

    expect(result.topologyChanged).toBe(true)
    expect(result.graphRefreshRequired).toBe(false)
    expect(result.formulaCellIndices).toEqual([])
    expect(engine.getCellValue('Sheet1', 'B3')).toEqual({ tag: ValueTag.Number, value: 8 })
    expect(engine.getCell('Sheet1', 'B3').formula).toBe('SUM(A1:A3)')
    expect(engine.getPerformanceCounters().cycleFormulaScans).toBe(0)
  })

  it('does not inspect unaffected prefix aggregate formulas above a deleted row', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'structure-delete-row-prefix-aggregates' })
    await engine.ready()
    engine.createSheet('Sheet1')
    for (let row = 1; row <= 6; row += 1) {
      engine.setCellValue('Sheet1', `A${row}`, row)
      engine.setCellFormula('Sheet1', `B${row}`, `SUM(A1:A${row})`)
    }

    engine.resetPerformanceCounters()
    const result = Effect.runSync(
      getStructureService(engine).applyStructuralAxisOp({
        kind: 'deleteRows',
        sheetName: 'Sheet1',
        start: 3,
        count: 1,
      }),
    )

    expect(result.graphRefreshRequired).toBe(false)
    expect(engine.getPerformanceCounters().structuralFormulaImpactCandidates).toBe(0)
    expect(result.formulaCellIndices).toEqual([])
    expect(engine.getCellValue('Sheet1', 'B5')).toEqual({ tag: ValueTag.Number, value: 17 })
  })

  it('deletes direct aggregate rows without forcing a full wasm program upload', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'structure-delete-row-aggregate-runtime-patch' })
    await engine.ready()
    engine.createSheet('Sheet1')
    for (let row = 1; row <= 6; row += 1) {
      engine.setCellValue('Sheet1', `A${row}`, row)
      engine.setCellFormula('Sheet1', `B${row}`, `SUM(A1:A${row})`)
    }

    engine.resetPerformanceCounters()
    engine.deleteRows('Sheet1', 2, 1)

    expect(engine.getCell('Sheet1', 'B5').formula).toBe('SUM(A1:A5)')
    expect(engine.getCellValue('Sheet1', 'B5')).toEqual({ tag: ValueTag.Number, value: 18 })
    expect(engine.getPerformanceCounters().wasmFullUploads).toBe(0)
    expect(engine.getPerformanceCounters().calcChainFullScans).toBe(0)
    expect(engine.getPerformanceCounters().structuralFormulaImpactCandidates).toBe(0)
  })

  it('does not scan cycle state for literal-only structural deletes', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'structure-delete-literal-column' })
    await engine.ready()
    engine.createSheet('Sheet1')
    for (let row = 1; row <= 8; row += 1) {
      engine.setCellValue('Sheet1', `A${row}`, row)
      engine.setCellValue('Sheet1', `B${row}`, row * 2)
      engine.setCellFormula('Sheet1', `C${row}`, `A${row}+B${row}`)
      engine.setCellFormula('Sheet1', `D${row}`, `C${row}*2`)
    }

    engine.resetPerformanceCounters()
    const result = Effect.runSync(
      getStructureService(engine).applyStructuralAxisOp({
        kind: 'deleteColumns',
        sheetName: 'Sheet1',
        start: 1,
        count: 1,
      }),
    )

    expect(result.graphRefreshRequired).toBe(false)
    expect(engine.getPerformanceCounters().cycleFormulaScans).toBe(0)
  })

  it('keeps graph refresh enabled when deleting formulas from an active cycle', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'structure-delete-cycle' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellFormula('Sheet1', 'A1', 'B1')
    engine.setCellFormula('Sheet1', 'B1', 'A1')

    const result = Effect.runSync(
      getStructureService(engine).applyStructuralAxisOp({
        kind: 'deleteRows',
        sheetName: 'Sheet1',
        start: 0,
        count: 1,
      }),
    )

    expect(result.topologyChanged).toBe(true)
    expect(result.graphRefreshRequired).toBe(true)
  })

  it('keeps repeated direct aggregate row moves off the graph refresh path', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'structure-move-row-aggregates' })
    await engine.ready()
    engine.createSheet('Sheet1')
    for (let row = 1; row <= 4; row += 1) {
      engine.setCellValue('Sheet1', `A${row}`, row)
      engine.setCellFormula('Sheet1', `B${row}`, `SUM(A1:A${row})`)
    }

    engine.resetPerformanceCounters()
    const result = Effect.runSync(
      getStructureService(engine).applyStructuralAxisOp({
        kind: 'moveRows',
        sheetName: 'Sheet1',
        start: 2,
        count: 1,
        target: 0,
      }),
    )

    expect(result.topologyChanged).toBe(false)
    expect(result.graphRefreshRequired).toBe(false)
    expect(engine.getCell('Sheet1', 'B1').formula).toBe('SUM(A1:A3)')
    expect(engine.getCell('Sheet1', 'B2').formula).toBe('SUM(A2:A2)')
    expect(engine.getCell('Sheet1', 'B3').formula).toBe('SUM(A2:A3)')
    expect(engine.getCell('Sheet1', 'B4').formula).toBe('SUM(A1:A4)')
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 6 })
    expect(engine.getCellValue('Sheet1', 'B2')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Sheet1', 'B3')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(engine.getCellValue('Sheet1', 'B4')).toEqual({ tag: ValueTag.Number, value: 10 })
    expect(result.formulaCellIndices).toEqual([])
    expect(engine.getPerformanceCounters().structuralFormulaImpactCandidates).toBe(0)
  })

  it('keeps repeated simple column families off the topology and dirty-formula path for inserts', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'structure-preserve-column-families' })
    await engine.ready()
    engine.createSheet('Sheet1')
    for (let row = 1; row <= 4; row += 1) {
      engine.setCellValue('Sheet1', `A${row}`, row)
      engine.setCellValue('Sheet1', `B${row}`, row * 2)
      engine.setCellFormula('Sheet1', `C${row}`, `A${row}+B${row}`)
      engine.setCellFormula('Sheet1', `D${row}`, `C${row}*2`)
    }

    const result = Effect.runSync(
      getStructureService(engine).applyStructuralAxisOp({
        kind: 'insertColumns',
        sheetName: 'Sheet1',
        start: 1,
        count: 1,
      }),
    )

    expect(result.topologyChanged).toBe(false)
    expect(result.formulaCellIndices).toEqual([])
    for (let row = 1; row <= 4; row += 1) {
      expect(engine.getCell('Sheet1', `D${row}`).formula).toBe(`A${row}+C${row}`)
      expect(engine.getCell('Sheet1', `E${row}`).formula).toBe(`D${row}*2`)
      expect(engine.getCellValue('Sheet1', `D${row}`)).toEqual({ tag: ValueTag.Number, value: row * 3 })
      expect(engine.getCellValue('Sheet1', `E${row}`)).toEqual({ tag: ValueTag.Number, value: row * 6 })
    }
  })

  it('keeps surviving range-backed row deletes off the topology churn path', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'structure-preserve-range-delete' })
    await engine.ready()
    engine.createSheet('Sheet1')
    for (let row = 1; row <= 4; row += 1) {
      engine.setCellValue('Sheet1', `A${row}`, row)
      engine.setCellValue('Sheet1', `B${row}`, row * 10)
      engine.setCellValue('Sheet1', `C${row}`, row * 100)
    }
    engine.setCellFormula('Sheet1', 'D5', 'SUM(A1:B4)+C4')

    const result = Effect.runSync(
      getStructureService(engine).applyStructuralAxisOp({
        kind: 'deleteRows',
        sheetName: 'Sheet1',
        start: 1,
        count: 1,
      }),
    )

    expect(result.topologyChanged).toBe(false)
    expect(engine.getCell('Sheet1', 'D4').formula).toBe('SUM(A1:B3)+C3')
    expect(result.formulaCellIndices).toEqual([engine.workbook.getCellIndex('Sheet1', 'D4')!])
  })
})
