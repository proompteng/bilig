import { Effect } from 'effect'
import { describe, expect, it } from 'vitest'
import { ValueTag } from '@bilig/protocol'
import { SpreadsheetEngine } from '../engine.js'
import { WorkbookStore } from '../workbook-store.js'
import { createEngineCellStateService, type EngineCellStateService } from '../engine/services/cell-state-service.js'

function isEngineCellStateService(value: unknown): value is EngineCellStateService {
  if (typeof value !== 'object' || value === null) {
    return false
  }
  return (
    typeof Reflect.get(value, 'restoreCellOps') === 'function' &&
    typeof Reflect.get(value, 'readRangeCells') === 'function' &&
    typeof Reflect.get(value, 'toCellStateOps') === 'function'
  )
}

function getCellStateService(engine: SpreadsheetEngine): EngineCellStateService {
  const runtime = Reflect.get(engine, 'runtime')
  if (typeof runtime !== 'object' || runtime === null) {
    throw new TypeError('Expected engine runtime')
  }
  const cellState = Reflect.get(runtime, 'cellState')
  if (!isEngineCellStateService(cellState)) {
    throw new TypeError('Expected engine cell state service')
  }
  return cellState
}

describe('EngineCellStateService', () => {
  it('translates relative formulas when materializing target cell ops', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'cell-state-formulas' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 2)
    engine.setCellValue('Sheet1', 'A2', 5)
    engine.setCellFormula('Sheet1', 'B1', 'A1*2')

    const snapshot = engine.getCell('Sheet1', 'B1')
    const ops = Effect.runSync(getCellStateService(engine).toCellStateOps('Sheet1', 'B2', snapshot, 'Sheet1', 'B1'))

    expect(ops).toContainEqual({
      kind: 'setCellFormula',
      sheetName: 'Sheet1',
      address: 'B2',
      formula: 'A2*2',
    })
  })

  it('restores inverse cell ops without duplicating format mutations', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'cell-state-restore' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 10)
    engine.setCellFormat('Sheet1', 'A1', '0.00')

    const ops = Effect.runSync(getCellStateService(engine).restoreCellOps('Sheet1', 'A1'))

    expect(ops).toContainEqual({
      kind: 'setCellValue',
      sheetName: 'Sheet1',
      address: 'A1',
      value: 10,
    })
    expect(ops.some((op) => op.kind === 'setCellFormat')).toBe(false)
    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Number, value: 10 })
  })

  it('materializes clear operations for empty and error snapshots and reads rectangular ranges', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'cell-state-edges' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'B1', 3)
    engine.setCellValue('Sheet1', 'A2', 'x')

    const cellState = getCellStateService(engine)

    expect(
      Effect.runSync(
        cellState.toCellStateOps('Sheet1', 'C3', {
          value: { tag: ValueTag.Empty, value: null },
          format: null,
        }),
      ),
    ).toEqual([{ kind: 'clearCell', sheetName: 'Sheet1', address: 'C3' }])

    expect(
      Effect.runSync(
        cellState.toCellStateOps('Sheet1', 'D4', {
          value: { tag: ValueTag.Error, code: 3 },
          format: '0.00',
        }),
      ),
    ).toEqual([
      { kind: 'clearCell', sheetName: 'Sheet1', address: 'D4' },
      { kind: 'setCellFormat', sheetName: 'Sheet1', address: 'D4', format: '0.00' },
    ])

    expect(
      Effect.runSync(
        cellState.readRangeCells({
          sheetName: 'Sheet1',
          startAddress: 'A1',
          endAddress: 'B2',
        }),
      ),
    ).toMatchObject([
      [
        { address: 'A1', sheetName: 'Sheet1', value: { tag: ValueTag.Empty } },
        { address: 'B1', sheetName: 'Sheet1', value: { tag: ValueTag.Number, value: 3 } },
      ],
      [
        { address: 'A2', sheetName: 'Sheet1', value: { tag: ValueTag.String, value: 'x' } },
        { address: 'B2', sheetName: 'Sheet1', value: { tag: ValueTag.Empty } },
      ],
    ])
  })

  it('clears an existing target format when replaying an unformatted empty snapshot', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'cell-state-clear-format' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellFormat('Sheet1', 'A1', '0.00')

    const snapshot = engine.getCell('Sheet1', 'B1')
    const ops = Effect.runSync(getCellStateService(engine).toCellStateOps('Sheet1', 'A1', snapshot, 'Sheet1', 'B1'))

    expect(ops).toEqual([
      { kind: 'clearCell', sheetName: 'Sheet1', address: 'A1' },
      { kind: 'setCellFormat', sheetName: 'Sheet1', address: 'A1', format: null },
    ])
  })

  it('replays explicit blank snapshots as null writes instead of clear ops', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'cell-state-explicit-blank' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.importSnapshot({
      version: 1,
      workbook: { name: 'cell-state-explicit-blank' },
      sheets: [
        {
          id: 1,
          name: 'Sheet1',
          order: 0,
          cells: [{ address: 'D3', value: null }],
        },
      ],
    })

    const cellState = getCellStateService(engine)

    expect(Effect.runSync(cellState.restoreCellOps('Sheet1', 'D3'))).toEqual([
      { kind: 'setCellValue', sheetName: 'Sheet1', address: 'D3', value: null },
    ])

    expect(
      Effect.runSync(
        cellState.toCellStateOps('Sheet1', 'E4', {
          address: 'D3',
          sheetName: 'Sheet1',
          value: { tag: ValueTag.Empty },
          flags: 0,
          version: 1,
        }),
      ),
    ).toEqual([{ kind: 'setCellValue', sheetName: 'Sheet1', address: 'E4', value: null }])
  })

  it('wraps service failures with contextual EngineCellStateError messages', () => {
    const workbook = new WorkbookStore('cell-state-errors')
    workbook.createSheet('Sheet1')
    const manual = createEngineCellStateService({
      state: { workbook },
      getCell: () => {
        throw 'read failure'
      },
      getCellByIndex: () => {
        throw 'capture failure'
      },
    })

    expect(() => Effect.runSync(manual.captureStoredCellOps(1, 'Sheet1', 'A1'))).toThrowError(
      'Failed to capture stored cell ops for Sheet1!A1',
    )
    expect(() =>
      Effect.runSync(
        manual.readRangeCells({
          sheetName: 'Sheet1',
          startAddress: 'A1',
          endAddress: 'A1',
        }),
      ),
    ).toThrowError('Failed to read range Sheet1!A1:A1')
  })
})
