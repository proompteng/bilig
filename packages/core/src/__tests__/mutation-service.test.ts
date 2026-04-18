import { Effect } from 'effect'
import { describe, expect, it } from 'vitest'
import { ValueTag, type CellSnapshot } from '@bilig/protocol'
import { createReplicaState } from '../replica-state.js'
import { SpreadsheetEngine } from '../engine.js'
import { createEngineMutationService } from '../engine/services/mutation-service.js'
import type { EngineMutationService } from '../engine/services/mutation-service.js'
import { WorkbookStore } from '../workbook-store.js'

function isEngineMutationService(value: unknown): value is EngineMutationService {
  if (typeof value !== 'object' || value === null) {
    return false
  }
  return (
    typeof Reflect.get(value, 'executeLocal') === 'function' &&
    typeof Reflect.get(value, 'captureUndoOps') === 'function' &&
    typeof Reflect.get(value, 'copyRange') === 'function' &&
    typeof Reflect.get(value, 'importSheetCsv') === 'function' &&
    typeof Reflect.get(value, 'renderCommit') === 'function'
  )
}

function getMutationService(engine: SpreadsheetEngine): EngineMutationService {
  const runtime = Reflect.get(engine, 'runtime')
  if (typeof runtime !== 'object' || runtime === null) {
    throw new TypeError('Expected engine runtime')
  }
  const mutation = Reflect.get(runtime, 'mutation')
  if (!isEngineMutationService(mutation)) {
    throw new TypeError('Expected engine mutation service')
  }
  return mutation
}

const EMPTY_CELL_SNAPSHOT: CellSnapshot = {
  sheetName: 'Sheet1',
  address: 'A1',
  value: { tag: ValueTag.Empty },
  flags: 0,
  version: 0,
}

describe('EngineMutationService', () => {
  it('clears redo history when a new local transaction lands', () => {
    const replicaState = createReplicaState('local')
    const workbook = new WorkbookStore('inverse')
    let replayDepth = 0
    const batches: EngineOpBatch[] = []
    const service = createEngineMutationService({
      state: {
        workbook,
        replicaState,
        undoStack: [],
        redoStack: [
          {
            forward: { kind: 'ops', ops: [] },
            inverse: { kind: 'ops', ops: [] },
          },
        ],
        getTransactionReplayDepth: () => replayDepth,
        setTransactionReplayDepth: (next) => {
          replayDepth = next
        },
      },
      captureSheetCellState: () => [],
      captureRowRangeCellState: () => [],
      captureColumnRangeCellState: () => [],
      restoreCellOps: () => [{ kind: 'upsertWorkbook', name: 'inverse' }],
      getCellByIndex: () => EMPTY_CELL_SNAPSHOT,
      applyBatchNow: (batch) => {
        batches.push(batch)
      },
    })

    const undoOps = Effect.runSync(service.executeLocal([{ kind: 'upsertWorkbook', name: 'forward' }]))

    expect(undoOps).toEqual([{ kind: 'upsertWorkbook', name: 'inverse' }])
    expect(batches).toHaveLength(1)
    expect(service).toBeDefined()
  })

  it('captures a single local transaction and clones the undo ops', () => {
    const replicaState = createReplicaState('local')
    const workbook = new WorkbookStore('inverse')
    let replayDepth = 0
    const state = {
      workbook,
      replicaState,
      undoStack: [] as Array<{
        forward: { kind: 'ops'; ops: unknown[] } | { kind: 'single-op'; op: unknown }
        inverse: { kind: 'ops'; ops: unknown[] } | { kind: 'single-op'; op: unknown }
      }>,
      redoStack: [] as Array<{
        forward: { kind: 'ops'; ops: unknown[] } | { kind: 'single-op'; op: unknown }
        inverse: { kind: 'ops'; ops: unknown[] } | { kind: 'single-op'; op: unknown }
      }>,
      getTransactionReplayDepth: () => replayDepth,
      setTransactionReplayDepth: (next: number) => {
        replayDepth = next
      },
    }
    const service = createEngineMutationService({
      state,
      captureSheetCellState: () => [],
      captureRowRangeCellState: () => [],
      captureColumnRangeCellState: () => [],
      restoreCellOps: () => [{ kind: 'upsertWorkbook', name: 'inverse' }],
      getCellByIndex: () => EMPTY_CELL_SNAPSHOT,
      applyBatchNow: () => {},
    })

    const captured = Effect.runSync(
      service.captureUndoOps(() => Effect.runSync(service.executeLocal([{ kind: 'upsertWorkbook', name: 'forward' }]))),
    )

    expect(captured.undoOps).toEqual([{ kind: 'upsertWorkbook', name: 'inverse' }])
  })

  it('drops malformed render commit records instead of forwarding partial engine ops', () => {
    const replicaState = createReplicaState('local')
    const workbook = new WorkbookStore('spec')
    let replayDepth = 0
    const batches: EngineOpBatch[] = []
    const service = createEngineMutationService({
      state: {
        workbook,
        replicaState,
        undoStack: [],
        redoStack: [],
        getTransactionReplayDepth: () => replayDepth,
        setTransactionReplayDepth: (next) => {
          replayDepth = next
        },
      },
      captureSheetCellState: () => [],
      captureRowRangeCellState: () => [],
      captureColumnRangeCellState: () => [],
      restoreCellOps: () => [],
      getCellByIndex: () => EMPTY_CELL_SNAPSHOT,
      applyBatchNow: (batch) => {
        batches.push(batch)
      },
    })

    Effect.runSync(
      service.renderCommit([
        { kind: 'renameSheet', oldName: 'Old' },
        { kind: 'upsertCell', sheetName: 'Sheet1' },
        { kind: 'upsertSheet', name: 'Sheet1', order: 0 },
        { kind: 'upsertCell', sheetName: 'Sheet1', addr: 'A1', value: 7, format: '0.00' },
      ]),
    )

    expect(batches).toHaveLength(1)
    expect(batches[0]?.ops).toEqual([
      { kind: 'upsertSheet', name: 'Sheet1', order: 0 },
      { kind: 'setCellValue', sheetName: 'Sheet1', address: 'A1', value: 7 },
      { kind: 'setCellFormat', sheetName: 'Sheet1', address: 'A1', format: '0.00' },
    ])
  })

  it('preserves undo history for workbook and sheet render commits', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'before-render-commit' })
    await engine.ready()

    engine.renderCommit([
      { kind: 'upsertWorkbook', name: 'after-render-commit' },
      { kind: 'upsertSheet', name: 'Sheet1', order: 0 },
      { kind: 'upsertCell', sheetName: 'Sheet1', addr: 'A1', value: 7 },
    ])

    expect(engine.exportSnapshot()).toMatchObject({
      workbook: { name: 'after-render-commit' },
      sheets: [
        {
          name: 'Sheet1',
          cells: [{ address: 'A1', value: 7 }],
        },
      ],
    })

    expect(engine.undo()).toBe(true)
    expect(engine.exportSnapshot()).toMatchObject({
      workbook: { name: 'before-render-commit' },
      sheets: [],
    })
  })

  it('preserves undo history for bulk render commit cell upserts', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'bulk-render-commit' })
    await engine.ready()

    engine.renderCommit([
      { kind: 'upsertSheet', name: 'Sheet1', order: 0 },
      { kind: 'upsertCell', sheetName: 'Sheet1', addr: 'A1', value: 1 },
      { kind: 'upsertCell', sheetName: 'Sheet1', addr: 'A2', value: 2 },
      { kind: 'upsertCell', sheetName: 'Sheet1', addr: 'A3', formula: 'A1+A2' },
      { kind: 'upsertCell', sheetName: 'Sheet1', addr: 'B1', value: 7, format: '0.00' },
    ])

    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Sheet1', 'A2')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 7 })
    expect(engine.exportSnapshot()).toMatchObject({
      sheets: [
        {
          name: 'Sheet1',
          cells: expect.arrayContaining([
            expect.objectContaining({ address: 'A3', formula: 'A1+A2' }),
            expect.objectContaining({ address: 'B1', value: 7, format: '0.00' }),
          ]),
        },
      ],
    })

    expect(engine.undo()).toBe(true)
    expect(engine.exportSnapshot()).toMatchObject({ sheets: [] })
  })

  it('routes simple render-commit cell upserts through the cell-mutation fast path when no replica batch is needed', () => {
    const workbook = new WorkbookStore('render-commit-fast-path')
    let replayDepth = 0
    const applyBatchCalls: Array<readonly unknown[]> = []
    const applyCellMutationCalls: Array<readonly unknown[]> = []
    const service = createEngineMutationService({
      state: {
        workbook,
        replicaState: createReplicaState('local'),
        undoStack: [],
        redoStack: [],
        trackReplicaVersions: false,
        getSyncClientConnection: () => null,
        batchListeners: new Set(),
        formulas: new Map(),
        getTransactionReplayDepth: () => replayDepth,
        setTransactionReplayDepth: (next) => {
          replayDepth = next
        },
      },
      captureSheetCellState: () => [],
      captureRowRangeCellState: () => [],
      captureColumnRangeCellState: () => [],
      captureStoredCellOps: () => [],
      restoreCellOps: () => [],
      readRangeCells: () => [],
      toCellStateOps: () => [],
      getCellByIndex: () => EMPTY_CELL_SNAPSHOT,
      applyBatchNow: (batch) => {
        applyBatchCalls.push(batch.ops)
        for (const op of batch.ops) {
          if (op.kind === 'upsertWorkbook') {
            workbook.workbookName = op.name
          } else if (op.kind === 'upsertSheet') {
            workbook.getOrCreateSheet(op.name, op.order)
          }
        }
      },
      applyCellMutationsAtBatchNow: (refs) => {
        applyCellMutationCalls.push(refs)
      },
    })

    Effect.runSync(
      service.renderCommit([
        { kind: 'upsertWorkbook', name: 'after-render-commit' },
        { kind: 'upsertSheet', name: 'Sheet1', order: 0 },
        { kind: 'upsertCell', sheetName: 'Sheet1', addr: 'A1', value: 1 },
        { kind: 'upsertCell', sheetName: 'Sheet1', addr: 'A2', value: 2 },
        { kind: 'upsertCell', sheetName: 'Sheet1', addr: 'A3', formula: 'A1+A2' },
      ]),
    )

    expect(applyBatchCalls).toEqual([
      [
        { kind: 'upsertWorkbook', name: 'after-render-commit' },
        { kind: 'upsertSheet', name: 'Sheet1', order: 0 },
      ],
    ])
    expect(applyCellMutationCalls).toHaveLength(1)
    expect(applyCellMutationCalls[0]).toEqual([
      { sheetId: 1, mutation: { kind: 'setCellValue', row: 0, col: 0, value: 1 } },
      { sheetId: 1, mutation: { kind: 'setCellValue', row: 1, col: 0, value: 2 } },
      { sheetId: 1, mutation: { kind: 'setCellFormula', row: 2, col: 0, formula: 'A1+A2' } },
    ])
  })

  it('captures sheet metadata and cells when building delete-sheet undo ops', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'undo-sheet' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 7)
    engine.setCellFormula('Sheet1', 'B1', 'A1*2')
    engine.setFreezePane('Sheet1', 1, 0)
    engine.setFilter('Sheet1', { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B3' })
    engine.setSort('Sheet1', { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B3' }, [{ keyAddress: 'B1', direction: 'asc' }])
    engine.setTable({
      name: 'Sales',
      sheetName: 'Sheet1',
      startAddress: 'A1',
      endAddress: 'B3',
      columnNames: ['Amount', 'Total'],
      headerRow: true,
      totalsRow: false,
    })

    const inverseOps = Effect.runSync(getMutationService(engine).executeLocal([{ kind: 'deleteSheet', name: 'Sheet1' }]))

    expect(inverseOps).not.toBeNull()
    expect(inverseOps).toContainEqual({ kind: 'upsertSheet', name: 'Sheet1', order: 0 })
    expect(inverseOps).toContainEqual({
      kind: 'setFreezePane',
      sheetName: 'Sheet1',
      rows: 1,
      cols: 0,
    })
    expect(inverseOps).toContainEqual({
      kind: 'setFilter',
      sheetName: 'Sheet1',
      range: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B3' },
    })
    expect(inverseOps).toContainEqual({
      kind: 'setSort',
      sheetName: 'Sheet1',
      range: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B3' },
      keys: [{ keyAddress: 'B1', direction: 'asc' }],
    })
    expect(inverseOps).toContainEqual({
      kind: 'upsertTable',
      table: {
        name: 'Sales',
        sheetName: 'Sheet1',
        startAddress: 'A1',
        endAddress: 'B3',
        columnNames: ['Amount', 'Total'],
        headerRow: true,
        totalsRow: false,
      },
    })
    expect(inverseOps).toContainEqual({
      kind: 'setCellValue',
      sheetName: 'Sheet1',
      address: 'A1',
      value: 7,
    })
    expect(inverseOps).toContainEqual({
      kind: 'setCellFormula',
      sheetName: 'Sheet1',
      address: 'B1',
      formula: 'A1*2',
    })
  })

  it('captures deleted row cells in reverse order-safe undo ops', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'undo-rows' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A2', 10)
    engine.setCellFormula('Sheet1', 'B2', 'A2*3')
    engine.updateRowMetadata('Sheet1', 1, 1, 24, false)

    const inverseOps = Effect.runSync(
      getMutationService(engine).executeLocal([{ kind: 'deleteRows', sheetName: 'Sheet1', start: 1, count: 1 }]),
    )

    expect(inverseOps).not.toBeNull()
    expect(inverseOps).toContainEqual({
      kind: 'insertRows',
      sheetName: 'Sheet1',
      start: 1,
      count: 1,
      entries: [{ id: 'row-1', index: 1, size: 24, hidden: false }],
    })
    expect(inverseOps).toContainEqual({
      kind: 'setCellValue',
      sheetName: 'Sheet1',
      address: 'A2',
      value: 10,
    })
    expect(inverseOps).toContainEqual({
      kind: 'setCellFormula',
      sheetName: 'Sheet1',
      address: 'B2',
      formula: 'A2*3',
    })
  })

  it('does not snapshot unrelated sheet cells in delete-row undo ops', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'undo-rows-narrow' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.createSheet('Other')
    engine.setCellValue('Sheet1', 'A2', 10)
    engine.setCellValue('Other', 'C3', 99)

    const inverseOps = Effect.runSync(
      getMutationService(engine).executeLocal([{ kind: 'deleteRows', sheetName: 'Sheet1', start: 1, count: 1 }]),
    )

    expect(inverseOps).not.toContainEqual({
      kind: 'setCellValue',
      sheetName: 'Other',
      address: 'C3',
      value: 99,
    })
  })

  it('does not snapshot unaffected formulas above a deleted row span', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'undo-rows-formula-narrow' })
    await engine.ready()
    engine.createSheet('Sheet1')
    for (let row = 1; row <= 4; row += 1) {
      engine.setCellValue('Sheet1', `A${row}`, row)
      engine.setCellFormula('Sheet1', `B${row}`, `SUM(A1:A${row})`)
    }

    const inverseOps = Effect.runSync(
      getMutationService(engine).executeLocal([{ kind: 'deleteRows', sheetName: 'Sheet1', start: 1, count: 1 }]),
    )

    const formulaOps = inverseOps?.filter(
      (op): op is Extract<(typeof inverseOps)[number], { kind: 'setCellFormula' }> => op.kind === 'setCellFormula',
    )
    expect(formulaOps).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ sheetName: 'Sheet1', address: 'B2', formula: 'SUM(A1:A2)' }),
        expect.objectContaining({ sheetName: 'Sheet1', address: 'B3', formula: 'SUM(A1:A3)' }),
        expect.objectContaining({ sheetName: 'Sheet1', address: 'B4', formula: 'SUM(A1:A4)' }),
      ]),
    )
    expect(formulaOps).not.toEqual(
      expect.arrayContaining([expect.objectContaining({ sheetName: 'Sheet1', address: 'B1', formula: 'SUM(A1:A1)' })]),
    )
  })

  it('fast-paths simple cell mutation history without restore callbacks', () => {
    const replicaState = createReplicaState('local')
    const workbook = new WorkbookStore('fast-history')
    const sheet = workbook.createSheet('Sheet1')
    const cell = workbook.ensureCellAt(sheet.id, 0, 0)
    workbook.cellStore.setValue(cell.cellIndex, { tag: ValueTag.Number, value: 7 }, 0)
    let replayDepth = 0
    const service = createEngineMutationService({
      state: {
        workbook,
        replicaState,
        undoStack: [],
        redoStack: [],
        getTransactionReplayDepth: () => replayDepth,
        setTransactionReplayDepth: (next) => {
          replayDepth = next
        },
      },
      captureSheetCellState: () => [],
      captureRowRangeCellState: () => [],
      captureColumnRangeCellState: () => [],
      restoreCellOps: () => {
        throw new Error('restoreCellOps should not be used for simple cell history')
      },
      getCellByIndex: () => ({
        sheetName: 'Sheet1',
        address: 'A1',
        value: { tag: ValueTag.Number, value: 7 },
        flags: 0,
        version: 0,
      }),
      applyCellMutationsAtBatchNow: () => {},
      applyBatchNow: () => {},
    })

    const inverseOps = Effect.runSync(service.executeLocal([{ kind: 'setCellValue', sheetName: 'Sheet1', address: 'A1', value: 9 }]))

    expect(inverseOps).toEqual([{ kind: 'setCellValue', sheetName: 'Sheet1', address: 'A1', value: 7 }])
  })

  it('can record simple cell mutation history without allocating undo ops', () => {
    const replicaState = createReplicaState('local')
    const workbook = new WorkbookStore('fast-history-no-return')
    const sheet = workbook.createSheet('Sheet1')
    const cell = workbook.ensureCellAt(sheet.id, 0, 0)
    workbook.cellStore.setValue(cell.cellIndex, { tag: ValueTag.Number, value: 7 }, 0)
    let replayDepth = 0
    const state = {
      workbook,
      replicaState,
      undoStack: [] as Array<{
        forward: { kind: 'ops'; ops: unknown[] } | { kind: 'single-op'; op: unknown }
        inverse: { kind: 'ops'; ops: unknown[] } | { kind: 'single-op'; op: unknown }
      }>,
      redoStack: [] as Array<{
        forward: { kind: 'ops'; ops: unknown[] } | { kind: 'single-op'; op: unknown }
        inverse: { kind: 'ops'; ops: unknown[] } | { kind: 'single-op'; op: unknown }
      }>,
      getTransactionReplayDepth: () => replayDepth,
      setTransactionReplayDepth: (next: number) => {
        replayDepth = next
      },
    }
    const service = createEngineMutationService({
      state,
      captureSheetCellState: () => [],
      captureRowRangeCellState: () => [],
      captureColumnRangeCellState: () => [],
      restoreCellOps: () => {
        throw new Error('restoreCellOps should not be used for simple cell history')
      },
      getCellByIndex: () => ({
        sheetName: 'Sheet1',
        address: 'A1',
        value: { tag: ValueTag.Number, value: 7 },
        flags: 0,
        version: 0,
      }),
      applyCellMutationsAtBatchNow: () => {},
      applyBatchNow: () => {},
    })

    const inverseOps = service.applyCellMutationsAtNow(
      [{ sheetId: sheet.id, mutation: { kind: 'setCellValue', row: 0, col: 0, value: 9 } }],
      {
        captureUndo: true,
        source: 'local',
        potentialNewCells: 1,
        returnUndoOps: false,
      },
    )

    expect(inverseOps).toBeNull()
    expect(state.undoStack).toHaveLength(1)
    expect(state.undoStack[0]?.inverse).toEqual({
      kind: 'single-op',
      op: { kind: 'setCellValue', sheetName: 'Sheet1', address: 'A1', value: 7 },
      potentialNewCells: 1,
    })
  })

  it('lazily materializes forward ops for multi-cell local history when no replica batch is needed', () => {
    const workbook = new WorkbookStore('lazy-forward')
    const sheet = workbook.getOrCreateSheet('Sheet1')
    let replayDepth = 0
    const state = {
      workbook,
      replicaState: createReplicaState('local'),
      undoStack: [] as Array<{
        forward: { kind: 'ops'; ops: unknown[] } | { kind: 'single-op'; op: unknown }
        inverse: { kind: 'ops'; ops: unknown[] } | { kind: 'single-op'; op: unknown }
      }>,
      redoStack: [] as Array<{
        forward: { kind: 'ops'; ops: unknown[] } | { kind: 'single-op'; op: unknown }
        inverse: { kind: 'ops'; ops: unknown[] } | { kind: 'single-op'; op: unknown }
      }>,
      trackReplicaVersions: false,
      getSyncClientConnection: () => null,
      batchListeners: new Set<() => void>(),
      formulas: new Map<number, unknown>(),
      getTransactionReplayDepth: () => replayDepth,
      setTransactionReplayDepth: (next: number) => {
        replayDepth = next
      },
    }
    const service = createEngineMutationService({
      state,
      captureSheetCellState: () => [],
      captureRowRangeCellState: () => [],
      captureColumnRangeCellState: () => [],
      restoreCellOps: () => [],
      getCellByIndex: (cellIndex) => ({
        sheetName: 'Sheet1',
        address: cellIndex === 0 ? 'A1' : 'A2',
        value: { tag: ValueTag.Empty },
        flags: 0,
        version: 0,
      }),
      applyCellMutationsAtBatchNow: () => {},
      applyBatchNow: () => {},
    })

    const inverseOps = service.applyCellMutationsAtNow(
      [
        { sheetId: sheet.id, mutation: { kind: 'setCellValue', row: 0, col: 0, value: 1 } },
        { sheetId: sheet.id, mutation: { kind: 'setCellValue', row: 1, col: 0, value: 2 } },
      ],
      {
        captureUndo: true,
        source: 'local',
        potentialNewCells: 2,
        returnUndoOps: false,
        reuseRefs: true,
      },
    )

    expect(inverseOps).toBeNull()
    expect(state.undoStack).toHaveLength(1)
    expect(state.undoStack[0]?.forward.kind).toBe('ops')
    expect(state.undoStack[0]?.forward.ops).toEqual([
      { kind: 'setCellValue', sheetName: 'Sheet1', address: 'A1', value: 1 },
      { kind: 'setCellValue', sheetName: 'Sheet1', address: 'A2', value: 2 },
    ])
  })

  it('does not synthesize blank column identities in delete undo ops', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'undo-columns' })
    await engine.ready()
    engine.createSheet('Sheet1')

    const inverseOps = Effect.runSync(
      getMutationService(engine).executeLocal([{ kind: 'deleteColumns', sheetName: 'Sheet1', start: 0, count: 1 }]),
    )

    expect(inverseOps).toContainEqual({
      kind: 'insertColumns',
      sheetName: 'Sheet1',
      start: 0,
      count: 1,
      entries: [],
    })
  })

  it('does not snapshot unrelated sheet cells in delete-column undo ops', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'undo-columns-narrow' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.createSheet('Other')
    engine.setCellValue('Sheet1', 'B1', 10)
    engine.setCellValue('Other', 'C3', 99)

    const inverseOps = Effect.runSync(
      getMutationService(engine).executeLocal([{ kind: 'deleteColumns', sheetName: 'Sheet1', start: 1, count: 1 }]),
    )

    expect(inverseOps).not.toContainEqual({
      kind: 'setCellValue',
      sheetName: 'Other',
      address: 'C3',
      value: 99,
    })
  })

  it('does not snapshot unaffected cross-sheet formulas left of a deleted column span', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'undo-columns-formula-narrow' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.createSheet('Other')
    engine.setCellValue('Other', 'A1', 7)
    engine.setCellFormula('Sheet1', 'A1', 'Other!A1')
    engine.setCellValue('Sheet1', 'A2', 1)
    engine.setCellValue('Sheet1', 'B2', 2)
    engine.setCellValue('Sheet1', 'C2', 3)
    engine.setCellFormula('Sheet1', 'D2', 'SUM(A2:C2)')

    const inverseOps = Effect.runSync(
      getMutationService(engine).executeLocal([{ kind: 'deleteColumns', sheetName: 'Sheet1', start: 1, count: 1 }]),
    )

    const formulaOps = inverseOps?.filter(
      (op): op is Extract<(typeof inverseOps)[number], { kind: 'setCellFormula' }> => op.kind === 'setCellFormula',
    )
    expect(formulaOps).toEqual(
      expect.arrayContaining([expect.objectContaining({ sheetName: 'Sheet1', address: 'D2', formula: 'SUM(A2:C2)' })]),
    )
    expect(formulaOps).not.toEqual(
      expect.arrayContaining([expect.objectContaining({ sheetName: 'Sheet1', address: 'A1', formula: 'Other!A1' })]),
    )
  })

  it('captures formulas that depend on whole-row ranges across deleted row spans', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'undo-rows-whole-range' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 1)
    engine.setCellValue('Sheet1', 'A2', 2)
    engine.setCellValue('Sheet1', 'A3', 3)
    engine.setCellFormula('Sheet1', 'B5', 'SUM(1:3)')

    const inverseOps = Effect.runSync(
      getMutationService(engine).executeLocal([{ kind: 'deleteRows', sheetName: 'Sheet1', start: 1, count: 1 }]),
    )

    expect(inverseOps).toContainEqual({
      kind: 'setCellFormula',
      sheetName: 'Sheet1',
      address: 'B5',
      formula: 'SUM(1:3)',
    })
  })

  it('captures formulas that depend on whole-column ranges across deleted column spans', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'undo-columns-whole-range' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'B1', 1)
    engine.setCellValue('Sheet1', 'C1', 2)
    engine.setCellValue('Sheet1', 'D1', 3)
    engine.setCellFormula('Sheet1', 'F2', 'SUM(B:D)')

    const inverseOps = Effect.runSync(
      getMutationService(engine).executeLocal([{ kind: 'deleteColumns', sheetName: 'Sheet1', start: 1, count: 1 }]),
    )

    expect(inverseOps).toContainEqual({
      kind: 'setCellFormula',
      sheetName: 'Sheet1',
      address: 'F2',
      formula: 'SUM(B:D)',
    })
  })

  it('copies ranges through the service and rewrites relative formulas', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'copy-range-service' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 5)
    engine.setCellValue('Sheet1', 'A2', 9)
    engine.setCellFormula('Sheet1', 'B1', 'A1*2')

    Effect.runSync(
      getMutationService(engine).copyRange(
        { sheetName: 'Sheet1', startAddress: 'B1', endAddress: 'B1' },
        { sheetName: 'Sheet1', startAddress: 'B2', endAddress: 'B2' },
      ),
    )

    expect(engine.getCell('Sheet1', 'B2').formula).toBe('A2*2')
    expect(engine.getCellValue('Sheet1', 'B2')).toEqual({ tag: ValueTag.Number, value: 18 })
  })

  it('imports csv through the service and replaces prior sheet contents', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'csv-import-service' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 1)
    engine.setCellValue('Sheet1', 'C3', 99)

    Effect.runSync(getMutationService(engine).importSheetCsv('Sheet1', '7,=A1*2\n"alpha,beta",'))

    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Number, value: 7 })
    expect(engine.getCell('Sheet1', 'B1').formula).toBe('A1*2')
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 14 })
    expect(engine.getCell('Sheet1', 'A2').value).toEqual({
      tag: ValueTag.String,
      value: 'alpha,beta',
      stringId: 1,
    })
    expect(engine.getCellValue('Sheet1', 'C3')).toEqual({ tag: ValueTag.Empty })
  })
})
