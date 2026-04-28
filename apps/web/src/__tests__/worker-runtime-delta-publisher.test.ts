import { describe, expect, it } from 'vitest'
import type { EngineEvent, RecalcMetrics } from '@bilig/protocol'
import { DirtyMaskV3 } from '../../../../packages/grid/src/renderer-v3/tile-damage-index.js'
import {
  WorkerRuntimeDeltaPublisher,
  buildWorkbookDeltaBatchesFromEngineEventV3,
  type WorkbookDeltaSheetIdentityV3,
} from '../worker-runtime-delta-publisher.js'
import type { WorkerEngine, WorkerSheet } from '../worker-runtime-support.js'

interface TestSheet extends WorkerSheet {
  readonly id: number
}

const metrics: RecalcMetrics = {
  batchId: 17,
  changedInputCount: 1,
  dirtyFormulaCount: 0,
  wasmFormulaCount: 0,
  jsFormulaCount: 0,
  rangeNodeVisits: 0,
  recalcMs: 0,
  compileMs: 0,
}

function createEngine(): WorkerEngine {
  const sheets = new Map<string, TestSheet>([
    ['Later', createSheet({ id: 8, name: 'Later', order: 1 })],
    ['Sheet1', createSheet({ id: 7, name: 'Sheet1', order: 0 })],
  ])
  const engine: WorkerEngine = {
    workbook: {
      workbookName: 'Test',
      cellStore: {
        sheetIds: Uint16Array.from([0, 7, 8]),
        rows: Uint32Array.from([0, 3, 12]),
        cols: Uint16Array.from([0, 4, 9]),
      },
      sheetsByName: sheets,
      getSheet: (sheetName) => sheets.get(sheetName),
      getSheetNameById: (sheetId) => [...sheets.values()].find((sheet) => sheet.id === sheetId)?.name ?? '',
      getQualifiedAddress: () => '',
    },
    ready: async () => undefined,
    createSheet: () => undefined,
    subscribe: () => () => undefined,
    subscribeBatches: () => () => undefined,
    getLastMetrics: () => metrics,
    getSyncState: () => 'local-only',
    getCell: () => {
      throw new Error('not implemented')
    },
    getCellStyle: () => undefined,
    setRangeNumberFormat: () => undefined,
    clearRangeNumberFormat: () => undefined,
    clearRange: () => undefined,
    setCellValue: () => undefined,
    setCellFormula: () => undefined,
    setRangeStyle: () => undefined,
    clearRangeStyle: () => undefined,
    clearCell: () => undefined,
    renderCommit: () => undefined,
    fillRange: () => undefined,
    copyRange: () => undefined,
    moveRange: () => undefined,
    insertRows: () => undefined,
    deleteRows: () => undefined,
    insertColumns: () => undefined,
    deleteColumns: () => undefined,
    updateRowMetadata: () => undefined,
    updateColumnMetadata: () => undefined,
    setFreezePane: () => undefined,
    getFreezePane: () => undefined,
    exportSnapshot: () => {
      throw new Error('not implemented')
    },
    exportReplicaSnapshot: () => {
      throw new Error('not implemented')
    },
    importSnapshot: () => undefined,
    importReplicaSnapshot: () => undefined,
    getColumnAxisEntries: () => [],
    getRowAxisEntries: () => [],
  }
  return engine
}

function createSheet(input: { readonly id: number; readonly name: string; readonly order: number }): TestSheet {
  return {
    id: input.id,
    name: input.name,
    order: input.order,
    grid: {
      forEachCellEntry: () => undefined,
    },
  }
}

function createEvent(): EngineEvent {
  return {
    kind: 'batch',
    invalidation: 'cells',
    changedCellIndices: Uint32Array.from([1, 2]),
    changedCells: [],
    invalidatedRanges: [{ sheetName: 'Sheet1', startAddress: 'B2', endAddress: 'C3' }],
    invalidatedRows: [{ sheetName: 'Later', startIndex: 32, endIndex: 33 }],
    invalidatedColumns: [{ sheetName: 'Sheet1', startIndex: 128, endIndex: 130 }],
    metrics,
  }
}

describe('WorkerRuntimeDeltaPublisher', () => {
  it('builds sheet-level workbook delta batches from engine impact', () => {
    const batches = buildWorkbookDeltaBatchesFromEngineEventV3({
      engine: createEngine(),
      event: createEvent(),
      allocateSeq: createSeqAllocator(),
    })

    const [sheet1, later] = batches
    if (!sheet1 || !later) {
      throw new Error('Expected two workbook delta batches')
    }
    expect(batches).toHaveLength(2)
    expect(sheet1).toMatchObject({
      seq: 1,
      source: 'workerAuthoritative',
      sheetId: 7,
      sheetOrdinal: 0,
      valueSeq: 17,
      calcSeq: 17,
    })
    expect([...sheet1.dirty.cellRanges]).toEqual([
      3,
      3,
      4,
      4,
      DirtyMaskV3.Value | DirtyMaskV3.Text,
      1,
      2,
      1,
      2,
      DirtyMaskV3.Value | DirtyMaskV3.Style | DirtyMaskV3.Text | DirtyMaskV3.Rect | DirtyMaskV3.Border,
    ])
    expect([...sheet1.dirty.axisX]).toEqual([128, 130, DirtyMaskV3.AxisX | DirtyMaskV3.Text | DirtyMaskV3.Rect])
    expect([...later.dirty.axisY]).toEqual([32, 33, DirtyMaskV3.AxisY | DirtyMaskV3.Text | DirtyMaskV3.Rect])
  })

  it('tracks publisher sequence across events', () => {
    const publisher = new WorkerRuntimeDeltaPublisher()

    expect(publisher.buildFromEngineEvent({ engine: createEngine(), event: createEvent() }).map((batch) => batch.seq)).toEqual([1, 2])
    expect(publisher.buildFromEngineEvent({ engine: createEngine(), event: createEvent() }).map((batch) => batch.seq)).toEqual([3, 4])

    publisher.reset()
    expect(publisher.buildFromEngineEvent({ engine: createEngine(), event: createEvent() }).map((batch) => batch.seq)).toEqual([1, 2])
  })

  it('allows explicit sheet identity resolution', () => {
    const batches = buildWorkbookDeltaBatchesFromEngineEventV3({
      engine: createEngine(),
      event: createEvent(),
      allocateSeq: createSeqAllocator(),
      resolveSheetIdentity: (sheetName): WorkbookDeltaSheetIdentityV3 | null =>
        sheetName === 'Sheet1' ? { sheetId: 99, sheetOrdinal: 3 } : null,
    })

    expect(batches).toHaveLength(1)
    expect(batches[0]).toMatchObject({
      sheetId: 99,
      sheetOrdinal: 3,
    })
  })
})

function createSeqAllocator(): () => number {
  let seq = 0
  return () => {
    seq += 1
    return seq
  }
}
