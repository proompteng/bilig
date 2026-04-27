import { describe, expect, it } from 'vitest'
import { ValueTag, type CellSnapshot, type EngineEvent, type RecalcMetrics } from '@bilig/protocol'
import { buildWorkerRenderTileDeltaBatch } from '../worker-runtime-render-tile-delta.js'

const emptyCell: CellSnapshot = {
  sheetName: 'Sheet1',
  address: 'A1',
  value: { tag: ValueTag.Empty },
  flags: 0,
  version: 0,
}

const engine = {
  workbook: {
    getSheet: () => undefined,
    getSheetNameById: () => 'Sheet1',
    cellStore: {
      sheetIds: new Uint16Array(),
      rows: new Uint32Array(),
      cols: new Uint16Array(),
    },
  },
  getCell: () => emptyCell,
  getCellStyle: () => undefined,
  getColumnAxisEntries: () => [],
  getRowAxisEntries: () => [],
  subscribeCells: () => () => undefined,
  getLastMetrics: () => ({ batchId: 3 }),
} as const

const metrics: RecalcMetrics = {
  batchId: 3,
  changedInputCount: 1,
  dirtyFormulaCount: 0,
  wasmFormulaCount: 0,
  jsFormulaCount: 0,
  rangeNodeVisits: 0,
  recalcMs: 0,
  compileMs: 0,
}

function createRangeInvalidationEvent(startAddress: string, endAddress = startAddress): EngineEvent {
  return {
    kind: 'batch',
    invalidation: 'cells',
    changedCellIndices: new Uint32Array(),
    changedCells: [],
    invalidatedRanges: [{ sheetName: 'Sheet1', startAddress, endAddress }],
    invalidatedRows: [],
    invalidatedColumns: [],
    metrics,
  }
}

describe('worker-runtime-render-tile-delta', () => {
  it('materializes fixed content tiles instead of frozen pane scene duplicates', () => {
    const batch = buildWorkerRenderTileDeltaBatch({
      engine,
      generation: 4,
      subscription: {
        sheetId: 7,
        sheetName: 'Sheet1',
        rowStart: 0,
        rowEnd: 33,
        colStart: 0,
        colEnd: 129,
        dprBucket: 2,
        cameraSeq: 17,
      },
    })

    const replacements = batch.mutations.filter((mutation) => mutation.kind === 'tileReplace')

    expect(batch).toMatchObject({ batchId: 3, cameraSeq: 17, sheetId: 7 })
    expect(replacements).toHaveLength(4)
    expect(replacements.map((mutation) => mutation.coord)).toEqual([
      expect.objectContaining({ paneKind: 'body', rowTile: 0, colTile: 0 }),
      expect.objectContaining({ paneKind: 'body', rowTile: 0, colTile: 1 }),
      expect.objectContaining({ paneKind: 'body', rowTile: 1, colTile: 0 }),
      expect.objectContaining({ paneKind: 'body', rowTile: 1, colTile: 1 }),
    ])
    expect(replacements.map((mutation) => mutation.bounds)).toEqual([
      { rowStart: 0, rowEnd: 31, colStart: 0, colEnd: 127 },
      { rowStart: 0, rowEnd: 31, colStart: 128, colEnd: 255 },
      { rowStart: 32, rowEnd: 63, colStart: 0, colEnd: 127 },
      { rowStart: 32, rowEnd: 63, colStart: 128, colEnd: 255 },
    ])
  })

  it('materializes only dirty visible tiles for event-driven batches', () => {
    const batch = buildWorkerRenderTileDeltaBatch({
      engine,
      event: createRangeInvalidationEvent('B2'),
      generation: 5,
      subscription: {
        sheetId: 7,
        sheetName: 'Sheet1',
        rowStart: 0,
        rowEnd: 63,
        colStart: 0,
        colEnd: 255,
        dprBucket: 1,
        cameraSeq: 18,
      },
    })

    const replacements = batch.mutations.filter((mutation) => mutation.kind === 'tileReplace')
    expect(replacements).toHaveLength(1)
    expect(replacements[0]).toMatchObject({
      coord: {
        rowTile: 0,
        colTile: 0,
      },
      bounds: {
        rowStart: 0,
        rowEnd: 31,
        colStart: 0,
        colEnd: 127,
      },
    })
  })

  it('skips event-driven tile materialization when dirty ranges miss the subscription', () => {
    const batch = buildWorkerRenderTileDeltaBatch({
      engine,
      event: createRangeInvalidationEvent('A1000'),
      generation: 6,
      subscription: {
        sheetId: 7,
        sheetName: 'Sheet1',
        rowStart: 0,
        rowEnd: 63,
        colStart: 0,
        colEnd: 255,
        dprBucket: 1,
        cameraSeq: 19,
      },
    })

    expect(batch.mutations).toEqual([])
  })
})
