import { describe, expect, it } from 'vitest'
import { ValueTag, type CellSnapshot, type EngineEvent, type RecalcMetrics } from '@bilig/protocol'
import { packTileKey53 } from '../../../../packages/grid/src/renderer-v3/tile-key.js'
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

function createColumnInvalidationEvent(startIndex: number, endIndex = startIndex): EngineEvent {
  return {
    kind: 'batch',
    invalidation: 'cells',
    changedCellIndices: new Uint32Array(),
    changedCells: [],
    invalidatedRanges: [],
    invalidatedRows: [],
    invalidatedColumns: [{ sheetName: 'Sheet1', startIndex, endIndex }],
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
    expect(replacements[0]?.dirtyLocalRows).toEqual(new Uint32Array([1, 1]))
    expect(replacements[0]?.dirtyLocalCols).toEqual(new Uint32Array([1, 1]))
    expect(replacements[0]?.dirtyMasks).toEqual(new Uint32Array([31]))
  })

  it('materializes warm tile interest on initial and dirty event-driven batches', () => {
    const warmTileKey = packTileKey53({
      colTile: 0,
      dprBucket: 1,
      rowTile: 1,
      sheetOrdinal: 7,
    })
    const subscription = {
      sheetId: 7,
      sheetName: 'Sheet1',
      rowStart: 0,
      rowEnd: 31,
      colStart: 0,
      colEnd: 127,
      dprBucket: 1,
      cameraSeq: 20,
      warmTileKeys: [warmTileKey],
    }

    const initialBatch = buildWorkerRenderTileDeltaBatch({
      engine,
      generation: 7,
      subscription,
    })
    const dirtyBatch = buildWorkerRenderTileDeltaBatch({
      engine,
      event: createRangeInvalidationEvent('A40'),
      generation: 8,
      subscription,
    })

    expect(initialBatch.mutations.filter((mutation) => mutation.kind === 'tileReplace').map((mutation) => mutation.bounds)).toEqual([
      { rowStart: 0, rowEnd: 31, colStart: 0, colEnd: 127 },
      { rowStart: 32, rowEnd: 63, colStart: 0, colEnd: 127 },
    ])
    expect(dirtyBatch.mutations.filter((mutation) => mutation.kind === 'tileReplace')).toHaveLength(1)
    expect(dirtyBatch.mutations[0]).toMatchObject({
      bounds: { rowStart: 32, rowEnd: 63, colStart: 0, colEnd: 127 },
      coord: { rowTile: 1, colTile: 0 },
    })
    expect(dirtyBatch.mutations[0]?.kind === 'tileReplace' ? dirtyBatch.mutations[0].dirtyLocalRows : null).toEqual(new Uint32Array([7, 7]))
    expect(dirtyBatch.mutations[0]?.kind === 'tileReplace' ? dirtyBatch.mutations[0].dirtyLocalCols : null).toEqual(new Uint32Array([0, 0]))
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

  it('clips axis dirty spans to interested render tiles without expanding to the whole sheet', () => {
    const batch = buildWorkerRenderTileDeltaBatch({
      engine,
      event: createColumnInvalidationEvent(130),
      generation: 9,
      subscription: {
        sheetId: 7,
        sheetName: 'Sheet1',
        rowStart: 0,
        rowEnd: 31,
        colStart: 0,
        colEnd: 255,
        dprBucket: 1,
        cameraSeq: 21,
      },
    })

    expect(batch.mutations).toHaveLength(1)
    expect(batch.mutations[0]).toMatchObject({
      kind: 'tileReplace',
      bounds: { rowStart: 0, rowEnd: 31, colStart: 128, colEnd: 255 },
      coord: { rowTile: 0, colTile: 1 },
    })
    expect(batch.mutations[0]?.kind === 'tileReplace' ? batch.mutations[0].dirtyLocalRows : null).toEqual(new Uint32Array([0, 31]))
    expect(batch.mutations[0]?.kind === 'tileReplace' ? batch.mutations[0].dirtyLocalCols : null).toEqual(new Uint32Array([2, 2]))
  })
})
