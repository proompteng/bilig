import { describe, expect, it } from 'vitest'
import { SpreadsheetEngine } from '@bilig/core'
import {
  VIEWPORT_TILE_COLUMN_COUNT,
  VIEWPORT_TILE_ROW_COUNT,
  ValueTag,
  type CellSnapshot,
  type EngineEvent,
  type RecalcMetrics,
} from '@bilig/protocol'
import { TextOverflowIndexV3 } from '../../../../packages/grid/src/renderer-v3/text-overflow-index.js'
import { GRID_RECT_INSTANCE_FLOAT_COUNT_V3 } from '../../../../packages/grid/src/renderer-v3/rect-instance-buffer.js'
import { DirtyMaskV3 } from '../../../../packages/grid/src/renderer-v3/tile-damage-index.js'
import { packTileKey53, unpackTileKey53 } from '../../../../packages/grid/src/renderer-v3/tile-key.js'
import { buildWorkerRenderTileDeltaBatch } from '../worker-runtime-render-tile-delta.js'
import { buildFreezeVersion } from '../worker-runtime-render-axis.js'
import type { WorkerEngine } from '../worker-runtime-support.js'

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

const RANGE_VISUAL_DIRTY_MASK = DirtyMaskV3.Value | DirtyMaskV3.Style | DirtyMaskV3.Text | DirtyMaskV3.Rect | DirtyMaskV3.Border

function hasOpaqueGreenFillRect(rectInstances: Float32Array, rectCount: number): boolean {
  for (let index = 0; index < rectCount; index += 1) {
    const offset = index * GRID_RECT_INSTANCE_FLOAT_COUNT_V3
    const red = rectInstances[offset + 4] ?? 1
    const green = rectInstances[offset + 5] ?? 0
    const blue = rectInstances[offset + 6] ?? 1
    const alpha = rectInstances[offset + 7] ?? 0
    const instanceKind = rectInstances[offset + 13] ?? -1
    if (instanceKind === 0 && red < 0.05 && green > 0.95 && blue < 0.05 && alpha > 0.95) {
      return true
    }
  }
  return false
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

function createChangedCellEvent(cellIndex: number): EngineEvent {
  return {
    kind: 'batch',
    invalidation: 'cells',
    changedCellIndices: Uint32Array.from([cellIndex]),
    changedCells: [],
    invalidatedRanges: [],
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

function createRowInvalidationEvent(startIndex: number, endIndex = startIndex): EngineEvent {
  return {
    kind: 'batch',
    invalidation: 'cells',
    changedCellIndices: new Uint32Array(),
    changedCells: [],
    invalidatedRanges: [],
    invalidatedRows: [{ sheetName: 'Sheet1', startIndex, endIndex }],
    invalidatedColumns: [],
    metrics,
  }
}

function createStringCell(address: string, value: string): CellSnapshot {
  return {
    sheetName: 'Sheet1',
    address,
    value: { tag: ValueTag.String, value, stringId: 1 },
    flags: 0,
    version: 1,
  }
}

function createEngineWithCells(cells: Map<string, CellSnapshot>) {
  return {
    ...engine,
    getCell: (_sheetName: string, address: string) => cells.get(address) ?? { ...emptyCell, address },
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

    expect(batch).toMatchObject({ batchId: 3, cameraSeq: 17, sheetId: 7, version: 4 })
    expect(replacements).toHaveLength(4)
    expect(replacements.every((mutation) => typeof mutation.rectSignature === 'string' && mutation.rectSignature.length > 0)).toBe(true)
    expect(replacements.every((mutation) => typeof mutation.textSignature === 'string' && mutation.textSignature.length > 0)).toBe(true)
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

  it('materializes authoritative snapshot style ranges on empty cells into initial render tiles', async () => {
    const seed = new SpreadsheetEngine({ workbookName: 'render-tile-style-range-seed' })
    await seed.ready()
    seed.createSheet('Sheet1')
    seed.setRangeStyle({ sheetName: 'Sheet1', startAddress: 'E6', endAddress: 'E6' }, { fill: { backgroundColor: '#00ff00' } })

    const restored = new SpreadsheetEngine({ workbookName: 'render-tile-style-range-restored' }) as SpreadsheetEngine & WorkerEngine
    await restored.ready()
    restored.importSnapshot(seed.exportSnapshot())

    const batch = buildWorkerRenderTileDeltaBatch({
      engine: restored,
      generation: 4,
      subscription: {
        sheetId: 7,
        sheetName: 'Sheet1',
        rowStart: 0,
        rowEnd: 31,
        colStart: 0,
        colEnd: 127,
        dprBucket: 1,
        cameraSeq: 17,
      },
    })
    const replacement = batch.mutations.find(
      (mutation) =>
        mutation.kind === 'tileReplace' &&
        mutation.bounds.rowStart <= 5 &&
        mutation.bounds.rowEnd >= 5 &&
        mutation.bounds.colStart <= 4 &&
        mutation.bounds.colEnd >= 4,
    )

    expect(restored.getCell('Sheet1', 'E6').styleId).toBeDefined()
    expect(replacement?.kind === 'tileReplace' ? hasOpaqueGreenFillRect(replacement.rectInstances, replacement.rectCount) : false).toBe(
      true,
    )
  })

  it('treats changed-cell render tile deltas as visual paint damage', () => {
    const changedEngine = {
      ...engine,
      workbook: {
        ...engine.workbook,
        cellStore: {
          sheetIds: Uint16Array.from([7]),
          rows: Uint32Array.from([5]),
          cols: Uint16Array.from([4]),
        },
      },
    }

    const batch = buildWorkerRenderTileDeltaBatch({
      engine: changedEngine,
      event: createChangedCellEvent(0),
      generation: 5,
      subscription: {
        sheetId: 7,
        sheetName: 'Sheet1',
        rowStart: 0,
        rowEnd: 31,
        colStart: 0,
        colEnd: 127,
        dprBucket: 1,
        cameraSeq: 18,
      },
    })
    const replacement = batch.mutations.find((mutation) => mutation.kind === 'tileReplace')

    expect(replacement?.kind === 'tileReplace' ? replacement.dirtyLocalRows : null).toEqual(new Uint32Array([5, 5]))
    expect(replacement?.kind === 'tileReplace' ? replacement.dirtyLocalCols : null).toEqual(new Uint32Array([4, 4]))
    expect(replacement?.kind === 'tileReplace' ? replacement.dirtyMasks : null).toEqual(new Uint32Array([RANGE_VISUAL_DIRTY_MASK]))
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
    expect(replacements[0]?.dirtyMasks).toEqual(new Uint32Array([RANGE_VISUAL_DIRTY_MASK]))
    expect(replacements[0]?.dirty.rectSpans).toEqual([{ offset: 0, length: replacements[0]?.rectCount ?? 0 }])
  })

  it('clips sheet-scale style range invalidations to interested render tiles', () => {
    const visibleTileKey = packTileKey53({
      colTile: 0,
      dprBucket: 1,
      rowTile: 0,
      sheetOrdinal: 7,
    })
    const warmTileKey = packTileKey53({
      colTile: 0,
      dprBucket: 1,
      rowTile: 1,
      sheetOrdinal: 7,
    })
    const batch = buildWorkerRenderTileDeltaBatch({
      engine,
      event: createRangeInvalidationEvent('B1', 'E1048576'),
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
        tileInterest: {
          seq: 11,
          sheetOrdinal: 7,
          axisSeqX: 1,
          axisSeqY: 1,
          freezeSeq: 1,
          visibleTileKeys: [visibleTileKey],
          warmTileKeys: [warmTileKey],
          pinnedTileKeys: [],
          reason: 'scroll' as const,
        },
      },
    })

    const replacements = batch.mutations.filter((mutation) => mutation.kind === 'tileReplace')
    expect(replacements).toHaveLength(2)
    expect(replacements.map((mutation) => mutation.bounds)).toEqual([
      { rowStart: 0, rowEnd: 31, colStart: 0, colEnd: 127 },
      { rowStart: 32, rowEnd: 63, colStart: 0, colEnd: 127 },
    ])
    replacements.forEach((replacement) => {
      expect(replacement.dirtyLocalRows).toEqual(new Uint32Array([0, 31]))
      expect(replacement.dirtyLocalCols).toEqual(new Uint32Array([1, 4]))
      expect(replacement.dirtyMasks).toEqual(new Uint32Array([RANGE_VISUAL_DIRTY_MASK]))
    })
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

  it('uses explicit V3 visible tile interest instead of rematerializing the whole viewport window', () => {
    const visibleTileKey = packTileKey53({
      colTile: 1,
      dprBucket: 1,
      rowTile: 0,
      sheetOrdinal: 7,
    })
    const warmTileKey = packTileKey53({
      colTile: 1,
      dprBucket: 1,
      rowTile: 1,
      sheetOrdinal: 7,
    })
    const subscription = {
      sheetId: 7,
      sheetName: 'Sheet1',
      rowStart: 0,
      rowEnd: 63,
      colStart: 0,
      colEnd: 255,
      dprBucket: 1,
      cameraSeq: 22,
      tileInterest: {
        seq: 9,
        sheetOrdinal: 7,
        axisSeqX: 1,
        axisSeqY: 1,
        freezeSeq: 1,
        visibleTileKeys: [visibleTileKey],
        warmTileKeys: [warmTileKey],
        pinnedTileKeys: [],
        reason: 'scroll' as const,
      },
    }

    const initialBatch = buildWorkerRenderTileDeltaBatch({
      engine,
      generation: 10,
      subscription,
    })
    const dirtyVisibleBatch = buildWorkerRenderTileDeltaBatch({
      engine,
      event: createRangeInvalidationEvent('DY2'),
      generation: 11,
      subscription,
    })
    const dirtyWarmBatch = buildWorkerRenderTileDeltaBatch({
      engine,
      event: createRangeInvalidationEvent('DY40'),
      generation: 12,
      subscription,
    })
    const dirtyViewportOnlyBatch = buildWorkerRenderTileDeltaBatch({
      engine,
      event: createRangeInvalidationEvent('B2'),
      generation: 13,
      subscription,
    })

    expect(initialBatch.mutations.filter((mutation) => mutation.kind === 'tileReplace').map((mutation) => mutation.bounds)).toEqual([
      { rowStart: 0, rowEnd: 31, colStart: 128, colEnd: 255 },
      { rowStart: 32, rowEnd: 63, colStart: 128, colEnd: 255 },
    ])
    expect(dirtyVisibleBatch.mutations).toHaveLength(1)
    expect(dirtyVisibleBatch.mutations[0]).toMatchObject({
      bounds: { rowStart: 0, rowEnd: 31, colStart: 128, colEnd: 255 },
      coord: { rowTile: 0, colTile: 1 },
    })
    expect(dirtyWarmBatch.mutations).toHaveLength(1)
    expect(dirtyWarmBatch.mutations[0]).toMatchObject({
      bounds: { rowStart: 32, rowEnd: 63, colStart: 128, colEnd: 255 },
      coord: { rowTile: 1, colTile: 1 },
    })
    expect(dirtyViewportOnlyBatch.mutations).toEqual([])
  })

  it('packs render tile keys with sheet ordinal instead of stable sheet id', () => {
    const visibleTileKey = packTileKey53({
      colTile: 0,
      dprBucket: 1,
      rowTile: 0,
      sheetOrdinal: 2,
    })
    const batch = buildWorkerRenderTileDeltaBatch({
      engine,
      generation: 14,
      subscription: {
        sheetId: 99,
        sheetOrdinal: 2,
        sheetName: 'Sheet1',
        rowStart: 0,
        rowEnd: 31,
        colStart: 0,
        colEnd: 127,
        dprBucket: 1,
        cameraSeq: 23,
        tileInterest: {
          seq: 10,
          sheetOrdinal: 2,
          axisSeqX: 1,
          axisSeqY: 1,
          freezeSeq: 1,
          visibleTileKeys: [visibleTileKey],
          warmTileKeys: [],
          pinnedTileKeys: [],
          reason: 'scroll',
        },
      },
    })

    const replacement = batch.mutations.find((mutation) => mutation.kind === 'tileReplace')
    expect(batch).toMatchObject({ sheetId: 99, sheetOrdinal: 2 })
    expect(replacement).toMatchObject({
      kind: 'tileReplace',
      coord: expect.objectContaining({ sheetId: 99, sheetOrdinal: 2 }),
      tileId: visibleTileKey,
    })
    expect(replacement?.kind === 'tileReplace' ? unpackTileKey53(replacement.tileId).sheetOrdinal : null).toBe(2)
  })

  it('versions worker-materialized tiles with the current freeze pane', () => {
    const batch = buildWorkerRenderTileDeltaBatch({
      engine: {
        ...engine,
        getFreezePane: () => ({ cols: 3, rows: 2 }),
      },
      generation: 15,
      subscription: {
        sheetId: 7,
        sheetName: 'Sheet1',
        rowStart: 0,
        rowEnd: 31,
        colStart: 0,
        colEnd: 127,
        dprBucket: 1,
        cameraSeq: 24,
      },
    })

    const replacement = batch.mutations.find((mutation) => mutation.kind === 'tileReplace')
    expect(replacement?.kind === 'tileReplace' ? replacement.version.freeze : null).toBe(buildFreezeVersion(2, 3))
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
    expect(batch.mutations[0]?.kind === 'tileReplace' ? batch.mutations[0].dirtyLocalCols : null).toEqual(new Uint32Array([2, 127]))
  })

  it('dirties source spill spans when a blocking cell changes', () => {
    const cells = new Map<string, CellSnapshot>([['A1', createStringCell('A1', 'spill text')]])
    const textOverflowIndex = new TextOverflowIndexV3()
    const subscription = {
      sheetId: 7,
      sheetName: 'Sheet1',
      rowStart: 0,
      rowEnd: 31,
      colStart: 0,
      colEnd: 127,
      dprBucket: 1,
      cameraSeq: 24,
    }

    buildWorkerRenderTileDeltaBatch({
      engine: createEngineWithCells(cells),
      generation: 15,
      subscription,
      textOverflowIndex,
    })
    cells.set('B1', createStringCell('B1', 'blocker'))
    const batch = buildWorkerRenderTileDeltaBatch({
      engine: createEngineWithCells(cells),
      event: createRangeInvalidationEvent('B1'),
      generation: 16,
      subscription,
      textOverflowIndex,
    })

    const replacement = batch.mutations[0]
    expect(replacement).toMatchObject({
      kind: 'tileReplace',
      bounds: { rowStart: 0, rowEnd: 31, colStart: 0, colEnd: 127 },
    })
    expect(replacement?.kind === 'tileReplace' ? replacement.dirtyLocalRows : null).toEqual(new Uint32Array([0, 0, 0, 0]))
    expect(replacement?.kind === 'tileReplace' ? replacement.dirtyLocalCols : null).toEqual(new Uint32Array([1, 1, 0, 127]))
    expect(replacement?.kind === 'tileReplace' ? replacement.dirtyMasks : null).toEqual(
      new Uint32Array([RANGE_VISUAL_DIRTY_MASK, DirtyMaskV3.Text]),
    )
  })

  it('dirties spill sources crossing resized columns', () => {
    const cells = new Map<string, CellSnapshot>([['A1', createStringCell('A1', 'spill text')]])
    const textOverflowIndex = new TextOverflowIndexV3()
    const subscription = {
      sheetId: 7,
      sheetName: 'Sheet1',
      rowStart: 0,
      rowEnd: 31,
      colStart: 0,
      colEnd: 127,
      dprBucket: 1,
      cameraSeq: 25,
    }

    buildWorkerRenderTileDeltaBatch({
      engine: createEngineWithCells(cells),
      generation: 17,
      subscription,
      textOverflowIndex,
    })
    const batch = buildWorkerRenderTileDeltaBatch({
      engine: createEngineWithCells(cells),
      event: createColumnInvalidationEvent(2),
      generation: 18,
      subscription,
      textOverflowIndex,
    })

    const replacement = batch.mutations[0]
    expect(replacement?.kind === 'tileReplace' ? replacement.dirtyLocalRows : null).toEqual(new Uint32Array([0, 31, 0, 0]))
    expect(replacement?.kind === 'tileReplace' ? replacement.dirtyLocalCols : null).toEqual(new Uint32Array([2, 127, 0, 127]))
    expect(replacement?.kind === 'tileReplace' ? replacement.dirtyMasks : null).toEqual(
      new Uint32Array([DirtyMaskV3.AxisX | DirtyMaskV3.Text | DirtyMaskV3.Rect, DirtyMaskV3.Text]),
    )
  })

  it('dirties every shifted visible tile after structural row and column invalidations', () => {
    const axisMask = DirtyMaskV3.AxisX | DirtyMaskV3.Text | DirtyMaskV3.Rect
    const rowAxisMask = DirtyMaskV3.AxisY | DirtyMaskV3.Text | DirtyMaskV3.Rect
    const columnSubscription = {
      sheetId: 7,
      sheetName: 'Sheet1',
      rowStart: 0,
      rowEnd: VIEWPORT_TILE_ROW_COUNT - 1,
      colStart: 0,
      colEnd: VIEWPORT_TILE_COLUMN_COUNT * 3 - 1,
      dprBucket: 1,
      cameraSeq: 26,
    }
    const rowSubscription = {
      sheetId: 7,
      sheetName: 'Sheet1',
      rowStart: 0,
      rowEnd: VIEWPORT_TILE_ROW_COUNT * 3 - 1,
      colStart: 0,
      colEnd: VIEWPORT_TILE_COLUMN_COUNT - 1,
      dprBucket: 1,
      cameraSeq: 27,
    }

    const columnBatch = buildWorkerRenderTileDeltaBatch({
      engine,
      event: createColumnInvalidationEvent(2),
      generation: 19,
      subscription: columnSubscription,
    })
    const rowBatch = buildWorkerRenderTileDeltaBatch({
      engine,
      event: createRowInvalidationEvent(2),
      generation: 20,
      subscription: rowSubscription,
    })

    const columnReplacements = columnBatch.mutations.filter((mutation) => mutation.kind === 'tileReplace')
    const rowReplacements = rowBatch.mutations.filter((mutation) => mutation.kind === 'tileReplace')

    expect(columnReplacements.map((mutation) => mutation.coord.colTile)).toEqual([0, 1, 2])
    expect(columnReplacements.map((mutation) => mutation.dirtyLocalCols)).toEqual([
      new Uint32Array([2, VIEWPORT_TILE_COLUMN_COUNT - 1]),
      new Uint32Array([0, VIEWPORT_TILE_COLUMN_COUNT - 1]),
      new Uint32Array([0, VIEWPORT_TILE_COLUMN_COUNT - 1]),
    ])
    expect(columnReplacements.map((mutation) => mutation.dirtyLocalRows)).toEqual([
      new Uint32Array([0, VIEWPORT_TILE_ROW_COUNT - 1]),
      new Uint32Array([0, VIEWPORT_TILE_ROW_COUNT - 1]),
      new Uint32Array([0, VIEWPORT_TILE_ROW_COUNT - 1]),
    ])
    expect(columnReplacements.map((mutation) => mutation.dirtyMasks)).toEqual([
      new Uint32Array([axisMask]),
      new Uint32Array([axisMask]),
      new Uint32Array([axisMask]),
    ])

    expect(rowReplacements.map((mutation) => mutation.coord.rowTile)).toEqual([0, 1, 2])
    expect(rowReplacements.map((mutation) => mutation.dirtyLocalRows)).toEqual([
      new Uint32Array([2, VIEWPORT_TILE_ROW_COUNT - 1]),
      new Uint32Array([0, VIEWPORT_TILE_ROW_COUNT - 1]),
      new Uint32Array([0, VIEWPORT_TILE_ROW_COUNT - 1]),
    ])
    expect(rowReplacements.map((mutation) => mutation.dirtyLocalCols)).toEqual([
      new Uint32Array([0, VIEWPORT_TILE_COLUMN_COUNT - 1]),
      new Uint32Array([0, VIEWPORT_TILE_COLUMN_COUNT - 1]),
      new Uint32Array([0, VIEWPORT_TILE_COLUMN_COUNT - 1]),
    ])
    expect(rowReplacements.map((mutation) => mutation.dirtyMasks)).toEqual([
      new Uint32Array([rowAxisMask]),
      new Uint32Array([rowAxisMask]),
      new Uint32Array([rowAxisMask]),
    ])
  })
})
