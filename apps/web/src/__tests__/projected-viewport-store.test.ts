import { describe, expect, it, vi } from 'vitest'
import { SpreadsheetEngine } from '@bilig/core'
import { getGridMetrics } from '@bilig/grid'
import { ValueTag, type RecalcMetrics } from '@bilig/protocol'
import { DirtyMaskV3 } from '../../../../packages/grid/src/renderer-v3/tile-damage-index.js'
import { GRID_RECT_INSTANCE_FLOAT_COUNT_V3 } from '../../../../packages/grid/src/renderer-v3/rect-instance-buffer.js'
import { buildLocalFixedRenderTiles } from '../../../../packages/grid/src/renderer-v3/local-render-tile-materializer.js'
import { encodeViewportPatch, type ViewportPatch, type WorkerEngineClient } from '@bilig/worker-transport'
import { DEFAULT_MAX_CACHED_CELLS_PER_SHEET } from '../projected-viewport-cell-cache.js'
import { ProjectedViewportStore } from '../projected-viewport-store.js'
import { buildViewportPatchFromEngine, DEFAULT_STYLE_ID } from '../worker-runtime-viewport.js'
import { OPTIMISTIC_CELL_SNAPSHOT_FLAG } from '../workbook-optimistic-cell-flags.js'
import { applyOptimisticClearRange } from '../workbook-optimistic-range.js'
import type { WorkerEngine } from '../worker-runtime-support.js'

const TEST_METRICS: RecalcMetrics = {
  batchId: 0,
  changedInputCount: 0,
  dirtyFormulaCount: 0,
  wasmFormulaCount: 0,
  jsFormulaCount: 0,
  rangeNodeVisits: 0,
  recalcMs: 0,
  compileMs: 0,
}
const LOCAL_CELL_VISUAL_DIRTY_MASK = DirtyMaskV3.Value | DirtyMaskV3.Style | DirtyMaskV3.Text | DirtyMaskV3.Rect | DirtyMaskV3.Border

function createPatch(styleId?: string): ViewportPatch {
  return {
    version: 1,
    full: false,
    freezeRows: 0,
    freezeCols: 0,
    viewport: {
      sheetName: 'Sheet1',
      rowStart: 3,
      rowEnd: 7,
      colStart: 2,
      colEnd: 4,
    },
    metrics: TEST_METRICS,
    styles: [],
    cells: [
      {
        row: 4,
        col: 3,
        snapshot: {
          sheetName: 'Sheet1',
          address: 'D5',
          value: { tag: ValueTag.Empty },
          flags: 0,
          version: 1,
          ...(styleId ? { styleId } : {}),
        },
        displayText: '',
        copyText: '',
        editorText: '',
        formatId: 0,
        styleId: styleId ?? 'style-0',
      },
    ],
    columns: [],
    rows: [],
  }
}

function createColumnPatch(size: number, hidden = false): ViewportPatch {
  return {
    ...createPatch(),
    cells: [],
    columns: [{ index: 0, size, hidden }],
  }
}

function createRowPatch(size: number, hidden = false): ViewportPatch {
  return {
    ...createPatch(),
    cells: [],
    rows: [{ index: 0, size, hidden }],
  }
}

function columnLabel(columnIndex: number): string {
  let index = columnIndex + 1
  let label = ''
  while (index > 0) {
    const remainder = (index - 1) % 26
    label = String.fromCharCode(65 + remainder) + label
    index = Math.floor((index - 1) / 26)
  }
  return label
}

function createLargePatch(rowCount: number, columnCount: number): ViewportPatch {
  return {
    version: 1,
    full: false,
    viewport: {
      sheetName: 'Sheet1',
      rowStart: 0,
      rowEnd: rowCount - 1,
      colStart: 0,
      colEnd: columnCount - 1,
    },
    metrics: TEST_METRICS,
    styles: [],
    cells: Array.from({ length: rowCount * columnCount }, (_, index) => {
      const row = Math.floor(index / columnCount)
      const col = index % columnCount
      return {
        row,
        col,
        snapshot: {
          sheetName: 'Sheet1',
          address: `${columnLabel(col)}${row + 1}`,
          value: { tag: ValueTag.Number, value: index },
          flags: 0,
          version: 1,
        },
        displayText: String(index),
        copyText: String(index),
        editorText: String(index),
        formatId: 0,
        styleId: 'style-0',
      }
    }),
    columns: [],
    rows: [],
  }
}

function countSheetCells(cache: ProjectedViewportStore, sheetName: string): number {
  let count = 0
  cache.workbook.getSheet(sheetName)?.grid.forEachCellEntry(() => {
    count += 1
  })
  return count
}

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

function createNoopWorkerEngineClient(): WorkerEngineClient {
  return {
    dispose: vi.fn(),
    invoke: vi.fn(async () => undefined),
    ready: vi.fn(async () => undefined),
    subscribe: vi.fn(() => () => undefined),
    subscribeBatches: vi.fn(() => () => undefined),
    subscribeRenderTileDeltas: vi.fn(() => () => undefined),
    subscribeViewportPatches: vi.fn(() => () => undefined),
    subscribeWorkbookDeltas: vi.fn(() => () => undefined),
  }
}

describe('ProjectedViewportStore', () => {
  it('renders range formatting for empty cells when the patch carries style ids outside the snapshot', () => {
    const cache = new ProjectedViewportStore()

    cache.applyViewportPatch({
      ...createPatch(),
      styles: [{ id: 'style-green', fill: { backgroundColor: '#00ff00' } }],
      cells: createPatch().cells.map((cell) => Object.assign({}, cell, { styleId: 'style-green' })),
    })

    expect(cache.getCell('Sheet1', 'D5').styleId).toBe('style-green')
    expect(cache.getCellStyle(cache.getCell('Sheet1', 'D5').styleId)).toEqual({
      id: 'style-green',
      fill: { backgroundColor: '#00ff00' },
    })
  })

  it('hydrates authoritative snapshot style ranges for empty cells into the viewport cache', async () => {
    const seed = new SpreadsheetEngine({ workbookName: 'viewport-style-range-seed' })
    await seed.ready()
    seed.createSheet('Sheet1')
    seed.setRangeStyle({ sheetName: 'Sheet1', startAddress: 'E6', endAddress: 'E6' }, { fill: { backgroundColor: '#00ff00' } })

    const restored = new SpreadsheetEngine({ workbookName: 'viewport-style-range-restored' }) as SpreadsheetEngine & WorkerEngine
    await restored.ready()
    restored.importSnapshot(seed.exportSnapshot())
    const cache = new ProjectedViewportStore()
    const patch = buildViewportPatchFromEngine({
      authoritativeRevision: 3,
      emptyCellSnapshot: (sheetName, address) => ({
        sheetName,
        address,
        flags: 0,
        value: { tag: ValueTag.Empty },
        version: 0,
      }),
      engine: restored,
      event: null,
      getFormatId: () => 0,
      getStyleRecord: (styleId) => restored.getCellStyle(styleId) ?? { id: DEFAULT_STYLE_ID },
      metrics: { ...TEST_METRICS, batchId: 3 },
      sheetImpact: null,
      state: {
        knownStyleIds: new Set(),
        lastCellSignatures: new Map(),
        lastColumnSignatures: new Map(),
        lastMergeSignatures: new Map(),
        lastRowSignatures: new Map(),
        lastStyleSignatures: new Map(),
        listener: () => undefined,
        nextVersion: 1,
        subscription: {
          sheetName: 'Sheet1',
          rowStart: 0,
          rowEnd: 31,
          colStart: 0,
          colEnd: 127,
        },
      },
    })

    cache.applyViewportPatch(patch)

    expect(restored.getCell('Sheet1', 'E6').styleId).toBeDefined()
    expect(cache.getCell('Sheet1', 'E6').styleId).toBe(restored.getCell('Sheet1', 'E6').styleId)
    expect(cache.getCellStyle(cache.getCell('Sheet1', 'E6').styleId)).toEqual({
      id: restored.getCell('Sheet1', 'E6').styleId,
      fill: { backgroundColor: '#00ff00' },
    })

    const tiles = buildLocalFixedRenderTiles({
      cameraSeq: 1,
      columnWidths: cache.getColumnWidths('Sheet1'),
      dprBucket: 1,
      engine: cache,
      generation: 3,
      gridMetrics: getGridMetrics(),
      rowHeights: cache.getRowHeights('Sheet1'),
      sheetId: 7,
      sheetOrdinal: 7,
      sheetName: 'Sheet1',
      sortedColumnWidthOverrides: [],
      sortedRowHeightOverrides: [],
      viewport: {
        sheetName: 'Sheet1',
        rowStart: 0,
        rowEnd: 31,
        colStart: 0,
        colEnd: 127,
      },
    })
    expect(tiles.some((tile) => hasOpaqueGreenFillRect(tile.rectInstances, tile.rectCount))).toBe(true)
  })

  it('optimistically styles visible empty cells and publishes full visual tile damage', () => {
    const cache = new ProjectedViewportStore(createNoopWorkerEngineClient())
    const deltaListener = vi.fn()
    cache.setSheetIdentities([{ id: 7, name: 'Sheet1', order: 3 }])
    const unsubscribeDeltas = cache.subscribeWorkbookDeltas(deltaListener)
    const unsubscribeViewport = cache.subscribeViewport(
      'Sheet1',
      {
        sheetName: 'Sheet1',
        rowStart: 0,
        rowEnd: 12,
        colStart: 0,
        colEnd: 12,
      },
      () => undefined,
      { initialPatch: 'none' },
    )

    const rollback = cache.setRangeStyle(
      { sheetName: 'Sheet1', startAddress: 'D5', endAddress: 'D5' },
      { fill: { backgroundColor: '#00ff00' } },
    )

    expect(rollback).toEqual(expect.any(Function))
    const styledCell = cache.getCell('Sheet1', 'D5')
    expect(styledCell.styleId).toMatch(/^style-local-/)
    expect(cache.getCellStyle(styledCell.styleId)).toEqual({
      id: styledCell.styleId,
      fill: { backgroundColor: '#00ff00' },
    })
    expect(deltaListener).toHaveBeenCalledWith(
      expect.objectContaining({
        dirty: {
          axisX: new Uint32Array(),
          axisY: new Uint32Array(),
          cellRanges: new Uint32Array([4, 4, 3, 3, LOCAL_CELL_VISUAL_DIRTY_MASK]),
        },
        source: 'localOptimistic',
      }),
    )

    const tiles = buildLocalFixedRenderTiles({
      cameraSeq: 1,
      columnWidths: cache.getColumnWidths('Sheet1'),
      dprBucket: 1,
      engine: cache,
      generation: 3,
      gridMetrics: getGridMetrics(),
      rowHeights: cache.getRowHeights('Sheet1'),
      sheetId: 7,
      sheetOrdinal: 7,
      sheetName: 'Sheet1',
      sortedColumnWidthOverrides: [],
      sortedRowHeightOverrides: [],
      viewport: {
        sheetName: 'Sheet1',
        rowStart: 0,
        rowEnd: 12,
        colStart: 0,
        colEnd: 12,
      },
    })
    expect(tiles.some((tile) => hasOpaqueGreenFillRect(tile.rectInstances, tile.rectCount))).toBe(true)

    deltaListener.mockClear()
    rollback?.()
    expect(countSheetCells(cache, 'Sheet1')).toBe(0)
    expect(deltaListener).toHaveBeenCalledWith(
      expect.objectContaining({
        dirty: {
          axisX: new Uint32Array(),
          axisY: new Uint32Array(),
          cellRanges: new Uint32Array([4, 4, 3, 3, LOCAL_CELL_VISUAL_DIRTY_MASK]),
        },
        source: 'localOptimistic',
      }),
    )
    unsubscribeViewport()
    unsubscribeDeltas()
  })

  it('resolves large range styles for cells that become visible after the mutation', () => {
    const cache = new ProjectedViewportStore(createNoopWorkerEngineClient())
    cache.setSheetIdentities([{ id: 7, name: 'Sheet1', order: 3 }])

    const rollback = cache.setRangeStyle(
      { sheetName: 'Sheet1', startAddress: 'D5', endAddress: 'F900' },
      { fill: { backgroundColor: '#00ff00' } },
    )

    expect(rollback).toEqual(expect.any(Function))
    expect(countSheetCells(cache, 'Sheet1')).toBe(0)
    const offscreenCell = cache.getCell('Sheet1', 'E700')
    expect(cache.getCellStyle(offscreenCell.styleId)).toEqual({
      id: offscreenCell.styleId,
      fill: { backgroundColor: '#00ff00' },
    })
    expect(countSheetCells(cache, 'Sheet1')).toBe(0)

    const unsubscribeViewport = cache.subscribeViewport(
      'Sheet1',
      {
        sheetName: 'Sheet1',
        rowStart: 695,
        rowEnd: 704,
        colStart: 3,
        colEnd: 5,
      },
      () => undefined,
      { initialPatch: 'none' },
    )

    expect(countSheetCells(cache, 'Sheet1')).toBeLessThanOrEqual(30)
    const materializedCell = cache.getCell('Sheet1', 'E700')
    expect(cache.getCellStyle(materializedCell.styleId)).toEqual({
      id: materializedCell.styleId,
      fill: { backgroundColor: '#00ff00' },
    })
    const tiles = buildLocalFixedRenderTiles({
      cameraSeq: 1,
      columnWidths: cache.getColumnWidths('Sheet1'),
      dprBucket: 1,
      engine: cache,
      generation: 3,
      gridMetrics: getGridMetrics(),
      rowHeights: cache.getRowHeights('Sheet1'),
      sheetId: 7,
      sheetOrdinal: 7,
      sheetName: 'Sheet1',
      sortedColumnWidthOverrides: [],
      sortedRowHeightOverrides: [],
      viewport: {
        sheetName: 'Sheet1',
        rowStart: 695,
        rowEnd: 704,
        colStart: 3,
        colEnd: 5,
      },
    })
    expect(tiles.some((tile) => hasOpaqueGreenFillRect(tile.rectInstances, tile.rectCount))).toBe(true)

    rollback?.()
    expect(countSheetCells(cache, 'Sheet1')).toBe(0)
    expect(cache.peekCell('Sheet1', 'E700')).toBeUndefined()
    unsubscribeViewport()
  })

  it('lets newer small style edits override an older large range overlay', () => {
    const cache = new ProjectedViewportStore(createNoopWorkerEngineClient())

    const rollbackLarge = cache.setRangeStyle(
      { sheetName: 'Sheet1', startAddress: 'D5', endAddress: 'F900' },
      { fill: { backgroundColor: '#00ff00' } },
    )
    const rollbackSmall = cache.setRangeStyle(
      { sheetName: 'Sheet1', startAddress: 'E700', endAddress: 'E700' },
      { fill: { backgroundColor: '#a4c2f4' } },
    )

    const overriddenCell = cache.getCell('Sheet1', 'E700')
    expect(cache.getCellStyle(overriddenCell.styleId)).toEqual({
      id: overriddenCell.styleId,
      fill: { backgroundColor: '#a4c2f4' },
    })
    const neighborCell = cache.getCell('Sheet1', 'D700')
    expect(cache.getCellStyle(neighborCell.styleId)).toEqual({
      id: neighborCell.styleId,
      fill: { backgroundColor: '#00ff00' },
    })

    rollbackSmall?.()
    const restoredCell = cache.getCell('Sheet1', 'E700')
    expect(cache.getCellStyle(restoredCell.styleId)).toEqual({
      id: restoredCell.styleId,
      fill: { backgroundColor: '#00ff00' },
    })
    rollbackLarge?.()
    expect(cache.peekCell('Sheet1', 'E700')).toBeUndefined()
  })

  it('keeps large clears visible for future viewports without materializing the whole range', () => {
    const cache = new ProjectedViewportStore(createNoopWorkerEngineClient())

    const rollback = applyOptimisticClearRange(cache, {
      sheetName: 'Sheet1',
      startAddress: 'A1',
      endAddress: 'D3000',
    })

    expect(rollback).toEqual(expect.any(Function))
    expect(countSheetCells(cache, 'Sheet1')).toBe(0)
    const optimisticClear = cache.peekCell('Sheet1', 'D2500')
    expect(optimisticClear).toMatchObject({
      flags: OPTIMISTIC_CELL_SNAPSHOT_FLAG,
      value: { tag: ValueTag.Empty },
    })
    expect(countSheetCells(cache, 'Sheet1')).toBe(0)

    const unsubscribeViewport = cache.subscribeViewport(
      'Sheet1',
      {
        sheetName: 'Sheet1',
        rowStart: 2495,
        rowEnd: 2504,
        colStart: 0,
        colEnd: 3,
      },
      () => undefined,
      { initialPatch: 'none' },
    )

    expect(countSheetCells(cache, 'Sheet1')).toBeLessThanOrEqual(40)
    expect(cache.getCell('Sheet1', 'D2500')).toMatchObject({
      flags: OPTIMISTIC_CELL_SNAPSHOT_FLAG,
      value: { tag: ValueTag.Empty },
    })

    rollback?.()
    expect(countSheetCells(cache, 'Sheet1')).toBe(0)
    expect(cache.peekCell('Sheet1', 'D2500')).toBeUndefined()
    unsubscribeViewport()
  })

  it('optimistically styles protected local edit snapshots instead of waiting for worker readback', () => {
    const cache = new ProjectedViewportStore()
    cache.setCellSnapshot({
      sheetName: 'Sheet1',
      address: 'D5',
      flags: OPTIMISTIC_CELL_SNAPSHOT_FLAG,
      input: 'moved-fill-proof',
      value: { tag: ValueTag.String, value: 'moved-fill-proof', stringId: 0 },
      version: 4,
    })

    const rollback = cache.setRangeStyle(
      { sheetName: 'Sheet1', startAddress: 'D5', endAddress: 'D5' },
      { fill: { backgroundColor: '#00ff00' } },
    )

    const styledCell = cache.getCell('Sheet1', 'D5')
    expect(styledCell).toMatchObject({
      flags: OPTIMISTIC_CELL_SNAPSHOT_FLAG,
      input: 'moved-fill-proof',
      value: { tag: ValueTag.String, value: 'moved-fill-proof' },
    })
    expect(cache.getCellStyle(styledCell.styleId)).toEqual({
      id: styledCell.styleId,
      fill: { backgroundColor: '#00ff00' },
    })

    rollback?.()
    expect(cache.getCell('Sheet1', 'D5').styleId).toBeUndefined()
  })

  it('accepts equal-version empty snapshots that clear stale styling', () => {
    const cache = new ProjectedViewportStore()

    cache.applyViewportPatch(createPatch('style-red'))
    expect(cache.getCell('Sheet1', 'D5').styleId).toBe('style-red')

    cache.applyViewportPatch(createPatch())

    expect(cache.getCell('Sheet1', 'D5').styleId).toBeUndefined()
  })

  it('exposes authoritative, projected, and tile-scene render revisions for visible proof checks', () => {
    const cache = new ProjectedViewportStore()

    cache.applyViewportPatch({
      ...createPatch(),
      authoritativeRevision: 17,
      metrics: {
        ...TEST_METRICS,
        batchId: 23,
      },
    })

    expect(cache.getRenderRevisionSnapshot()).toEqual({
      authoritativeRevision: 17,
      localRevision: 0,
      projectedRevision: 23,
      tileSceneCameraSeq: null,
      tileSceneRevision: null,
    })
  })

  it('rejects stale authoritative viewport patches before they regress rendered cells, axes, or proof revisions', () => {
    const cache = new ProjectedViewportStore()

    cache.applyViewportPatch({
      ...createPatch(),
      authoritativeRevision: 17,
      metrics: {
        ...TEST_METRICS,
        batchId: 23,
      },
      columns: [{ index: 0, size: 93, hidden: true }],
      rows: [{ index: 0, size: 44, hidden: true }],
      cells: [
        {
          row: 4,
          col: 3,
          snapshot: {
            sheetName: 'Sheet1',
            address: 'D5',
            value: { tag: ValueTag.String, value: 'fresh', stringId: 1 },
            input: 'fresh',
            flags: 0,
            version: 17,
          },
          displayText: 'fresh',
          copyText: 'fresh',
          editorText: 'fresh',
          formatId: 0,
          styleId: 'style-0',
        },
      ],
    })

    const staleDamage = cache.applyViewportPatch({
      ...createPatch(),
      authoritativeRevision: 12,
      metrics: {
        ...TEST_METRICS,
        batchId: 18,
      },
      freezeRows: 4,
      freezeCols: 3,
      columns: [{ index: 0, size: 68, hidden: false }],
      rows: [{ index: 0, size: 30, hidden: false }],
      cells: [
        {
          row: 4,
          col: 3,
          snapshot: {
            sheetName: 'Sheet1',
            address: 'D5',
            value: { tag: ValueTag.String, value: 'stale', stringId: 2 },
            input: 'stale',
            flags: 0,
            version: 99,
          },
          displayText: 'stale',
          copyText: 'stale',
          editorText: 'stale',
          formatId: 0,
          styleId: 'style-0',
        },
      ],
    })

    expect(staleDamage).toEqual([])
    expect(cache.getCell('Sheet1', 'D5')).toMatchObject({
      value: { tag: ValueTag.String, value: 'fresh', stringId: 1 },
      input: 'fresh',
      version: 17,
    })
    expect(cache.getColumnWidths('Sheet1')[0]).toBe(0)
    expect(cache.getColumnSizes('Sheet1')[0]).toBe(93)
    expect(cache.getHiddenColumns('Sheet1')[0]).toBe(true)
    expect(cache.getRowHeights('Sheet1')[0]).toBe(0)
    expect(cache.getRowSizes('Sheet1')[0]).toBe(44)
    expect(cache.getHiddenRows('Sheet1')[0]).toBe(true)
    expect(cache.getFreezeRows('Sheet1')).toBe(0)
    expect(cache.getFreezeCols('Sheet1')).toBe(0)
    expect(cache.getRenderRevisionSnapshot()).toEqual({
      authoritativeRevision: 17,
      localRevision: 0,
      projectedRevision: 23,
      tileSceneCameraSeq: null,
      tileSceneRevision: null,
    })
  })

  it('increments the local render revision for optimistic cell snapshots', () => {
    const cache = new ProjectedViewportStore()

    expect(cache.getRenderRevisionSnapshot().localRevision).toBe(0)

    cache.setCellSnapshot({
      sheetName: 'Sheet1',
      address: 'D5',
      flags: 0,
      input: 'local',
      value: { tag: ValueTag.String, value: 'local', stringId: 0 },
      version: 1,
    })

    expect(cache.getRenderRevisionSnapshot().localRevision).toBe(1)
  })

  it('hydrates selected-cell cache without publishing render tile deltas', () => {
    const cache = new ProjectedViewportStore(createNoopWorkerEngineClient())
    const listener = vi.fn()
    cache.setSheetIdentities([{ id: 7, name: 'Sheet1', order: 0 }])
    const unsubscribeDeltas = cache.subscribeWorkbookDeltas(listener)

    expect(cache.getRenderRevisionSnapshot().localRevision).toBe(0)

    cache.setCellSnapshot(
      {
        sheetName: 'Sheet1',
        address: 'D5',
        flags: 0,
        input: 'hydrated',
        value: { tag: ValueTag.String, value: 'hydrated', stringId: 0 },
        version: 12,
      },
      { emitLocalDelta: false },
    )

    expect(cache.getCell('Sheet1', 'D5').input).toBe('hydrated')
    expect(cache.getRenderRevisionSnapshot().localRevision).toBe(0)
    expect(listener).not.toHaveBeenCalled()

    unsubscribeDeltas()
  })

  it('can force authoritative selected-cell hydration over stale optimistic input', () => {
    const cache = new ProjectedViewportStore()

    cache.setCellSnapshot({
      sheetName: 'Sheet1',
      address: 'D5',
      flags: OPTIMISTIC_CELL_SNAPSHOT_FLAG,
      input: 'ghost-content',
      value: { tag: ValueTag.String, value: 'ghost-content', stringId: 0 },
      version: 8,
    })

    cache.setCellSnapshot(
      {
        sheetName: 'Sheet1',
        address: 'D5',
        flags: 0,
        value: { tag: ValueTag.Empty },
        version: 0,
      },
      { force: true, forceOptimistic: true },
    )

    expect(cache.getCell('Sheet1', 'D5')).toEqual({
      sheetName: 'Sheet1',
      address: 'D5',
      flags: 0,
      value: { tag: ValueTag.Empty },
      version: 0,
    })
  })

  it('accepts partial reset-empty patches that clear newer local input after structural deletes', () => {
    const cache = new ProjectedViewportStore()

    cache.setCellSnapshot({
      sheetName: 'Sheet1',
      address: 'D5',
      value: { tag: ValueTag.String, value: 'stale-tail', stringId: 1 },
      input: 'stale-tail',
      flags: 0,
      version: 7,
    })

    cache.applyViewportPatch({
      ...createPatch(),
      authoritativeRevision: 0,
      full: false,
      cells: [
        {
          row: 4,
          col: 3,
          snapshot: {
            sheetName: 'Sheet1',
            address: 'D5',
            value: { tag: ValueTag.Empty },
            flags: 0,
            version: 0,
          },
          displayText: '',
          copyText: '',
          editorText: '',
          formatId: 0,
          styleId: 'style-0',
        },
      ],
    })

    expect(cache.getCell('Sheet1', 'D5').value).toEqual({ tag: ValueTag.Empty })
    expect(cache.getCell('Sheet1', 'D5').input).toBeUndefined()
  })

  it('keeps newer local input when a full reset-empty patch arrives during hydration', () => {
    const cache = new ProjectedViewportStore()

    cache.setCellSnapshot({
      sheetName: 'Sheet1',
      address: 'D5',
      value: { tag: ValueTag.String, value: 'local-edit', stringId: 1 },
      input: 'local-edit',
      flags: 0,
      version: 7,
    })

    cache.applyViewportPatch({
      ...createPatch(),
      full: true,
      cells: [
        {
          row: 4,
          col: 3,
          snapshot: {
            sheetName: 'Sheet1',
            address: 'D5',
            value: { tag: ValueTag.Empty },
            flags: 0,
            version: 0,
          },
          displayText: '',
          copyText: '',
          editorText: '',
          formatId: 0,
          styleId: 'style-0',
        },
      ],
    })

    expect(cache.getCell('Sheet1', 'D5')).toMatchObject({
      value: { tag: ValueTag.String, value: 'local-edit' },
      input: 'local-edit',
      version: 7,
    })
  })

  it('keeps an equal-version local formula snapshot when a later patch drops the formula', () => {
    const cache = new ProjectedViewportStore()

    cache.applyViewportPatch({
      ...createPatch(),
      cells: [
        {
          row: 4,
          col: 3,
          snapshot: {
            sheetName: 'Sheet1',
            address: 'D5',
            value: { tag: ValueTag.Boolean, value: true },
            input: '=A1="HELLO"',
            formula: 'A1="HELLO"',
            flags: 0,
            version: 3,
          },
          displayText: 'TRUE',
          copyText: '=A1="HELLO"',
          editorText: '=A1="HELLO"',
          formatId: 0,
          styleId: 'style-0',
        },
      ],
    })

    cache.applyViewportPatch({
      ...createPatch(),
      cells: [
        {
          row: 4,
          col: 3,
          snapshot: {
            sheetName: 'Sheet1',
            address: 'D5',
            value: { tag: ValueTag.Boolean, value: false },
            flags: 0,
            version: 3,
          },
          displayText: 'FALSE',
          copyText: 'FALSE',
          editorText: 'FALSE',
          formatId: 0,
          styleId: 'style-0',
        },
      ],
    })

    expect(cache.getCell('Sheet1', 'D5')).toMatchObject({
      value: { tag: ValueTag.Boolean, value: true },
      formula: 'A1="HELLO"',
      version: 3,
    })
  })

  it('keeps a local formula snapshot when a newer eval-only patch drops source metadata', () => {
    const cache = new ProjectedViewportStore()

    cache.applyViewportPatch({
      ...createPatch(),
      cells: [
        {
          row: 4,
          col: 3,
          snapshot: {
            sheetName: 'Sheet1',
            address: 'D5',
            value: { tag: ValueTag.Boolean, value: true },
            input: '=A1="HELLO"',
            formula: 'A1="HELLO"',
            flags: 0,
            version: 3,
          },
          displayText: 'TRUE',
          copyText: '=A1="HELLO"',
          editorText: '=A1="HELLO"',
          formatId: 0,
          styleId: 'style-0',
        },
      ],
    })

    cache.applyViewportPatch({
      ...createPatch(),
      cells: [
        {
          row: 4,
          col: 3,
          snapshot: {
            sheetName: 'Sheet1',
            address: 'D5',
            value: { tag: ValueTag.Boolean, value: false },
            flags: 0,
            version: 4,
          },
          displayText: 'FALSE',
          copyText: 'FALSE',
          editorText: 'FALSE',
          formatId: 0,
          styleId: 'style-0',
        },
      ],
    })

    expect(cache.getCell('Sheet1', 'D5')).toMatchObject({
      value: { tag: ValueTag.Boolean, value: true },
      formula: 'A1="HELLO"',
      version: 3,
    })
  })

  it('keeps a local formula snapshot when a direct cell refresh drops source metadata', () => {
    const cache = new ProjectedViewportStore()

    cache.setCellSnapshot({
      sheetName: 'Sheet1',
      address: 'D5',
      value: { tag: ValueTag.Boolean, value: true },
      input: '=A1="HELLO"',
      formula: 'A1="HELLO"',
      flags: 0,
      version: 3,
    })

    cache.setCellSnapshot({
      sheetName: 'Sheet1',
      address: 'D5',
      value: { tag: ValueTag.Boolean, value: false },
      flags: 0,
      version: 4,
    })

    expect(cache.getCell('Sheet1', 'D5')).toMatchObject({
      value: { tag: ValueTag.Boolean, value: true },
      formula: 'A1="HELLO"',
      version: 3,
    })
  })

  it('keeps an optimistic formula snapshot when an eval-only patch drops source metadata', () => {
    const cache = new ProjectedViewportStore()

    cache.setCellSnapshot({
      sheetName: 'Sheet1',
      address: 'A2',
      value: { tag: ValueTag.Boolean, value: true },
      input: '=A1="HELLO"',
      formula: 'A1="HELLO"',
      flags: OPTIMISTIC_CELL_SNAPSHOT_FLAG,
      version: 1,
    })

    cache.applyViewportPatch({
      ...createPatch(),
      cells: [
        {
          row: 1,
          col: 0,
          snapshot: {
            sheetName: 'Sheet1',
            address: 'A2',
            value: { tag: ValueTag.Boolean, value: false },
            flags: 0,
            version: 1,
          },
          displayText: 'FALSE',
          copyText: 'FALSE',
          editorText: 'FALSE',
          formatId: 0,
          styleId: 'style-0',
        },
      ],
    })

    expect(cache.getCell('Sheet1', 'A2')).toMatchObject({
      value: { tag: ValueTag.Boolean, value: true },
      formula: 'A1="HELLO"',
      flags: OPTIMISTIC_CELL_SNAPSHOT_FLAG,
      version: 1,
    })
  })

  it('accepts a newer literal snapshot when the source input is present', () => {
    const cache = new ProjectedViewportStore()

    cache.applyViewportPatch({
      ...createPatch(),
      cells: [
        {
          row: 4,
          col: 3,
          snapshot: {
            sheetName: 'Sheet1',
            address: 'D5',
            value: { tag: ValueTag.Boolean, value: true },
            input: '=A1="HELLO"',
            formula: 'A1="HELLO"',
            flags: 0,
            version: 3,
          },
          displayText: 'TRUE',
          copyText: '=A1="HELLO"',
          editorText: '=A1="HELLO"',
          formatId: 0,
          styleId: 'style-0',
        },
      ],
    })

    cache.applyViewportPatch({
      ...createPatch(),
      cells: [
        {
          row: 4,
          col: 3,
          snapshot: {
            sheetName: 'Sheet1',
            address: 'D5',
            value: { tag: ValueTag.Boolean, value: false },
            input: false,
            flags: 0,
            version: 4,
          },
          displayText: 'FALSE',
          copyText: 'FALSE',
          editorText: 'FALSE',
          formatId: 0,
          styleId: 'style-0',
        },
      ],
    })

    const snapshot = cache.getCell('Sheet1', 'D5')
    expect(snapshot).toMatchObject({
      value: { tag: ValueTag.Boolean, value: false },
      input: false,
      version: 4,
    })
    expect('formula' in snapshot).toBe(false)
  })

  it('reports damage when a style record changes without a newer cell snapshot', () => {
    const cache = new ProjectedViewportStore()

    cache.applyViewportPatch({
      ...createPatch('style-fill'),
      styles: [{ id: 'style-fill', fill: { backgroundColor: '#c9daf8' } }],
    })

    const damage = cache.applyViewportPatch({
      ...createPatch('style-fill'),
      styles: [{ id: 'style-fill', fill: { backgroundColor: '#a4c2f4' } }],
    })

    expect(damage).toEqual([{ cell: [3, 4] }])
    expect(cache.getCellStyle('style-fill')).toEqual({
      id: 'style-fill',
      fill: { backgroundColor: '#a4c2f4' },
    })
  })

  it('tracks freeze pane metadata from viewport patches', () => {
    const cache = new ProjectedViewportStore()

    cache.applyViewportPatch({
      ...createPatch(),
      freezeRows: 2,
      freezeCols: 1,
    })

    expect(cache.getFreezeRows('Sheet1')).toBe(2)
    expect(cache.getFreezeCols('Sheet1')).toBe(1)
  })

  it('does not notify freeze subscribers when a patch confirms the default unfrozen state', () => {
    const cache = new ProjectedViewportStore()
    const freezeListener = vi.fn()

    const unsubscribeFreeze = cache.subscribeSheetChannel('Sheet1', 'freeze', freezeListener)

    cache.applyViewportPatch({
      ...createPatch(),
      freezeRows: 0,
      freezeCols: 0,
    })

    expect(freezeListener).not.toHaveBeenCalled()

    unsubscribeFreeze()
  })

  it('clears stale viewport cells on full patches without dropping cells outside the viewport', () => {
    const cache = new ProjectedViewportStore()

    cache.setCellSnapshot({
      sheetName: 'Sheet1',
      address: 'A1',
      value: { tag: ValueTag.String, value: 'pinned', stringId: 1 },
      flags: 0,
      version: 1,
    })
    cache.applyViewportPatch({ ...createPatch(), full: true })

    const damage = cache.applyViewportPatch({
      ...createPatch(),
      full: true,
      cells: [],
    })

    expect(damage).toEqual([{ cell: [3, 4] }])
    expect(cache.peekCell('Sheet1', 'D5')).toBeUndefined()
    expect(cache.getCell('Sheet1', 'A1').value).toEqual({
      tag: ValueTag.String,
      value: 'pinned',
      stringId: 1,
    })
  })

  it('drops stale sheet cache entries when sheets disappear', () => {
    const cache = new ProjectedViewportStore()

    cache.applyViewportPatch(createPatch())
    expect(cache.peekCell('Sheet1', 'D5')).toBeDefined()

    cache.setKnownSheets(['Sheet2'])

    expect(cache.peekCell('Sheet1', 'D5')).toBeUndefined()
  })

  it('resets same-sheet projected state before installing a replacement authoritative snapshot', () => {
    const cache = new ProjectedViewportStore()

    cache.applyViewportPatch(createPatch('style-red'))
    cache.setColumnWidth('Sheet1', 0, 68)
    cache.setRowHeight('Sheet1', 0, 240)

    cache.resetProjectionState()

    expect(cache.peekCell('Sheet1', 'D5')).toBeUndefined()
    expect(cache.workbook.getSheet('Sheet1')).toBeDefined()
    expect(cache.getColumnWidths('Sheet1')[0]).toBeUndefined()
    expect(cache.getRowHeights('Sheet1')[0]).toBeUndefined()
  })

  it('publishes sparse local column axis deltas without materializing default axis sizes', () => {
    const cache = new ProjectedViewportStore(createNoopWorkerEngineClient())
    const events: string[] = []
    const unsubscribeRenderTiles = cache.subscribeRenderTileDeltas(
      {
        sheetId: 7,
        sheetName: 'Sheet1',
        sheetOrdinal: 3,
        rowStart: 0,
        rowEnd: 31,
        colStart: 0,
        colEnd: 63,
      },
      () => undefined,
    )
    const unsubscribeAxis = cache.subscribeSheetChannel('Sheet1', 'columnWidths', () => {
      events.push(`axis:${cache.getColumnWidths('Sheet1')[2]}`)
    })
    const unsubscribeDeltas = cache.subscribeWorkbookDeltas((batch) => {
      events.push(`delta:${cache.getColumnWidths('Sheet1')[2]}:${[...batch.dirty.axisX].join(':')}:${batch.dirty.axisY.length}`)
    })

    cache.setColumnWidth('Sheet1', 2, 144)

    expect(events).toEqual(['axis:144', 'delta:144:2:2:44:0'])
    expect(cache.getColumnWidths('Sheet1')).toEqual({ 2: 144 })

    unsubscribeAxis()
    unsubscribeDeltas()
    unsubscribeRenderTiles()
  })

  it('publishes sparse local row axis deltas without materializing default axis sizes', () => {
    const cache = new ProjectedViewportStore(createNoopWorkerEngineClient())
    const events: string[] = []
    const unsubscribeRenderTiles = cache.subscribeRenderTileDeltas(
      {
        sheetId: 7,
        sheetName: 'Sheet1',
        sheetOrdinal: 3,
        rowStart: 0,
        rowEnd: 31,
        colStart: 0,
        colEnd: 63,
      },
      () => undefined,
    )
    const unsubscribeAxis = cache.subscribeSheetChannel('Sheet1', 'rowHeights', () => {
      events.push(`axis:${cache.getRowHeights('Sheet1')[3]}`)
    })
    const unsubscribeDeltas = cache.subscribeWorkbookDeltas((batch) => {
      events.push(`delta:${cache.getRowHeights('Sheet1')[3]}:${batch.dirty.axisX.length}:${[...batch.dirty.axisY].join(':')}`)
    })

    cache.setRowHeight('Sheet1', 3, 48)

    expect(events).toEqual(['axis:48', 'delta:48:0:3:3:76'])
    expect(cache.getRowHeights('Sheet1')).toEqual({ 3: 48 })

    unsubscribeAxis()
    unsubscribeDeltas()
    unsubscribeRenderTiles()
  })

  it('rejects invalid local axis mutations before state and dirty deltas diverge', () => {
    const cache = new ProjectedViewportStore(createNoopWorkerEngineClient())
    const listener = vi.fn()

    cache.setSheetIdentities([{ id: 7, name: 'Sheet1', order: 3 }])
    const unsubscribeDeltas = cache.subscribeWorkbookDeltas(listener)

    expect(() => cache.setColumnWidth('Sheet1', -1, 144)).toThrow('Invalid projected column index: -1')
    expect(() => cache.setRowHeight('Sheet1', 1.5, 48)).toThrow('Invalid projected row index: 1.5')
    expect(() => cache.setColumnHidden('Sheet1', 0, true, Number.NaN)).toThrow('Invalid projected column size: NaN')
    expect(() => cache.rollbackRowHidden('Sheet1', 0, { hidden: false, size: Number.POSITIVE_INFINITY })).toThrow(
      'Invalid projected row size: Infinity',
    )

    expect(cache.getColumnWidths('Sheet1')).toEqual({})
    expect(cache.getRowHeights('Sheet1')).toEqual({})
    expect(listener).not.toHaveBeenCalled()

    unsubscribeDeltas()
  })

  it('drops stale sheet identities when runtime state publishes a replacement sheet list', () => {
    const cache = new ProjectedViewportStore(createNoopWorkerEngineClient())
    const listener = vi.fn()

    cache.setSheetIdentities([{ id: 7, name: 'Sheet1', order: 0 }])
    cache.setSheetIdentities([{ id: 8, name: 'Sheet2', order: 0 }])
    const unsubscribeDeltas = cache.subscribeWorkbookDeltas(listener)

    cache.setCellSnapshot({
      address: 'A1',
      flags: 0,
      sheetName: 'Sheet1',
      value: { tag: ValueTag.Number, value: 17 },
      version: 12,
    })

    expect(listener).not.toHaveBeenCalled()

    cache.setCellSnapshot({
      address: 'A1',
      flags: 0,
      sheetName: 'Sheet2',
      value: { tag: ValueTag.Number, value: 18 },
      version: 13,
    })

    expect(listener).toHaveBeenCalledWith(
      expect.objectContaining({
        sheetId: 8,
        sheetOrdinal: 0,
        source: 'localOptimistic',
      }),
    )

    unsubscribeDeltas()
  })

  it('clears a pending local column width once the authoritative patch matches it', () => {
    const cache = new ProjectedViewportStore()

    cache.setColumnWidth('Sheet1', 0, 68)
    cache.applyViewportPatch(createColumnPatch(68))
    cache.applyViewportPatch(createColumnPatch(93))

    expect(cache.getColumnWidths('Sheet1')[0]).toBe(93)
  })

  it('keeps pending local column width visible across older authoritative patches', () => {
    const cache = new ProjectedViewportStore()

    cache.setColumnWidth('Sheet1', 0, 68)
    cache.applyViewportPatch(createColumnPatch(93))

    expect(cache.getColumnWidths('Sheet1')[0]).toBe(68)
    expect(cache.getColumnSizes('Sheet1')[0]).toBe(93)

    cache.applyViewportPatch(createColumnPatch(68))
    cache.applyViewportPatch(createColumnPatch(104))

    expect(cache.getColumnWidths('Sheet1')[0] ?? 104).toBe(104)
    expect(cache.getColumnSizes('Sheet1')[0] ?? 104).toBe(104)
  })

  it('keeps default authoritative axis sizes sparse after a local resize', () => {
    const cache = new ProjectedViewportStore()

    cache.applyViewportPatch({
      ...createPatch(),
      cells: [],
      columns: Array.from({ length: 8 }, (_, index) => ({ index, size: 104, hidden: false })),
      rows: Array.from({ length: 8 }, (_, index) => ({ index, size: 22, hidden: false })),
    })
    cache.setColumnWidth('Sheet1', 2, 144)
    cache.setRowHeight('Sheet1', 3, 48)

    expect(cache.getColumnWidths('Sheet1')).toEqual({ 2: 144 })
    expect(cache.getRowHeights('Sheet1')).toEqual({ 3: 48 })
  })

  it('rolls back a failed local column width mutation without leaving a pending width behind', () => {
    const cache = new ProjectedViewportStore()

    cache.setColumnWidth('Sheet1', 0, 68)
    cache.rollbackColumnWidth('Sheet1', 0, undefined)
    cache.applyViewportPatch(createColumnPatch(104))

    expect(cache.getColumnWidths('Sheet1')[0]).toBeUndefined()
  })

  it('clears a pending local row height once the authoritative patch matches it', () => {
    const cache = new ProjectedViewportStore()

    cache.setRowHeight('Sheet1', 0, 30)
    cache.applyViewportPatch(createRowPatch(30))
    cache.applyViewportPatch(createRowPatch(44))

    expect(cache.getRowHeights('Sheet1')[0]).toBe(44)
  })

  it('keeps pending local row height visible across older authoritative patches', () => {
    const cache = new ProjectedViewportStore()

    cache.setRowHeight('Sheet1', 0, 30)
    cache.applyViewportPatch(createRowPatch(22))

    expect(cache.getRowHeights('Sheet1')[0]).toBe(30)
    expect(cache.getRowSizes('Sheet1')[0] ?? 22).toBe(22)

    cache.applyViewportPatch(createRowPatch(30))
    cache.applyViewportPatch(createRowPatch(44))

    expect(cache.getRowHeights('Sheet1')[0]).toBe(44)
    expect(cache.getRowSizes('Sheet1')[0]).toBe(44)
  })

  it('rolls back a failed local row height mutation without leaving a pending height behind', () => {
    const cache = new ProjectedViewportStore()

    cache.setRowHeight('Sheet1', 0, 30)
    cache.rollbackRowHeight('Sheet1', 0, undefined)
    cache.applyViewportPatch(createRowPatch(22))

    expect(cache.getRowHeights('Sheet1')[0]).toBeUndefined()
  })

  it('preserves hidden column metadata and collapses hidden columns from the visible axis map', () => {
    const cache = new ProjectedViewportStore()

    cache.applyViewportPatch(createColumnPatch(93, true))

    expect(cache.getColumnWidths('Sheet1')[0]).toBe(0)
    expect(cache.getColumnSizes('Sheet1')[0]).toBe(93)
    expect(cache.getHiddenColumns('Sheet1')[0]).toBe(true)

    cache.applyViewportPatch(createColumnPatch(93, false))

    expect(cache.getColumnWidths('Sheet1')[0]).toBe(93)
    expect(cache.getHiddenColumns('Sheet1')[0]).toBeUndefined()
  })

  it('preserves hidden row metadata and collapses hidden rows from the visible axis map', () => {
    const cache = new ProjectedViewportStore()

    cache.applyViewportPatch(createRowPatch(44, true))

    expect(cache.getRowHeights('Sheet1')[0]).toBe(0)
    expect(cache.getRowSizes('Sheet1')[0]).toBe(44)
    expect(cache.getHiddenRows('Sheet1')[0]).toBe(true)

    cache.applyViewportPatch(createRowPatch(44, false))

    expect(cache.getRowHeights('Sheet1')[0]).toBe(44)
    expect(cache.getHiddenRows('Sheet1')[0]).toBeUndefined()
  })

  it('supports optimistic column hide and rollback using preserved raw sizes', () => {
    const cache = new ProjectedViewportStore()

    cache.applyViewportPatch(createColumnPatch(93))
    cache.setColumnHidden('Sheet1', 0, true, 93)

    expect(cache.getColumnWidths('Sheet1')[0]).toBe(0)
    expect(cache.getColumnSizes('Sheet1')[0]).toBe(93)
    expect(cache.getHiddenColumns('Sheet1')[0]).toBe(true)

    cache.rollbackColumnHidden('Sheet1', 0, { hidden: false, size: 93 })

    expect(cache.getColumnWidths('Sheet1')[0]).toBe(93)
    expect(cache.getColumnSizes('Sheet1')[0]).toBe(93)
    expect(cache.getHiddenColumns('Sheet1')[0]).toBeUndefined()
  })

  it('supports optimistic row hide and rollback using preserved raw sizes', () => {
    const cache = new ProjectedViewportStore()

    cache.applyViewportPatch(createRowPatch(44))
    cache.setRowHidden('Sheet1', 0, true, 44)

    expect(cache.getRowHeights('Sheet1')[0]).toBe(0)
    expect(cache.getRowSizes('Sheet1')[0]).toBe(44)
    expect(cache.getHiddenRows('Sheet1')[0]).toBe(true)

    cache.rollbackRowHidden('Sheet1', 0, { hidden: false, size: 44 })

    expect(cache.getRowHeights('Sheet1')[0]).toBe(44)
    expect(cache.getRowSizes('Sheet1')[0]).toBe(44)
    expect(cache.getHiddenRows('Sheet1')[0]).toBeUndefined()
  })

  it('prunes back to the cache cap after the last viewport unsubscribes', () => {
    const cache = new ProjectedViewportStore({
      invoke: async () => undefined,
      ready: async () => undefined,
      subscribe: () => () => undefined,
      subscribeBatches: () => () => undefined,
      subscribeViewportPatches: () => () => undefined,
      dispose: () => undefined,
    })

    const unsubscribe = cache.subscribeViewport('Sheet1', { rowStart: 0, rowEnd: 600, colStart: 0, colEnd: 9 }, () => undefined)

    cache.applyViewportPatch(createLargePatch(601, 10))

    expect(countSheetCells(cache, 'Sheet1')).toBe(6010)

    unsubscribe()

    expect(countSheetCells(cache, 'Sheet1')).toBe(DEFAULT_MAX_CACHED_CELLS_PER_SHEET)
  })

  it('supports a configurable viewport cache cap', () => {
    const cache = new ProjectedViewportStore(
      {
        invoke: async () => undefined,
        ready: async () => undefined,
        subscribe: () => () => undefined,
        subscribeBatches: () => () => undefined,
        subscribeViewportPatches: () => () => undefined,
        dispose: () => undefined,
      },
      {
        maxCachedCellsPerSheet: 1000,
      },
    )
    const unsubscribe = cache.subscribeViewport('Sheet1', { rowStart: 0, rowEnd: 600, colStart: 0, colEnd: 9 }, () => undefined)

    cache.applyViewportPatch(createLargePatch(601, 10))

    expect(countSheetCells(cache, 'Sheet1')).toBe(6010)

    unsubscribe()

    expect(countSheetCells(cache, 'Sheet1')).toBe(1000)
  })

  it('supports fractional cache caps in viewport-store options', () => {
    const cache = new ProjectedViewportStore(
      {
        invoke: async () => undefined,
        ready: async () => undefined,
        subscribe: () => () => undefined,
        subscribeBatches: () => () => undefined,
        subscribeViewportPatches: () => () => undefined,
        dispose: () => undefined,
      },
      {
        maxCachedCellsPerSheet: 15.4,
      },
    )
    const unsubscribe = cache.subscribeViewport('Sheet1', { rowStart: 0, rowEnd: 2, colStart: 0, colEnd: 9 }, () => undefined)

    cache.applyViewportPatch(createLargePatch(3, 10))

    expect(countSheetCells(cache, 'Sheet1')).toBe(30)

    unsubscribe()

    expect(countSheetCells(cache, 'Sheet1')).toBe(15)
  })

  it('falls back to the default cache cap for non-finite values', () => {
    const cache = new ProjectedViewportStore(
      {
        invoke: async () => undefined,
        ready: async () => undefined,
        subscribe: () => () => undefined,
        subscribeBatches: () => () => undefined,
        subscribeViewportPatches: () => () => undefined,
        dispose: () => undefined,
      },
      {
        maxCachedCellsPerSheet: Number.NaN,
      },
    )
    const unsubscribe = cache.subscribeViewport(
      'Sheet1',
      {
        rowStart: 0,
        rowEnd: DEFAULT_MAX_CACHED_CELLS_PER_SHEET,
        colStart: 0,
        colEnd: 0,
      },
      () => undefined,
    )

    cache.applyViewportPatch(createLargePatch(DEFAULT_MAX_CACHED_CELLS_PER_SHEET + 1, 1))

    expect(countSheetCells(cache, 'Sheet1')).toBe(DEFAULT_MAX_CACHED_CELLS_PER_SHEET + 1)

    unsubscribe()

    expect(countSheetCells(cache, 'Sheet1')).toBe(DEFAULT_MAX_CACHED_CELLS_PER_SHEET)
  })

  it('keeps pinned cell subscriptions while pruning after viewport teardown', () => {
    const cache = new ProjectedViewportStore({
      invoke: async () => undefined,
      ready: async () => undefined,
      subscribe: () => () => undefined,
      subscribeBatches: () => () => undefined,
      subscribeViewportPatches: () => () => undefined,
      dispose: () => undefined,
    })

    const unsubscribeViewport = cache.subscribeViewport('Sheet1', { rowStart: 0, rowEnd: 600, colStart: 0, colEnd: 9 }, () => undefined)
    const unsubscribeCell = cache.subscribeCells('Sheet1', ['A1'], () => undefined)

    cache.applyViewportPatch(createLargePatch(601, 10))
    unsubscribeViewport()

    expect(cache.peekCell('Sheet1', 'A1')).toBeDefined()
    expect(countSheetCells(cache, 'Sheet1')).toBe(DEFAULT_MAX_CACHED_CELLS_PER_SHEET)

    unsubscribeCell()
  })

  it('allows auxiliary selection viewports to skip duplicate initial hydration', () => {
    const typedPatch: ViewportPatch = {
      ...createPatch(),
      full: true,
      cells: [
        {
          row: 4,
          col: 3,
          snapshot: {
            sheetName: 'Sheet1',
            address: 'D5',
            value: { tag: ValueTag.String, value: 'future selection update', stringId: 1 },
            input: 'future selection update',
            flags: 0,
            version: 2,
          },
          displayText: 'future selection update',
          copyText: 'future selection update',
          editorText: 'future selection update',
          formatId: 0,
          styleId: 'style-0',
        },
      ],
    }
    const encodedPatch = encodeViewportPatch(typedPatch)
    let publishPatch: ((bytes: Uint8Array) => void) | null = null
    const subscribeViewportPatches = vi.fn((subscription, listener: (bytes: Uint8Array) => void) => {
      expect(subscription).toMatchObject({ initialPatch: 'none' })
      publishPatch = listener
      return () => undefined
    })
    const cache = new ProjectedViewportStore({
      invoke: vi.fn(),
      ready: async () => undefined,
      subscribe: () => () => undefined,
      subscribeBatches: () => () => undefined,
      subscribeViewportPatches,
      dispose: () => undefined,
    })
    const cellListener = vi.fn()

    const unsubscribeCell = cache.subscribeCell('Sheet1', 'D5', cellListener)
    const unsubscribeViewport = cache.subscribeAuxiliaryViewport(
      'Sheet1',
      { rowStart: 4, rowEnd: 4, colStart: 3, colEnd: 3 },
      () => undefined,
      { initialPatch: 'none' },
    )

    expect(cache.peekCell('Sheet1', 'D5')).toBeUndefined()
    expect(cellListener).not.toHaveBeenCalled()

    publishPatch?.(encodedPatch)

    expect(cache.getCell('Sheet1', 'D5').input).toBe('future selection update')
    expect(cellListener).toHaveBeenCalledTimes(1)

    unsubscribeViewport()
    unsubscribeCell()
  })

  it('notifies selected-cell listeners only for the addressed cell', () => {
    const cache = new ProjectedViewportStore()
    const selectedCellListener = vi.fn()
    const unrelatedCellListener = vi.fn()

    const unsubscribeSelected = cache.subscribeCell('Sheet1', 'A1', selectedCellListener)
    const unsubscribeUnrelated = cache.subscribeCell('Sheet1', 'B2', unrelatedCellListener)

    cache.setCellSnapshot({
      sheetName: 'Sheet1',
      address: 'A1',
      value: { tag: ValueTag.Number, value: 42 },
      flags: 0,
      version: 1,
    })

    expect(selectedCellListener).toHaveBeenCalledTimes(1)
    expect(unrelatedCellListener).not.toHaveBeenCalled()

    unsubscribeSelected()
    unsubscribeUnrelated()
  })

  it('notifies only the subscribed sheet-axis channels for viewport axis patches', () => {
    const cache = new ProjectedViewportStore()
    const freezeListener = vi.fn()
    const columnsListener = vi.fn()
    const rowsListener = vi.fn()

    const unsubscribeFreeze = cache.subscribeSheetChannel('Sheet1', 'freeze', freezeListener)
    const unsubscribeColumns = cache.subscribeSheetChannel('Sheet1', 'columnWidths', columnsListener)
    const unsubscribeRows = cache.subscribeSheetChannel('Sheet1', 'rowHeights', rowsListener)

    cache.applyViewportPatch({
      viewport: { sheetName: 'Sheet1', rowStart: 0, rowEnd: 2, colStart: 0, colEnd: 2 },
      full: false,
      cells: [],
      styles: [],
      columns: [{ index: 1, size: 140, hidden: false }],
      rows: [],
      freezeRows: 1,
      freezeCols: 1,
    })

    expect(freezeListener).toHaveBeenCalledTimes(1)
    expect(columnsListener).toHaveBeenCalledTimes(1)
    expect(rowsListener).not.toHaveBeenCalled()

    unsubscribeFreeze()
    unsubscribeColumns()
    unsubscribeRows()
  })

  it('exposes viewport merge metadata through the grid engine interface', () => {
    const cache = new ProjectedViewportStore()
    const mergeListener = vi.fn()
    const cellListener = vi.fn()
    const unsubscribeMerge = cache.subscribeSheetChannel('Sheet1', 'merges', mergeListener)
    const unsubscribeCell = cache.subscribeCells('Sheet1', ['C5'], cellListener)

    const damage = cache.applyViewportPatch({
      ...createPatch(),
      cells: [],
      viewport: { sheetName: 'Sheet1', rowStart: 4, rowEnd: 4, colStart: 2, colEnd: 3 },
      merges: [{ sheetName: 'Sheet1', startAddress: 'C5', endAddress: 'D5' }],
    })

    expect(damage).toEqual([{ cell: [2, 4] }, { cell: [3, 4] }])
    expect(cache.getMergeRange('Sheet1', 'C5')).toEqual({
      sheetName: 'Sheet1',
      startAddress: 'C5',
      endAddress: 'D5',
    })
    expect(cache.getMergeRange('Sheet1', 'D5')).toEqual({
      sheetName: 'Sheet1',
      startAddress: 'C5',
      endAddress: 'D5',
    })
    expect(cache.listMergeRanges('Sheet1')).toEqual([
      {
        sheetName: 'Sheet1',
        startAddress: 'C5',
        endAddress: 'D5',
      },
    ])
    expect(mergeListener).toHaveBeenCalledTimes(1)
    expect(cellListener).toHaveBeenCalledTimes(1)

    unsubscribeMerge()
    unsubscribeCell()
  })
})
