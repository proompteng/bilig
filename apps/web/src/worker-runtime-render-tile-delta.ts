import type { GridEngineLike } from '@bilig/grid'
import {
  VIEWPORT_TILE_COLUMN_COUNT,
  VIEWPORT_TILE_ROW_COUNT,
  type EngineEvent,
  MAX_COLS,
  MAX_ROWS,
  type RecalcMetrics,
  type Viewport,
  type WorkbookAxisEntrySnapshot,
  type WorkbookFreezePaneSnapshot,
} from '@bilig/protocol'
import { parseCellAddress } from '@bilig/formula'
import type { RenderTileDeltaBatch, RenderTileDeltaSubscription, RenderTileReplaceMutation } from '@bilig/worker-transport'
import { getGridMetrics } from '../../../packages/grid/src/gridMetrics.js'
import { materializeGridRenderTileV3 } from '../../../packages/grid/src/renderer-v3/grid-tile-materializer.js'
import { resolveGridRenderTileDirtySpansV3 } from '../../../packages/grid/src/renderer-v3/render-tile-dirty-spans.js'
import type { GridRenderTile } from '../../../packages/grid/src/renderer-v3/render-tile-source.js'
import { DirtyMaskV3 } from '../../../packages/grid/src/renderer-v3/tile-damage-index.js'
import { packTileKey53, tileKeysForViewport, unpackTileKey53, type TileKey53 } from '../../../packages/grid/src/renderer-v3/tile-key.js'
import type { TextOverflowIndexV3, TextOverflowSourceV3 } from '../../../packages/grid/src/renderer-v3/text-overflow-index.js'
import { buildFreezeVersion, buildRenderedAxisState } from './worker-runtime-render-axis.js'
import { listViewportTileBounds } from './worker-viewport-tile-store.js'

interface RenderTileDeltaEngineLike extends GridEngineLike {
  workbook: GridEngineLike['workbook'] & {
    readonly cellStore?:
      | {
          readonly sheetIds: Uint16Array
          readonly rows: Uint32Array
          readonly cols: Uint16Array
        }
      | undefined
    getSheetNameById?: ((sheetId: number) => string) | undefined
  }
  getColumnAxisEntries(sheetName: string): readonly WorkbookAxisEntrySnapshot[]
  getRowAxisEntries(sheetName: string): readonly WorkbookAxisEntrySnapshot[]
  getFreezePane?(sheetName: string): WorkbookFreezePaneSnapshot | undefined
  getLastMetrics(): Pick<RecalcMetrics, 'batchId'>
}

const CHANGED_CELL_DIRTY_MASK = DirtyMaskV3.Value | DirtyMaskV3.Text
const INVALIDATED_RANGE_DIRTY_MASK = DirtyMaskV3.Value | DirtyMaskV3.Style | DirtyMaskV3.Text | DirtyMaskV3.Rect | DirtyMaskV3.Border
const AXIS_X_DIRTY_MASK = DirtyMaskV3.AxisX | DirtyMaskV3.Text | DirtyMaskV3.Rect
const AXIS_Y_DIRTY_MASK = DirtyMaskV3.AxisY | DirtyMaskV3.Text | DirtyMaskV3.Rect

interface MaterializedTileViewport {
  readonly viewport: Viewport
  readonly dirtyLocalRows?: Uint32Array | undefined
  readonly dirtyLocalCols?: Uint32Array | undefined
  readonly dirtyMasks?: Uint32Array | undefined
}

interface DirtyLocalSpan {
  readonly rowStart: number
  readonly rowEnd: number
  readonly colStart: number
  readonly colEnd: number
  readonly mask: number
}

export function buildWorkerRenderTileDeltaBatch(input: {
  engine: RenderTileDeltaEngineLike
  subscription: RenderTileDeltaSubscription
  generation: number
  event?: EngineEvent | undefined
  textOverflowIndex?: TextOverflowIndexV3 | undefined
}): RenderTileDeltaBatch {
  const { engine, event, subscription, generation } = input
  const metrics = engine.getLastMetrics()
  const batchId = Number.isInteger(metrics.batchId) && metrics.batchId >= 0 ? metrics.batchId : 0
  const gridMetrics = getGridMetrics()
  const columnAxis = buildRenderedAxisState(engine.getColumnAxisEntries(subscription.sheetName), gridMetrics.columnWidth)
  const rowAxis = buildRenderedAxisState(engine.getRowAxisEntries(subscription.sheetName), gridMetrics.rowHeight)
  const freezePane = engine.getFreezePane?.(subscription.sheetName)
  const freezeVersion = buildFreezeVersion(freezePane?.rows ?? 0, freezePane?.cols ?? 0)
  const materializedViewports = resolveMaterializedTileViewports({
    batchId,
    engine,
    event,
    subscription,
    textOverflowIndex: input.textOverflowIndex,
  })
  const materializedTiles = materializedViewports.map((viewport) =>
    materializeGridRenderTileV3({
      axisSeqX: columnAxis.version,
      axisSeqY: rowAxis.version,
      cameraSeq: subscription.cameraSeq ?? batchId,
      columnWidths: columnAxis.sizes,
      dprBucket: subscription.dprBucket ?? 1,
      engine,
      dirtyLocalCols: viewport.dirtyLocalCols,
      dirtyLocalRows: viewport.dirtyLocalRows,
      dirtyMasks: viewport.dirtyMasks,
      freezeSeq: freezeVersion,
      glyphAtlasSeq: 0,
      gridMetrics,
      materializedAtSeq: generation,
      packetSeq: batchId,
      rectSeq: batchId,
      rowHeights: rowAxis.sizes,
      sheetId: subscription.sheetId,
      sheetOrdinal: resolveSubscriptionSheetOrdinal(subscription),
      sheetName: subscription.sheetName,
      sortedColumnWidthOverrides: columnAxis.sortedOverrides,
      sortedRowHeightOverrides: rowAxis.sortedOverrides,
      styleSeq: batchId,
      textSeq: batchId,
      valueSeq: batchId,
      viewport: viewport.viewport,
    }),
  )
  materializedTiles.forEach((tile) => {
    input.textOverflowIndex?.replaceTileRuns({
      sheetOrdinal: tile.coord.sheetOrdinal,
      textRuns: tile.textRuns,
      viewport: tile.bounds,
    })
  })

  return {
    magic: 'bilig.render.tile.delta',
    version: 3,
    sheetId: subscription.sheetId,
    sheetOrdinal: resolveSubscriptionSheetOrdinal(subscription),
    batchId,
    cameraSeq: subscription.cameraSeq ?? 0,
    mutations: materializedTiles.map(buildRenderTileReplaceMutation),
  }
}

function resolveMaterializedTileViewports(input: {
  readonly engine: RenderTileDeltaEngineLike
  readonly subscription: RenderTileDeltaSubscription
  readonly event: EngineEvent | undefined
  readonly batchId: number
  readonly textOverflowIndex?: TextOverflowIndexV3 | undefined
}): readonly MaterializedTileViewport[] {
  const { engine, event, subscription } = input
  if (!event || event.invalidation === 'full') {
    return resolveInterestedTileViewports(subscription).map((viewport) => ({ viewport }))
  }

  const dprBucket = subscription.dprBucket ?? 1
  const dirtySpansByTile = collectEventDirtyTileSpans({
    engine,
    event,
    subscription,
    dprBucket,
    sheetOrdinal: resolveSubscriptionSheetOrdinal(subscription),
    sheetName: subscription.sheetName,
    textOverflowIndex: input.textOverflowIndex,
  })

  if (dirtySpansByTile.size === 0) {
    return []
  }

  const interestedTileKeys = new Set(resolveInterestedTileKeys(subscription))
  return resolveInterestedTileViewports(subscription).flatMap((viewport) => {
    const tileKey = tileKeyFromTileViewport(subscription, viewport)
    if (!interestedTileKeys.has(tileKey)) {
      return []
    }
    const dirtySpans = dirtySpansByTile.get(tileKey)
    if (!dirtySpans) {
      return []
    }
    return [
      {
        viewport,
        ...packDirtyLocalSpans(dirtySpans),
      },
    ]
  })
}

function resolveInterestedTileKeys(subscription: RenderTileDeltaSubscription): readonly TileKey53[] {
  const dprBucket = subscription.dprBucket ?? 1
  const explicitInterest = subscription.tileInterest
  const visibleKeys =
    explicitInterest && explicitInterest.visibleTileKeys.length > 0
      ? explicitInterest.visibleTileKeys.filter((key) => isSubscriptionTileKey(subscription, key))
      : tileKeysForViewport({
          dprBucket,
          sheetOrdinal: resolveSubscriptionSheetOrdinal(subscription),
          viewport: subscription,
        })
  const keys = new Set<number>(visibleKeys)
  explicitInterest?.warmTileKeys.forEach((key) => {
    if (!isSubscriptionTileKey(subscription, key)) {
      return
    }
    keys.add(key)
  })
  explicitInterest?.pinnedTileKeys.forEach((key) => {
    if (!isSubscriptionTileKey(subscription, key)) {
      return
    }
    keys.add(key)
  })
  subscription.warmTileKeys?.forEach((key) => {
    if (!isSubscriptionTileKey(subscription, key)) {
      return
    }
    keys.add(key)
  })
  return [...keys]
}

function resolveInterestedTileViewports(subscription: RenderTileDeltaSubscription): readonly Viewport[] {
  if (subscription.tileInterest?.visibleTileKeys.length) {
    return resolveInterestedTileKeys(subscription)
      .map((key) => tileViewportFromKey(subscription, key))
      .filter((viewport): viewport is Viewport => viewport !== null)
  }
  const viewportsByKey = new Map<number, Viewport>()
  for (const viewport of listViewportTileBounds(subscription)) {
    viewportsByKey.set(tileKeyFromTileViewport(subscription, viewport), viewport)
  }
  subscription.warmTileKeys?.forEach((key) => {
    const viewport = tileViewportFromKey(subscription, key)
    if (viewport) {
      viewportsByKey.set(key, viewport)
    }
  })
  return [...viewportsByKey.values()].toSorted((a, b) => a.rowStart - b.rowStart || a.colStart - b.colStart)
}

function tileViewportFromKey(subscription: RenderTileDeltaSubscription, key: TileKey53): Viewport | null {
  if (!isSubscriptionTileKey(subscription, key)) {
    return null
  }
  const fields = unpackTileKey53(key)
  const rowStart = fields.rowTile * VIEWPORT_TILE_ROW_COUNT
  const colStart = fields.colTile * VIEWPORT_TILE_COLUMN_COUNT
  return {
    colEnd: Math.min(colStart + VIEWPORT_TILE_COLUMN_COUNT - 1, MAX_COLS - 1),
    colStart,
    rowEnd: Math.min(rowStart + VIEWPORT_TILE_ROW_COUNT - 1, MAX_ROWS - 1),
    rowStart,
  }
}

function isSubscriptionTileKey(subscription: RenderTileDeltaSubscription, key: TileKey53): boolean {
  const fields = unpackTileKey53(key)
  return fields.sheetOrdinal === resolveSubscriptionSheetOrdinal(subscription) && fields.dprBucket === (subscription.dprBucket ?? 1)
}

function tileKeyFromTileViewport(subscription: RenderTileDeltaSubscription, viewport: Viewport): TileKey53 {
  return packTileKey53({
    colTile: Math.floor(viewport.colStart / VIEWPORT_TILE_COLUMN_COUNT),
    dprBucket: subscription.dprBucket ?? 1,
    rowTile: Math.floor(viewport.rowStart / VIEWPORT_TILE_ROW_COUNT),
    sheetOrdinal: resolveSubscriptionSheetOrdinal(subscription),
  })
}

function resolveSubscriptionSheetOrdinal(subscription: RenderTileDeltaSubscription): number {
  return subscription.sheetOrdinal ?? subscription.tileInterest?.sheetOrdinal ?? subscription.sheetId
}

function collectEventDirtyTileSpans(input: {
  readonly engine: RenderTileDeltaEngineLike
  readonly event: EngineEvent
  readonly subscription: RenderTileDeltaSubscription
  readonly sheetName: string
  readonly sheetOrdinal: number
  readonly dprBucket: number
  readonly textOverflowIndex?: TextOverflowIndexV3 | undefined
}): Map<TileKey53, DirtyLocalSpan[]> {
  const { dprBucket, engine, event, sheetName, sheetOrdinal, subscription, textOverflowIndex } = input
  const spansByTile = new Map<TileKey53, DirtyLocalSpan[]>()
  const interestedTileKeys = resolveInterestedTileKeys(subscription)
  const cellStore = engine.workbook.cellStore
  const getSheetNameById = engine.workbook.getSheetNameById
  if (cellStore && getSheetNameById) {
    for (let index = 0; index < event.changedCellIndices.length; index += 1) {
      const cellIndex = event.changedCellIndices[index]!
      const sheetId = cellStore.sheetIds[cellIndex]
      if (sheetId === undefined || getSheetNameById(sheetId) !== sheetName) {
        continue
      }
      const row = cellStore.rows[cellIndex] ?? 0
      const col = cellStore.cols[cellIndex] ?? 0
      markDirtyRangeSpans(spansByTile, {
        colEnd: col,
        colStart: col,
        candidateTileKeys: interestedTileKeys,
        dprBucket,
        mask: CHANGED_CELL_DIRTY_MASK,
        rowEnd: row,
        rowStart: row,
        sheetOrdinal,
      })
      markTextOverflowDependencySpans(spansByTile, {
        dprBucket,
        sheetOrdinal,
        sourceCallback: (callback) => textOverflowIndex?.markDependenciesForCell({ col, row, sheetOrdinal }, callback),
      })
    }
  }

  event.invalidatedRanges.forEach((range) => {
    if (range.sheetName !== sheetName) {
      return
    }
    const start = parseCellAddress(range.startAddress, range.sheetName)
    const end = parseCellAddress(range.endAddress, range.sheetName)
    const colStart = Math.min(start.col, end.col)
    const colEnd = Math.max(start.col, end.col)
    const rowStart = Math.min(start.row, end.row)
    const rowEnd = Math.max(start.row, end.row)
    markDirtyRangeSpans(spansByTile, {
      colEnd,
      colStart,
      candidateTileKeys: interestedTileKeys,
      dprBucket,
      mask: INVALIDATED_RANGE_DIRTY_MASK,
      rowEnd,
      rowStart,
      sheetOrdinal,
    })
    markTextOverflowDependencySpans(spansByTile, {
      dprBucket,
      sheetOrdinal,
      sourceCallback: (callback) =>
        textOverflowIndex?.markDependenciesForCellRange({ colEnd, colStart, rowEnd, rowStart, sheetOrdinal }, callback),
    })
  })

  event.invalidatedColumns.forEach((column) => {
    if (column.sheetName !== sheetName) {
      return
    }
    markDirtyAxisXSpans(spansByTile, {
      colEnd: column.endIndex,
      colStart: column.startIndex,
      dprBucket,
      mask: AXIS_X_DIRTY_MASK,
      sheetOrdinal,
      subscription,
    })
    markTextOverflowDependencySpans(spansByTile, {
      dprBucket,
      sheetOrdinal,
      sourceCallback: (callback) =>
        textOverflowIndex?.markDependenciesForAxisX({ colEnd: column.endIndex, colStart: column.startIndex, sheetOrdinal }, callback),
    })
  })

  event.invalidatedRows.forEach((row) => {
    if (row.sheetName !== sheetName) {
      return
    }
    markDirtyAxisYSpans(spansByTile, {
      dprBucket,
      mask: AXIS_Y_DIRTY_MASK,
      rowEnd: row.endIndex,
      rowStart: row.startIndex,
      sheetOrdinal,
      subscription,
    })
  })
  return spansByTile
}

function markTextOverflowDependencySpans(
  spansByTile: Map<TileKey53, DirtyLocalSpan[]>,
  input: {
    readonly sheetOrdinal: number
    readonly dprBucket: number
    readonly sourceCallback: (callback: (source: TextOverflowSourceV3) => void) => void
  },
): void {
  input.sourceCallback((source) => {
    markDirtyRangeSpans(spansByTile, {
      colEnd: source.spillColEnd,
      colStart: source.col,
      dprBucket: input.dprBucket,
      mask: DirtyMaskV3.Text,
      rowEnd: source.row,
      rowStart: source.row,
      sheetOrdinal: input.sheetOrdinal,
    })
  })
}

function markDirtyAxisXSpans(
  spansByTile: Map<TileKey53, DirtyLocalSpan[]>,
  input: {
    readonly subscription: RenderTileDeltaSubscription
    readonly sheetOrdinal: number
    readonly dprBucket: number
    readonly colStart: number
    readonly colEnd: number
    readonly mask: number
  },
): void {
  const colStart = Math.max(0, Math.min(MAX_COLS - 1, input.colStart))
  const colEnd = Math.max(colStart, Math.min(MAX_COLS - 1, input.colEnd))
  resolveInterestedTileKeys(input.subscription).forEach((tileKey) => {
    const fields = unpackTileKey53(tileKey)
    const tileColStart = fields.colTile * VIEWPORT_TILE_COLUMN_COUNT
    const tileColEnd = Math.min(MAX_COLS - 1, tileColStart + VIEWPORT_TILE_COLUMN_COUNT - 1)
    if (tileColEnd < colStart || tileColStart > colEnd) {
      return
    }
    const localColStart = Math.max(colStart, tileColStart) - tileColStart
    const spans = spansByTile.get(tileKey) ?? []
    spans.push({
      colEnd: tileColEnd - tileColStart,
      colStart: localColStart,
      mask: input.mask,
      rowEnd: VIEWPORT_TILE_ROW_COUNT - 1,
      rowStart: 0,
    })
    spansByTile.set(tileKey, spans)
  })
}

function markDirtyAxisYSpans(
  spansByTile: Map<TileKey53, DirtyLocalSpan[]>,
  input: {
    readonly subscription: RenderTileDeltaSubscription
    readonly sheetOrdinal: number
    readonly dprBucket: number
    readonly rowStart: number
    readonly rowEnd: number
    readonly mask: number
  },
): void {
  const rowStart = Math.max(0, Math.min(MAX_ROWS - 1, input.rowStart))
  const rowEnd = Math.max(rowStart, Math.min(MAX_ROWS - 1, input.rowEnd))
  resolveInterestedTileKeys(input.subscription).forEach((tileKey) => {
    const fields = unpackTileKey53(tileKey)
    const tileRowStart = fields.rowTile * VIEWPORT_TILE_ROW_COUNT
    const tileRowEnd = Math.min(MAX_ROWS - 1, tileRowStart + VIEWPORT_TILE_ROW_COUNT - 1)
    if (tileRowEnd < rowStart || tileRowStart > rowEnd) {
      return
    }
    const localRowStart = Math.max(rowStart, tileRowStart) - tileRowStart
    const spans = spansByTile.get(tileKey) ?? []
    spans.push({
      colEnd: VIEWPORT_TILE_COLUMN_COUNT - 1,
      colStart: 0,
      mask: input.mask,
      rowEnd: tileRowEnd - tileRowStart,
      rowStart: localRowStart,
    })
    spansByTile.set(tileKey, spans)
  })
}

function markDirtyRangeSpans(
  spansByTile: Map<TileKey53, DirtyLocalSpan[]>,
  input: {
    readonly sheetOrdinal: number
    readonly dprBucket: number
    readonly rowStart: number
    readonly rowEnd: number
    readonly colStart: number
    readonly colEnd: number
    readonly mask: number
    readonly candidateTileKeys?: readonly TileKey53[] | undefined
  },
): void {
  const rowStart = Math.max(0, Math.min(MAX_ROWS - 1, input.rowStart))
  const rowEnd = Math.max(rowStart, Math.min(MAX_ROWS - 1, input.rowEnd))
  const colStart = Math.max(0, Math.min(MAX_COLS - 1, input.colStart))
  const colEnd = Math.max(colStart, Math.min(MAX_COLS - 1, input.colEnd))
  const rowTileStart = Math.floor(rowStart / VIEWPORT_TILE_ROW_COUNT)
  const rowTileEnd = Math.floor(rowEnd / VIEWPORT_TILE_ROW_COUNT)
  const colTileStart = Math.floor(colStart / VIEWPORT_TILE_COLUMN_COUNT)
  const colTileEnd = Math.floor(colEnd / VIEWPORT_TILE_COLUMN_COUNT)
  const tileArea = (rowTileEnd - rowTileStart + 1) * (colTileEnd - colTileStart + 1)
  if (input.candidateTileKeys && tileArea > input.candidateTileKeys.length) {
    input.candidateTileKeys.forEach((tileKey) => {
      const fields = unpackTileKey53(tileKey)
      if (fields.sheetOrdinal !== input.sheetOrdinal || fields.dprBucket !== input.dprBucket) {
        return
      }
      appendDirtyRangeTileSpan(spansByTile, {
        colEnd,
        colStart,
        colTile: fields.colTile,
        dprBucket: input.dprBucket,
        mask: input.mask,
        rowEnd,
        rowStart,
        rowTile: fields.rowTile,
        sheetOrdinal: input.sheetOrdinal,
      })
    })
    return
  }
  for (let rowTile = rowTileStart; rowTile <= rowTileEnd; rowTile += 1) {
    for (let colTile = colTileStart; colTile <= colTileEnd; colTile += 1) {
      appendDirtyRangeTileSpan(spansByTile, {
        colEnd,
        colStart,
        colTile,
        dprBucket: input.dprBucket,
        mask: input.mask,
        rowEnd,
        rowStart,
        rowTile,
        sheetOrdinal: input.sheetOrdinal,
      })
    }
  }
}

function appendDirtyRangeTileSpan(
  spansByTile: Map<TileKey53, DirtyLocalSpan[]>,
  input: {
    readonly sheetOrdinal: number
    readonly dprBucket: number
    readonly rowStart: number
    readonly rowEnd: number
    readonly colStart: number
    readonly colEnd: number
    readonly rowTile: number
    readonly colTile: number
    readonly mask: number
  },
): void {
  const tileRowStart = input.rowTile * VIEWPORT_TILE_ROW_COUNT
  const tileRowEnd = Math.min(MAX_ROWS - 1, tileRowStart + VIEWPORT_TILE_ROW_COUNT - 1)
  const tileColStart = input.colTile * VIEWPORT_TILE_COLUMN_COUNT
  const tileColEnd = Math.min(MAX_COLS - 1, tileColStart + VIEWPORT_TILE_COLUMN_COUNT - 1)
  if (tileRowEnd < input.rowStart || tileRowStart > input.rowEnd || tileColEnd < input.colStart || tileColStart > input.colEnd) {
    return
  }
  const tileKey = packTileKey53({
    colTile: input.colTile,
    dprBucket: input.dprBucket,
    rowTile: input.rowTile,
    sheetOrdinal: input.sheetOrdinal,
  })
  const spans = spansByTile.get(tileKey) ?? []
  spans.push({
    colEnd: Math.min(input.colEnd, tileColEnd) - tileColStart,
    colStart: Math.max(input.colStart, tileColStart) - tileColStart,
    mask: input.mask,
    rowEnd: Math.min(input.rowEnd, tileRowEnd) - tileRowStart,
    rowStart: Math.max(input.rowStart, tileRowStart) - tileRowStart,
  })
  spansByTile.set(tileKey, spans)
}

function packDirtyLocalSpans(spans: readonly DirtyLocalSpan[]): {
  readonly dirtyLocalRows: Uint32Array
  readonly dirtyLocalCols: Uint32Array
  readonly dirtyMasks: Uint32Array
} {
  const dirtyLocalRows = new Uint32Array(spans.length * 2)
  const dirtyLocalCols = new Uint32Array(spans.length * 2)
  const dirtyMasks = new Uint32Array(spans.length)
  spans.forEach((span, index) => {
    const rowOffset = index * 2
    dirtyLocalRows[rowOffset] = span.rowStart
    dirtyLocalRows[rowOffset + 1] = span.rowEnd
    dirtyLocalCols[rowOffset] = span.colStart
    dirtyLocalCols[rowOffset + 1] = span.colEnd
    dirtyMasks[index] = span.mask
  })
  return { dirtyLocalCols, dirtyLocalRows, dirtyMasks }
}

function buildRenderTileReplaceMutation(tile: GridRenderTile): RenderTileReplaceMutation {
  return {
    kind: 'tileReplace',
    tileId: tile.tileId,
    coord: tile.coord,
    version: tile.version,
    bounds: tile.bounds,
    rectInstances: tile.rectInstances,
    rectCount: tile.rectCount,
    textMetrics: tile.textMetrics,
    glyphRefs: new Uint32Array(0),
    textRuns: tile.textRuns.map((run) => ({
      align: run.align,
      col: run.col,
      clipHeight: run.clipHeight,
      clipWidth: run.clipWidth,
      clipX: run.clipX,
      clipY: run.clipY,
      color: run.color,
      font: run.font,
      fontSize: run.fontSize,
      height: run.height,
      row: run.row,
      strike: run.strike,
      spillColEnd: run.spillColEnd,
      text: run.text,
      underline: run.underline,
      width: run.width,
      wrap: run.wrap,
      x: run.x,
      y: run.y,
    })),
    textCount: tile.textCount,
    dirty: resolveGridRenderTileDirtySpansV3(tile),
    dirtyLocalCols: tile.dirtyLocalCols,
    dirtyLocalRows: tile.dirtyLocalRows,
    dirtyMasks: tile.dirtyMasks,
  }
}
