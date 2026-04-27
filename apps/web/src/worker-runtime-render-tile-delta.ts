import type { GridEngineLike } from '@bilig/grid'
import {
  VIEWPORT_TILE_COLUMN_COUNT,
  VIEWPORT_TILE_ROW_COUNT,
  type EngineEvent,
  type RecalcMetrics,
  type Viewport,
  type WorkbookAxisEntrySnapshot,
} from '@bilig/protocol'
import { parseCellAddress } from '@bilig/formula'
import type { RenderTileDeltaBatch, RenderTileDeltaSubscription, RenderTileReplaceMutation } from '@bilig/worker-transport'
import { getGridMetrics } from '../../../packages/grid/src/gridMetrics.js'
import { materializeGridRenderTileV3 } from '../../../packages/grid/src/renderer-v3/grid-tile-materializer.js'
import type { GridRenderTile } from '../../../packages/grid/src/renderer-v3/render-tile-source.js'
import { DirtyMaskV3, DirtyTileIndexV3 } from '../../../packages/grid/src/renderer-v3/tile-damage-index.js'
import { packTileKey53, tileKeysForViewport } from '../../../packages/grid/src/renderer-v3/tile-key.js'
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
  getLastMetrics(): Pick<RecalcMetrics, 'batchId'>
}

const FULL_SPAN_START = 0
const CHANGED_CELL_DIRTY_MASK = DirtyMaskV3.Value | DirtyMaskV3.Text | DirtyMaskV3.Rect
const INVALIDATED_RANGE_DIRTY_MASK = DirtyMaskV3.Value | DirtyMaskV3.Style | DirtyMaskV3.Text | DirtyMaskV3.Rect | DirtyMaskV3.Border
const AXIS_X_DIRTY_MASK = DirtyMaskV3.AxisX | DirtyMaskV3.Text | DirtyMaskV3.Rect
const AXIS_Y_DIRTY_MASK = DirtyMaskV3.AxisY | DirtyMaskV3.Text | DirtyMaskV3.Rect

export function buildWorkerRenderTileDeltaBatch(input: {
  engine: RenderTileDeltaEngineLike
  subscription: RenderTileDeltaSubscription
  generation: number
  event?: EngineEvent | undefined
}): RenderTileDeltaBatch {
  const { engine, event, subscription, generation } = input
  const metrics = engine.getLastMetrics()
  const batchId = Number.isInteger(metrics.batchId) && metrics.batchId >= 0 ? metrics.batchId : 0
  const gridMetrics = getGridMetrics()
  const columnAxis = buildRenderedAxisState(engine.getColumnAxisEntries(subscription.sheetName), gridMetrics.columnWidth)
  const rowAxis = buildRenderedAxisState(engine.getRowAxisEntries(subscription.sheetName), gridMetrics.rowHeight)
  const freezeVersion = buildFreezeVersion(0, 0)
  const materializedViewports = resolveMaterializedTileViewports({
    batchId,
    engine,
    event,
    subscription,
  })

  return {
    magic: 'bilig.render.tile.delta',
    version: 1,
    sheetId: subscription.sheetId,
    batchId,
    cameraSeq: subscription.cameraSeq ?? 0,
    mutations: materializedViewports.map((viewport) =>
      buildRenderTileReplaceMutation(
        materializeGridRenderTileV3({
          axisSeqX: columnAxis.version,
          axisSeqY: rowAxis.version,
          cameraSeq: subscription.cameraSeq ?? batchId,
          columnWidths: columnAxis.sizes,
          dprBucket: subscription.dprBucket ?? 1,
          engine,
          freezeSeq: freezeVersion,
          glyphAtlasSeq: 0,
          gridMetrics,
          materializedAtSeq: generation,
          packetSeq: batchId,
          rectSeq: batchId,
          rowHeights: rowAxis.sizes,
          sheetId: subscription.sheetId,
          sheetName: subscription.sheetName,
          sortedColumnWidthOverrides: columnAxis.sortedOverrides,
          sortedRowHeightOverrides: rowAxis.sortedOverrides,
          styleSeq: batchId,
          textSeq: batchId,
          valueSeq: batchId,
          viewport,
        }),
      ),
    ),
  }
}

function resolveMaterializedTileViewports(input: {
  readonly engine: RenderTileDeltaEngineLike
  readonly subscription: RenderTileDeltaSubscription
  readonly event: EngineEvent | undefined
  readonly batchId: number
}): Viewport[] {
  const { engine, event, subscription } = input
  if (!event || event.invalidation === 'full') {
    return listViewportTileBounds(subscription)
  }

  const dprBucket = subscription.dprBucket ?? 1
  const dirtyIndex = new DirtyTileIndexV3()
  markEventDirtyTiles({
    dirtyIndex,
    engine,
    event,
    dprBucket,
    sheetOrdinal: subscription.sheetId,
    sheetName: subscription.sheetName,
  })

  const dirtyTileKeys = new Set(
    dirtyIndex.consumeVisible(
      tileKeysForViewport({
        dprBucket,
        sheetOrdinal: subscription.sheetId,
        viewport: subscription,
      }),
    ),
  )
  if (dirtyTileKeys.size === 0) {
    return []
  }

  return listViewportTileBounds(subscription).filter((viewport) =>
    dirtyTileKeys.has(
      packTileKey53({
        colTile: Math.floor(viewport.colStart / VIEWPORT_TILE_COLUMN_COUNT),
        dprBucket,
        rowTile: Math.floor(viewport.rowStart / VIEWPORT_TILE_ROW_COUNT),
        sheetOrdinal: subscription.sheetId,
      }),
    ),
  )
}

function markEventDirtyTiles(input: {
  readonly dirtyIndex: DirtyTileIndexV3
  readonly engine: RenderTileDeltaEngineLike
  readonly event: EngineEvent
  readonly sheetName: string
  readonly sheetOrdinal: number
  readonly dprBucket: number
}): void {
  const { dirtyIndex, dprBucket, engine, event, sheetName, sheetOrdinal } = input
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
      dirtyIndex.markCellRange({
        colEnd: col,
        colStart: col,
        dprBucket,
        mask: CHANGED_CELL_DIRTY_MASK,
        rowEnd: row,
        rowStart: row,
        sheetOrdinal,
      })
    }
  }

  event.invalidatedRanges.forEach((range) => {
    if (range.sheetName !== sheetName) {
      return
    }
    const start = parseCellAddress(range.startAddress, range.sheetName)
    const end = parseCellAddress(range.endAddress, range.sheetName)
    dirtyIndex.markCellRange({
      colEnd: Math.max(start.col, end.col),
      colStart: Math.min(start.col, end.col),
      dprBucket,
      mask: INVALIDATED_RANGE_DIRTY_MASK,
      rowEnd: Math.max(start.row, end.row),
      rowStart: Math.min(start.row, end.row),
      sheetOrdinal,
    })
  })

  event.invalidatedColumns.forEach((column) => {
    if (column.sheetName !== sheetName) {
      return
    }
    dirtyIndex.markAxisX({
      colEnd: column.endIndex,
      colStart: column.startIndex,
      dprBucket,
      mask: AXIS_X_DIRTY_MASK,
      sheetOrdinal,
    })
  })

  event.invalidatedRows.forEach((row) => {
    if (row.sheetName !== sheetName) {
      return
    }
    dirtyIndex.markAxisY({
      dprBucket,
      mask: AXIS_Y_DIRTY_MASK,
      rowEnd: row.endIndex,
      rowStart: row.startIndex,
      sheetOrdinal,
    })
  })
}

export function buildRenderTileReplaceMutation(tile: GridRenderTile): RenderTileReplaceMutation {
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
      clipHeight: run.clipHeight,
      clipWidth: run.clipWidth,
      clipX: run.clipX,
      clipY: run.clipY,
      color: run.color,
      font: run.font,
      fontSize: run.fontSize,
      height: run.height,
      strike: run.strike,
      text: run.text,
      underline: run.underline,
      width: run.width,
      x: run.x,
      y: run.y,
    })),
    textCount: tile.textCount,
    dirty: {
      rectSpans: tile.rectCount > 0 ? [{ offset: FULL_SPAN_START, length: tile.rectCount }] : [],
      textSpans: tile.textCount > 0 ? [{ offset: FULL_SPAN_START, length: tile.textCount }] : [],
      glyphSpans: [],
    },
  }
}
