import type { GridEngineLike } from '@bilig/grid'
import type { RecalcMetrics, Viewport, WorkbookAxisEntrySnapshot } from '@bilig/protocol'
import type {
  RenderTileDeltaBatch,
  RenderTileDeltaSubscription,
  RenderTileReplaceMutation,
  RenderTileTextRun,
} from '@bilig/worker-transport'
import { buildGridGpuScene } from '../../../packages/grid/src/gridGpuScene.js'
import {
  getGridMetrics,
  getResolvedColumnWidth,
  getResolvedRowHeight,
  resolveRowOffset,
  type GridMetrics,
} from '../../../packages/grid/src/gridMetrics.js'
import { buildGridTextScene } from '../../../packages/grid/src/gridTextScene.js'
import type { GridSelection, Item, Rectangle } from '../../../packages/grid/src/gridTypes.js'
import { CompactSelection } from '../../../packages/grid/src/gridTypes.js'
import { collectViewportItems } from '../../../packages/grid/src/gridViewportItems.js'
import {
  createGridTileKeyV2,
  packGridScenePacketV2,
  type GridScenePacketV2,
  type GridSceneTextRun,
} from '../../../packages/grid/src/renderer-v2/scene-packet-v2.js'
import { validateGridScenePacketV2 } from '../../../packages/grid/src/renderer-v2/scene-packet-validator.js'
import { packTileKey53 } from '../../../packages/grid/src/renderer-v3/tile-key.js'
import { resolveColumnOffset } from '../../../packages/grid/src/workbookGridViewport.js'
import type { WorkbookPaneScenePacket } from './resident-pane-scene-types.js'
import { buildFreezeVersion, buildRenderedAxisState } from './worker-runtime-render-axis.js'
import { listViewportTileBounds } from './worker-viewport-tile-store.js'

interface RenderTileDeltaEngineLike extends GridEngineLike {
  getColumnAxisEntries(sheetName: string): readonly WorkbookAxisEntrySnapshot[]
  getRowAxisEntries(sheetName: string): readonly WorkbookAxisEntrySnapshot[]
  getLastMetrics(): Pick<RecalcMetrics, 'batchId'>
}

const FULL_SPAN_START = 0
const STATIC_TILE_SELECTED_CELL: Item = Object.freeze([-1, -1] as const)
const STATIC_TILE_GRID_SELECTION: GridSelection = Object.freeze({
  columns: CompactSelection.empty(),
  current: undefined,
  rows: CompactSelection.empty(),
})

export function buildWorkerRenderTileDeltaBatch(input: {
  engine: RenderTileDeltaEngineLike
  subscription: RenderTileDeltaSubscription
  generation: number
}): RenderTileDeltaBatch {
  const { engine, subscription, generation } = input
  const metrics = engine.getLastMetrics()
  const batchId = Number.isInteger(metrics.batchId) && metrics.batchId >= 0 ? metrics.batchId : 0
  const gridMetrics = getGridMetrics()
  const columnAxis = buildRenderedAxisState(engine.getColumnAxisEntries(subscription.sheetName), gridMetrics.columnWidth)
  const rowAxis = buildRenderedAxisState(engine.getRowAxisEntries(subscription.sheetName), gridMetrics.rowHeight)
  const freezeVersion = buildFreezeVersion(0, 0)
  const generatedAt = Date.now()

  return {
    magic: 'bilig.render.tile.delta',
    version: 1,
    sheetId: subscription.sheetId,
    batchId,
    cameraSeq: subscription.cameraSeq ?? 0,
    mutations: listViewportTileBounds(subscription).map((viewport) =>
      buildRenderTileReplaceMutation(
        subscription.sheetId,
        buildContentTileScenePacket({
          batchId,
          columnAxis,
          engine,
          freezeVersion,
          generatedAt,
          generation,
          gridMetrics,
          rowAxis,
          subscription,
          viewport,
        }),
      ),
    ),
  }
}

function buildContentTileScenePacket(input: {
  readonly batchId: number
  readonly columnAxis: ReturnType<typeof buildRenderedAxisState>
  readonly engine: RenderTileDeltaEngineLike
  readonly freezeVersion: number
  readonly generatedAt: number
  readonly generation: number
  readonly gridMetrics: GridMetrics
  readonly rowAxis: ReturnType<typeof buildRenderedAxisState>
  readonly subscription: RenderTileDeltaSubscription
  readonly viewport: Viewport
}): GridScenePacketV2 {
  const { batchId, columnAxis, engine, freezeVersion, generatedAt, generation, gridMetrics, rowAxis, subscription, viewport } = input
  const surfaceSize = resolveViewportSurfaceSize(viewport, {
    gridMetrics,
    sortedColumnWidthOverrides: columnAxis.sortedOverrides,
    sortedRowHeightOverrides: rowAxis.sortedOverrides,
  })
  const visibleItems = collectViewportItems(viewport)
  const visibleRegion = {
    range: {
      x: viewport.colStart,
      y: viewport.rowStart,
      width: viewport.colEnd - viewport.colStart + 1,
      height: viewport.rowEnd - viewport.rowStart + 1,
    },
    tx: 0,
    ty: 0,
    freezeRows: 0,
    freezeCols: 0,
  }
  const getCellBounds = createTileCellBoundsResolver({
    columnWidths: columnAxis.sizes,
    gridMetrics,
    rowHeights: rowAxis.sizes,
    sortedColumnWidthOverrides: columnAxis.sortedOverrides,
    sortedRowHeightOverrides: rowAxis.sortedOverrides,
    viewport,
  })
  const packet = packGridScenePacketV2({
    cameraSeq: subscription.cameraSeq ?? batchId,
    generatedAt,
    generation,
    gpuScene: buildGridGpuScene({
      activeHeaderDrag: null,
      columnWidths: columnAxis.sizes,
      contentMode: 'data',
      engine,
      getCellBounds,
      gridMetrics,
      gridSelection: STATIC_TILE_GRID_SELECTION,
      hostBounds: { left: 0, top: 0 },
      hoveredCell: null,
      hoveredHeader: null,
      resizeGuideColumn: null,
      resizeGuideRow: null,
      rowHeights: rowAxis.sizes,
      selectedCell: STATIC_TILE_SELECTED_CELL,
      selectionRange: null,
      sheetName: subscription.sheetName,
      visibleItems,
      visibleRegion,
    }),
    key: createGridTileKeyV2({
      axisVersionX: columnAxis.version,
      axisVersionY: rowAxis.version,
      dprBucket: subscription.dprBucket ?? 1,
      freezeVersion,
      paneId: 'body',
      selectionIndependentVersion: batchId,
      sheetName: subscription.sheetName,
      styleVersion: batchId,
      valueVersion: batchId,
      viewport,
    }),
    paneId: 'body',
    requestSeq: batchId,
    sheetName: subscription.sheetName,
    surfaceSize,
    textScene: buildGridTextScene({
      activeHeaderDrag: null,
      columnWidths: columnAxis.sizes,
      contentMode: 'data',
      editingCell: null,
      engine,
      getCellBounds,
      gridMetrics,
      hostBounds: {
        height: surfaceSize.height,
        left: 0,
        top: 0,
        width: surfaceSize.width,
      },
      hoveredHeader: null,
      resizeGuideColumn: null,
      rowHeights: rowAxis.sizes,
      selectedCell: STATIC_TILE_SELECTED_CELL,
      selectedCellSnapshot: null,
      selectionRange: null,
      sheetName: subscription.sheetName,
      visibleItems,
      visibleRegion,
    }),
    viewport,
  })
  const validation = validateGridScenePacketV2(packet)
  if (!validation.ok) {
    throw new Error(`Invalid worker render tile packet: ${validation.reason}`)
  }
  return packet
}

export function buildRenderTileDeltaBatchFromResidentPaneScenes(input: {
  sheetId: number
  batchId: number
  cameraSeq: number
  scenes: readonly WorkbookPaneScenePacket[]
}): RenderTileDeltaBatch {
  return {
    magic: 'bilig.render.tile.delta',
    version: 1,
    sheetId: input.sheetId,
    batchId: input.batchId,
    cameraSeq: input.cameraSeq,
    mutations: input.scenes.map((scene) => buildRenderTileReplaceMutation(input.sheetId, scene.packedScene)),
  }
}

export function buildRenderTileReplaceMutation(sheetId: number, packet: GridScenePacketV2): RenderTileReplaceMutation {
  return {
    kind: 'tileReplace',
    tileId: packTileKey53({
      sheetOrdinal: sheetId,
      rowTile: packet.key.rowTile,
      colTile: packet.key.colTile,
      dprBucket: packet.key.dprBucket,
    }),
    coord: {
      sheetId,
      paneKind: packet.key.paneKind,
      rowTile: packet.key.rowTile,
      colTile: packet.key.colTile,
      dprBucket: packet.key.dprBucket,
    },
    version: {
      axisX: packet.key.axisVersionX,
      axisY: packet.key.axisVersionY,
      values: packet.key.valueVersion,
      styles: packet.key.styleVersion,
      text: packet.key.textEpoch,
      freeze: packet.key.freezeVersion,
    },
    bounds: packet.viewport,
    rectInstances: packet.rectInstances,
    rectCount: packet.rectCount,
    textMetrics: packet.textMetrics,
    glyphRefs: new Uint32Array(0),
    textRuns: packet.textRuns.map(mapTextRun),
    textCount: packet.textCount,
    dirty: {
      rectSpans: packet.rectCount > 0 ? [{ offset: FULL_SPAN_START, length: packet.rectCount }] : [],
      textSpans: packet.textCount > 0 ? [{ offset: FULL_SPAN_START, length: packet.textCount }] : [],
      glyphSpans: [],
    },
  }
}

function createTileCellBoundsResolver(input: {
  readonly columnWidths: Readonly<Record<number, number>>
  readonly gridMetrics: GridMetrics
  readonly rowHeights: Readonly<Record<number, number>>
  readonly sortedColumnWidthOverrides: readonly (readonly [number, number])[]
  readonly sortedRowHeightOverrides: readonly (readonly [number, number])[]
  readonly viewport: Viewport
}): (col: number, row: number) => Rectangle | undefined {
  const baseX = resolveColumnOffset(input.viewport.colStart, input.sortedColumnWidthOverrides, input.gridMetrics.columnWidth)
  const baseY = resolveRowOffset(input.viewport.rowStart, input.sortedRowHeightOverrides, input.gridMetrics.rowHeight)
  return (col, row) => {
    if (col < input.viewport.colStart || col > input.viewport.colEnd || row < input.viewport.rowStart || row > input.viewport.rowEnd) {
      return undefined
    }
    return {
      height: getResolvedRowHeight(input.rowHeights, row, input.gridMetrics.rowHeight),
      width: getResolvedColumnWidth(input.columnWidths, col, input.gridMetrics.columnWidth),
      x: resolveColumnOffset(col, input.sortedColumnWidthOverrides, input.gridMetrics.columnWidth) - baseX,
      y: resolveRowOffset(row, input.sortedRowHeightOverrides, input.gridMetrics.rowHeight) - baseY,
    }
  }
}

function resolveViewportSurfaceSize(
  viewport: Viewport,
  input: {
    readonly gridMetrics: GridMetrics
    readonly sortedColumnWidthOverrides: readonly (readonly [number, number])[]
    readonly sortedRowHeightOverrides: readonly (readonly [number, number])[]
  },
): { readonly width: number; readonly height: number } {
  return {
    height:
      resolveRowOffset(viewport.rowEnd + 1, input.sortedRowHeightOverrides, input.gridMetrics.rowHeight) -
      resolveRowOffset(viewport.rowStart, input.sortedRowHeightOverrides, input.gridMetrics.rowHeight),
    width:
      resolveColumnOffset(viewport.colEnd + 1, input.sortedColumnWidthOverrides, input.gridMetrics.columnWidth) -
      resolveColumnOffset(viewport.colStart, input.sortedColumnWidthOverrides, input.gridMetrics.columnWidth),
  }
}

function mapTextRun(run: GridSceneTextRun): RenderTileTextRun {
  return {
    text: run.text,
    x: run.x,
    y: run.y,
    width: run.width,
    height: run.height,
    clipX: run.clipX,
    clipY: run.clipY,
    clipWidth: run.clipWidth,
    clipHeight: run.clipHeight,
    font: run.font,
    fontSize: run.fontSize,
    color: run.color,
    underline: run.underline,
    strike: run.strike,
  }
}
