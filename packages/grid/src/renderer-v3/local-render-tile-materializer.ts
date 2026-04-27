import { MAX_COLS, MAX_ROWS, VIEWPORT_TILE_COLUMN_COUNT, VIEWPORT_TILE_ROW_COUNT, type Viewport } from '@bilig/protocol'
import { buildGridGpuScene } from '../gridGpuScene.js'
import { getResolvedColumnWidth, getResolvedRowHeight, resolveRowOffset, type GridMetrics } from '../gridMetrics.js'
import { buildGridTextScene } from '../gridTextScene.js'
import type { GridEngineLike } from '../grid-engine.js'
import { CompactSelection, type GridSelection, type Item, type Rectangle } from '../gridTypes.js'
import { collectViewportItems } from '../gridViewportItems.js'
import { resolveColumnOffset } from '../workbookGridViewport.js'
import {
  createGridTileKeyV2,
  packGridScenePacketV2,
  type GridScenePacketV2,
  type GridSceneTextRun,
} from '../renderer-v2/scene-packet-v2.js'
import { validateGridScenePacketV2 } from '../renderer-v2/scene-packet-validator.js'
import { packTileKey53, unpackTileKey53, tileKeysForViewport } from './tile-key.js'
import type { GridRenderTile, GridRenderTileTextRun } from './render-tile-source.js'

interface RenderTileAxisInput {
  readonly columnWidths: Readonly<Record<number, number>>
  readonly rowHeights: Readonly<Record<number, number>>
  readonly sortedColumnWidthOverrides: readonly (readonly [number, number])[]
  readonly sortedRowHeightOverrides: readonly (readonly [number, number])[]
  readonly gridMetrics: GridMetrics
}

const STATIC_TILE_SELECTED_CELL: Item = Object.freeze([-1, -1] as const)
const STATIC_TILE_GRID_SELECTION: GridSelection = Object.freeze({
  columns: CompactSelection.empty(),
  current: undefined,
  rows: CompactSelection.empty(),
})

export function buildLocalFixedRenderTiles(input: {
  readonly engine: GridEngineLike
  readonly sheetName: string
  readonly sheetId: number
  readonly viewport: Viewport
  readonly columnWidths: Readonly<Record<number, number>>
  readonly rowHeights: Readonly<Record<number, number>>
  readonly sortedColumnWidthOverrides: readonly (readonly [number, number])[]
  readonly sortedRowHeightOverrides: readonly (readonly [number, number])[]
  readonly gridMetrics: GridMetrics
  readonly dprBucket: number
  readonly generation: number
  readonly cameraSeq: number
}): readonly GridRenderTile[] {
  const axisVersionX = hashAxisOverrides(input.sortedColumnWidthOverrides)
  const axisVersionY = hashAxisOverrides(input.sortedRowHeightOverrides)
  return tileKeysForViewport({
    dprBucket: input.dprBucket,
    sheetOrdinal: input.sheetId,
    viewport: input.viewport,
  }).map((tileId) => {
    const key = unpackTileKey53(tileId)
    const tileViewport = viewportFromTileKey(key.rowTile, key.colTile)
    const packet = buildLocalFixedRenderTilePacket({
      ...input,
      axisVersionX,
      axisVersionY,
      tileViewport,
    })
    return renderTileFromPacket(input.sheetId, packet)
  })
}

function buildLocalFixedRenderTilePacket(
  input: RenderTileAxisInput & {
    readonly engine: GridEngineLike
    readonly sheetName: string
    readonly sheetId: number
    readonly tileViewport: Viewport
    readonly dprBucket: number
    readonly generation: number
    readonly cameraSeq: number
    readonly axisVersionX: number
    readonly axisVersionY: number
  },
): GridScenePacketV2 {
  const surfaceSize = resolveTileSurfaceSize(input.tileViewport, input)
  const visibleItems = collectViewportItems(input.tileViewport)
  const visibleRegion = {
    range: {
      x: input.tileViewport.colStart,
      y: input.tileViewport.rowStart,
      width: input.tileViewport.colEnd - input.tileViewport.colStart + 1,
      height: input.tileViewport.rowEnd - input.tileViewport.rowStart + 1,
    },
    tx: 0,
    ty: 0,
    freezeRows: 0,
    freezeCols: 0,
  }
  const getCellBounds = createTileCellBoundsResolver(input)
  const packet = packGridScenePacketV2({
    cameraSeq: input.cameraSeq,
    generatedAt: Date.now(),
    generation: input.generation,
    gpuScene: buildGridGpuScene({
      activeHeaderDrag: null,
      columnWidths: input.columnWidths,
      contentMode: 'data',
      engine: input.engine,
      getCellBounds,
      gridMetrics: input.gridMetrics,
      gridSelection: STATIC_TILE_GRID_SELECTION,
      hostBounds: { left: 0, top: 0 },
      hoveredCell: null,
      hoveredHeader: null,
      resizeGuideColumn: null,
      resizeGuideRow: null,
      rowHeights: input.rowHeights,
      selectedCell: STATIC_TILE_SELECTED_CELL,
      selectionRange: null,
      sheetName: input.sheetName,
      visibleItems,
      visibleRegion,
    }),
    key: createGridTileKeyV2({
      axisVersionX: input.axisVersionX,
      axisVersionY: input.axisVersionY,
      colTile: Math.floor(input.tileViewport.colStart / VIEWPORT_TILE_COLUMN_COUNT),
      dprBucket: input.dprBucket,
      freezeVersion: 0,
      paneId: 'body',
      rowTile: Math.floor(input.tileViewport.rowStart / VIEWPORT_TILE_ROW_COUNT),
      selectionIndependentVersion: input.generation,
      sheetName: input.sheetName,
      styleVersion: input.generation,
      textEpoch: input.generation,
      valueVersion: input.generation,
      viewport: input.tileViewport,
    }),
    paneId: 'body',
    requestSeq: input.generation,
    sheetName: input.sheetName,
    surfaceSize,
    textScene: buildGridTextScene({
      activeHeaderDrag: null,
      columnWidths: input.columnWidths,
      contentMode: 'data',
      editingCell: null,
      engine: input.engine,
      getCellBounds,
      gridMetrics: input.gridMetrics,
      hostBounds: {
        height: surfaceSize.height,
        left: 0,
        top: 0,
        width: surfaceSize.width,
      },
      hoveredHeader: null,
      resizeGuideColumn: null,
      rowHeights: input.rowHeights,
      selectedCell: STATIC_TILE_SELECTED_CELL,
      selectedCellSnapshot: null,
      selectionRange: null,
      sheetName: input.sheetName,
      visibleItems,
      visibleRegion,
    }),
    viewport: input.tileViewport,
  })
  const validation = validateGridScenePacketV2(packet)
  if (!validation.ok) {
    throw new Error(`Invalid local render tile packet: ${validation.reason}`)
  }
  return packet
}

function renderTileFromPacket(sheetId: number, packet: GridScenePacketV2): GridRenderTile {
  const tileId = packTileKey53({
    colTile: packet.key.colTile,
    dprBucket: packet.key.dprBucket,
    rowTile: packet.key.rowTile,
    sheetOrdinal: sheetId,
  })
  return {
    bounds: packet.viewport,
    coord: {
      colTile: packet.key.colTile,
      dprBucket: packet.key.dprBucket,
      paneKind: 'body',
      rowTile: packet.key.rowTile,
      sheetId,
    },
    lastBatchId: packet.requestSeq,
    lastCameraSeq: packet.cameraSeq,
    rectCount: packet.rectCount,
    rectInstances: packet.rectInstances,
    textCount: packet.textCount,
    textMetrics: packet.textMetrics,
    textRuns: packet.textRuns.map(mapTextRun),
    tileId,
    version: {
      axisX: packet.key.axisVersionX,
      axisY: packet.key.axisVersionY,
      freeze: packet.key.freezeVersion,
      styles: packet.key.styleVersion,
      text: packet.key.textEpoch,
      values: packet.key.valueVersion,
    },
  }
}

function mapTextRun(run: GridSceneTextRun): GridRenderTileTextRun {
  return {
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
  }
}

function createTileCellBoundsResolver(input: RenderTileAxisInput & { readonly tileViewport: Viewport }) {
  const baseX = resolveColumnOffset(input.tileViewport.colStart, input.sortedColumnWidthOverrides, input.gridMetrics.columnWidth)
  const baseY = resolveRowOffset(input.tileViewport.rowStart, input.sortedRowHeightOverrides, input.gridMetrics.rowHeight)
  return (col: number, row: number): Rectangle | undefined => {
    if (
      col < input.tileViewport.colStart ||
      col > input.tileViewport.colEnd ||
      row < input.tileViewport.rowStart ||
      row > input.tileViewport.rowEnd
    ) {
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

function resolveTileSurfaceSize(viewport: Viewport, input: RenderTileAxisInput): { readonly width: number; readonly height: number } {
  return {
    height:
      resolveRowOffset(viewport.rowEnd + 1, input.sortedRowHeightOverrides, input.gridMetrics.rowHeight) -
      resolveRowOffset(viewport.rowStart, input.sortedRowHeightOverrides, input.gridMetrics.rowHeight),
    width:
      resolveColumnOffset(viewport.colEnd + 1, input.sortedColumnWidthOverrides, input.gridMetrics.columnWidth) -
      resolveColumnOffset(viewport.colStart, input.sortedColumnWidthOverrides, input.gridMetrics.columnWidth),
  }
}

function viewportFromTileKey(rowTile: number, colTile: number): Viewport {
  const rowStart = rowTile * VIEWPORT_TILE_ROW_COUNT
  const colStart = colTile * VIEWPORT_TILE_COLUMN_COUNT
  return {
    colEnd: Math.min(MAX_COLS - 1, colStart + VIEWPORT_TILE_COLUMN_COUNT - 1),
    colStart,
    rowEnd: Math.min(MAX_ROWS - 1, rowStart + VIEWPORT_TILE_ROW_COUNT - 1),
    rowStart,
  }
}

function hashAxisOverrides(entries: readonly (readonly [number, number])[]): number {
  if (entries.length === 0) {
    return 0
  }
  let hash = 2_166_136_261
  for (const [index, size] of entries) {
    hash = mixRevisionInteger(hash, index)
    hash = mixRevisionInteger(hash, Math.round(size * 1_000))
  }
  return hash >>> 0
}

function mixRevisionInteger(hash: number, value: number): number {
  return Math.imul((hash ^ value) >>> 0, 16_777_619) >>> 0
}
