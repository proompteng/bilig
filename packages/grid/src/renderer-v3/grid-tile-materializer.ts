import { VIEWPORT_TILE_COLUMN_COUNT, VIEWPORT_TILE_ROW_COUNT, type Viewport } from '@bilig/protocol'
import { buildGridGpuScene } from '../gridGpuScene.js'
import { getResolvedColumnWidth, getResolvedRowHeight, resolveRowOffset, type GridMetrics } from '../gridMetrics.js'
import { buildGridTextScene } from '../gridTextScene.js'
import type { GridEngineLike } from '../grid-engine.js'
import { CompactSelection, type GridSelection, type Item, type Rectangle } from '../gridTypes.js'
import { collectViewportItems } from '../gridViewportItems.js'
import { resolveColumnOffset } from '../workbookGridViewport.js'
import { packGridRectBufferV3 } from './rect-instance-buffer.js'
import type { GridRenderTile } from './render-tile-source.js'
import { createGridTilePacketV3 } from './tile-packet-v3.js'
import { packTileKey53 } from './tile-key.js'
import { packGridTextBufferV3 } from './text-run-buffer.js'

export interface GridTileMaterializerAxisInputV3 {
  readonly columnWidths: Readonly<Record<number, number>>
  readonly rowHeights: Readonly<Record<number, number>>
  readonly sortedColumnWidthOverrides: readonly (readonly [number, number])[]
  readonly sortedRowHeightOverrides: readonly (readonly [number, number])[]
  readonly gridMetrics: GridMetrics
}

export interface MaterializeGridRenderTileInputV3 extends GridTileMaterializerAxisInputV3 {
  readonly engine: GridEngineLike
  readonly sheetName: string
  readonly sheetId: number
  readonly viewport: Viewport
  readonly dprBucket: number
  readonly packetSeq: number
  readonly materializedAtSeq: number
  readonly cameraSeq: number
  readonly valueSeq: number
  readonly styleSeq: number
  readonly textSeq: number
  readonly rectSeq: number
  readonly axisSeqX: number
  readonly axisSeqY: number
  readonly freezeSeq: number
  readonly glyphAtlasSeq?: number | undefined
  readonly dirtyLocalRows?: Uint32Array | undefined
  readonly dirtyLocalCols?: Uint32Array | undefined
  readonly dirtyMasks?: Uint32Array | undefined
}

const STATIC_TILE_SELECTED_CELL: Item = Object.freeze([-1, -1] as const)
const STATIC_TILE_GRID_SELECTION: GridSelection = Object.freeze({
  columns: CompactSelection.empty(),
  current: undefined,
  rows: CompactSelection.empty(),
})

export function materializeGridRenderTileV3(input: MaterializeGridRenderTileInputV3): GridRenderTile {
  const rowTile = Math.floor(input.viewport.rowStart / VIEWPORT_TILE_ROW_COUNT)
  const colTile = Math.floor(input.viewport.colStart / VIEWPORT_TILE_COLUMN_COUNT)
  const tileId = packTileKey53({
    colTile,
    dprBucket: input.dprBucket,
    rowTile,
    sheetOrdinal: input.sheetId,
  })
  const surfaceSize = resolveTileSurfaceSizeV3(input.viewport, input)
  const visibleItems = collectViewportItems(input.viewport)
  const visibleRegion = {
    range: {
      x: input.viewport.colStart,
      y: input.viewport.rowStart,
      width: input.viewport.colEnd - input.viewport.colStart + 1,
      height: input.viewport.rowEnd - input.viewport.rowStart + 1,
    },
    tx: 0,
    ty: 0,
    freezeRows: 0,
    freezeCols: 0,
  }
  const getCellBounds = createTileCellBoundsResolverV3(input)
  const gpuScene = buildGridGpuScene({
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
  })
  const textScene = buildGridTextScene({
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
  })
  const rectBuffer = packGridRectBufferV3(gpuScene, surfaceSize)
  const textBuffer = packGridTextBufferV3(textScene)
  const packet = createGridTilePacketV3({
    axisSeqX: input.axisSeqX,
    axisSeqY: input.axisSeqY,
    cellCount: visibleItems.length,
    dirtyLocalCols: input.dirtyLocalCols,
    dirtyLocalRows: input.dirtyLocalRows,
    dirtyMasks: input.dirtyMasks,
    freezeSeq: input.freezeSeq,
    glyphAtlasSeq: input.glyphAtlasSeq ?? 0,
    materializedAtSeq: input.materializedAtSeq,
    packetSeq: input.packetSeq,
    rectInstanceCount: rectBuffer.rectCount,
    rectInstances: rectBuffer.rectInstances,
    rectSeq: input.rectSeq,
    sheetId: input.sheetId,
    styleSeq: input.styleSeq,
    textRunCount: textBuffer.textCount,
    textRuns: textBuffer.textMetrics,
    textSeq: input.textSeq,
    tileKey: tileId,
    valueSeq: input.valueSeq,
  })
  return {
    bounds: {
      colEnd: packet.colEnd,
      colStart: packet.colStart,
      rowEnd: packet.rowEnd,
      rowStart: packet.rowStart,
    },
    coord: {
      colTile,
      dprBucket: input.dprBucket,
      paneKind: 'body',
      rowTile,
      sheetId: input.sheetId,
    },
    lastBatchId: packet.packetSeq,
    lastCameraSeq: input.cameraSeq,
    packet,
    rectCount: rectBuffer.rectCount,
    rectInstances: rectBuffer.rectInstances,
    textCount: textBuffer.textCount,
    textMetrics: textBuffer.textMetrics,
    textRuns: textBuffer.textRuns,
    dirtyLocalCols: packet.dirtyLocalCols,
    dirtyLocalRows: packet.dirtyLocalRows,
    dirtyMasks: packet.dirtyMasks,
    tileId,
    version: {
      axisX: packet.axisSeqX,
      axisY: packet.axisSeqY,
      freeze: packet.freezeSeq,
      styles: packet.styleSeq,
      text: packet.textSeq,
      values: packet.valueSeq,
    },
  }
}

export function createTileCellBoundsResolverV3(
  input: GridTileMaterializerAxisInputV3 & { readonly viewport: Viewport },
): (col: number, row: number) => Rectangle | undefined {
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

export function resolveTileSurfaceSizeV3(
  viewport: Viewport,
  input: GridTileMaterializerAxisInputV3,
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
