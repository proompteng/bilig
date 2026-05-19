import { formatAddress } from '@bilig/formula'
import { ValueTag, type CellStyleRecord, type Viewport } from '@bilig/protocol'

import type { GridEngineLike } from '../grid-engine.js'
import { snapshotToRenderCell } from '../gridCells.js'
import { buildGridGpuScene } from '../gridGpuScene.js'
import type { GridMetrics } from '../gridMetrics.js'
import { parseGpuColor, type GridGpuColor } from '../gridGpuPrimitives.js'
import { CompactSelection, type GridSelection, type Item, type Rectangle } from '../gridTypes.js'
import { collectViewportItems } from '../gridViewportItems.js'
import { GRID_RECT_INSTANCE_FLOAT_COUNT_V3, packGridRectBufferV3 } from '../renderer-v3/rect-instance-buffer.js'
import { createTileCellBoundsResolverV3, resolveTileSurfaceSizeV3 } from '../renderer-v3/grid-tile-materializer.js'
import type { GridRenderTile } from '../renderer-v3/render-tile-source.js'

interface VisibleTextRefreshCacheInput {
  readonly columnWidths: Readonly<Record<number, number>>
  readonly engine: GridEngineLike
  readonly gridMetrics: GridMetrics
  readonly rowHeights: Readonly<Record<number, number>>
  readonly sceneRevision: number
  readonly sheetName: string
  readonly sortedColumnWidthOverrides: readonly (readonly [number, number])[]
  readonly sortedRowHeightOverrides: readonly (readonly [number, number])[]
  readonly visibleViewport: Viewport
}

interface VisibleCellFillExpectation {
  readonly bounds: Rectangle
  readonly colorKey: string | null
  readonly allowsPartialFill: boolean
}

interface TileFillRect {
  readonly x: number
  readonly y: number
  readonly width: number
  readonly height: number
  readonly colorKey: string
}

const CELL_FILL_COVERAGE_EPSILON = 0.5
const STATIC_TILE_SELECTED_CELL: Item = Object.freeze([-1, -1] as const)
const STATIC_TILE_GRID_SELECTION: GridSelection = Object.freeze({
  columns: CompactSelection.empty(),
  current: undefined,
  rows: CompactSelection.empty(),
})

interface VisibleTextRefreshCacheEntry extends VisibleTextRefreshCacheInput {
  readonly needsLocalRefresh: boolean
  readonly renderRevisionKey: string
  readonly tile: GridRenderTile
  readonly visibleColEnd: number
  readonly visibleColStart: number
  readonly visibleRowEnd: number
  readonly visibleRowStart: number
}

export class GridVisibleTextRefreshCache {
  private readonly entries = new Map<number, VisibleTextRefreshCacheEntry>()

  needsLocalRefresh(tileKey: number, tile: GridRenderTile | null, input: VisibleTextRefreshCacheInput): boolean {
    if (!tile) {
      this.entries.delete(tileKey)
      return false
    }
    const visibleRowStart = Math.max(tile.bounds.rowStart, input.visibleViewport.rowStart)
    const visibleRowEnd = Math.min(tile.bounds.rowEnd, input.visibleViewport.rowEnd)
    const visibleColStart = Math.max(tile.bounds.colStart, input.visibleViewport.colStart)
    const visibleColEnd = Math.min(tile.bounds.colEnd, input.visibleViewport.colEnd)
    if (visibleRowStart > visibleRowEnd || visibleColStart > visibleColEnd) {
      this.entries.delete(tileKey)
      return false
    }

    const renderRevisionKey = resolveRenderRevisionKey(input.engine)
    const cached = this.entries.get(tileKey)
    if (
      cached &&
      cached.tile === tile &&
      cached.columnWidths === input.columnWidths &&
      cached.engine === input.engine &&
      cached.gridMetrics === input.gridMetrics &&
      cached.rowHeights === input.rowHeights &&
      cached.sheetName === input.sheetName &&
      cached.sceneRevision === input.sceneRevision &&
      cached.renderRevisionKey === renderRevisionKey &&
      cached.sortedColumnWidthOverrides === input.sortedColumnWidthOverrides &&
      cached.sortedRowHeightOverrides === input.sortedRowHeightOverrides &&
      cached.visibleRowStart === visibleRowStart &&
      cached.visibleRowEnd === visibleRowEnd &&
      cached.visibleColStart === visibleColStart &&
      cached.visibleColEnd === visibleColEnd
    ) {
      return cached.needsLocalRefresh
    }

    const needsLocalRefresh = tileVisibleTextNeedsLocalRefresh(tile, input, {
      visibleColEnd,
      visibleColStart,
      visibleRowEnd,
      visibleRowStart,
    })
    this.entries.set(tileKey, {
      ...input,
      needsLocalRefresh,
      renderRevisionKey,
      tile,
      visibleColEnd,
      visibleColStart,
      visibleRowEnd,
      visibleRowStart,
    })
    return needsLocalRefresh
  }
}

function tileVisibleTextNeedsLocalRefresh(
  tile: GridRenderTile,
  input: Pick<
    VisibleTextRefreshCacheInput,
    'columnWidths' | 'engine' | 'gridMetrics' | 'rowHeights' | 'sheetName' | 'sortedColumnWidthOverrides' | 'sortedRowHeightOverrides'
  >,
  visibleBounds: {
    readonly visibleColEnd: number
    readonly visibleColStart: number
    readonly visibleRowEnd: number
    readonly visibleRowStart: number
  },
): boolean {
  if (tile.textRuns.length !== tile.textCount) {
    return true
  }

  const textRunsByCell = new Map<string, GridRenderTile['textRuns'][number]>()
  for (const run of tile.textRuns) {
    if (run.text.length === 0) {
      continue
    }
    const row = run.row
    const col = run.col
    if (
      !Number.isInteger(row) ||
      !Number.isInteger(col) ||
      row === undefined ||
      col === undefined ||
      row < tile.bounds.rowStart ||
      row > tile.bounds.rowEnd ||
      col < tile.bounds.colStart ||
      col > tile.bounds.colEnd
    ) {
      return true
    }
    if (
      row >= visibleBounds.visibleRowStart &&
      row <= visibleBounds.visibleRowEnd &&
      col >= visibleBounds.visibleColStart &&
      col <= visibleBounds.visibleColEnd
    ) {
      const key = `${row}:${col}`
      if (textRunsByCell.has(key)) {
        return true
      }
      textRunsByCell.set(key, run)
    }
  }

  const projectedRevision = resolveProjectedRevision(input.engine)
  const tileStyleRevisionIsBehind = projectedRevision !== null && tile.version.styles < projectedRevision
  let needsCurrentStyleProof = tileStyleRevisionIsBehind && tileHasAuthoredPaintRects(tile)
  const expectedVisibleFills: VisibleCellFillExpectation[] = []
  const getCellBounds = createTileCellBoundsResolverV3({
    columnWidths: input.columnWidths,
    gridMetrics: input.gridMetrics,
    rowHeights: input.rowHeights,
    sortedColumnWidthOverrides: input.sortedColumnWidthOverrides,
    sortedRowHeightOverrides: input.sortedRowHeightOverrides,
    viewport: tile.bounds,
  })

  for (let row = visibleBounds.visibleRowStart; row <= visibleBounds.visibleRowEnd; row += 1) {
    for (let col = visibleBounds.visibleColStart; col <= visibleBounds.visibleColEnd; col += 1) {
      const snapshot = input.engine.getCell(input.sheetName, formatAddress(row, col))
      const style = input.engine.getCellStyle(snapshot.styleId)
      const bounds = getCellBounds(col, row)
      if (!bounds) {
        return true
      }
      expectedVisibleFills.push({
        allowsPartialFill: snapshot.value.tag === ValueTag.Boolean,
        bounds,
        colorKey: style?.fill?.backgroundColor ? gpuColorKey(parseGpuColor(style.fill.backgroundColor)) : null,
      })
      const renderCell = snapshotToRenderCell(snapshot, style)
      const visibleTileRun = textRunsByCell.get(`${row}:${col}`) ?? null
      if ((visibleTileRun?.text ?? '') !== renderCell.displayText) {
        return true
      }
      if (visibleTileRun && textRunStyleDiffersFromRenderCell(visibleTileRun, renderCell)) {
        return true
      }
      needsCurrentStyleProof = needsCurrentStyleProof || (tileStyleRevisionIsBehind && styleAffectsVisibleGridPaint(style))
    }
  }
  if (!tileMatchesExpectedVisibleRectScene(tile, input, expectedVisibleFills)) {
    return true
  }
  return needsCurrentStyleProof
}

function resolveRenderRevisionKey(engine: GridEngineLike): string {
  const revision = engine.getRenderRevisionSnapshot?.()
  if (!revision) {
    return 'untracked'
  }
  return [
    revision.authoritativeRevision ?? 'none',
    revision.localRevision ?? 'none',
    revision.projectedRevision,
    revision.tileSceneRevision ?? 'none',
    revision.tileSceneCameraSeq ?? 'none',
  ].join(':')
}

function resolveProjectedRevision(engine: GridEngineLike): number | null {
  const revision = engine.getRenderRevisionSnapshot?.().projectedRevision
  return typeof revision === 'number' && Number.isInteger(revision) && revision >= 0 ? revision : null
}

function styleAffectsVisibleGridPaint(style: CellStyleRecord | undefined): boolean {
  return Boolean(style?.fill?.backgroundColor || style?.borders)
}

function tileHasAuthoredPaintRects(tile: GridRenderTile): boolean {
  const readableRectCount = Math.min(tile.rectCount, Math.floor(tile.rectInstances.length / GRID_RECT_INSTANCE_FLOAT_COUNT_V3))
  for (let index = 0; index < readableRectCount; index += 1) {
    const offset = index * GRID_RECT_INSTANCE_FLOAT_COUNT_V3
    const instanceKind = tile.rectInstances[offset + 13] ?? -1
    const fillAlpha = tile.rectInstances[offset + 7] ?? 0
    if (instanceKind === 0 && fillAlpha > 0.01) {
      return true
    }
  }
  return tile.rectCount > expectedBaseGridLineRectCount(tile)
}

function tileMatchesExpectedVisibleRectScene(
  tile: GridRenderTile,
  input: Pick<
    VisibleTextRefreshCacheInput,
    'columnWidths' | 'engine' | 'gridMetrics' | 'rowHeights' | 'sheetName' | 'sortedColumnWidthOverrides' | 'sortedRowHeightOverrides'
  >,
  expectedVisibleFills: readonly VisibleCellFillExpectation[],
): boolean {
  if (!tile.rectSignature) {
    return tileMatchesExpectedVisibleFillCoverage(tile, expectedVisibleFills)
  }
  const getCellBounds = createTileCellBoundsResolverV3({
    columnWidths: input.columnWidths,
    gridMetrics: input.gridMetrics,
    rowHeights: input.rowHeights,
    sortedColumnWidthOverrides: input.sortedColumnWidthOverrides,
    sortedRowHeightOverrides: input.sortedRowHeightOverrides,
    viewport: tile.bounds,
  })
  const visibleItems = collectViewportItems(tile.bounds)
  const expectedRectBuffer = packGridRectBufferV3(
    buildGridGpuScene({
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
      includeLeadingGridLines: false,
      resizeGuideColumn: null,
      resizeGuideRow: null,
      rowHeights: input.rowHeights,
      selectedCell: STATIC_TILE_SELECTED_CELL,
      selectionRange: null,
      sheetName: input.sheetName,
      visibleItems,
      visibleRegion: {
        range: {
          height: tile.bounds.rowEnd - tile.bounds.rowStart + 1,
          width: tile.bounds.colEnd - tile.bounds.colStart + 1,
          x: tile.bounds.colStart,
          y: tile.bounds.rowStart,
        },
        freezeCols: 0,
        freezeRows: 0,
        tx: 0,
        ty: 0,
      },
    }),
    resolveTileSurfaceSizeV3(tile.bounds, input),
  )
  return tile.rectSignature === expectedRectBuffer.rectSignature
}

function tileMatchesExpectedVisibleFillCoverage(
  tile: GridRenderTile,
  expectedVisibleFills: readonly VisibleCellFillExpectation[],
): boolean {
  const tileFillRects = collectTileFillRects(tile)
  for (const expectation of expectedVisibleFills) {
    const intersectingRects = tileFillRects.filter((rect) => fillRectIntersectsCell(rect, expectation.bounds))
    const coveringRects = tileFillRects.filter((rect) => fillRectCoversCell(rect, expectation.bounds))
    if (!expectation.colorKey) {
      if ((expectation.allowsPartialFill ? coveringRects : intersectingRects).length > 0) {
        return false
      }
      continue
    }
    if (!coveringRects.some((rect) => rect.colorKey === expectation.colorKey)) {
      return false
    }
    if (coveringRects.some((rect) => rect.colorKey !== expectation.colorKey)) {
      return false
    }
    if (!expectation.allowsPartialFill && intersectingRects.some((rect) => rect.colorKey !== expectation.colorKey)) {
      return false
    }
  }
  return true
}

function collectTileFillRects(tile: GridRenderTile): TileFillRect[] {
  const readableRectCount = Math.min(tile.rectCount, Math.floor(tile.rectInstances.length / GRID_RECT_INSTANCE_FLOAT_COUNT_V3))
  const rects: TileFillRect[] = []
  for (let index = 0; index < readableRectCount; index += 1) {
    const offset = index * GRID_RECT_INSTANCE_FLOAT_COUNT_V3
    const instanceKind = tile.rectInstances[offset + 13] ?? -1
    const fillAlpha = tile.rectInstances[offset + 7] ?? 0
    if (instanceKind !== 0 || fillAlpha <= 0.01) {
      continue
    }
    rects.push({
      colorKey: gpuColorKey({
        a: fillAlpha,
        b: tile.rectInstances[offset + 6] ?? 0,
        g: tile.rectInstances[offset + 5] ?? 0,
        r: tile.rectInstances[offset + 4] ?? 0,
      }),
      height: tile.rectInstances[offset + 3] ?? 0,
      width: tile.rectInstances[offset + 2] ?? 0,
      x: tile.rectInstances[offset + 0] ?? 0,
      y: tile.rectInstances[offset + 1] ?? 0,
    })
  }
  return rects
}

function fillRectCoversCell(rect: TileFillRect, bounds: Rectangle): boolean {
  return (
    rect.x <= bounds.x + CELL_FILL_COVERAGE_EPSILON &&
    rect.y <= bounds.y + CELL_FILL_COVERAGE_EPSILON &&
    rect.x + rect.width >= bounds.x + bounds.width - CELL_FILL_COVERAGE_EPSILON &&
    rect.y + rect.height >= bounds.y + bounds.height - CELL_FILL_COVERAGE_EPSILON
  )
}

function fillRectIntersectsCell(rect: TileFillRect, bounds: Rectangle): boolean {
  return (
    rect.x < bounds.x + bounds.width - CELL_FILL_COVERAGE_EPSILON &&
    rect.x + rect.width > bounds.x + CELL_FILL_COVERAGE_EPSILON &&
    rect.y < bounds.y + bounds.height - CELL_FILL_COVERAGE_EPSILON &&
    rect.y + rect.height > bounds.y + CELL_FILL_COVERAGE_EPSILON
  )
}

function gpuColorKey(color: GridGpuColor): string {
  return [color.r, color.g, color.b, color.a].map((channel) => String(Math.round(channel * 1000))).join(':')
}

function expectedBaseGridLineRectCount(tile: GridRenderTile): number {
  const rowCount = Math.max(0, tile.bounds.rowEnd - tile.bounds.rowStart + 1)
  const colCount = Math.max(0, tile.bounds.colEnd - tile.bounds.colStart + 1)
  return rowCount + colCount
}

function textRunStyleDiffersFromRenderCell(
  run: GridRenderTile['textRuns'][number],
  renderCell: ReturnType<typeof snapshotToRenderCell>,
): boolean {
  return (
    (run.font ?? '') !== renderCell.font ||
    (run.color ?? '') !== renderCell.color ||
    (run.align ?? 'left') !== renderCell.align ||
    (run.wrap ?? false) !== renderCell.wrap ||
    run.underline !== renderCell.underline
  )
}
