import { indexToColumn } from '@bilig/formula'
import { buildGridGpuHeaderScene } from '../gridGpuHeaderScene.js'
import { getResolvedCellFontFamily } from '../gridCells.js'
import { getVisibleColumnBounds, getVisibleRowBounds, type GridMetrics } from '../gridMetrics.js'
import { buildHeaderPaneStates, type GridHeaderPaneState } from '../gridHeaderPanes.js'
import { CompactSelection, type GridSelection, type Item, type Rectangle } from '../gridTypes.js'
import type { GridTextItem, GridTextScene } from '../gridTextScene.js'
import { collectVisibleColumnBounds, collectVisibleRowBounds } from '../visibleGridAxes.js'
import { workbookThemeColors } from '../workbookTheme.js'
import { parseGpuColor, type GridGpuScene } from '../gridGpuScene.js'

const STATIC_SELECTED_CELL: Item = Object.freeze([-1, -1] as const)
const STATIC_GRID_SELECTION: GridSelection = Object.freeze({
  columns: CompactSelection.empty(),
  current: undefined,
  rows: CompactSelection.empty(),
})
const DEFAULT_HEADER_FONT_SIZE = 11

export interface WorkbookHeaderPaneInputV3 {
  readonly columnWidths: Readonly<Record<number, number>>
  readonly freezeCols: number
  readonly freezeRows: number
  readonly frozenColumnWidth: number
  readonly frozenRowHeight: number
  readonly getHeaderCellLocalBounds: (col: number, row: number) => Rectangle | undefined
  readonly gridMetrics: GridMetrics
  readonly hostClientHeight: number
  readonly hostClientWidth: number
  readonly residentBodyHeight: number
  readonly residentBodyWidth: number
  readonly residentHeaderItems: readonly Item[]
  readonly residentHeaderRegion: {
    readonly range: Pick<Rectangle, 'x' | 'y' | 'width' | 'height'>
    readonly tx: number
    readonly ty: number
    readonly freezeRows: number
    readonly freezeCols: number
  }
  readonly residentViewport: {
    readonly rowStart: number
    readonly rowEnd: number
    readonly colStart: number
    readonly colEnd: number
  }
  readonly rowHeights: Readonly<Record<number, number>>
  readonly sheetName: string
}

export function buildWorkbookHeaderPaneStatesV3(input: WorkbookHeaderPaneInputV3): readonly GridHeaderPaneState[] {
  return buildHeaderPaneStates({
    gpuScene: buildWorkbookHeaderGpuSceneV3(input),
    textScene: buildWorkbookHeaderTextSceneV3(input),
    sheetName: input.sheetName,
    residentViewport: input.residentViewport,
    freezeCols: input.freezeCols,
    freezeRows: input.freezeRows,
    hostWidth: input.hostClientWidth,
    hostHeight: input.hostClientHeight,
    gridMetrics: input.gridMetrics,
    frozenColumnWidth: input.frozenColumnWidth,
    frozenRowHeight: input.frozenRowHeight,
    residentBodyWidth: input.residentBodyWidth,
    residentBodyHeight: input.residentBodyHeight,
  })
}

export function buildWorkbookHeaderGpuSceneV3(input: WorkbookHeaderPaneInputV3): GridGpuScene {
  return buildGridGpuHeaderScene({
    palette: {
      gridLineColor: parseGpuColor(workbookThemeColors.gridBorder),
      headerDragAnchorFillColor: parseGpuColor('rgba(33, 86, 58, 0.18)'),
      headerFillColor: parseGpuColor(workbookThemeColors.surfaceSubtle),
      headerHoverFillColor: parseGpuColor(workbookThemeColors.muted),
      headerSelectedFillColor: parseGpuColor(workbookThemeColors.accentSoft),
      resizeGuideColor: parseGpuColor('rgba(33, 86, 58, 0.72)'),
      resizeGuideGlowColor: parseGpuColor('rgba(191, 213, 196, 0.28)'),
      selectionFillColor: parseGpuColor('rgba(33, 86, 58, 0.08)'),
    },
    activeHeaderDrag: null,
    columnWidths: input.columnWidths,
    getCellBounds: input.getHeaderCellLocalBounds,
    gridMetrics: input.gridMetrics,
    gridSelection: STATIC_GRID_SELECTION,
    hoveredHeader: null,
    resizeGuideColumn: null,
    resizeGuideRow: null,
    rowHeights: input.rowHeights,
    selectedCell: STATIC_SELECTED_CELL,
    selectionRange: null,
    visibleItems: input.residentHeaderItems,
    visibleRegion: input.residentHeaderRegion,
  })
}

export function buildWorkbookHeaderTextSceneV3(input: WorkbookHeaderPaneInputV3): GridTextScene {
  const items: GridTextItem[] = []
  const headerFontSize = DEFAULT_HEADER_FONT_SIZE
  const headerFont = `500 ${headerFontSize}px ${getResolvedCellFontFamily()}`
  const hasFrozenAxes = input.freezeRows > 0 || input.freezeCols > 0
  const visibleColumns = hasFrozenAxes
    ? collectVisibleColumnBounds(input.residentHeaderItems, input.getHeaderCellLocalBounds, input.gridMetrics)
    : getVisibleColumnBounds(
        input.residentHeaderRegion.range,
        input.gridMetrics.rowMarkerWidth - input.residentHeaderRegion.tx,
        Number.MAX_SAFE_INTEGER,
        input.columnWidths,
        input.gridMetrics.columnWidth,
      )

  for (const column of visibleColumns) {
    items.push({
      align: 'center',
      clipInsetBottom: 0,
      clipInsetLeft: Math.max(0, input.gridMetrics.rowMarkerWidth - column.left),
      clipInsetRight: 0,
      clipInsetTop: 0,
      color: workbookThemeColors.textMuted,
      font: headerFont,
      fontSize: headerFontSize,
      height: input.gridMetrics.headerHeight,
      strike: false,
      text: indexToColumn(column.index),
      underline: false,
      width: column.width,
      wrap: false,
      x: column.left,
      y: 0,
    })
  }

  const visibleRows = hasFrozenAxes
    ? collectVisibleRowBounds(input.residentHeaderItems, input.getHeaderCellLocalBounds, input.gridMetrics)
    : getVisibleRowBounds(
        input.residentHeaderRegion.range,
        input.gridMetrics.headerHeight - input.residentHeaderRegion.ty,
        Number.MAX_SAFE_INTEGER,
        input.rowHeights,
        input.gridMetrics.rowHeight,
      )

  for (const row of visibleRows) {
    items.push({
      align: 'right',
      clipInsetBottom: 0,
      clipInsetLeft: 0,
      clipInsetRight: 0,
      clipInsetTop: Math.max(0, input.gridMetrics.headerHeight - row.top),
      color: workbookThemeColors.textMuted,
      font: headerFont,
      fontSize: headerFontSize,
      height: row.height,
      strike: false,
      text: String(row.index + 1),
      underline: false,
      width: input.gridMetrics.rowMarkerWidth,
      wrap: false,
      x: 0,
      y: row.top,
    })
  }

  return { items }
}
