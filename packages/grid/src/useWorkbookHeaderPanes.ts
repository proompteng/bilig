import { useMemo } from 'react'
import type { Viewport } from '@bilig/protocol'
import { buildGridGpuScene, type GridGpuScene } from './gridGpuScene.js'
import { buildGridTextScene, type GridTextScene } from './gridTextScene.js'
import { buildHeaderPaneStates, type GridHeaderPaneState } from './gridHeaderPanes.js'
import type { GridEngineLike } from './grid-engine.js'
import type { GridMetrics } from './gridMetrics.js'
import { CompactSelection, type GridSelection, type Item, type Rectangle } from './gridTypes.js'
import type { WorkbookRenderTilePaneState } from './renderer-v3/render-tile-pane-state.js'

const STATIC_SCENE_SELECTED_CELL: Item = Object.freeze([-1, -1] as const)
const STATIC_SCENE_GRID_SELECTION: GridSelection = Object.freeze({
  columns: CompactSelection.empty(),
  current: undefined,
  rows: CompactSelection.empty(),
})

function noteHeaderPaneBuild(): void {
  if (typeof window === 'undefined') {
    return
  }
  ;(window as Window & { __biligScrollPerf?: { noteHeaderPaneBuild?: () => void } }).__biligScrollPerf?.noteHeaderPaneBuild?.()
}

export function useWorkbookHeaderPanes(input: {
  readonly columnWidths: Readonly<Record<number, number>>
  readonly engine: GridEngineLike
  readonly freezeCols: number
  readonly freezeRows: number
  readonly frozenColumnWidth: number
  readonly frozenRowHeight: number
  readonly getHeaderCellLocalBounds: (col: number, row: number) => Rectangle | undefined
  readonly gridMetrics: GridMetrics
  readonly hostClientHeight: number
  readonly hostClientWidth: number
  readonly hostElement: HTMLDivElement | null
  readonly residentBodyPane: WorkbookRenderTilePaneState | null
  readonly residentHeaderItems: readonly Item[]
  readonly residentHeaderRegion: {
    readonly range: Pick<Rectangle, 'x' | 'y' | 'width' | 'height'>
    readonly tx: number
    readonly ty: number
    readonly freezeRows: number
    readonly freezeCols: number
  }
  readonly residentViewport: Viewport
  readonly rowHeights: Readonly<Record<number, number>>
  readonly sheetName: string
}): readonly GridHeaderPaneState[] {
  const {
    columnWidths,
    engine,
    freezeCols,
    freezeRows,
    frozenColumnWidth,
    frozenRowHeight,
    getHeaderCellLocalBounds,
    gridMetrics,
    hostClientHeight,
    hostClientWidth,
    hostElement,
    residentBodyPane,
    residentHeaderItems,
    residentHeaderRegion,
    residentViewport,
    rowHeights,
    sheetName,
  } = input
  const emptyGpuScene = useMemo<GridGpuScene>(() => ({ borderRects: [], fillRects: [] }), [])
  const emptyTextScene = useMemo<GridTextScene>(() => ({ items: [] }), [])

  const headerGpuScene = useMemo<GridGpuScene>(() => {
    if (!hostElement) {
      return emptyGpuScene
    }
    return buildGridGpuScene({
      contentMode: 'headers',
      engine,
      columnWidths,
      rowHeights,
      gridMetrics,
      gridSelection: STATIC_SCENE_GRID_SELECTION,
      activeHeaderDrag: null,
      hoveredCell: null,
      hoveredHeader: null,
      resizeGuideColumn: null,
      resizeGuideRow: null,
      selectedCell: STATIC_SCENE_SELECTED_CELL,
      selectionRange: null,
      sheetName,
      visibleItems: residentHeaderItems,
      visibleRegion: residentHeaderRegion,
      hostBounds: {
        left: 0,
        top: 0,
      },
      getCellBounds: getHeaderCellLocalBounds,
    })
  }, [
    columnWidths,
    emptyGpuScene,
    engine,
    getHeaderCellLocalBounds,
    gridMetrics,
    hostElement,
    rowHeights,
    sheetName,
    residentHeaderItems,
    residentHeaderRegion,
  ])

  const headerTextScene = useMemo<GridTextScene>(() => {
    if (!hostElement) {
      return emptyTextScene
    }
    return buildGridTextScene({
      contentMode: 'headers',
      engine,
      columnWidths,
      rowHeights,
      editingCell: null,
      gridMetrics,
      activeHeaderDrag: null,
      hoveredHeader: null,
      resizeGuideColumn: null,
      selectedCell: STATIC_SCENE_SELECTED_CELL,
      selectedCellSnapshot: null,
      selectionRange: null,
      sheetName,
      visibleItems: residentHeaderItems,
      visibleRegion: residentHeaderRegion,
      hostBounds: {
        left: 0,
        top: 0,
        width: 0,
        height: 0,
      },
      getCellBounds: getHeaderCellLocalBounds,
    })
  }, [
    columnWidths,
    emptyTextScene,
    engine,
    getHeaderCellLocalBounds,
    gridMetrics,
    hostElement,
    rowHeights,
    sheetName,
    residentHeaderItems,
    residentHeaderRegion,
  ])

  const headerPanes = useMemo(() => {
    noteHeaderPaneBuild()
    return buildHeaderPaneStates({
      gpuScene: headerGpuScene,
      textScene: headerTextScene,
      sheetName,
      residentViewport,
      freezeCols,
      freezeRows,
      hostWidth: hostClientWidth,
      hostHeight: hostClientHeight,
      gridMetrics,
      frozenColumnWidth,
      frozenRowHeight,
      residentBodyWidth: residentBodyPane?.surfaceSize.width ?? 0,
      residentBodyHeight: residentBodyPane?.surfaceSize.height ?? 0,
    })
  }, [
    frozenColumnWidth,
    frozenRowHeight,
    gridMetrics,
    headerGpuScene,
    headerTextScene,
    hostClientHeight,
    hostClientWidth,
    sheetName,
    residentViewport,
    freezeCols,
    freezeRows,
    residentBodyPane?.surfaceSize.height,
    residentBodyPane?.surfaceSize.width,
  ])

  return useMemo(
    () =>
      headerPanes.map((pane) =>
        pane.paneId === 'top-body'
          ? { ...pane, contentOffset: { x: residentBodyPane?.contentOffset.x ?? 0, y: 0 } }
          : pane.paneId === 'left-body'
            ? { ...pane, contentOffset: { x: 0, y: residentBodyPane?.contentOffset.y ?? 0 } }
            : pane,
      ),
    [headerPanes, residentBodyPane?.contentOffset.x, residentBodyPane?.contentOffset.y],
  )
}
