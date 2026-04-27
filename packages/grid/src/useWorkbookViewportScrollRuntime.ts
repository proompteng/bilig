import { useCallback, useEffect, useLayoutEffect, useRef, type Dispatch, type SetStateAction } from 'react'
import type { Viewport } from '@bilig/protocol'
import { createGridGeometrySnapshotFromAxes } from './gridGeometry.js'
import type { GridMetrics } from './gridMetrics.js'
import type { VisibleRegionState } from './gridPointer.js'
import { noteGridScrollInput } from './grid-render-counters.js'
import { resolveGridRenderScrollTransform, sameViewportBounds, sameVisibleRegionWindow } from './gridViewportController.js'
import type { GridAxisWorldIndex } from './gridAxisWorldIndex.js'
import type { Item } from './gridTypes.js'
import type { GridCameraStore } from './runtime/gridCameraStore.js'
import type { GridRuntimeHost } from './runtime/gridRuntimeHost.js'
import { viewportFromVisibleRegion } from './useGridCameraState.js'
import { resolveResidentViewport, hasSelectionTargetChanged } from './workbookGridViewport.js'
import type { WorkbookGridScrollSnapshot, WorkbookGridScrollStore } from './workbookGridScrollStore.js'

type MutableRef<T> = {
  current: T
}

function noteVisibleWindowChange(): void {
  if (typeof window === 'undefined') {
    return
  }
  ;(window as Window & { __biligScrollPerf?: { noteVisibleWindowChange?: () => void } }).__biligScrollPerf?.noteVisibleWindowChange?.()
}

export function shouldCommitWorkbookVisibleRegion(input: {
  readonly current: VisibleRegionState
  readonly next: VisibleRegionState
  readonly requiresLiveViewportState: boolean
}): boolean {
  const { current, next, requiresLiveViewportState } = input
  if (requiresLiveViewportState) {
    return !sameVisibleRegionWindow(current, next)
  }
  const currentResidentViewport = resolveResidentViewport(viewportFromVisibleRegion(current))
  const targetResidentViewport = resolveResidentViewport(viewportFromVisibleRegion(next))
  return !(
    current.freezeCols === next.freezeCols &&
    current.freezeRows === next.freezeRows &&
    sameViewportBounds(currentResidentViewport, targetResidentViewport)
  )
}

export function useWorkbookViewportScrollRuntime(input: {
  readonly columnAxis: GridAxisWorldIndex
  readonly freezeCols: number
  readonly freezeRows: number
  readonly gridCameraStore: GridCameraStore
  readonly gridMetrics: GridMetrics
  readonly gridRuntimeHost: GridRuntimeHost
  readonly hostElement: HTMLDivElement | null
  readonly liveVisibleRegionRef: MutableRef<VisibleRegionState>
  readonly onVisibleViewportChange?: ((viewport: Viewport) => void) | undefined
  readonly requiresLiveViewportState: boolean
  readonly restoreViewportTarget?:
    | {
        readonly token: number
        readonly viewport: Viewport
      }
    | undefined
  readonly rowAxis: GridAxisWorldIndex
  readonly scrollTransformRef: MutableRef<WorkbookGridScrollSnapshot>
  readonly scrollTransformStore: WorkbookGridScrollStore
  readonly scrollViewportRef: MutableRef<HTMLDivElement | null>
  readonly selectedCell: Item
  readonly setVisibleRegion: Dispatch<SetStateAction<VisibleRegionState>>
  readonly sheetName: string
  readonly sortedColumnWidthOverrides: readonly (readonly [number, number])[]
  readonly sortedRowHeightOverrides: readonly (readonly [number, number])[]
  readonly syncRuntimeAxes: () => void
  readonly viewport: Viewport
}): void {
  const {
    columnAxis,
    freezeCols,
    freezeRows,
    gridCameraStore,
    gridMetrics,
    gridRuntimeHost,
    hostElement,
    liveVisibleRegionRef,
    onVisibleViewportChange,
    requiresLiveViewportState,
    restoreViewportTarget,
    rowAxis,
    scrollTransformRef,
    scrollTransformStore,
    scrollViewportRef,
    selectedCell,
    setVisibleRegion,
    sheetName,
    sortedColumnWidthOverrides,
    sortedRowHeightOverrides,
    syncRuntimeAxes,
    viewport,
  } = input
  const autoScrollSelectionRef = useRef<{ sheetName: string; col: number; row: number } | null>(null)
  const restoredViewportTokenRef = useRef<number | null>(null)
  const scrollSyncFrameRef = useRef<number | null>(null)

  const syncVisibleRegion = useCallback(() => {
    const scrollViewport = scrollViewportRef.current
    if (!scrollViewport) {
      return
    }
    syncRuntimeAxes()
    const camera = gridRuntimeHost.updateCamera({
      scrollLeft: scrollViewport.scrollLeft,
      scrollTop: scrollViewport.scrollTop,
      viewportWidth: scrollViewport.clientWidth,
      viewportHeight: scrollViewport.clientHeight,
      dpr: window.devicePixelRatio || 1,
      freezeRows,
      freezeCols,
      gridMetrics,
    })
    gridCameraStore.setSnapshot(
      createGridGeometrySnapshotFromAxes({
        columns: columnAxis,
        dpr: window.devicePixelRatio || 1,
        freezeCols,
        freezeRows,
        gridMetrics,
        hostHeight: scrollViewport.clientHeight,
        hostWidth: scrollViewport.clientWidth,
        previousCamera: gridCameraStore.getSnapshot()?.camera ?? null,
        rows: rowAxis,
        scrollLeft: scrollViewport.scrollLeft,
        scrollTop: scrollViewport.scrollTop,
        seq: camera.seq,
        sheetName,
      }),
    )
    const next = camera.visibleRegion
    const { renderTx, renderTy } = resolveGridRenderScrollTransform({
      nextVisibleRegion: next,
      renderViewport: viewport,
      sortedColumnWidthOverrides,
      sortedRowHeightOverrides,
      defaultColumnWidth: gridMetrics.columnWidth,
      defaultRowHeight: gridMetrics.rowHeight,
    })
    scrollTransformRef.current = {
      renderTx,
      renderTy,
      scrollLeft: scrollViewport.scrollLeft,
      scrollTop: scrollViewport.scrollTop,
      tx: next.tx,
      ty: next.ty,
    }
    liveVisibleRegionRef.current = next
    scrollTransformStore.setSnapshot(scrollTransformRef.current)
    onVisibleViewportChange?.(viewportFromVisibleRegion(next))
    setVisibleRegion((current) => {
      if (!shouldCommitWorkbookVisibleRegion({ current, next, requiresLiveViewportState })) {
        return current
      }
      noteVisibleWindowChange()
      return next
    })
  }, [
    columnAxis,
    freezeCols,
    freezeRows,
    gridRuntimeHost,
    gridCameraStore,
    gridMetrics,
    liveVisibleRegionRef,
    onVisibleViewportChange,
    requiresLiveViewportState,
    rowAxis,
    scrollTransformRef,
    scrollTransformStore,
    scrollViewportRef,
    setVisibleRegion,
    sheetName,
    sortedColumnWidthOverrides,
    sortedRowHeightOverrides,
    syncRuntimeAxes,
    viewport,
  ])

  useEffect(() => {
    if (!requiresLiveViewportState) {
      return
    }
    setVisibleRegion((current) => {
      const next = liveVisibleRegionRef.current
      if (sameVisibleRegionWindow(current, next)) {
        return current
      }
      return next
    })
  }, [liveVisibleRegionRef, requiresLiveViewportState, setVisibleRegion])

  useLayoutEffect(() => {
    const scrollViewport = scrollViewportRef.current
    if (!scrollViewport) {
      return
    }

    syncVisibleRegion()
    const scheduleVisibleRegionSync = () => {
      if (scrollSyncFrameRef.current !== null) {
        return
      }
      scrollSyncFrameRef.current = window.requestAnimationFrame(() => {
        scrollSyncFrameRef.current = null
        syncVisibleRegion()
      })
    }
    const handleScroll = () => {
      noteGridScrollInput()
      syncVisibleRegion()
    }
    scrollViewport.addEventListener('scroll', handleScroll, { passive: true })
    const observer = new ResizeObserver(() => {
      scheduleVisibleRegionSync()
    })
    observer.observe(scrollViewport)
    return () => {
      if (scrollSyncFrameRef.current !== null) {
        window.cancelAnimationFrame(scrollSyncFrameRef.current)
        scrollSyncFrameRef.current = null
      }
      observer.disconnect()
      scrollViewport.removeEventListener('scroll', handleScroll)
    }
  }, [hostElement, scrollViewportRef, syncVisibleRegion])

  useLayoutEffect(() => {
    const scrollViewport = scrollViewportRef.current
    if (!scrollViewport) {
      return
    }
    const previousAutoScrollSelection = autoScrollSelectionRef.current
    const nextAutoScrollSelection = {
      sheetName,
      col: selectedCell[0],
      row: selectedCell[1],
    }
    if (!hasSelectionTargetChanged(previousAutoScrollSelection, nextAutoScrollSelection)) {
      return
    }
    autoScrollSelectionRef.current = nextAutoScrollSelection
    syncRuntimeAxes()
    const nextScrollPosition = gridRuntimeHost.resolveScrollForCellIntoView({
      cell: selectedCell,
      freezeRows,
      freezeCols,
      gridMetrics,
      scrollLeft: scrollViewport.scrollLeft,
      scrollTop: scrollViewport.scrollTop,
      viewportHeight: scrollViewport.clientHeight,
      viewportWidth: scrollViewport.clientWidth,
    })
    scrollViewport.scrollLeft = nextScrollPosition.scrollLeft
    scrollViewport.scrollTop = nextScrollPosition.scrollTop
  }, [freezeCols, freezeRows, gridRuntimeHost, gridMetrics, scrollViewportRef, selectedCell, sheetName, syncRuntimeAxes])

  useLayoutEffect(() => {
    const scrollViewport = scrollViewportRef.current
    if (!scrollViewport || !restoreViewportTarget) {
      return
    }
    if (restoredViewportTokenRef.current === restoreViewportTarget.token) {
      return
    }
    restoredViewportTokenRef.current = restoreViewportTarget.token
    syncRuntimeAxes()
    const { scrollLeft: nextScrollLeft, scrollTop: nextScrollTop } = gridRuntimeHost.resolveScrollPositionForViewport({
      viewport: restoreViewportTarget.viewport,
      freezeRows,
      freezeCols,
    })
    scrollViewport.scrollLeft = nextScrollLeft
    scrollViewport.scrollTop = nextScrollTop
  }, [freezeCols, freezeRows, gridRuntimeHost, restoreViewportTarget, scrollViewportRef, syncRuntimeAxes])
}
