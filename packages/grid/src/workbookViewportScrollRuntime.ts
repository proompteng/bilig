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
import { resolveResidentViewport } from './workbookGridViewport.js'
import type { WorkbookGridScrollSnapshot, WorkbookGridScrollStore } from './workbookGridScrollStore.js'

type MutableRef<T> = {
  current: T
}

type VisibleRegionUpdater = VisibleRegionState | ((current: VisibleRegionState) => VisibleRegionState)

type SetVisibleRegion = (updater: VisibleRegionUpdater) => void

export interface WorkbookViewportScrollRuntimeInput {
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
  readonly setVisibleRegion: SetVisibleRegion
  readonly sheetName: string
  readonly sortedColumnWidthOverrides: readonly (readonly [number, number])[]
  readonly sortedRowHeightOverrides: readonly (readonly [number, number])[]
  readonly syncRuntimeAxes: () => void
  readonly viewport: Viewport
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

function shouldNotifyWorkbookVisibleViewportChange(input: {
  readonly previous: Viewport | null
  readonly next: Viewport
  readonly requiresLiveViewportState: boolean
}): boolean {
  const { previous, next, requiresLiveViewportState } = input
  if (!previous) {
    return true
  }
  if (requiresLiveViewportState) {
    return !sameViewportBounds(previous, next)
  }
  return !sameViewportBounds(resolveResidentViewport(previous), resolveResidentViewport(next))
}

export class WorkbookViewportScrollRuntime {
  private autoScrollSelection: WorkbookAutoScrollSelectionTarget | null = null
  private input: WorkbookViewportScrollRuntimeInput | null = null
  private lastNotifiedViewport: Viewport | null = null
  private restoredViewportToken: number | null = null
  private scrollSyncFrame: number | null = null

  updateInput(input: WorkbookViewportScrollRuntimeInput): void {
    this.input = input
  }

  syncVisibleRegion(): void {
    const input = this.input
    const scrollViewport = input?.scrollViewportRef.current
    if (!input || !scrollViewport) {
      return
    }
    input.syncRuntimeAxes()
    const dpr = typeof window === 'undefined' ? 1 : window.devicePixelRatio || 1
    const camera = input.gridRuntimeHost.updateCamera({
      dpr,
      freezeCols: input.freezeCols,
      freezeRows: input.freezeRows,
      gridMetrics: input.gridMetrics,
      scrollLeft: scrollViewport.scrollLeft,
      scrollTop: scrollViewport.scrollTop,
      viewportHeight: scrollViewport.clientHeight,
      viewportWidth: scrollViewport.clientWidth,
    })
    input.gridCameraStore.setSnapshot(
      createGridGeometrySnapshotFromAxes({
        columns: input.columnAxis,
        dpr,
        freezeCols: input.freezeCols,
        freezeRows: input.freezeRows,
        gridMetrics: input.gridMetrics,
        hostHeight: scrollViewport.clientHeight,
        hostWidth: scrollViewport.clientWidth,
        previousCamera: input.gridCameraStore.getSnapshot()?.camera ?? null,
        rows: input.rowAxis,
        scrollLeft: scrollViewport.scrollLeft,
        scrollTop: scrollViewport.scrollTop,
        seq: camera.seq,
        sheetName: input.sheetName,
      }),
    )
    const next = camera.visibleRegion
    const { renderTx, renderTy } = resolveGridRenderScrollTransform({
      defaultColumnWidth: input.gridMetrics.columnWidth,
      defaultRowHeight: input.gridMetrics.rowHeight,
      nextVisibleRegion: next,
      renderViewport: input.viewport,
      sortedColumnWidthOverrides: input.sortedColumnWidthOverrides,
      sortedRowHeightOverrides: input.sortedRowHeightOverrides,
    })
    input.scrollTransformRef.current = {
      renderTx,
      renderTy,
      scrollLeft: scrollViewport.scrollLeft,
      scrollTop: scrollViewport.scrollTop,
      tx: next.tx,
      ty: next.ty,
    }
    input.liveVisibleRegionRef.current = next
    input.scrollTransformStore.setSnapshot(input.scrollTransformRef.current)
    const nextViewport = viewportFromVisibleRegion(next)
    if (
      shouldNotifyWorkbookVisibleViewportChange({
        next: nextViewport,
        previous: this.lastNotifiedViewport,
        requiresLiveViewportState: input.requiresLiveViewportState,
      })
    ) {
      this.lastNotifiedViewport = nextViewport
      input.onVisibleViewportChange?.(nextViewport)
    }
    input.setVisibleRegion((current) => {
      if (
        !shouldCommitWorkbookVisibleRegion({
          current,
          next,
          requiresLiveViewportState: input.requiresLiveViewportState,
        })
      ) {
        return current
      }
      noteVisibleWindowChange()
      return next
    })
  }

  attachScrollViewport(): () => void {
    const input = this.input
    const scrollViewport = input?.scrollViewportRef.current
    if (!input || !scrollViewport) {
      return () => undefined
    }

    this.syncVisibleRegion()
    const scheduleVisibleRegionSync = () => {
      if (this.scrollSyncFrame !== null || typeof window === 'undefined') {
        return
      }
      this.scrollSyncFrame = window.requestAnimationFrame(() => {
        this.scrollSyncFrame = null
        this.syncVisibleRegion()
      })
    }
    const handleScroll = () => {
      noteGridScrollInput()
      this.syncVisibleRegion()
    }
    scrollViewport.addEventListener('scroll', handleScroll, { passive: true })
    const observer = typeof ResizeObserver === 'undefined' ? null : new ResizeObserver(scheduleVisibleRegionSync)
    observer?.observe(scrollViewport)
    return () => {
      this.cancelScheduledSync()
      observer?.disconnect()
      scrollViewport.removeEventListener('scroll', handleScroll)
    }
  }

  syncLiveVisibleRegionForOverlay(): void {
    const input = this.input
    if (!input?.requiresLiveViewportState) {
      return
    }
    input.setVisibleRegion((current) => {
      const next = input.liveVisibleRegionRef.current
      if (sameVisibleRegionWindow(current, next)) {
        return current
      }
      return next
    })
  }

  autoScrollSelectionIntoView(): void {
    const input = this.input
    const scrollViewport = input?.scrollViewportRef.current
    if (!input || !scrollViewport) {
      return
    }
    input.syncRuntimeAxes()
    const axisSnapshot = input.gridRuntimeHost.snapshot()
    const nextAutoScrollSelection = {
      sheetName: input.sheetName,
      col: input.selectedCell[0],
      row: input.selectedCell[1],
      freezeCols: input.freezeCols,
      freezeRows: input.freezeRows,
      viewportWidth: scrollViewport.clientWidth,
      viewportHeight: scrollViewport.clientHeight,
      axisSeqX: axisSnapshot.axisSeqX,
      axisSeqY: axisSnapshot.axisSeqY,
    }
    if (!hasAutoScrollSelectionTargetChanged(this.autoScrollSelection, nextAutoScrollSelection)) {
      return
    }
    this.autoScrollSelection = nextAutoScrollSelection
    const nextScrollPosition = input.gridRuntimeHost.resolveScrollForCellIntoView({
      cell: input.selectedCell,
      freezeCols: input.freezeCols,
      freezeRows: input.freezeRows,
      gridMetrics: input.gridMetrics,
      scrollLeft: scrollViewport.scrollLeft,
      scrollTop: scrollViewport.scrollTop,
      viewportHeight: scrollViewport.clientHeight,
      viewportWidth: scrollViewport.clientWidth,
    })
    scrollViewport.scrollLeft = nextScrollPosition.scrollLeft
    scrollViewport.scrollTop = nextScrollPosition.scrollTop
  }

  restoreViewportTarget(): void {
    const input = this.input
    const scrollViewport = input?.scrollViewportRef.current
    if (!input || !scrollViewport || !input.restoreViewportTarget) {
      return
    }
    if (this.restoredViewportToken === input.restoreViewportTarget.token) {
      return
    }
    this.restoredViewportToken = input.restoreViewportTarget.token
    input.syncRuntimeAxes()
    const { scrollLeft, scrollTop } = input.gridRuntimeHost.resolveScrollPositionForViewport({
      freezeCols: input.freezeCols,
      freezeRows: input.freezeRows,
      viewport: input.restoreViewportTarget.viewport,
    })
    scrollViewport.scrollLeft = scrollLeft
    scrollViewport.scrollTop = scrollTop
  }

  dispose(): void {
    this.cancelScheduledSync()
    this.input = null
    this.lastNotifiedViewport = null
  }

  private cancelScheduledSync(): void {
    if (this.scrollSyncFrame === null || typeof window === 'undefined') {
      this.scrollSyncFrame = null
      return
    }
    window.cancelAnimationFrame(this.scrollSyncFrame)
    this.scrollSyncFrame = null
  }
}

interface WorkbookAutoScrollSelectionTarget {
  readonly sheetName: string
  readonly col: number
  readonly row: number
  readonly freezeCols: number
  readonly freezeRows: number
  readonly viewportWidth: number
  readonly viewportHeight: number
  readonly axisSeqX: number
  readonly axisSeqY: number
}

function hasAutoScrollSelectionTargetChanged(
  previous: WorkbookAutoScrollSelectionTarget | null,
  next: WorkbookAutoScrollSelectionTarget,
): boolean {
  return (
    previous === null ||
    previous.sheetName !== next.sheetName ||
    previous.col !== next.col ||
    previous.row !== next.row ||
    previous.freezeCols !== next.freezeCols ||
    previous.freezeRows !== next.freezeRows ||
    previous.viewportWidth !== next.viewportWidth ||
    previous.viewportHeight !== next.viewportHeight ||
    previous.axisSeqX !== next.axisSeqX ||
    previous.axisSeqY !== next.axisSeqY
  )
}
