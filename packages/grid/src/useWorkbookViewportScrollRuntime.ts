import { useEffect, useLayoutEffect, useRef, type Dispatch, type SetStateAction } from 'react'
import type { Viewport } from '@bilig/protocol'
import type { GridMetrics } from './gridMetrics.js'
import type { VisibleRegionState } from './gridPointer.js'
import type { GridAxisWorldIndex } from './gridAxisWorldIndex.js'
import type { Item } from './gridTypes.js'
import type { GridCameraStore } from './runtime/gridCameraStore.js'
import type { GridRuntimeHost } from './runtime/gridRuntimeHost.js'
import type { WorkbookGridScrollSnapshot, WorkbookGridScrollStore } from './workbookGridScrollStore.js'
import { WorkbookViewportScrollRuntime } from './workbookViewportScrollRuntime.js'

export { shouldCommitWorkbookVisibleRegion } from './workbookViewportScrollRuntime.js'

type MutableRef<T> = {
  current: T
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
  const runtimeRef = useRef<WorkbookViewportScrollRuntime | null>(null)
  if (!runtimeRef.current) {
    runtimeRef.current = new WorkbookViewportScrollRuntime()
  }
  const runtime = runtimeRef.current
  runtime.updateInput(input)

  useEffect(() => {
    return () => {
      runtime.dispose()
    }
  }, [runtime])

  useEffect(() => {
    runtime.syncLiveVisibleRegionForOverlay()
  }, [input.liveVisibleRegionRef, input.requiresLiveViewportState, input.setVisibleRegion, runtime])

  useLayoutEffect(() => {
    return runtime.attachScrollViewport()
  }, [input.hostElement, input.scrollViewportRef, runtime])

  useLayoutEffect(() => {
    runtime.autoScrollSelectionIntoView()
  }, [
    input.freezeCols,
    input.freezeRows,
    input.gridRuntimeHost,
    input.gridMetrics,
    input.scrollViewportRef,
    input.selectedCell,
    input.sheetName,
    input.syncRuntimeAxes,
    runtime,
  ])

  useLayoutEffect(() => {
    runtime.restoreViewportTarget()
  }, [
    input.freezeCols,
    input.freezeRows,
    input.gridRuntimeHost,
    input.restoreViewportTarget,
    input.scrollViewportRef,
    input.syncRuntimeAxes,
    runtime,
  ])
}
