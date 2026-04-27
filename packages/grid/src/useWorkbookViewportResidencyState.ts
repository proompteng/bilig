import { useCallback, useEffect, useMemo, useRef, useState } from 'react'
import type { GridEngineLike } from './grid-engine.js'
import type { VisibleRegionState } from './gridPointer.js'
import { WorkbookViewportResidencyRuntime, type WorkbookViewportResidencyState } from './workbookViewportResidencyRuntime.js'

export type { WorkbookResidentHeaderRegion, WorkbookViewportResidencyState } from './workbookViewportResidencyRuntime.js'

export function useWorkbookViewportResidencyState(input: {
  readonly engine: GridEngineLike
  readonly freezeCols: number
  readonly freezeRows: number
  readonly sheetName: string
  readonly shouldUseRemoteRenderTileSource: boolean
  readonly visibleRegion: VisibleRegionState
}): WorkbookViewportResidencyState {
  const { engine, freezeCols, freezeRows, sheetName, shouldUseRemoteRenderTileSource, visibleRegion } = input
  const [sceneRevision, setSceneRevision] = useState(0)
  const runtimeRef = useRef<WorkbookViewportResidencyRuntime | null>(null)
  if (!runtimeRef.current) {
    runtimeRef.current = new WorkbookViewportResidencyRuntime()
  }
  const runtime = runtimeRef.current
  const state = useMemo(
    () =>
      runtime.resolve({
        freezeCols,
        freezeRows,
        sceneRevision,
        visibleRegion,
      }),
    [freezeCols, freezeRows, runtime, sceneRevision, visibleRegion],
  )
  const { visibleAddresses } = state
  const invalidateScene = useCallback(() => {
    setSceneRevision((current) => current + 1)
  }, [])

  useEffect(() => {
    if (shouldUseRemoteRenderTileSource) {
      return
    }
    return engine.subscribeCells(sheetName, visibleAddresses, invalidateScene)
  }, [engine, invalidateScene, sheetName, shouldUseRemoteRenderTileSource, visibleAddresses])

  return state
}
