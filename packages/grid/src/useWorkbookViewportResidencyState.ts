import { useCallback, useEffect, useMemo, useState } from 'react'
import type { GridEngineLike } from './grid-engine.js'
import type { VisibleRegionState } from './gridPointer.js'
import type { GridRuntimeHost } from './runtime/gridRuntimeHost.js'
import type { GridResidentHeaderRegion, GridViewportResidencyState } from './runtime/gridViewportResidencyRuntime.js'

export type WorkbookResidentHeaderRegion = GridResidentHeaderRegion
export type WorkbookViewportResidencyState = GridViewportResidencyState

export function useWorkbookViewportResidencyState(input: {
  readonly engine: GridEngineLike
  readonly freezeCols: number
  readonly freezeRows: number
  readonly gridRuntimeHost: GridRuntimeHost
  readonly sheetName: string
  readonly shouldUseRemoteRenderTileSource: boolean
  readonly visibleRegion: VisibleRegionState
}): GridViewportResidencyState {
  const { engine, freezeCols, freezeRows, gridRuntimeHost, sheetName, shouldUseRemoteRenderTileSource, visibleRegion } = input
  const [sceneRevision, setSceneRevision] = useState(0)
  const state = useMemo(
    () =>
      gridRuntimeHost.resolveViewportResidency({
        freezeCols,
        freezeRows,
        sceneRevision,
        visibleRegion,
      }),
    [freezeCols, freezeRows, gridRuntimeHost, sceneRevision, visibleRegion],
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
