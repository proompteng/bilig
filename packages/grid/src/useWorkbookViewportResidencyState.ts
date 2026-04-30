import { useEffect, useMemo, useSyncExternalStore } from 'react'
import type { GridEngineLike } from './grid-engine.js'
import type { VisibleRegionState } from './gridPointer.js'
import type { GridRuntimeHost } from './runtime/gridRuntimeHost.js'

export type WorkbookViewportResidencyState = ReturnType<GridRuntimeHost['resolveViewportResidency']>
export type WorkbookResidentHeaderRegion = WorkbookViewportResidencyState['residentHeaderRegion']

export function useWorkbookViewportResidencyState(input: {
  readonly engine: GridEngineLike
  readonly freezeCols: number
  readonly freezeRows: number
  readonly gridRuntimeHost: GridRuntimeHost
  readonly sheetName: string
  readonly shouldUseRemoteRenderTileSource: boolean
  readonly visibleRegion: VisibleRegionState
}): WorkbookViewportResidencyState {
  const { engine, freezeCols, freezeRows, gridRuntimeHost, sheetName, shouldUseRemoteRenderTileSource, visibleRegion } = input
  const sceneRevision = useSyncExternalStore(
    (listener) => gridRuntimeHost.subscribeViewportResidencySceneRevision(listener),
    () => gridRuntimeHost.snapshotViewportResidencySceneRevision(),
    () => gridRuntimeHost.snapshotViewportResidencySceneRevision(),
  )
  const state = useMemo(() => {
    void sceneRevision
    return gridRuntimeHost.resolveViewportResidency({
      freezeCols,
      freezeRows,
      visibleRegion,
    })
  }, [freezeCols, freezeRows, gridRuntimeHost, sceneRevision, visibleRegion])
  const { visibleAddresses } = state

  useEffect(() => {
    return gridRuntimeHost.connectViewportResidencyInvalidation({
      engine,
      sheetName,
      shouldUseRemoteRenderTileSource,
      visibleAddresses,
    })
  }, [engine, gridRuntimeHost, sheetName, shouldUseRemoteRenderTileSource, visibleAddresses])

  return state
}
