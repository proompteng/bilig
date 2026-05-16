import { memo, useCallback, useEffect, useLayoutEffect, useRef, useSyncExternalStore } from 'react'
import type { GridGeometrySnapshot } from '../gridGeometry.js'
import type { GridHeaderPaneState } from '../gridHeaderPanes.js'
import type { GridCameraStore } from '../runtime/gridCameraStore.js'
import type { WorkbookGridScrollStore } from '../workbookGridScrollStore.js'
import { WorkbookPaneCanvasFallbackV3 } from './WorkbookPaneCanvasFallbackV3.js'
export { TYPEGPU_V3_ACTIVE_RESOURCE_DEFER_MS, GridDrawSchedulerV3, shouldDeferTypeGpuV3PreloadSync } from './draw-scheduler.js'
export { resolveTypeGpuV3DrawScrollSnapshot } from './workbook-pane-renderer-runtime.js'
import type { DynamicGridOverlayBatchV3 } from './dynamic-overlay-batch.js'
import type { WorkbookRenderTilePaneState } from './render-tile-pane-state.js'
import { WorkbookPaneRendererHostRuntimeV3 } from './workbook-pane-renderer-host-runtime.js'
import type { WorkbookPaneSurfaceBackendStatusV3 } from './workbook-pane-surface-runtime.js'

export interface WorkbookPaneRendererV3Props {
  readonly active: boolean
  readonly host: HTMLDivElement | null
  readonly geometry: GridGeometrySnapshot | null
  readonly cameraStore?: GridCameraStore | null
  readonly enableCanvasFallback?: boolean | undefined
  readonly headerPanes?: readonly GridHeaderPaneState[] | undefined
  readonly tilePanes: readonly WorkbookRenderTilePaneState[]
  readonly preloadTilePanes?: readonly WorkbookRenderTilePaneState[] | undefined
  readonly overlayBuilder?: ((geometry: GridGeometrySnapshot) => DynamicGridOverlayBatchV3 | null | undefined) | undefined
  readonly overlay?: DynamicGridOverlayBatchV3 | undefined
  readonly scrollTransformStore?: WorkbookGridScrollStore | null
}

export const WorkbookPaneRendererV3 = memo(function WorkbookPaneRendererV3({
  active,
  cameraStore = null,
  enableCanvasFallback = false,
  geometry,
  headerPanes = [],
  host,
  overlay,
  overlayBuilder,
  preloadTilePanes = [],
  scrollTransformStore = null,
  tilePanes,
}: WorkbookPaneRendererV3Props) {
  const hostRuntimeRef = useRef<WorkbookPaneRendererHostRuntimeV3 | null>(null)
  const hostRuntimeLifetimeRef = useRef(0)
  if (!hostRuntimeRef.current) {
    hostRuntimeRef.current = new WorkbookPaneRendererHostRuntimeV3()
  }
  const hostRuntime = hostRuntimeRef.current
  const backendStatus = useSyncExternalStore(
    hostRuntime.subscribeBackendStatus,
    hostRuntime.getBackendStatusSnapshot,
    hostRuntime.getBackendStatusSnapshot,
  )
  const frameProofStatus = useSyncExternalStore(
    hostRuntime.subscribeFrameProofStatus,
    hostRuntime.getFrameProofStatusSnapshot,
    hostRuntime.getFrameProofStatusSnapshot,
  )
  const hasPresentedFrame = useSyncExternalStore(
    hostRuntime.subscribeFrameProofStatus,
    hostRuntime.getHasPresentedFrameSnapshot,
    hostRuntime.getHasPresentedFrameSnapshot,
  )

  const setCanvasRef = useCallback(
    (canvas: HTMLCanvasElement | null) => {
      hostRuntime.setCanvas(canvas)
    },
    [hostRuntime],
  )
  useLayoutEffect(() => {
    hostRuntime.updateProps({
      active,
      cameraStore,
      geometry,
      headerPanes,
      host,
      overlay: overlay ?? null,
      overlayBuilder: overlayBuilder ?? null,
      preloadTilePanes,
      scrollTransformStore,
      tilePanes,
    })
  }, [
    active,
    cameraStore,
    geometry,
    headerPanes,
    host,
    hostRuntime,
    overlay,
    overlayBuilder,
    preloadTilePanes,
    scrollTransformStore,
    tilePanes,
  ])

  useEffect(() => {
    const lifetime = hostRuntimeLifetimeRef.current + 1
    hostRuntimeLifetimeRef.current = lifetime
    return () => {
      queueMicrotask(() => {
        if (hostRuntimeLifetimeRef.current !== lifetime) {
          return
        }
        hostRuntime.dispose()
        if (hostRuntimeRef.current === hostRuntime) {
          hostRuntimeRef.current = null
        }
      })
    }
  }, [hostRuntime])

  if (!active || !host) {
    return null
  }
  const showCanvasFallback = shouldMountWorkbookCanvasProofLayerV3({
    backendStatus,
    enableCanvasFallback,
    frameProofStatus,
    hasPresentedFrame,
    headerPaneCount: headerPanes.length,
    overlayRectCount: overlay?.rectCount ?? 0,
    tilePaneCount: tilePanes.length,
  })
  const showTypeGpuCanvas = backendStatus !== 'unavailable'
  const typeGpuCanvasOpacity = showCanvasFallback ? 0 : 1
  const tileSceneRevision = resolveWorkbookPaneTileSceneRevisionV3(tilePanes)
  const tileSceneCameraSeq = resolveWorkbookPaneTileSceneCameraSeqV3(tilePanes)
  const visibleRenderRevision = frameProofStatus === 'presented' ? tileSceneRevision : null
  const visibleRenderCameraSeq = frameProofStatus === 'presented' ? tileSceneCameraSeq : null
  const headerTextRunCount = headerPanes.reduce((total, pane) => total + pane.textRuns.length, 0)
  const tileTextRunCount = tilePanes.reduce((total, pane) => total + pane.tile.textRuns.length, 0)

  return (
    <>
      {showCanvasFallback ? (
        <WorkbookPaneCanvasFallbackV3
          active={active}
          cameraStore={cameraStore}
          geometry={geometry}
          headerPanes={headerPanes}
          host={host}
          overlay={overlay ?? null}
          overlayBuilder={overlayBuilder ?? null}
          scrollTransformStore={scrollTransformStore}
          tilePanes={tilePanes}
        />
      ) : null}
      {showTypeGpuCanvas ? (
        <canvas
          aria-hidden="true"
          className="pointer-events-none absolute inset-0 z-10"
          data-pane-renderer="workbook-pane-renderer-v3"
          data-renderer-mode="typegpu-v3"
          data-testid="grid-pane-renderer"
          data-v3-backend-status={backendStatus}
          data-v3-body-world-x={geometry?.camera.bodyWorldX ?? 0}
          data-v3-body-world-y={geometry?.camera.bodyWorldY ?? 0}
          data-v3-canvas-proof-layer={showCanvasFallback ? 'mounted' : 'not-mounted'}
          data-v3-frame-proof-status={frameProofStatus}
          data-v3-header-pane-count={headerPanes.length}
          data-v3-header-text-run-count={headerTextRunCount}
          data-v3-preload-pane-count={preloadTilePanes.length}
          data-v3-text-run-count={tileTextRunCount}
          data-v3-tile-scene-camera-seq={tileSceneCameraSeq ?? ''}
          data-v3-tile-scene-revision={tileSceneRevision ?? ''}
          data-v3-tile-pane-count={tilePanes.length}
          data-v3-visible-render-camera-seq={visibleRenderCameraSeq ?? ''}
          data-v3-visible-render-revision={visibleRenderRevision ?? ''}
          ref={setCanvasRef}
          style={{ backgroundColor: 'transparent', contain: 'strict', height: '100%', opacity: typeGpuCanvasOpacity, width: '100%' }}
        />
      ) : null}
    </>
  )
})

export function shouldMountWorkbookCanvasProofLayerV3(input: {
  readonly backendStatus: WorkbookPaneSurfaceBackendStatusV3
  readonly enableCanvasFallback?: boolean | undefined
  readonly frameProofStatus?: 'idle' | 'pending' | 'presented' | undefined
  readonly hasPresentedFrame?: boolean | undefined
  readonly headerPaneCount: number
  readonly overlayRectCount?: number | undefined
  readonly tilePaneCount: number
}): boolean {
  if (input.enableCanvasFallback || input.backendStatus !== 'ready') {
    return true
  }
  if (input.hasPresentedFrame) {
    return false
  }
  const hasVisiblePaneContent = input.tilePaneCount > 0 || input.headerPaneCount > 0 || (input.overlayRectCount ?? 0) > 0
  return hasVisiblePaneContent && input.frameProofStatus !== 'presented'
}

export function resolveWorkbookPaneTileSceneRevisionV3(tilePanes: readonly WorkbookRenderTilePaneState[]): number | null {
  return maxTilePaneField(tilePanes, (pane) => pane.tile.lastBatchId)
}

export function resolveWorkbookPaneTileSceneCameraSeqV3(tilePanes: readonly WorkbookRenderTilePaneState[]): number | null {
  return maxTilePaneField(tilePanes, (pane) => pane.tile.lastCameraSeq)
}

function maxTilePaneField(
  tilePanes: readonly WorkbookRenderTilePaneState[],
  readValue: (pane: WorkbookRenderTilePaneState) => number,
): number | null {
  let result: number | null = null
  for (const pane of tilePanes) {
    const value = readValue(pane)
    if (!Number.isFinite(value)) {
      continue
    }
    result = result === null ? value : Math.max(result, value)
  }
  return result
}
