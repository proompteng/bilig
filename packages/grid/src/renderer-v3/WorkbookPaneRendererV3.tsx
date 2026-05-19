import { memo, useCallback, useEffect, useLayoutEffect, useRef, useSyncExternalStore } from 'react'
import type { GridRenderRevisionSnapshot } from '../grid-engine.js'
import type { GridGeometrySnapshot } from '../gridGeometry.js'
import type { GridHeaderPaneState } from '../gridHeaderPanes.js'
import type { GridCameraStore } from '../runtime/gridCameraStore.js'
import type { WorkbookGridScrollStore } from '../workbookGridScrollStore.js'
import { WorkbookPaneCanvasFallbackV3 } from './WorkbookPaneCanvasFallbackV3.js'
export { TYPEGPU_V3_ACTIVE_RESOURCE_DEFER_MS, GridDrawSchedulerV3, shouldDeferTypeGpuV3PreloadSync } from './draw-scheduler.js'
export { resolveTypeGpuV3DrawScrollSnapshot } from './workbook-pane-renderer-runtime.js'
import type { DynamicGridOverlayBatchV3 } from './dynamic-overlay-batch.js'
import type { WorkbookRenderTilePaneState } from './render-tile-pane-state.js'
import { WorkbookPaneNativeTextLayerV3, type SuppressedNativeTextCellV3 } from './WorkbookPaneNativeTextLayerV3.js'
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
  readonly renderRevisionSnapshot?: GridRenderRevisionSnapshot | null | undefined
  readonly overlayBuilder?: ((geometry: GridGeometrySnapshot) => DynamicGridOverlayBatchV3 | null | undefined) | undefined
  readonly overlay?: DynamicGridOverlayBatchV3 | undefined
  readonly scrollTransformStore?: WorkbookGridScrollStore | null
  readonly suppressedTextCell?: SuppressedNativeTextCellV3 | null | undefined
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
  renderRevisionSnapshot = null,
  scrollTransformStore = null,
  suppressedTextCell = null,
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
  const frameProofSignature = useSyncExternalStore(
    hostRuntime.subscribeFrameProofStatus,
    hostRuntime.getFrameProofSignatureSnapshot,
    hostRuntime.getFrameProofSignatureSnapshot,
  )
  const hasPresentedFrame = useSyncExternalStore(
    hostRuntime.subscribeFrameProofStatus,
    hostRuntime.getHasPresentedFrameSnapshot,
    hostRuntime.getHasPresentedFrameSnapshot,
  )
  const presentedFrameProofSignature = useSyncExternalStore(
    hostRuntime.subscribeFrameProofStatus,
    hostRuntime.getPresentedFrameProofSignatureSnapshot,
    hostRuntime.getPresentedFrameProofSignatureSnapshot,
  )
  const presentedVisualFrame = useSyncExternalStore(
    hostRuntime.subscribeFrameProofStatus,
    hostRuntime.getPresentedVisualFrameSnapshot,
    hostRuntime.getPresentedVisualFrameSnapshot,
  )
  const hasPresentedVisibleFrame = presentedFrameProofSignature.length > 0

  const setCanvasRef = useCallback(
    (canvas: HTMLCanvasElement | null) => {
      hostRuntime.setCanvas(canvas)
    },
    [hostRuntime],
  )
  const headerTextRunCount = headerPanes.reduce((total, pane) => total + pane.textRuns.length, 0)
  const tileTextRunCount = tilePanes.reduce((total, pane) => total + pane.tile.textRuns.length, 0)
  const hasNativeTextRuns = headerTextRunCount + tileTextRunCount > 0
  const showNativeTextLayer = active && hasNativeTextRuns
  const hasVisiblePaneContent = hasWorkbookPaneVisibleContentV3({
    headerPaneCount: headerPanes.length,
    overlayRectCount: overlay?.rectCount ?? 0,
    tilePaneCount: tilePanes.length,
  })
  const showCanvasFallback = shouldMountWorkbookCanvasProofLayerV3({
    backendStatus,
    enableCanvasFallback,
    frameProofStatus,
    hasPresentedVisibleFrame,
    headerPaneCount: headerPanes.length,
    overlayRectCount: overlay?.rectCount ?? 0,
    tilePaneCount: tilePanes.length,
  })
  const showTypeGpuCanvas = backendStatus !== 'unavailable'
  const showCanvasGridFloor = showTypeGpuCanvas && hasVisiblePaneContent

  useLayoutEffect(() => {
    hostRuntime.updateProps({
      active,
      cameraStore,
      drawText: !showNativeTextLayer,
      geometry,
      headerPanes,
      host,
      overlay: overlay ?? null,
      overlayBuilder: overlayBuilder ?? null,
      preloadTilePanes,
      renderRevisionSnapshot,
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
    renderRevisionSnapshot,
    scrollTransformStore,
    showNativeTextLayer,
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
  const typeGpuCanvasOpacity = resolveWorkbookPaneTypeGpuCanvasOpacityV3({
    frameProofStatus,
    hasPresentedVisibleFrame,
    showCanvasFallback,
  })
  const tileSceneRevision = resolveWorkbookPaneTileSceneRevisionV3(tilePanes)
  const tileSceneCameraSeq = resolveWorkbookPaneTileSceneCameraSeqV3(tilePanes)
  const visibleRenderRevision = resolveWorkbookPanePresentedRevisionV3(frameProofStatus, tileSceneRevision)
  const visibleRenderCameraSeq = resolveWorkbookPanePresentedRevisionV3(frameProofStatus, tileSceneCameraSeq)
  const visibleProjectedRenderRevision = resolveWorkbookPanePresentedRevisionV3(frameProofStatus, renderRevisionSnapshot?.projectedRevision)
  const visibleLocalRenderRevision = resolveWorkbookPanePresentedRevisionV3(frameProofStatus, renderRevisionSnapshot?.localRevision)
  const visibleAuthoritativeRenderRevision = resolveWorkbookPanePresentedRevisionV3(
    frameProofStatus,
    renderRevisionSnapshot?.authoritativeRevision,
  )

  return (
    <>
      {showCanvasGridFloor ? (
        <WorkbookPaneCanvasFallbackV3
          active={active}
          cameraStore={cameraStore}
          drawText={false}
          geometry={geometry}
          headerPanes={headerPanes}
          host={host}
          layer="grid-floor"
          overlay={overlay ?? null}
          overlayBuilder={overlayBuilder ?? null}
          scrollTransformStore={scrollTransformStore}
          tilePanes={tilePanes}
        />
      ) : null}
      {showCanvasFallback ? (
        <WorkbookPaneCanvasFallbackV3
          active={active}
          cameraStore={cameraStore}
          drawText={!showNativeTextLayer}
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
          data-v3-draw-text={showNativeTextLayer ? 'false' : 'true'}
          data-v3-frame-proof-status={frameProofStatus}
          data-v3-frame-proof-signature={frameProofSignature}
          data-v3-has-presented-frame={hasPresentedFrame ? 'true' : 'false'}
          data-v3-has-presented-visible-frame={hasPresentedVisibleFrame ? 'true' : 'false'}
          data-v3-header-pane-count={headerPanes.length}
          data-v3-header-text-run-count={headerTextRunCount}
          data-v3-authoritative-render-revision={renderRevisionSnapshot?.authoritativeRevision ?? ''}
          data-v3-local-render-revision={renderRevisionSnapshot?.localRevision ?? ''}
          data-v3-presented-frame-proof-signature={presentedFrameProofSignature}
          data-v3-presented-camera-seq={presentedVisualFrame?.cameraSeq ?? ''}
          data-v3-presented-overlay-camera-seq={presentedVisualFrame?.overlayCameraSeq ?? ''}
          data-v3-presented-overlay-seq={presentedVisualFrame?.overlaySeq ?? ''}
          data-v3-presented-render-tx={presentedVisualFrame?.scrollSnapshot.renderTx ?? presentedVisualFrame?.scrollSnapshot.tx ?? ''}
          data-v3-presented-render-ty={presentedVisualFrame?.scrollSnapshot.renderTy ?? presentedVisualFrame?.scrollSnapshot.ty ?? ''}
          data-v3-presented-scroll-left={presentedVisualFrame?.scrollSnapshot.scrollLeft ?? ''}
          data-v3-presented-scroll-top={presentedVisualFrame?.scrollSnapshot.scrollTop ?? ''}
          data-v3-preload-pane-count={preloadTilePanes.length}
          data-v3-projected-render-revision={renderRevisionSnapshot?.projectedRevision ?? ''}
          data-v3-text-run-count={tileTextRunCount}
          data-v3-tile-scene-camera-seq={tileSceneCameraSeq ?? ''}
          data-v3-tile-scene-revision={tileSceneRevision ?? ''}
          data-v3-tile-pane-count={tilePanes.length}
          data-v3-visible-authoritative-render-revision={visibleAuthoritativeRenderRevision ?? ''}
          data-v3-visible-local-render-revision={visibleLocalRenderRevision ?? ''}
          data-v3-visible-projected-render-revision={visibleProjectedRenderRevision ?? ''}
          data-v3-visible-render-camera-seq={visibleRenderCameraSeq ?? ''}
          data-v3-visible-render-revision={visibleRenderRevision ?? ''}
          ref={setCanvasRef}
          style={{ backgroundColor: 'transparent', contain: 'strict', height: '100%', opacity: typeGpuCanvasOpacity, width: '100%' }}
        />
      ) : null}
      {showNativeTextLayer ? (
        <WorkbookPaneNativeTextLayerV3
          active={active}
          cameraStore={cameraStore}
          geometry={geometry}
          headerPanes={headerPanes}
          presentedScrollSnapshot={presentedVisualFrame?.scrollSnapshot ?? null}
          scrollTransformStore={scrollTransformStore}
          suppressedTextCell={suppressedTextCell}
          tilePanes={tilePanes}
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
  readonly hasPresentedVisibleFrame?: boolean | undefined
  readonly headerPaneCount: number
  readonly overlayRectCount?: number | undefined
  readonly tilePaneCount: number
}): boolean {
  if (input.enableCanvasFallback) {
    return true
  }
  if (input.backendStatus !== 'ready') {
    return true
  }
  const hasVisiblePaneContent = hasWorkbookPaneVisibleContentV3(input)
  if (input.frameProofStatus === 'pending') {
    return hasVisiblePaneContent
  }
  if (input.frameProofStatus === 'presented') {
    return false
  }
  if (input.hasPresentedVisibleFrame || input.hasPresentedFrame) {
    return false
  }
  return hasVisiblePaneContent
}

export function resolveWorkbookPaneTypeGpuCanvasOpacityV3(input: {
  readonly frameProofStatus: 'idle' | 'pending' | 'presented'
  readonly hasPresentedVisibleFrame: boolean
  readonly showCanvasFallback: boolean
}): number {
  if (!input.showCanvasFallback) {
    return 1
  }
  if (input.frameProofStatus === 'pending' && input.hasPresentedVisibleFrame) {
    return 1
  }
  return 0
}

export function hasWorkbookPaneVisibleContentV3(input: {
  readonly headerPaneCount: number
  readonly overlayRectCount?: number | undefined
  readonly tilePaneCount: number
}): boolean {
  return input.tilePaneCount > 0 || input.headerPaneCount > 0 || (input.overlayRectCount ?? 0) > 0
}

export function resolveWorkbookPaneTileSceneRevisionV3(tilePanes: readonly WorkbookRenderTilePaneState[]): number | null {
  return maxTilePaneField(tilePanes, (pane) => pane.tile.lastBatchId)
}

export function resolveWorkbookPanePresentedRevisionV3(
  frameProofStatus: 'idle' | 'pending' | 'presented',
  revision: number | null | undefined,
): number | null {
  return frameProofStatus === 'presented' && revision !== null && revision !== undefined ? revision : null
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
