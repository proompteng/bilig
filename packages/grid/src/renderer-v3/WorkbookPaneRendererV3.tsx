import { memo, useCallback, useEffect, useMemo, useRef } from 'react'
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

export interface WorkbookPaneRendererV3Props {
  readonly active: boolean
  readonly host: HTMLDivElement | null
  readonly geometry: GridGeometrySnapshot | null
  readonly cameraStore?: GridCameraStore | null
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
  if (!hostRuntimeRef.current) {
    hostRuntimeRef.current = new WorkbookPaneRendererHostRuntimeV3()
  }
  const hostRuntime = hostRuntimeRef.current

  const setCanvasRef = useCallback(
    (canvas: HTMLCanvasElement | null) => {
      hostRuntime.setCanvas(canvas)
    },
    [hostRuntime],
  )
  const fallbackOverlay = useMemo(() => {
    if (overlay) {
      return overlay
    }
    return overlayBuilder && geometry ? (overlayBuilder(geometry) ?? null) : null
  }, [geometry, overlay, overlayBuilder])

  useEffect(() => {
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
    return () => {
      hostRuntime.dispose()
    }
  }, [hostRuntime])

  if (!active || !host) {
    return null
  }

  return (
    <>
      <WorkbookPaneCanvasFallbackV3
        active={active}
        geometry={geometry}
        headerPanes={headerPanes}
        host={host}
        overlay={fallbackOverlay}
        scrollTransformStore={scrollTransformStore}
        tilePanes={tilePanes}
      />
      <canvas
        aria-hidden="true"
        className="pointer-events-none absolute inset-0 z-10"
        data-pane-renderer="workbook-pane-renderer-v3"
        data-renderer-mode="typegpu-v3"
        data-testid="grid-pane-renderer"
        data-v3-body-world-x={geometry?.camera.bodyWorldX ?? 0}
        data-v3-body-world-y={geometry?.camera.bodyWorldY ?? 0}
        data-v3-header-pane-count={headerPanes.length}
        data-v3-preload-pane-count={preloadTilePanes.length}
        data-v3-tile-pane-count={tilePanes.length}
        ref={setCanvasRef}
        style={{ contain: 'strict', height: '100%', width: '100%' }}
      />
    </>
  )
})
