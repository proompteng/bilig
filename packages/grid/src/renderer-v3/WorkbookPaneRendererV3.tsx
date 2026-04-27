import { memo, useEffect, useRef, useState } from 'react'
import type { GridGeometrySnapshot } from '../gridGeometry.js'
import type { GridHeaderPaneState } from '../gridHeaderPanes.js'
import type { GridCameraStore } from '../runtime/gridCameraStore.js'
import type { WorkbookGridScrollStore } from '../workbookGridScrollStore.js'
export { TYPEGPU_V3_ACTIVE_RESOURCE_DEFER_MS, GridDrawSchedulerV3, shouldDeferTypeGpuV3PreloadSync } from './draw-scheduler.js'
export { resolveTypeGpuV3DrawScrollSnapshot } from './workbook-pane-renderer-runtime.js'
import type { DynamicGridOverlayBatchV3 } from './dynamic-overlay-batch.js'
import type { WorkbookRenderTilePaneState } from './render-tile-pane-state.js'
import { WorkbookPaneRendererRuntimeV3 } from './workbook-pane-renderer-runtime.js'
import { EMPTY_WORKBOOK_PANE_SURFACE_SNAPSHOT_V3, WorkbookPaneSurfaceRuntimeV3 } from './workbook-pane-surface-runtime.js'

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
  const canvasRef = useRef<HTMLCanvasElement | null>(null)
  const rendererRuntimeRef = useRef<WorkbookPaneRendererRuntimeV3 | null>(null)
  const surfaceRuntimeRef = useRef<WorkbookPaneSurfaceRuntimeV3 | null>(null)
  const [surfaceSnapshot, setSurfaceSnapshot] = useState(EMPTY_WORKBOOK_PANE_SURFACE_SNAPSHOT_V3)
  if (!rendererRuntimeRef.current) {
    rendererRuntimeRef.current = new WorkbookPaneRendererRuntimeV3()
  }
  if (!surfaceRuntimeRef.current) {
    surfaceRuntimeRef.current = new WorkbookPaneSurfaceRuntimeV3()
  }
  const rendererRuntime = rendererRuntimeRef.current
  const surfaceRuntime = surfaceRuntimeRef.current

  useEffect(() => {
    return surfaceRuntime.subscribe(setSurfaceSnapshot)
  }, [surfaceRuntime])

  useEffect(() => {
    surfaceRuntime.setHost(host)
    return () => surfaceRuntime.setHost(null)
  }, [host, surfaceRuntime])

  useEffect(() => {
    surfaceRuntime.setActive(active)
    return () => surfaceRuntime.setActive(false)
  }, [active, surfaceRuntime])

  useEffect(() => {
    surfaceRuntime.setCanvas(active && host ? canvasRef.current : null)
    return () => surfaceRuntime.setCanvas(null)
  }, [active, host, surfaceRuntime])

  useEffect(() => {
    rendererRuntime.updateState({
      active,
      backend: surfaceSnapshot.backend,
      cameraStore,
      geometry,
      headerPanes,
      overlay: overlay ?? null,
      overlayBuilder: overlayBuilder ?? null,
      preloadTilePanes,
      scrollTransformStore,
      surface: surfaceSnapshot.surface,
      tilePanes,
      webGpuReady: surfaceSnapshot.webGpuReady,
    })
    rendererRuntime.drawNow()
    rendererRuntime.requestDraw()
  }, [
    active,
    cameraStore,
    geometry,
    headerPanes,
    overlay,
    overlayBuilder,
    preloadTilePanes,
    rendererRuntime,
    scrollTransformStore,
    surfaceSnapshot,
    tilePanes,
  ])

  useEffect(() => {
    const canvas = canvasRef.current
    return () => {
      surfaceRuntime.dispose()
      rendererRuntime.dispose()
      if (canvas) {
        canvas.width = 0
        canvas.height = 0
      }
    }
  }, [rendererRuntime, surfaceRuntime])

  if (!active || !host) {
    return null
  }

  return (
    <canvas
      aria-hidden="true"
      className="pointer-events-none absolute inset-0 z-10"
      data-pane-renderer="workbook-pane-renderer-v3"
      data-renderer-mode="typegpu-v3"
      data-testid="grid-pane-renderer"
      data-v3-body-world-x={geometry?.camera.bodyWorldX ?? 0}
      data-v3-body-world-y={geometry?.camera.bodyWorldY ?? 0}
      ref={canvasRef}
      style={{ contain: 'strict' }}
    />
  )
})
