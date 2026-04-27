import { memo, useEffect, useRef, useState } from 'react'
import type { GridGeometrySnapshot } from '../gridGeometry.js'
import type { GridHeaderPaneState } from '../gridHeaderPanes.js'
import type { GridCameraStore } from '../runtime/gridCameraStore.js'
import {
  createWorkbookTypeGpuBackendV3,
  destroyWorkbookTypeGpuBackendV3,
  syncWorkbookTypeGpuSurfaceV3,
  type WorkbookTypeGpuBackendV3,
} from './typegpu-workbook-backend-v3.js'
import type { WorkbookGridScrollStore } from '../workbookGridScrollStore.js'
export { TYPEGPU_V3_ACTIVE_RESOURCE_DEFER_MS, GridDrawSchedulerV3, shouldDeferTypeGpuV3PreloadSync } from './draw-scheduler.js'
export { resolveTypeGpuV3DrawScrollSnapshot } from './workbook-pane-renderer-runtime.js'
import type { DynamicGridOverlayBatchV3 } from './dynamic-overlay-batch.js'
import type { WorkbookRenderTilePaneState } from './render-tile-pane-state.js'
import { WorkbookPaneRendererRuntimeV3, type TypeGpuSurfaceSizeV3 } from './workbook-pane-renderer-runtime.js'

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

function resolveSurfaceSize(host: HTMLElement): TypeGpuSurfaceSizeV3 {
  const width = Math.max(0, Math.floor(host.clientWidth))
  const height = Math.max(0, Math.floor(host.clientHeight))
  const dpr = Math.max(1, window.devicePixelRatio || 1)
  return {
    dpr,
    height,
    pixelHeight: Math.max(1, Math.floor(height * dpr)),
    pixelWidth: Math.max(1, Math.floor(width * dpr)),
    width,
  }
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
  const backendRef = useRef<WorkbookTypeGpuBackendV3 | null>(null)
  const rendererRuntimeRef = useRef<WorkbookPaneRendererRuntimeV3 | null>(null)
  const [webGpuReady, setWebGpuReady] = useState(false)
  const [surfaceSize, setSurfaceSize] = useState<TypeGpuSurfaceSizeV3>({ dpr: 1, height: 0, pixelHeight: 0, pixelWidth: 0, width: 0 })
  if (!rendererRuntimeRef.current) {
    rendererRuntimeRef.current = new WorkbookPaneRendererRuntimeV3()
  }
  const rendererRuntime = rendererRuntimeRef.current

  useEffect(() => {
    if (!host) {
      return
    }
    const update = () => {
      const size = resolveSurfaceSize(host)
      setSurfaceSize(size)
    }
    update()
    const observer = new ResizeObserver(update)
    observer.observe(host)
    return () => observer.disconnect()
  }, [host])

  useEffect(() => {
    let cancelled = false
    let effectBackend: WorkbookTypeGpuBackendV3 | null = null

    async function init() {
      if (!active || !canvasRef.current) {
        setWebGpuReady(false)
        return
      }
      const backend = await createWorkbookTypeGpuBackendV3(canvasRef.current)
      if (cancelled) {
        if (backend) {
          destroyWorkbookTypeGpuBackendV3(backend)
        }
        return
      }
      if (!backend) {
        setWebGpuReady(false)
        return
      }
      effectBackend = backend
      backendRef.current = backend
      setWebGpuReady(true)
    }

    void init()
    return () => {
      cancelled = true
      setWebGpuReady(false)
      if (effectBackend) {
        destroyWorkbookTypeGpuBackendV3(effectBackend)
      }
      backendRef.current = null
    }
  }, [active])

  useEffect(() => {
    if (!active || !webGpuReady) {
      return
    }
    const backend = backendRef.current
    const canvas = canvasRef.current
    if (!backend || !canvas) {
      return
    }
    syncWorkbookTypeGpuSurfaceV3({
      backend,
      canvas,
      size: surfaceSize,
    })
  }, [active, surfaceSize, webGpuReady])

  useEffect(() => {
    rendererRuntime.updateState({
      active,
      backend: backendRef.current,
      cameraStore,
      geometry,
      headerPanes,
      overlay: overlay ?? null,
      overlayBuilder: overlayBuilder ?? null,
      preloadTilePanes,
      scrollTransformStore,
      surface: surfaceSize,
      tilePanes,
      webGpuReady,
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
    surfaceSize,
    tilePanes,
    webGpuReady,
  ])

  useEffect(() => {
    if (!active || !scrollTransformStore) {
      return
    }
    const scheduleDraw = () => {
      rendererRuntime.noteInputSignalAndRequestDraw()
    }
    return scrollTransformStore.subscribe(scheduleDraw)
  }, [active, rendererRuntime, scrollTransformStore])

  useEffect(() => {
    if (!active || !cameraStore) {
      return
    }
    const scheduleDraw = () => {
      rendererRuntime.noteInputSignalAndRequestDraw()
    }
    return cameraStore.subscribe(scheduleDraw)
  }, [active, cameraStore, rendererRuntime])

  useEffect(() => {
    const canvas = canvasRef.current
    return () => {
      rendererRuntime.dispose()
      if (canvas) {
        canvas.width = 0
        canvas.height = 0
      }
    }
  }, [rendererRuntime])

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
