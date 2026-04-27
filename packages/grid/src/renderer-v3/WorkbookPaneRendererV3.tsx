import { memo, useEffect, useRef, useState } from 'react'
import type { GridGeometrySnapshot } from '../gridGeometry.js'
import type { GridHeaderPaneState } from '../gridHeaderPanes.js'
import type { GridCameraStore } from '../renderer-v2/gridCameraStore.js'
import { GridRenderLoop } from '../renderer-v2/gridRenderLoop.js'
import {
  createWorkbookTypeGpuBackend,
  destroyWorkbookTypeGpuBackend,
  drawWorkbookTypeGpuFrame,
  syncWorkbookTypeGpuSurface,
  type WorkbookTypeGpuBackend,
} from '../renderer-v2/workbook-typegpu-backend.js'
import type { WorkbookGridScrollSnapshot, WorkbookGridScrollStore } from '../workbookGridScrollStore.js'
import type { DynamicGridOverlayBatchV3 } from './dynamic-overlay-batch.js'
import type { WorkbookRenderTilePaneState } from './render-tile-pane-state.js'

export const TYPEGPU_V3_ACTIVE_RESOURCE_DEFER_MS = 48
const TYPEGPU_V3_IDLE_PRELOAD_RETRY_MS = 64

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

interface TypeGpuSurfaceSize {
  readonly width: number
  readonly height: number
  readonly pixelWidth: number
  readonly pixelHeight: number
  readonly dpr: number
}

function resolveSurfaceSize(host: HTMLElement): TypeGpuSurfaceSize {
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

export function resolveTypeGpuV3DrawScrollSnapshot(input: {
  readonly fallback: WorkbookGridScrollSnapshot
  readonly geometry: GridGeometrySnapshot | null
  readonly panes: readonly WorkbookRenderTilePaneState[]
}): WorkbookGridScrollSnapshot {
  const bodyPane = input.panes.find((pane) => pane.paneId === 'body')
  if (!input.geometry || !bodyPane) {
    return input.fallback
  }

  const bodyWorldX = input.geometry.camera.frozenWidth + (input.fallback.scrollLeft ?? input.geometry.camera.bodyScrollX)
  const bodyWorldY = input.geometry.camera.frozenHeight + (input.fallback.scrollTop ?? input.geometry.camera.bodyScrollY)
  return {
    ...input.fallback,
    renderTx: bodyWorldX - input.geometry.columns.offsetOf(bodyPane.viewport.colStart),
    renderTy: bodyWorldY - input.geometry.rows.offsetOf(bodyPane.viewport.rowStart),
  }
}

export function shouldDeferTypeGpuV3PreloadSync(input: {
  readonly now: number
  readonly lastScrollSignalAt: number
  readonly camera: {
    readonly updatedAt: number
    readonly velocityX: number
    readonly velocityY: number
  } | null
}): boolean {
  const hasMovingCamera =
    input.camera !== null &&
    Math.abs(input.camera.velocityX) + Math.abs(input.camera.velocityY) > 0.01 &&
    input.now - input.camera.updatedAt < TYPEGPU_V3_ACTIVE_RESOURCE_DEFER_MS
  return hasMovingCamera || input.now - input.lastScrollSignalAt < TYPEGPU_V3_ACTIVE_RESOURCE_DEFER_MS
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
  const backendRef = useRef<WorkbookTypeGpuBackend | null>(null)
  const renderLoopRef = useRef<GridRenderLoop | null>(null)
  const idlePreloadRetryRef = useRef<number | null>(null)
  const lastScrollSignalAtRef = useRef(0)
  const drawFrameRef = useRef<() => void>(() => {})
  const activeRef = useRef(active)
  const webGpuReadyRef = useRef(false)
  const surfaceSizeRef = useRef<TypeGpuSurfaceSize>({ dpr: 1, height: 0, pixelHeight: 0, pixelWidth: 0, width: 0 })
  const headerPanePayloadsRef = useRef<readonly GridHeaderPaneState[]>([])
  const tilePanePayloadsRef = useRef<readonly WorkbookRenderTilePaneState[]>([])
  const preloadTilePanePayloadsRef = useRef<readonly WorkbookRenderTilePaneState[]>([])
  const overlayBuilderRef = useRef<typeof overlayBuilder>(overlayBuilder)
  const overlayRef = useRef<typeof overlay>(overlay)
  const geometryRef = useRef<GridGeometrySnapshot | null>(geometry)
  const cameraStoreRef = useRef<GridCameraStore | null>(cameraStore)
  const scrollTransformStoreRef = useRef<WorkbookGridScrollStore | null>(scrollTransformStore)
  const [webGpuReady, setWebGpuReady] = useState(false)
  const [surfaceSize, setSurfaceSize] = useState<TypeGpuSurfaceSize>({ dpr: 1, height: 0, pixelHeight: 0, pixelWidth: 0, width: 0 })

  useEffect(() => {
    if (!host) {
      return
    }
    const update = () => {
      const size = resolveSurfaceSize(host)
      surfaceSizeRef.current = size
      setSurfaceSize(size)
    }
    update()
    const observer = new ResizeObserver(update)
    observer.observe(host)
    return () => observer.disconnect()
  }, [host])

  useEffect(() => {
    let cancelled = false
    let effectBackend: WorkbookTypeGpuBackend | null = null

    async function init() {
      if (!active || !canvasRef.current) {
        setWebGpuReady(false)
        return
      }
      const backend = await createWorkbookTypeGpuBackend(canvasRef.current)
      if (cancelled) {
        if (backend) {
          destroyWorkbookTypeGpuBackend(backend)
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
        destroyWorkbookTypeGpuBackend(effectBackend)
      }
      backendRef.current = null
    }
  }, [active])

  useEffect(() => {
    activeRef.current = active
  }, [active])

  useEffect(() => {
    webGpuReadyRef.current = webGpuReady
  }, [webGpuReady])

  useEffect(() => {
    surfaceSizeRef.current = surfaceSize
  }, [surfaceSize])

  useEffect(() => {
    geometryRef.current = geometry
  }, [geometry])

  useEffect(() => {
    cameraStoreRef.current = cameraStore
  }, [cameraStore])

  useEffect(() => {
    headerPanePayloadsRef.current = headerPanes
  }, [headerPanes])

  useEffect(() => {
    tilePanePayloadsRef.current = tilePanes
  }, [tilePanes])

  useEffect(() => {
    preloadTilePanePayloadsRef.current = preloadTilePanes
  }, [preloadTilePanes])

  useEffect(() => {
    overlayBuilderRef.current = overlayBuilder
  }, [overlayBuilder])

  useEffect(() => {
    overlayRef.current = overlay
  }, [overlay])

  useEffect(() => {
    scrollTransformStoreRef.current = scrollTransformStore
  }, [scrollTransformStore])

  useEffect(() => {
    if (!active || !webGpuReady) {
      return
    }
    const backend = backendRef.current
    const canvas = canvasRef.current
    if (!backend || !canvas) {
      return
    }
    syncWorkbookTypeGpuSurface({
      backend,
      canvas,
      size: surfaceSize,
    })
  }, [active, surfaceSize, webGpuReady])

  useEffect(() => {
    drawFrameRef.current = () => {
      if (!activeRef.current || !webGpuReadyRef.current) {
        return
      }

      const backend = backendRef.current
      const surface = surfaceSizeRef.current
      const headerPanePayloads = headerPanePayloadsRef.current
      const baseTilePanePayloads = tilePanePayloadsRef.current
      const preloadTilePanePayloads = preloadTilePanePayloadsRef.current
      if (!backend || surface.width <= 0 || surface.height <= 0) {
        return
      }
      const latestGeometry = cameraStoreRef.current?.getSnapshot() ?? geometryRef.current
      const camera = latestGeometry?.camera ?? null
      const now = performance.now()
      const deferPreloadSync = shouldDeferTypeGpuV3PreloadSync({
        camera,
        lastScrollSignalAt: lastScrollSignalAtRef.current,
        now,
      })
      if (deferPreloadSync) {
        if (idlePreloadRetryRef.current !== null) {
          window.clearTimeout(idlePreloadRetryRef.current)
        }
        idlePreloadRetryRef.current = window.setTimeout(() => {
          idlePreloadRetryRef.current = null
          renderLoopRef.current ??= new GridRenderLoop()
          renderLoopRef.current.requestDraw(drawFrameRef.current)
        }, TYPEGPU_V3_IDLE_PRELOAD_RETRY_MS)
      }
      const overlayBatch = overlayBuilderRef.current && latestGeometry ? overlayBuilderRef.current(latestGeometry) : overlayRef.current

      drawWorkbookTypeGpuFrame({
        backend,
        headerPanes: headerPanePayloads,
        overlay: overlayBatch ?? null,
        panes: [],
        preloadTilePanes: preloadTilePanePayloads,
        syncPreloadPanes: !deferPreloadSync,
        tilePanes: baseTilePanePayloads,
        scrollSnapshot: resolveTypeGpuV3DrawScrollSnapshot({
          fallback: scrollTransformStoreRef.current?.getSnapshot() ?? { tx: 0, ty: 0 },
          geometry: latestGeometry,
          panes: baseTilePanePayloads,
        }),
        surface,
      })
    }

    drawFrameRef.current()
    renderLoopRef.current ??= new GridRenderLoop()
    renderLoopRef.current.requestDraw(drawFrameRef.current)
  }, [active, headerPanes, overlay, overlayBuilder, preloadTilePanes, surfaceSize, tilePanes, webGpuReady])

  useEffect(() => {
    if (!active || !scrollTransformStore) {
      return
    }
    const scheduleDraw = () => {
      lastScrollSignalAtRef.current = performance.now()
      renderLoopRef.current ??= new GridRenderLoop()
      renderLoopRef.current.requestDraw(drawFrameRef.current)
    }
    return scrollTransformStore.subscribe(scheduleDraw)
  }, [active, scrollTransformStore])

  useEffect(() => {
    if (!active || !cameraStore) {
      return
    }
    const scheduleDraw = () => {
      lastScrollSignalAtRef.current = performance.now()
      renderLoopRef.current ??= new GridRenderLoop()
      renderLoopRef.current.requestDraw(drawFrameRef.current)
    }
    return cameraStore.subscribe(scheduleDraw)
  }, [active, cameraStore])

  useEffect(() => {
    const canvas = canvasRef.current
    return () => {
      renderLoopRef.current?.cancel()
      renderLoopRef.current = null
      if (idlePreloadRetryRef.current !== null) {
        window.clearTimeout(idlePreloadRetryRef.current)
        idlePreloadRetryRef.current = null
      }
      if (canvas) {
        canvas.width = 0
        canvas.height = 0
      }
    }
  }, [])

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
