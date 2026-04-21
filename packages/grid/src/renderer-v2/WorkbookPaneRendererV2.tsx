import { memo, useEffect, useMemo, useRef, useState } from 'react'
import type { GridGeometrySnapshot } from '../gridGeometry.js'
import type { GridGpuScene } from '../gridGpuScene.js'
import type { GridTextScene } from '../gridTextScene.js'
import type { WorkbookGridScrollSnapshot, WorkbookGridScrollStore } from '../workbookGridScrollStore.js'
import type { WorkbookRenderPaneState } from '../renderer/pane-scene-types.js'
import type { GridCameraStore } from './gridCameraStore.js'
import { GridRenderLoop } from './gridRenderLoop.js'
import {
  createWorkbookTypeGpuBackend,
  destroyWorkbookTypeGpuBackend,
  drawWorkbookTypeGpuFrame,
  syncWorkbookTypeGpuSurface,
  type WorkbookTypeGpuBackend,
} from './workbook-typegpu-backend.js'

export interface WorkbookPaneRendererV2Props {
  readonly active: boolean
  readonly host: HTMLDivElement | null
  readonly geometry: GridGeometrySnapshot | null
  readonly cameraStore?: GridCameraStore | null
  readonly panes: readonly WorkbookRenderPaneState[]
  readonly overlay?:
    | {
        readonly gpuScene: GridGpuScene
        readonly textScene: GridTextScene
      }
    | undefined
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

export function resolveTypeGpuV2DrawScrollSnapshot(input: {
  readonly fallback: WorkbookGridScrollSnapshot
  readonly geometry: GridGeometrySnapshot | null
  readonly panes: readonly WorkbookRenderPaneState[]
}): WorkbookGridScrollSnapshot {
  const bodyPane = input.panes.find((pane) => pane.paneId === 'body' && pane.viewport)
  if (!input.geometry || !bodyPane?.viewport) {
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

export const WorkbookPaneRendererV2 = memo(function WorkbookPaneRendererV2({
  active,
  cameraStore = null,
  geometry,
  host,
  overlay,
  panes,
  scrollTransformStore = null,
}: WorkbookPaneRendererV2Props) {
  const canvasRef = useRef<HTMLCanvasElement | null>(null)
  const backendRef = useRef<WorkbookTypeGpuBackend | null>(null)
  const renderLoopRef = useRef<GridRenderLoop | null>(null)
  const drawFrameRef = useRef<() => void>(() => {})
  const activeRef = useRef(active)
  const webGpuReadyRef = useRef(false)
  const surfaceSizeRef = useRef<TypeGpuSurfaceSize>({ dpr: 1, height: 0, pixelHeight: 0, pixelWidth: 0, width: 0 })
  const panePayloadsRef = useRef<readonly WorkbookRenderPaneState[]>([])
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

  const panePayloads = useMemo<readonly WorkbookRenderPaneState[]>(() => {
    const next: WorkbookRenderPaneState[] = [...panes]
    if (overlay) {
      next.push({
        contentOffset: { x: 0, y: 0 },
        frame: { x: 0, y: 0, width: surfaceSize.width, height: surfaceSize.height },
        generation: -1,
        gpuScene: overlay.gpuScene,
        paneId: 'overlay',
        scrollAxes: { x: false, y: false },
        surfaceSize: { width: surfaceSize.width, height: surfaceSize.height },
        textScene: overlay.textScene,
      })
    }
    return next
  }, [overlay, panes, surfaceSize.height, surfaceSize.width])

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
    panePayloadsRef.current = panePayloads
  }, [panePayloads])

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
      const resolvedPanePayloads = panePayloadsRef.current
      if (!backend || surface.width <= 0 || surface.height <= 0) {
        return
      }

      drawWorkbookTypeGpuFrame({
        backend,
        panes: resolvedPanePayloads,
        scrollSnapshot: resolveTypeGpuV2DrawScrollSnapshot({
          fallback: scrollTransformStoreRef.current?.getSnapshot() ?? { tx: 0, ty: 0 },
          geometry: cameraStoreRef.current?.getSnapshot() ?? geometryRef.current,
          panes: resolvedPanePayloads,
        }),
        surface,
      })
    }

    drawFrameRef.current()
  }, [active, panePayloads, surfaceSize, webGpuReady])

  useEffect(() => {
    if (!active || !scrollTransformStore) {
      return
    }
    const scheduleDraw = () => {
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
      data-pane-renderer="workbook-pane-renderer-v2"
      data-renderer-mode="typegpu-v2"
      data-testid="grid-pane-renderer"
      data-v2-body-world-x={geometry?.camera.bodyWorldX ?? 0}
      data-v2-body-world-y={geometry?.camera.bodyWorldY ?? 0}
      ref={canvasRef}
      style={{ contain: 'strict' }}
    />
  )
})
