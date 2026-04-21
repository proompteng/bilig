import { memo, useEffect, useMemo, useRef, useState } from 'react'
import type { GridGeometrySnapshot } from '../gridGeometry.js'
import type { GridGpuScene } from '../gridGpuScene.js'
import type { GridTextScene } from '../gridTextScene.js'
import type { WorkbookGridScrollSnapshot, WorkbookGridScrollStore } from '../workbookGridScrollStore.js'
import { createGlyphAtlas } from '../renderer/glyph-atlas.js'
import { WorkbookPaneBufferCache } from '../renderer/pane-buffer-cache.js'
import type { WorkbookRenderPaneState } from '../renderer/pane-scene-types.js'
import { drawTypeGpuPanes } from '../renderer/typegpu-draw-pass.js'
import { GridRenderScheduler } from '../renderer/grid-render-scheduler.js'
import { syncTypeGpuPaneResources } from '../renderer/typegpu-resource-cache.js'
import {
  createTypeGpuRenderer,
  destroyTypeGpuRenderer,
  syncTypeGpuAtlasResources,
  type TypeGpuRendererArtifacts,
} from '../renderer/typegpu-renderer.js'
import { createTypeGpuSurfaceState, syncTypeGpuCanvasSurface } from '../renderer/typegpu-surface-manager.js'
import type { GridCameraStore } from './gridCameraStore.js'

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
  const artifactsRef = useRef<TypeGpuRendererArtifacts | null>(null)
  const paneBuffersRef = useRef(new WorkbookPaneBufferCache())
  const atlasRef = useRef(createGlyphAtlas())
  const surfaceStateRef = useRef(createTypeGpuSurfaceState())
  const renderSchedulerRef = useRef<GridRenderScheduler | null>(null)
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
    let effectArtifacts: TypeGpuRendererArtifacts | null = null
    const paneBufferCache = paneBuffersRef.current

    async function init() {
      if (!active || !canvasRef.current) {
        setWebGpuReady(false)
        return
      }
      const artifacts = await createTypeGpuRenderer(canvasRef.current)
      if (cancelled) {
        if (artifacts) {
          destroyTypeGpuRenderer(artifacts)
        }
        return
      }
      if (!artifacts) {
        setWebGpuReady(false)
        return
      }
      effectArtifacts = artifacts
      artifactsRef.current = artifacts
      setWebGpuReady(true)
    }

    void init()
    return () => {
      cancelled = true
      setWebGpuReady(false)
      paneBufferCache.dispose()
      if (effectArtifacts) {
        destroyTypeGpuRenderer(effectArtifacts)
      }
      artifactsRef.current = null
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
    const artifacts = artifactsRef.current
    const canvas = canvasRef.current
    if (!artifacts || !canvas) {
      return
    }
    syncTypeGpuCanvasSurface({
      artifacts,
      canvas,
      size: surfaceSize,
      state: surfaceStateRef.current,
    })
  }, [active, surfaceSize, webGpuReady])

  useEffect(() => {
    drawFrameRef.current = () => {
      if (!activeRef.current || !webGpuReadyRef.current) {
        return
      }

      const artifacts = artifactsRef.current
      const surface = surfaceSizeRef.current
      const resolvedPanePayloads = panePayloadsRef.current
      if (!artifacts || surface.width <= 0 || surface.height <= 0) {
        return
      }

      const atlas = atlasRef.current
      syncTypeGpuPaneResources({
        artifacts,
        atlas,
        paneBuffers: paneBuffersRef.current,
        panes: resolvedPanePayloads,
      })

      syncTypeGpuAtlasResources(artifacts, atlas)
      drawTypeGpuPanes({
        artifacts,
        paneBuffers: paneBuffersRef.current,
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
      renderSchedulerRef.current ??= new GridRenderScheduler()
      renderSchedulerRef.current.requestDraw(drawFrameRef.current)
    }
    return scrollTransformStore.subscribe(scheduleDraw)
  }, [active, scrollTransformStore])

  useEffect(() => {
    if (!active || !cameraStore) {
      return
    }
    const scheduleDraw = () => {
      renderSchedulerRef.current ??= new GridRenderScheduler()
      renderSchedulerRef.current.requestDraw(drawFrameRef.current)
    }
    return cameraStore.subscribe(scheduleDraw)
  }, [active, cameraStore])

  useEffect(() => {
    const canvas = canvasRef.current
    return () => {
      renderSchedulerRef.current?.cancel()
      renderSchedulerRef.current = null
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
