import { memo, useEffect, useMemo, useRef, useState } from 'react'
import type { GridGpuScene } from '../gridGpuScene.js'
import type { GridTextScene } from '../gridTextScene.js'
import type { WorkbookRenderPaneState } from './pane-scene-types.js'
import { WorkbookPaneBufferCache } from './pane-buffer-cache.js'
import { createGlyphAtlas } from './glyph-atlas.js'
import type { WorkbookGridScrollStore } from '../workbookGridScrollStore.js'
import {
  createTypeGpuRenderer,
  destroyTypeGpuRenderer,
  syncTypeGpuAtlasResources,
  type TypeGpuRendererArtifacts,
} from './typegpu-renderer.js'
import { drawTypeGpuPanes } from './typegpu-draw-pass.js'
import { GridRenderScheduler } from './grid-render-scheduler.js'
import { syncTypeGpuPaneResources } from './typegpu-resource-cache.js'
import { createTypeGpuSurfaceState, syncTypeGpuCanvasSurface } from './typegpu-surface-manager.js'

interface WorkbookPaneRendererProps {
  readonly active: boolean
  readonly host: HTMLDivElement | null
  readonly panes: readonly WorkbookRenderPaneState[]
  readonly overlay?: {
    readonly gpuScene: GridGpuScene
    readonly textScene: GridTextScene
  }
  readonly scrollTransformStore?: WorkbookGridScrollStore | null
  readonly onActiveChange?: ((active: boolean) => void) | undefined
}

interface SurfaceSize {
  readonly width: number
  readonly height: number
  readonly pixelWidth: number
  readonly pixelHeight: number
  readonly dpr: number
}

function resolveSurfaceSize(host: HTMLElement): SurfaceSize {
  const width = Math.max(0, Math.floor(host.clientWidth))
  const height = Math.max(0, Math.floor(host.clientHeight))
  const dpr = Math.max(1, window.devicePixelRatio || 1)
  return {
    width,
    height,
    pixelWidth: Math.max(1, Math.floor(width * dpr)),
    pixelHeight: Math.max(1, Math.floor(height * dpr)),
    dpr,
  }
}

export const WorkbookPaneRenderer = memo(function WorkbookPaneRenderer({
  active,
  host,
  panes,
  overlay,
  scrollTransformStore = null,
  onActiveChange,
}: WorkbookPaneRendererProps) {
  const canvasRef = useRef<HTMLCanvasElement | null>(null)
  const artifactsRef = useRef<TypeGpuRendererArtifacts | null>(null)
  const paneBuffersRef = useRef(new WorkbookPaneBufferCache())
  const atlasRef = useRef(createGlyphAtlas())
  const surfaceStateRef = useRef(createTypeGpuSurfaceState())
  const renderSchedulerRef = useRef<GridRenderScheduler | null>(null)
  const drawFrameRef = useRef<() => void>(() => {})
  const [webGpuReady, setWebGpuReady] = useState(false)
  const [surfaceSize, setSurfaceSize] = useState<SurfaceSize>({ width: 0, height: 0, pixelWidth: 0, pixelHeight: 0, dpr: 1 })

  const activeRef = useRef(active)
  const webGpuReadyRef = useRef(webGpuReady)
  const surfaceSizeRef = useRef(surfaceSize)
  const panePayloadsRef = useRef<readonly WorkbookRenderPaneState[]>([])
  const scrollTransformStoreRef = useRef<WorkbookGridScrollStore | null>(scrollTransformStore)

  useEffect(() => {
    onActiveChange?.(webGpuReady)
  }, [onActiveChange, webGpuReady])

  useEffect(() => {
    if (!host) return
    const update = () => {
      const size = resolveSurfaceSize(host)
      setSurfaceSize(size)
      surfaceSizeRef.current = size
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
        paneId: 'overlay',
        generation: -1,
        frame: { x: 0, y: 0, width: surfaceSize.width, height: surfaceSize.height },
        surfaceSize: { width: surfaceSize.width, height: surfaceSize.height },
        contentOffset: { x: 0, y: 0 },
        scrollAxes: { x: false, y: false },
        gpuScene: overlay.gpuScene,
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
      const canvas = canvasRef.current
      const surface = surfaceSizeRef.current
      const resolvedPanePayloads = panePayloadsRef.current
      if (!artifacts || !canvas || surface.width <= 0 || surface.height <= 0) {
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
      const scrollSnapshot = scrollTransformStoreRef.current?.getSnapshot() ?? { tx: 0, ty: 0 }
      drawTypeGpuPanes({
        artifacts,
        paneBuffers: paneBuffersRef.current,
        panes: resolvedPanePayloads,
        scrollSnapshot,
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

  if (!host || !active) return null
  return (
    <canvas
      aria-hidden="true"
      className="pointer-events-none absolute inset-0 z-10"
      data-pane-renderer="workbook-pane-renderer"
      data-testid="grid-pane-renderer"
      ref={canvasRef}
      style={{ contain: 'strict' }}
    />
  )
})
