import { memo, useEffect, useMemo, useRef, useState } from 'react'
import { parseGpuColor, type GridGpuScene } from '../gridGpuScene.js'
import type { GridTextScene } from '../gridTextScene.js'
import type { Rectangle } from '../gridTypes.js'
import type { WorkbookRenderPaneState } from './pane-scene-types.js'
import { WorkbookPaneBufferCache, type WorkbookPaneBufferEntry } from './pane-buffer-cache.js'
import { createGlyphAtlas } from './glyph-atlas.js'
import { buildTextDecorationRectsFromScene, buildTextQuadsFromScene, type TextDecorationRect } from './text-quad-buffer.js'
import type { WorkbookGridScrollStore } from '../workbookGridScrollStore.js'
import {
  WORKBOOK_RECT_INSTANCE_LAYOUT,
  WORKBOOK_TEXT_INSTANCE_LAYOUT,
  WORKBOOK_UNIT_QUAD_LAYOUT,
  createTypeGpuSurfaceBindGroup,
  createTypeGpuSurfaceUniform,
  createTypeGpuTextBindGroup,
  createTypeGpuRenderer,
  destroyTypeGpuRenderer,
  ensureTypeGpuVertexBuffer,
  syncTypeGpuAtlasResources,
  type TypeGpuRendererArtifacts,
  updateTypeGpuSurfaceUniform,
  writeTypeGpuVertexBuffer,
} from './typegpu-renderer.js'

const RECT_INSTANCE_FLOAT_COUNT = 20

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

function resolveClampedScissorRect(frame: Rectangle, surface: SurfaceSize): { x: number; y: number; width: number; height: number } | null {
  const dpr = surface.dpr
  const x0 = Math.max(0, Math.min(surface.pixelWidth, Math.floor(frame.x * dpr)))
  const y0 = Math.max(0, Math.min(surface.pixelHeight, Math.floor(frame.y * dpr)))
  const x1 = Math.max(x0, Math.min(surface.pixelWidth, Math.ceil((frame.x + frame.width) * dpr)))
  const y1 = Math.max(y0, Math.min(surface.pixelHeight, Math.ceil((frame.y + frame.height) * dpr)))
  if (x0 >= x1 || y0 >= y1) return null
  return { x: x0, y: y0, width: x1 - x0, height: y1 - y0 }
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

function buildRectInstanceData(input: {
  frame: Rectangle
  scene: GridGpuScene
  decorationRects?: readonly TextDecorationRect[]
}): Float32Array {
  const rects = input.scene.fillRects
  const borders = input.scene.borderRects
  const decorationRects = input.decorationRects ?? []
  const total = rects.length + borders.length + decorationRects.length
  const floats = new Float32Array(Math.max(1, total) * RECT_INSTANCE_FLOAT_COUNT)

  const clipX = 0
  const clipY = 0
  const clipX1 = input.frame.width
  const clipY1 = input.frame.height

  let offset = 0
  for (const rect of rects) {
    floats[offset + 0] = rect.x
    floats[offset + 1] = rect.y
    floats[offset + 2] = rect.width
    floats[offset + 3] = rect.height
    floats[offset + 4] = rect.color.r
    floats[offset + 5] = rect.color.g
    floats[offset + 6] = rect.color.b
    floats[offset + 7] = rect.color.a
    floats[offset + 8] = 0
    floats[offset + 9] = 0
    floats[offset + 10] = 0
    floats[offset + 11] = 0
    floats[offset + 12] = rect.color.a < 0.2 ? 2 : 0
    floats[offset + 13] = 0
    floats[offset + 14] = 0
    floats[offset + 15] = 0
    floats[offset + 16] = clipX
    floats[offset + 17] = clipY
    floats[offset + 18] = clipX1
    floats[offset + 19] = clipY1
    offset += RECT_INSTANCE_FLOAT_COUNT
  }
  for (const rect of borders) {
    floats[offset + 0] = rect.x
    floats[offset + 1] = rect.y
    floats[offset + 2] = rect.width
    floats[offset + 3] = rect.height
    floats[offset + 4] = 0
    floats[offset + 5] = 0
    floats[offset + 6] = 0
    floats[offset + 7] = 0
    floats[offset + 8] = rect.color.r
    floats[offset + 9] = rect.color.g
    floats[offset + 10] = rect.color.b
    floats[offset + 11] = rect.color.a
    floats[offset + 12] = 0
    floats[offset + 13] = 1
    floats[offset + 14] = 0
    floats[offset + 15] = 0
    floats[offset + 16] = clipX
    floats[offset + 17] = clipY
    floats[offset + 18] = clipX1
    floats[offset + 19] = clipY1
    offset += RECT_INSTANCE_FLOAT_COUNT
  }

  for (const rect of decorationRects) {
    const color = parseGpuColor(rect.color)
    floats[offset + 0] = rect.x
    floats[offset + 1] = rect.y
    floats[offset + 2] = rect.width
    floats[offset + 3] = rect.height
    floats[offset + 4] = color.r
    floats[offset + 5] = color.g
    floats[offset + 6] = color.b
    floats[offset + 7] = color.a
    floats[offset + 8] = 0
    floats[offset + 9] = 0
    floats[offset + 10] = 0
    floats[offset + 11] = 0
    floats[offset + 12] = 0
    floats[offset + 13] = 0
    floats[offset + 14] = 0
    floats[offset + 15] = 0
    floats[offset + 16] = clipX
    floats[offset + 17] = clipY
    floats[offset + 18] = clipX1
    floats[offset + 19] = clipY1
    offset += RECT_INSTANCE_FLOAT_COUNT
  }

  return floats
}

function buildTextInstanceData(input: { textScene: GridTextScene; atlas: ReturnType<typeof createGlyphAtlas> }): {
  floats: Float32Array
  quadCount: number
} {
  return buildTextQuadsFromScene(input.textScene.items, input.atlas)
}

function resolvePaneOrigin(pane: WorkbookRenderPaneState): { x: number; y: number } {
  return {
    x: pane.frame.x,
    y: pane.frame.y,
  }
}

function resolvePaneRenderOffset(
  pane: WorkbookRenderPaneState,
  scrollSnapshot: { readonly tx: number; readonly ty: number },
): { x: number; y: number } {
  return {
    x: pane.contentOffset.x - (pane.scrollAxes.x ? scrollSnapshot.tx : 0),
    y: pane.contentOffset.y - (pane.scrollAxes.y ? scrollSnapshot.ty : 0),
  }
}

function ensurePaneSurfaceBindings(artifacts: TypeGpuRendererArtifacts, paneCache: WorkbookPaneBufferEntry): void {
  if (!paneCache.surfaceUniform) {
    paneCache.surfaceUniform = createTypeGpuSurfaceUniform(artifacts.root)
  }
  if (!paneCache.surfaceBindGroup) {
    paneCache.surfaceBindGroup = createTypeGpuSurfaceBindGroup(artifacts.root, paneCache.surfaceUniform)
  }

  if (!artifacts.atlasTexture) {
    paneCache.textBindGroup = null
    paneCache.textBindGroupAtlasVersion = -1
    return
  }

  if (!paneCache.textBindGroup || paneCache.textBindGroupAtlasVersion !== artifacts.atlasVersion) {
    paneCache.textBindGroup = createTypeGpuTextBindGroup(artifacts, paneCache.surfaceUniform)
    paneCache.textBindGroupAtlasVersion = artifacts.atlasVersion
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
  const drawFrameRef = useRef<() => void>(() => {})
  const scheduledDrawFrameRef = useRef<number | null>(null)
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

      if (canvas.width !== surface.pixelWidth) {
        canvas.width = surface.pixelWidth
      }
      if (canvas.height !== surface.pixelHeight) {
        canvas.height = surface.pixelHeight
      }
      canvas.style.width = `${surface.width}px`
      canvas.style.height = `${surface.height}px`
      artifacts.context.configure({
        alphaMode: 'premultiplied',
        device: artifacts.device,
        format: artifacts.format,
      })

      const atlas = atlasRef.current
      paneBuffersRef.current.pruneExcept(new Set(resolvedPanePayloads.map((pane) => pane.paneId)))

      resolvedPanePayloads.forEach((pane) => {
        const paneCache = paneBuffersRef.current.get(pane.paneId)
        const textSceneChanged = paneCache.textScene !== pane.textScene
        if (textSceneChanged) {
          paneCache.decorationRects = buildTextDecorationRectsFromScene(pane.textScene.items, atlas)
          const textPayload = buildTextInstanceData({
            textScene: pane.textScene,
            atlas,
          })
          const textBuffer = ensureTypeGpuVertexBuffer(
            artifacts.root,
            WORKBOOK_TEXT_INSTANCE_LAYOUT,
            paneCache.textBuffer,
            paneCache.textCapacity,
            textPayload.quadCount,
          )
          paneCache.textBuffer = textBuffer.buffer
          paneCache.textCapacity = textBuffer.capacity
          paneCache.textCount = textPayload.quadCount
          writeTypeGpuVertexBuffer(paneCache.textBuffer, textPayload.floats)
          paneCache.textScene = pane.textScene
        }

        if (paneCache.rectScene !== pane.gpuScene || textSceneChanged) {
          const decorationRects = paneCache.decorationRects ?? []
          const rectFloats = buildRectInstanceData({
            frame: pane.frame,
            scene: pane.gpuScene,
            ...(decorationRects.length > 0 ? { decorationRects } : {}),
          })
          const rectBuffer = ensureTypeGpuVertexBuffer(
            artifacts.root,
            WORKBOOK_RECT_INSTANCE_LAYOUT,
            paneCache.rectBuffer,
            paneCache.rectCapacity,
            pane.gpuScene.fillRects.length + pane.gpuScene.borderRects.length + decorationRects.length,
          )
          paneCache.rectBuffer = rectBuffer.buffer
          paneCache.rectCapacity = rectBuffer.capacity
          paneCache.rectCount = pane.gpuScene.fillRects.length + pane.gpuScene.borderRects.length + decorationRects.length
          writeTypeGpuVertexBuffer(paneCache.rectBuffer, rectFloats)
          paneCache.rectScene = pane.gpuScene
        }
      })

      syncTypeGpuAtlasResources(artifacts, atlas)
      const scrollSnapshot = scrollTransformStoreRef.current?.getSnapshot() ?? { tx: 0, ty: 0 }

      const commandEncoder = artifacts.device.createCommandEncoder()
      const pass = commandEncoder.beginRenderPass({
        colorAttachments: [
          {
            view: artifacts.context.getCurrentTexture().createView(),
            clearValue: { r: 0, g: 0, b: 0, a: 0 },
            loadOp: 'clear',
            storeOp: 'store',
          },
        ],
      })

      resolvedPanePayloads.forEach((pane) => {
        const paneCache = paneBuffersRef.current.get(pane.paneId)
        const scissorRect = resolveClampedScissorRect(pane.frame, surface)
        if (!scissorRect) {
          return
        }

        pass.setScissorRect(scissorRect.x, scissorRect.y, scissorRect.width, scissorRect.height)
        const paneOrigin = resolvePaneOrigin(pane)
        const paneRenderOffset = resolvePaneRenderOffset(pane, scrollSnapshot)

        if (paneCache.rectCount > 0 || paneCache.textCount > 0) {
          ensurePaneSurfaceBindings(artifacts, paneCache)
          updateTypeGpuSurfaceUniform(paneCache.surfaceUniform!, surface, paneOrigin, paneRenderOffset)
        }

        if (paneCache.rectCount > 0 && paneCache.rectBuffer && paneCache.surfaceBindGroup) {
          const rectRenderer = artifacts.rectPipeline.with(pass).with(paneCache.surfaceBindGroup)
          rectRenderer
            .with(WORKBOOK_UNIT_QUAD_LAYOUT, artifacts.quadBuffer)
            .with(WORKBOOK_RECT_INSTANCE_LAYOUT, paneCache.rectBuffer)
            .draw(6, paneCache.rectCount)
        }

        if (paneCache.textCount > 0 && paneCache.textBuffer && paneCache.textBindGroup) {
          const textRenderer = artifacts.textPipeline.with(pass).with(paneCache.textBindGroup)
          textRenderer
            .with(WORKBOOK_UNIT_QUAD_LAYOUT, artifacts.quadBuffer)
            .with(WORKBOOK_TEXT_INSTANCE_LAYOUT, paneCache.textBuffer)
            .draw(6, paneCache.textCount)
        }
      })

      pass.end()
      artifacts.device.queue.submit([commandEncoder.finish()])
    }

    drawFrameRef.current()
  }, [active, panePayloads, surfaceSize, webGpuReady])

  useEffect(() => {
    if (!active || !scrollTransformStore) {
      return
    }

    const scheduleDraw = () => {
      if (scheduledDrawFrameRef.current !== null) {
        return
      }
      scheduledDrawFrameRef.current = window.requestAnimationFrame(() => {
        scheduledDrawFrameRef.current = null
        drawFrameRef.current()
      })
    }

    return scrollTransformStore.subscribe(scheduleDraw)
  }, [active, scrollTransformStore])

  useEffect(() => {
    const canvas = canvasRef.current
    return () => {
      if (scheduledDrawFrameRef.current !== null) {
        window.cancelAnimationFrame(scheduledDrawFrameRef.current)
        scheduledDrawFrameRef.current = null
      }
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
