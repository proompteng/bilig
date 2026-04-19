import { memo, useEffect, useMemo, useRef, useState } from 'react'
import type { GridGpuScene } from '../gridGpuScene.js'
import type { GridTextScene } from '../gridTextScene.js'
import type { Rectangle } from '../gridTypes.js'
import type { WorkbookPaneRenderState } from './pane-scene-types.js'
import { WorkbookPaneBufferCache } from './pane-buffer-cache.js'
import { createGlyphAtlas } from './glyph-atlas.js'
import { buildTextDecorationRectsFromScene, buildTextQuadsFromScene } from './text-quad-buffer.js'
import { WORKBOOK_RECT_SHADER, WORKBOOK_TEXT_SHADER } from './workbook-pane-shaders.js'

const GPU_BUFFER_USAGE_COPY_DST = 0x0008
const GPU_BUFFER_USAGE_UNIFORM = 0x0040
const GPU_BUFFER_USAGE_VERTEX = 0x0020
const GPU_TEXTURE_USAGE_TEXTURE_BINDING = 0x0004
const GPU_TEXTURE_USAGE_COPY_DST = 0x0008
const GPU_TEXTURE_USAGE_RENDER_ATTACHMENT = 0x0010
const SURFACE_UNIFORM_FLOAT_COUNT = 4
const RECT_INSTANCE_FLOAT_COUNT = 8
const TEXT_INSTANCE_FLOAT_COUNT = 16
const UNIT_QUAD = new Float32Array([0, 0, 1, 0, 0, 1, 0, 1, 1, 0, 1, 1])

interface WorkbookPaneRendererProps {
  readonly active: boolean
  readonly host: HTMLDivElement | null
  readonly panes: readonly WorkbookPaneRenderState[]
  readonly overlay?: {
    readonly gpuScene: GridGpuScene
    readonly textScene: GridTextScene
  }
  readonly onActiveChange?: ((active: boolean) => void) | undefined
}

interface SurfaceSize {
  readonly width: number
  readonly height: number
  readonly pixelWidth: number
  readonly pixelHeight: number
  readonly dpr: number
}

interface RendererArtifacts {
  readonly device: GPUDevice
  readonly context: GPUCanvasContext
  readonly format: GPUTextureFormat
  readonly rectPipeline: GPURenderPipeline
  readonly textPipeline: GPURenderPipeline
  readonly quadBuffer: GPUBuffer
  readonly uniformBuffer: GPUBuffer
  readonly rectBindGroup: GPUBindGroup
  readonly sampler: GPUSampler
  atlasTexture: GPUTexture | null
  textBindGroup: GPUBindGroup | null
  atlasVersion: number
}

type WorkbookRenderPaneId = WorkbookPaneRenderState['paneId'] | 'overlay'

interface PanePayload {
  readonly paneId: WorkbookRenderPaneId
  readonly frame: Rectangle
  readonly contentOffset: {
    readonly x: number
    readonly y: number
  }
  readonly gpuScene: GridGpuScene
  readonly textScene: GridTextScene
}

interface ColorParseCacheState {
  readonly map: Map<string, readonly [number, number, number, number]>
  readonly canvas: HTMLCanvasElement | null
  readonly context: CanvasRenderingContext2D | null
}

function noteCanvasSurfaceMount(kind: 'canvas' | 'dom'): void {
  if (typeof window === 'undefined') {
    return
  }
  ;(
    window as Window & {
      __biligScrollPerf?: {
        noteCanvasSurfaceMount?: (kind: 'canvas' | 'dom') => void
      }
    }
  ).__biligScrollPerf?.noteCanvasSurfaceMount?.(kind)
}

function noteCanvasPaint(layer: string): void {
  if (typeof window === 'undefined') {
    return
  }
  ;(window as Window & { __biligScrollPerf?: { noteCanvasPaint?: (layer: string) => void } }).__biligScrollPerf?.noteCanvasPaint?.(layer)
}

let colorParseCacheState: ColorParseCacheState | null = null

function getColorParseCacheState(): ColorParseCacheState {
  if (colorParseCacheState) {
    return colorParseCacheState
  }
  if (typeof document === 'undefined') {
    colorParseCacheState = {
      map: new Map(),
      canvas: null,
      context: null,
    }
    return colorParseCacheState
  }
  const canvas = document.createElement('canvas')
  canvas.width = 1
  canvas.height = 1
  colorParseCacheState = {
    map: new Map(),
    canvas,
    context: canvas.getContext('2d'),
  }
  return colorParseCacheState
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

function isGpuCanvasContext(value: unknown): value is GPUCanvasContext {
  return typeof value === 'object' && value !== null && 'configure' in value && 'getCurrentTexture' in value
}

function parseCssColor(color: string): readonly [number, number, number, number] {
  const cache = getColorParseCacheState()
  const cached = cache.map.get(color)
  if (cached) {
    return cached
  }
  if (!cache.context || !cache.canvas) {
    return [0, 0, 0, 1]
  }
  cache.context.clearRect(0, 0, 1, 1)
  cache.context.fillStyle = color
  cache.context.fillRect(0, 0, 1, 1)
  const data = cache.context.getImageData(0, 0, 1, 1).data
  const parsed = [data[0]! / 255, data[1]! / 255, data[2]! / 255, data[3]! / 255] as const
  cache.map.set(color, parsed)
  return parsed
}

function buildRectInstanceData(input: {
  frame: Rectangle
  contentOffset: { readonly x: number; readonly y: number }
  scene: GridGpuScene
  decorationRects?: readonly { x: number; y: number; width: number; height: number; color: string }[]
}): Float32Array {
  const rects = [...input.scene.fillRects, ...input.scene.borderRects]
  const decorationRects = input.decorationRects ?? []
  const totalRectCount = rects.length + decorationRects.length
  const floats = new Float32Array(Math.max(1, totalRectCount) * RECT_INSTANCE_FLOAT_COUNT)
  rects.forEach((rect, index) => {
    const base = index * RECT_INSTANCE_FLOAT_COUNT
    floats[base + 0] = input.frame.x + input.contentOffset.x + rect.x
    floats[base + 1] = input.frame.y + input.contentOffset.y + rect.y
    floats[base + 2] = rect.width
    floats[base + 3] = rect.height
    floats[base + 4] = rect.color.r
    floats[base + 5] = rect.color.g
    floats[base + 6] = rect.color.b
    floats[base + 7] = rect.color.a
  })
  decorationRects.forEach((rect, index) => {
    const base = (rects.length + index) * RECT_INSTANCE_FLOAT_COUNT
    const [r, g, b, a] = parseCssColor(rect.color)
    floats[base + 0] = input.frame.x + input.contentOffset.x + rect.x
    floats[base + 1] = input.frame.y + input.contentOffset.y + rect.y
    floats[base + 2] = rect.width
    floats[base + 3] = rect.height
    floats[base + 4] = r
    floats[base + 5] = g
    floats[base + 6] = b
    floats[base + 7] = a
  })
  return floats
}

function buildTextInstanceData(input: {
  paneId: WorkbookRenderPaneId
  frame: Rectangle
  contentOffset: { readonly x: number; readonly y: number }
  textScene: GridTextScene
  atlas: ReturnType<typeof createGlyphAtlas>
}): { floats: Float32Array; quadCount: number } {
  const quads = buildTextQuadsFromScene(input.textScene.items, input.atlas)
  const floats = new Float32Array(Math.max(1, quads.length) * TEXT_INSTANCE_FLOAT_COUNT)
  quads.forEach((quad, index) => {
    const base = index * TEXT_INSTANCE_FLOAT_COUNT
    const [r, g, b, a] = parseCssColor(quad.color)
    const quadX = quad.x + input.frame.x + input.contentOffset.x
    const quadY = quad.y + input.frame.y + input.contentOffset.y
    const clipLeft = quad.clipX + input.frame.x + input.contentOffset.x
    const clipTop = quad.clipY + input.frame.y + input.contentOffset.y
    floats[base + 0] = quadX
    floats[base + 1] = quadY
    floats[base + 2] = quad.width
    floats[base + 3] = quad.height
    floats[base + 4] = quad.u0
    floats[base + 5] = quad.v0
    floats[base + 6] = quad.u1
    floats[base + 7] = quad.v1
    floats[base + 8] = r
    floats[base + 9] = g
    floats[base + 10] = b
    floats[base + 11] = a
    floats[base + 12] = clipLeft
    floats[base + 13] = clipTop
    floats[base + 14] = clipLeft + quad.clipWidth
    floats[base + 15] = clipTop + quad.clipHeight
  })
  if (quads.length > 0) {
    noteCanvasPaint(`text:${input.paneId}`)
  }
  return { floats, quadCount: quads.length }
}

function ensureBuffer(
  device: GPUDevice,
  current: GPUBuffer | null,
  currentCapacity: number,
  nextCapacity: number,
): { buffer: GPUBuffer; capacity: number } {
  if (current && currentCapacity >= nextCapacity) {
    return { buffer: current, capacity: currentCapacity }
  }
  current?.destroy()
  return {
    buffer: device.createBuffer({
      size: nextCapacity,
      usage: GPU_BUFFER_USAGE_VERTEX | GPU_BUFFER_USAGE_COPY_DST,
    }),
    capacity: nextCapacity,
  }
}

function buildRectPipeline(device: GPUDevice, format: GPUTextureFormat): GPURenderPipeline {
  const shader = device.createShaderModule({ code: WORKBOOK_RECT_SHADER })
  return device.createRenderPipeline({
    layout: 'auto',
    vertex: {
      module: shader,
      entryPoint: 'vs_main',
      buffers: [
        {
          arrayStride: 2 * Float32Array.BYTES_PER_ELEMENT,
          attributes: [{ shaderLocation: 0, offset: 0, format: 'float32x2' }],
        },
        {
          arrayStride: RECT_INSTANCE_FLOAT_COUNT * Float32Array.BYTES_PER_ELEMENT,
          stepMode: 'instance',
          attributes: [
            { shaderLocation: 1, offset: 0, format: 'float32x2' },
            { shaderLocation: 2, offset: 2 * Float32Array.BYTES_PER_ELEMENT, format: 'float32x2' },
            { shaderLocation: 3, offset: 4 * Float32Array.BYTES_PER_ELEMENT, format: 'float32x4' },
          ],
        },
      ],
    },
    fragment: {
      module: shader,
      entryPoint: 'fs_main',
      targets: [{ format }],
    },
    primitive: {
      topology: 'triangle-list',
    },
  })
}

function buildTextPipeline(device: GPUDevice, format: GPUTextureFormat): GPURenderPipeline {
  const shader = device.createShaderModule({ code: WORKBOOK_TEXT_SHADER })
  return device.createRenderPipeline({
    layout: 'auto',
    vertex: {
      module: shader,
      entryPoint: 'vs_main',
      buffers: [
        {
          arrayStride: 2 * Float32Array.BYTES_PER_ELEMENT,
          attributes: [{ shaderLocation: 0, offset: 0, format: 'float32x2' }],
        },
        {
          arrayStride: TEXT_INSTANCE_FLOAT_COUNT * Float32Array.BYTES_PER_ELEMENT,
          stepMode: 'instance',
          attributes: [
            { shaderLocation: 1, offset: 0, format: 'float32x2' },
            { shaderLocation: 2, offset: 2 * Float32Array.BYTES_PER_ELEMENT, format: 'float32x2' },
            { shaderLocation: 3, offset: 4 * Float32Array.BYTES_PER_ELEMENT, format: 'float32x2' },
            { shaderLocation: 4, offset: 6 * Float32Array.BYTES_PER_ELEMENT, format: 'float32x2' },
            { shaderLocation: 5, offset: 8 * Float32Array.BYTES_PER_ELEMENT, format: 'float32x4' },
            { shaderLocation: 6, offset: 12 * Float32Array.BYTES_PER_ELEMENT, format: 'float32x4' },
          ],
        },
      ],
    },
    fragment: {
      module: shader,
      entryPoint: 'fs_main',
      targets: [
        {
          format,
          blend: {
            color: {
              srcFactor: 'src-alpha',
              dstFactor: 'one-minus-src-alpha',
              operation: 'add',
            },
            alpha: {
              srcFactor: 'one',
              dstFactor: 'one-minus-src-alpha',
              operation: 'add',
            },
          },
        },
      ],
    },
    primitive: {
      topology: 'triangle-list',
    },
  })
}

export const WorkbookPaneRenderer = memo(function WorkbookPaneRenderer({
  active,
  host,
  panes,
  overlay,
  onActiveChange,
}: WorkbookPaneRendererProps) {
  const canvasRef = useRef<HTMLCanvasElement | null>(null)
  const artifactsRef = useRef<RendererArtifacts | null>(null)
  const paneBuffersRef = useRef(new WorkbookPaneBufferCache())
  const atlasRef = useRef(createGlyphAtlas())
  const [webGpuReady, setWebGpuReady] = useState(false)
  const [surfaceSize, setSurfaceSize] = useState<SurfaceSize>({
    width: 0,
    height: 0,
    pixelWidth: 0,
    pixelHeight: 0,
    dpr: 1,
  })

  useEffect(() => {
    onActiveChange?.(webGpuReady)
  }, [onActiveChange, webGpuReady])

  useEffect(() => {
    if (!host) {
      setSurfaceSize({
        width: 0,
        height: 0,
        pixelWidth: 0,
        pixelHeight: 0,
        dpr: 1,
      })
      return
    }
    const update = () => setSurfaceSize(resolveSurfaceSize(host))
    update()
    if (typeof ResizeObserver === 'undefined') {
      const frame = window.requestAnimationFrame(update)
      return () => {
        window.cancelAnimationFrame(frame)
      }
    }
    const observer = new ResizeObserver(update)
    observer.observe(host)
    return () => {
      observer.disconnect()
    }
  }, [host])

  useEffect(() => {
    let cancelled = false
    const paneBuffers = paneBuffersRef.current

    async function initialize() {
      if (!active || !canvasRef.current || !('gpu' in navigator)) {
        setWebGpuReady(false)
        return
      }
      noteCanvasSurfaceMount('canvas')
      const adapter = await navigator.gpu.requestAdapter()
      if (!adapter || cancelled) {
        setWebGpuReady(false)
        return
      }
      const device = await adapter.requestDevice()
      if (cancelled) {
        device.destroy()
        return
      }
      const context = canvasRef.current.getContext('webgpu')
      if (!isGpuCanvasContext(context)) {
        device.destroy()
        setWebGpuReady(false)
        return
      }
      const format = navigator.gpu.getPreferredCanvasFormat()
      const rectPipeline = buildRectPipeline(device, format)
      const textPipeline = buildTextPipeline(device, format)
      const quadBuffer = device.createBuffer({
        size: UNIT_QUAD.byteLength,
        usage: GPU_BUFFER_USAGE_VERTEX | GPU_BUFFER_USAGE_COPY_DST,
      })
      device.queue.writeBuffer(quadBuffer, 0, UNIT_QUAD)
      const uniformBuffer = device.createBuffer({
        size: SURFACE_UNIFORM_FLOAT_COUNT * Float32Array.BYTES_PER_ELEMENT,
        usage: GPU_BUFFER_USAGE_UNIFORM | GPU_BUFFER_USAGE_COPY_DST,
      })
      const rectBindGroup = device.createBindGroup({
        layout: rectPipeline.getBindGroupLayout(0),
        entries: [{ binding: 0, resource: { buffer: uniformBuffer } }],
      })
      const sampler = device.createSampler({
        magFilter: 'linear',
        minFilter: 'linear',
      })
      artifactsRef.current = {
        device,
        context,
        format,
        rectPipeline,
        textPipeline,
        quadBuffer,
        uniformBuffer,
        rectBindGroup,
        sampler,
        atlasTexture: null,
        textBindGroup: null,
        atlasVersion: -1,
      }
      setWebGpuReady(true)
    }

    void initialize()

    return () => {
      cancelled = true
      setWebGpuReady(false)
      paneBuffers.dispose()
      const artifacts = artifactsRef.current
      artifacts?.atlasTexture?.destroy()
      artifacts?.quadBuffer.destroy()
      artifacts?.uniformBuffer.destroy()
      artifacts?.device.destroy()
      artifactsRef.current = null
    }
  }, [active])

  const panePayloads = useMemo<readonly PanePayload[]>(() => {
    const next: PanePayload[] = panes.map((pane) => ({
      paneId: pane.paneId,
      frame: pane.frame,
      contentOffset: pane.contentOffset,
      gpuScene: pane.gpuScene,
      textScene: pane.textScene,
    }))
    if (overlay) {
      next.push({
        paneId: 'overlay',
        frame: {
          x: 0,
          y: 0,
          width: surfaceSize.width,
          height: surfaceSize.height,
        },
        contentOffset: { x: 0, y: 0 },
        gpuScene: overlay.gpuScene,
        textScene: overlay.textScene,
      })
    }
    return next
  }, [overlay, panes, surfaceSize.height, surfaceSize.width])

  useEffect(() => {
    if (!active || !webGpuReady) {
      return
    }
    const artifacts = artifactsRef.current
    const canvas = canvasRef.current
    if (!artifacts || !canvas || surfaceSize.width <= 0 || surfaceSize.height <= 0) {
      return
    }
    if (canvas.width !== surfaceSize.pixelWidth) {
      canvas.width = surfaceSize.pixelWidth
    }
    if (canvas.height !== surfaceSize.pixelHeight) {
      canvas.height = surfaceSize.pixelHeight
    }
    canvas.style.width = `${surfaceSize.width}px`
    canvas.style.height = `${surfaceSize.height}px`
    artifacts.context.configure({
      device: artifacts.device,
      format: artifacts.format,
      alphaMode: 'premultiplied',
    })
    artifacts.device.queue.writeBuffer(artifacts.uniformBuffer, 0, new Float32Array([surfaceSize.width, surfaceSize.height, 0, 0]))

    const atlas = atlasRef.current
    const textPayloads = panePayloads.map((pane) => ({
      paneId: pane.paneId,
      frame: pane.frame,
      contentOffset: pane.contentOffset,
      textScene: pane.textScene,
    }))
    const textBuffers = textPayloads.map((pane) => ({
      paneId: pane.paneId,
      payload: buildTextInstanceData({
        paneId: pane.paneId,
        frame: pane.frame,
        contentOffset: pane.contentOffset,
        textScene: pane.textScene,
        atlas,
      }),
    }))
    const decorationRectsByPane = new Map<WorkbookRenderPaneId, ReturnType<typeof buildTextDecorationRectsFromScene>>(
      panePayloads.map((pane) => [pane.paneId, buildTextDecorationRectsFromScene(pane.textScene.items, atlas)]),
    )

    const atlasCanvas = atlas.getCanvas()
    const atlasVersion = atlas.getVersion()
    if (atlasCanvas && artifacts.atlasVersion !== atlasVersion) {
      const atlasSize = atlas.getSize()
      artifacts.atlasTexture?.destroy()
      artifacts.atlasTexture = artifacts.device.createTexture({
        size: {
          width: atlasSize.width,
          height: atlasSize.height,
          depthOrArrayLayers: 1,
        },
        format: 'rgba8unorm',
        usage: GPU_TEXTURE_USAGE_TEXTURE_BINDING | GPU_TEXTURE_USAGE_COPY_DST | GPU_TEXTURE_USAGE_RENDER_ATTACHMENT,
      })
      artifacts.device.queue.copyExternalImageToTexture(
        { source: atlasCanvas },
        { texture: artifacts.atlasTexture },
        { width: atlasSize.width, height: atlasSize.height },
      )
      artifacts.textBindGroup = artifacts.device.createBindGroup({
        layout: artifacts.textPipeline.getBindGroupLayout(0),
        entries: [
          { binding: 0, resource: { buffer: artifacts.uniformBuffer } },
          { binding: 1, resource: artifacts.sampler },
          { binding: 2, resource: artifacts.atlasTexture.createView() },
        ],
      })
      artifacts.atlasVersion = atlasVersion
    }

    const activePaneIds = new Set(panePayloads.map((pane) => pane.paneId))
    paneBuffersRef.current.pruneExcept(activePaneIds)

    const encoder = artifacts.device.createCommandEncoder()
    const pass = encoder.beginRenderPass({
      colorAttachments: [
        {
          view: artifacts.context.getCurrentTexture().createView(),
          clearValue: { r: 0, g: 0, b: 0, a: 0 },
          loadOp: 'clear',
          storeOp: 'store',
        },
      ],
    })
    pass.setVertexBuffer(0, artifacts.quadBuffer)
    const pixelScale = surfaceSize.dpr

    panePayloads.forEach((pane, paneIndex) => {
      const paneCache = paneBuffersRef.current.get(pane.paneId)
      const decorationRects = decorationRectsByPane.get(pane.paneId)
      const rectFloats = buildRectInstanceData({
        frame: pane.frame,
        contentOffset: pane.contentOffset,
        scene: pane.gpuScene,
        ...(decorationRects ? { decorationRects } : {}),
      })
      const rectBuffer = ensureBuffer(
        artifacts.device,
        paneCache.rectBuffer,
        paneCache.rectCapacity,
        rectFloats.byteLength || RECT_INSTANCE_FLOAT_COUNT * Float32Array.BYTES_PER_ELEMENT,
      )
      paneCache.rectBuffer = rectBuffer.buffer
      paneCache.rectCapacity = rectBuffer.capacity
      paneCache.rectCount =
        pane.gpuScene.fillRects.length + pane.gpuScene.borderRects.length + (decorationRectsByPane.get(pane.paneId)?.length ?? 0)
      artifacts.device.queue.writeBuffer(paneCache.rectBuffer, 0, rectFloats)

      const textPayload = textBuffers[paneIndex]!.payload
      const textBuffer = ensureBuffer(
        artifacts.device,
        paneCache.textBuffer,
        paneCache.textCapacity,
        textPayload.floats.byteLength || TEXT_INSTANCE_FLOAT_COUNT * Float32Array.BYTES_PER_ELEMENT,
      )
      paneCache.textBuffer = textBuffer.buffer
      paneCache.textCapacity = textBuffer.capacity
      paneCache.textCount = textPayload.quadCount
      artifacts.device.queue.writeBuffer(paneCache.textBuffer, 0, textPayload.floats)

      pass.setScissorRect(
        Math.max(0, Math.floor(pane.frame.x * pixelScale)),
        Math.max(0, Math.floor(pane.frame.y * pixelScale)),
        Math.max(1, Math.floor(pane.frame.width * pixelScale)),
        Math.max(1, Math.floor(pane.frame.height * pixelScale)),
      )

      if (paneCache.rectCount > 0) {
        noteCanvasPaint(`gpu:${pane.paneId}`)
        pass.setPipeline(artifacts.rectPipeline)
        pass.setBindGroup(0, artifacts.rectBindGroup)
        pass.setVertexBuffer(1, paneCache.rectBuffer)
        pass.draw(6, paneCache.rectCount)
      }

      if (paneCache.textCount > 0 && artifacts.textBindGroup) {
        pass.setPipeline(artifacts.textPipeline)
        pass.setBindGroup(0, artifacts.textBindGroup)
        pass.setVertexBuffer(1, paneCache.textBuffer)
        pass.draw(6, paneCache.textCount)
      }
    })

    pass.end()
    artifacts.device.queue.submit([encoder.finish()])
  }, [active, panePayloads, surfaceSize, webGpuReady])

  if (!host || !active) {
    return null
  }

  return (
    <canvas
      aria-hidden="true"
      className="pointer-events-none absolute inset-0 z-10"
      data-pane-renderer="workbook-pane-renderer"
      data-testid="grid-text-overlay"
      ref={canvasRef}
    />
  )
})
