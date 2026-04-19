import { useEffect, useMemo, useRef, useState } from 'react'
import type { GridGpuRect, GridGpuScene } from './gridGpuScene.js'

interface GridGpuSurfaceProps {
  readonly scene: GridGpuScene
  readonly host: HTMLDivElement | null
  readonly onActiveChange?: ((active: boolean) => void) | undefined
}

const GPU_BUFFER_USAGE_COPY_DST = 0x0008
const GPU_BUFFER_USAGE_UNIFORM = 0x0040
const GPU_BUFFER_USAGE_VERTEX = 0x0020
const SURFACE_UNIFORM_FLOAT_COUNT = 4
const RECT_INSTANCE_FLOAT_COUNT = 8
const UNIT_QUAD = new Float32Array([0, 0, 1, 0, 0, 1, 0, 1, 1, 0, 1, 1])
const EMPTY_SCENE: GridGpuScene = Object.freeze({
  fillRects: Object.freeze([]),
  borderRects: Object.freeze([]),
})
const GRID_GPU_SHADER = /* wgsl */ `
struct SurfaceUniforms {
  size: vec2f,
  _padding: vec2f,
};

@group(0) @binding(0) var<uniform> surface: SurfaceUniforms;

struct VertexOut {
  @builtin(position) position: vec4f,
  @location(0) color: vec4f,
};

@vertex
fn vs_main(
  @location(0) quad: vec2f,
  @location(1) rect_origin: vec2f,
  @location(2) rect_size: vec2f,
  @location(3) rect_color: vec4f,
) -> VertexOut {
  let pixel = rect_origin + quad * rect_size;
  let clip = vec2f(
    (pixel.x / surface.size.x) * 2.0 - 1.0,
    1.0 - (pixel.y / surface.size.y) * 2.0,
  );

  var out: VertexOut;
  out.position = vec4f(clip, 0.0, 1.0);
  out.color = rect_color;
  return out;
}

@fragment
fn fs_main(@location(0) color: vec4f) -> @location(0) vec4f {
  return color;
}
`

interface SurfaceSize {
  readonly width: number
  readonly height: number
  readonly pixelWidth: number
  readonly pixelHeight: number
}

interface WebGpuArtifacts {
  readonly device: GPUDevice
  readonly context: GPUCanvasContext
  readonly format: GPUTextureFormat
  readonly pipeline: GPURenderPipeline
  readonly quadBuffer: GPUBuffer
  readonly uniformBuffer: GPUBuffer
  readonly uniformBindGroup: GPUBindGroup
}

interface MutableRenderState {
  instanceBuffer: GPUBuffer | null
  instanceCapacity: number
}

type SurfaceMode = 'inactive' | 'webgpu'

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

export function GridGpuSurface({ scene, host, onActiveChange }: GridGpuSurfaceProps) {
  const canvasRef = useRef<HTMLCanvasElement | null>(null)
  const artifactsRef = useRef<WebGpuArtifacts | null>(null)
  const renderStateRef = useRef<MutableRenderState>({ instanceBuffer: null, instanceCapacity: 0 })
  const [mode, setMode] = useState<SurfaceMode>('inactive')
  const [surfaceSize, setSurfaceSize] = useState<SurfaceSize>({
    width: 0,
    height: 0,
    pixelWidth: 0,
    pixelHeight: 0,
  })

  useEffect(() => {
    onActiveChange?.(mode === 'webgpu')
  }, [mode, onActiveChange])

  useEffect(() => {
    if (!host) {
      setSurfaceSize({ width: 0, height: 0, pixelWidth: 0, pixelHeight: 0 })
      return
    }

    const updateSurfaceSize = () => {
      const next = resolveSurfaceSize(host)
      setSurfaceSize((current) =>
        current.width === next.width &&
        current.height === next.height &&
        current.pixelWidth === next.pixelWidth &&
        current.pixelHeight === next.pixelHeight
          ? current
          : next,
      )
    }

    updateSurfaceSize()
    const observer = new ResizeObserver(() => {
      updateSurfaceSize()
    })
    observer.observe(host)
    return () => {
      observer.disconnect()
    }
  }, [host])

  useEffect(() => {
    let cancelled = false
    const renderState = renderStateRef.current

    async function initialize() {
      if (!host || !canvasRef.current) {
        setMode('inactive')
        return
      }

      if (!('gpu' in navigator)) {
        setMode('inactive')
        return
      }

      const adapter = await navigator.gpu.requestAdapter()
      if (!adapter || cancelled) {
        setMode('inactive')
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
        setMode('inactive')
        return
      }

      const format = navigator.gpu.getPreferredCanvasFormat()
      const pipeline = createGridGpuPipeline(device, format)
      const quadBuffer = device.createBuffer({
        size: UNIT_QUAD.byteLength,
        usage: GPU_BUFFER_USAGE_VERTEX | GPU_BUFFER_USAGE_COPY_DST,
      })
      device.queue.writeBuffer(quadBuffer, 0, UNIT_QUAD)

      const uniformBuffer = device.createBuffer({
        size: SURFACE_UNIFORM_FLOAT_COUNT * Float32Array.BYTES_PER_ELEMENT,
        usage: GPU_BUFFER_USAGE_UNIFORM | GPU_BUFFER_USAGE_COPY_DST,
      })
      const uniformBindGroup = device.createBindGroup({
        layout: pipeline.getBindGroupLayout(0),
        entries: [
          {
            binding: 0,
            resource: {
              buffer: uniformBuffer,
            },
          },
        ],
      })

      artifactsRef.current = {
        device,
        context,
        format,
        pipeline,
        quadBuffer,
        uniformBuffer,
        uniformBindGroup,
      }
      setMode('webgpu')
    }

    void initialize()
    if (host) {
      noteCanvasSurfaceMount('canvas')
    }

    return () => {
      cancelled = true
      setMode('inactive')
      cleanupRenderState(renderState)
      const artifacts = artifactsRef.current
      artifacts?.quadBuffer.destroy()
      artifacts?.uniformBuffer.destroy()
      artifacts?.device.destroy()
      artifactsRef.current = null
    }
  }, [host])

  const activeScene = useMemo(() => (mode === 'webgpu' ? scene : EMPTY_SCENE), [mode, scene])
  useEffect(() => {
    if (mode !== 'webgpu') {
      return
    }

    const canvas = canvasRef.current
    const artifacts = artifactsRef.current
    if (!canvas || !artifacts) {
      return
    }

    configureCanvasElement(canvas, surfaceSize)
    configureCanvasContext(artifacts.context, artifacts.device, artifacts.format)
    noteCanvasPaint('gpu:overlay')
    renderRects({
      artifacts,
      rects: [...activeScene.fillRects, ...activeScene.borderRects],
      renderState: renderStateRef.current,
      surfaceSize,
    })
  }, [activeScene, mode, surfaceSize])

  if (!host) {
    return null
  }

  return <canvas ref={canvasRef} aria-hidden="true" className="pointer-events-none absolute inset-0 z-10" />
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
  }
}

function configureCanvasElement(canvas: HTMLCanvasElement, surfaceSize: SurfaceSize): void {
  if (canvas.width !== surfaceSize.pixelWidth) {
    canvas.width = surfaceSize.pixelWidth
  }
  if (canvas.height !== surfaceSize.pixelHeight) {
    canvas.height = surfaceSize.pixelHeight
  }
  const cssWidth = `${surfaceSize.width}px`
  const cssHeight = `${surfaceSize.height}px`
  if (canvas.style.width !== cssWidth) {
    canvas.style.width = cssWidth
  }
  if (canvas.style.height !== cssHeight) {
    canvas.style.height = cssHeight
  }
}

function configureCanvasContext(context: GPUCanvasContext, device: GPUDevice, format: GPUTextureFormat): void {
  context.configure({
    alphaMode: 'premultiplied',
    device,
    format,
  })
}

function createGridGpuPipeline(device: GPUDevice, format: GPUTextureFormat): GPURenderPipeline {
  const shaderModule = device.createShaderModule({ code: GRID_GPU_SHADER })

  return device.createRenderPipeline({
    layout: 'auto',
    vertex: {
      module: shaderModule,
      entryPoint: 'vs_main',
      buffers: [
        {
          arrayStride: 2 * Float32Array.BYTES_PER_ELEMENT,
          attributes: [
            {
              shaderLocation: 0,
              offset: 0,
              format: 'float32x2',
            },
          ],
        },
        {
          arrayStride: RECT_INSTANCE_FLOAT_COUNT * Float32Array.BYTES_PER_ELEMENT,
          stepMode: 'instance',
          attributes: [
            {
              shaderLocation: 1,
              offset: 0,
              format: 'float32x2',
            },
            {
              shaderLocation: 2,
              offset: 2 * Float32Array.BYTES_PER_ELEMENT,
              format: 'float32x2',
            },
            {
              shaderLocation: 3,
              offset: 4 * Float32Array.BYTES_PER_ELEMENT,
              format: 'float32x4',
            },
          ],
        },
      ],
    },
    fragment: {
      module: shaderModule,
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

function renderRects({
  artifacts,
  rects,
  renderState,
  surfaceSize,
}: {
  readonly artifacts: WebGpuArtifacts
  readonly rects: readonly GridGpuRect[]
  readonly renderState: MutableRenderState
  readonly surfaceSize: SurfaceSize
}): void {
  const { context, device, pipeline, quadBuffer, uniformBuffer, uniformBindGroup } = artifacts
  const uniformData = new Float32Array([Math.max(1, surfaceSize.width), Math.max(1, surfaceSize.height), 0, 0])
  device.queue.writeBuffer(uniformBuffer, 0, uniformData)

  const instanceData = buildInstanceData(rects)
  const instanceBuffer = ensureInstanceBuffer(device, renderState, instanceData.byteLength)
  if (instanceData.byteLength > 0) {
    device.queue.writeBuffer(instanceBuffer, 0, instanceData)
  }

  const encoder = device.createCommandEncoder()
  const pass = encoder.beginRenderPass({
    colorAttachments: [
      {
        view: context.getCurrentTexture().createView(),
        clearValue: { r: 0, g: 0, b: 0, a: 0 },
        loadOp: 'clear',
        storeOp: 'store',
      },
    ],
  })
  pass.setPipeline(pipeline)
  pass.setBindGroup(0, uniformBindGroup)
  pass.setVertexBuffer(0, quadBuffer)
  pass.setVertexBuffer(1, instanceBuffer)
  pass.draw(UNIT_QUAD.length / 2, rects.length)
  pass.end()
  device.queue.submit([encoder.finish()])
}

function buildInstanceData(rects: readonly GridGpuRect[]): Float32Array {
  if (rects.length === 0) {
    return new Float32Array(0)
  }

  const floats = new Float32Array(rects.length * RECT_INSTANCE_FLOAT_COUNT)
  let offset = 0
  for (const rect of rects) {
    floats[offset] = rect.x
    floats[offset + 1] = rect.y
    floats[offset + 2] = rect.width
    floats[offset + 3] = rect.height
    floats[offset + 4] = rect.color.r
    floats[offset + 5] = rect.color.g
    floats[offset + 6] = rect.color.b
    floats[offset + 7] = rect.color.a
    offset += RECT_INSTANCE_FLOAT_COUNT
  }
  return floats
}

function ensureInstanceBuffer(device: GPUDevice, renderState: MutableRenderState, byteLength: number): GPUBuffer {
  if (renderState.instanceBuffer && renderState.instanceCapacity >= byteLength) {
    return renderState.instanceBuffer
  }

  renderState.instanceBuffer?.destroy()
  const nextCapacity = Math.max(RECT_INSTANCE_FLOAT_COUNT * Float32Array.BYTES_PER_ELEMENT, byteLength)
  const instanceBuffer = device.createBuffer({
    size: nextCapacity,
    usage: GPU_BUFFER_USAGE_VERTEX | GPU_BUFFER_USAGE_COPY_DST,
  })
  renderState.instanceBuffer = instanceBuffer
  renderState.instanceCapacity = nextCapacity
  return instanceBuffer
}

function cleanupRenderState(renderState: MutableRenderState): void {
  renderState.instanceBuffer?.destroy()
  renderState.instanceBuffer = null
  renderState.instanceCapacity = 0
}

function isGpuCanvasContext(value: RenderingContext | ImageBitmapRenderingContext | null): value is GPUCanvasContext {
  return value !== null && typeof value === 'object' && 'configure' in value && 'getCurrentTexture' in value
}
