import { memo, useEffect, useRef, useState } from 'react'
import type { GridGpuRect, GridGpuScene } from './gridGpuScene.js'
import type { Rectangle } from './gridTypes.js'

interface GridGpuPaneSurfaceProps {
  readonly paneId: string
  readonly active: boolean
  readonly frame: Rectangle
  readonly surfaceSize: {
    readonly width: number
    readonly height: number
  }
  readonly contentOffset: {
    readonly x: number
    readonly y: number
  }
  readonly scene: GridGpuScene
}

const GPU_BUFFER_USAGE_COPY_DST = 0x0008
const GPU_BUFFER_USAGE_UNIFORM = 0x0040
const GPU_BUFFER_USAGE_VERTEX = 0x0020
const SURFACE_UNIFORM_FLOAT_COUNT = 4
const RECT_INSTANCE_FLOAT_COUNT = 8
const UNIT_QUAD = new Float32Array([0, 0, 1, 0, 0, 1, 0, 1, 1, 0, 1, 1])
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

export const GridGpuPaneSurface = memo(function GridGpuPaneSurface({
  paneId,
  active,
  frame,
  surfaceSize,
  contentOffset,
  scene,
}: GridGpuPaneSurfaceProps) {
  const canvasRef = useRef<HTMLCanvasElement | null>(null)
  const artifactsRef = useRef<WebGpuArtifacts | null>(null)
  const renderStateRef = useRef<MutableRenderState>({ instanceBuffer: null, instanceCapacity: 0 })
  const [webGpuReady, setWebGpuReady] = useState(false)

  useEffect(() => {
    if (!active) {
      return
    }
    noteCanvasSurfaceMount('canvas')
  }, [active])

  useEffect(() => {
    let cancelled = false
    const renderState = renderStateRef.current

    async function initialize() {
      if (!active || !canvasRef.current) {
        setWebGpuReady(false)
        return
      }
      if (!('gpu' in navigator)) {
        setWebGpuReady(false)
        return
      }
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
        entries: [{ binding: 0, resource: { buffer: uniformBuffer } }],
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
      setWebGpuReady(true)
    }

    void initialize()

    return () => {
      cancelled = true
      setWebGpuReady(false)
      cleanupRenderState(renderState)
      const artifacts = artifactsRef.current
      artifacts?.quadBuffer.destroy()
      artifacts?.uniformBuffer.destroy()
      artifacts?.device.destroy()
      artifactsRef.current = null
    }
  }, [active])

  useEffect(() => {
    if (!active || !webGpuReady) {
      return
    }
    const canvas = canvasRef.current
    const artifacts = artifactsRef.current
    if (!canvas || !artifacts) {
      return
    }
    const resolvedSurfaceSize = resolveSurfaceSize(surfaceSize)
    configureCanvasElement(canvas, resolvedSurfaceSize)
    configureCanvasContext(artifacts.context, artifacts.device, artifacts.format)
    noteCanvasPaint(`gpu:${paneId}`)
    renderRects({
      artifacts,
      rects: [...scene.fillRects, ...scene.borderRects],
      renderState: renderStateRef.current,
      surfaceSize: resolvedSurfaceSize,
    })
  }, [active, paneId, scene, surfaceSize, webGpuReady])

  if (!active || frame.width <= 0 || frame.height <= 0 || surfaceSize.width <= 0 || surfaceSize.height <= 0) {
    return null
  }

  return (
    <div
      aria-hidden="true"
      className="pointer-events-none absolute z-10 overflow-hidden"
      style={{
        height: frame.height,
        left: frame.x,
        top: frame.y,
        width: frame.width,
      }}
    >
      <canvas
        className="absolute"
        data-pane-id={paneId}
        data-testid={`grid-gpu-pane-${paneId}`}
        ref={canvasRef}
        style={{
          left: contentOffset.x,
          top: contentOffset.y,
          width: surfaceSize.width,
          height: surfaceSize.height,
        }}
      />
    </div>
  )
})

function resolveSurfaceSize(size: { readonly width: number; readonly height: number }): SurfaceSize {
  const dpr = Math.max(1, window.devicePixelRatio || 1)
  return {
    width: Math.max(0, Math.floor(size.width)),
    height: Math.max(0, Math.floor(size.height)),
    pixelWidth: Math.max(1, Math.floor(size.width * dpr)),
    pixelHeight: Math.max(1, Math.floor(size.height * dpr)),
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
    device,
    format,
    alphaMode: 'premultiplied',
  })
}

function renderRects(input: {
  artifacts: WebGpuArtifacts
  rects: readonly GridGpuRect[]
  renderState: MutableRenderState
  surfaceSize: SurfaceSize
}): void {
  const { artifacts, rects, renderState, surfaceSize } = input
  const { device, context, pipeline, quadBuffer, uniformBuffer, uniformBindGroup } = artifacts
  const instanceCount = rects.length
  const instanceCapacity = Math.max(1, instanceCount)
  const instanceBufferSize = instanceCapacity * RECT_INSTANCE_FLOAT_COUNT * Float32Array.BYTES_PER_ELEMENT
  if (!renderState.instanceBuffer || renderState.instanceCapacity < instanceCapacity) {
    renderState.instanceBuffer?.destroy()
    renderState.instanceBuffer = device.createBuffer({
      size: instanceBufferSize,
      usage: GPU_BUFFER_USAGE_VERTEX | GPU_BUFFER_USAGE_COPY_DST,
    })
    renderState.instanceCapacity = instanceCapacity
  }
  const instanceFloats = new Float32Array(instanceCapacity * RECT_INSTANCE_FLOAT_COUNT)
  rects.forEach((rect, index) => {
    const base = index * RECT_INSTANCE_FLOAT_COUNT
    instanceFloats[base + 0] = rect.x
    instanceFloats[base + 1] = rect.y
    instanceFloats[base + 2] = rect.width
    instanceFloats[base + 3] = rect.height
    instanceFloats[base + 4] = rect.color.r
    instanceFloats[base + 5] = rect.color.g
    instanceFloats[base + 6] = rect.color.b
    instanceFloats[base + 7] = rect.color.a
  })
  device.queue.writeBuffer(renderState.instanceBuffer, 0, instanceFloats)
  device.queue.writeBuffer(uniformBuffer, 0, new Float32Array([surfaceSize.width, surfaceSize.height, 0, 0]))

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
  pass.setVertexBuffer(1, renderState.instanceBuffer)
  pass.draw(6, instanceCount)
  pass.end()
  device.queue.submit([encoder.finish()])
}

function cleanupRenderState(renderState: MutableRenderState): void {
  renderState.instanceBuffer?.destroy()
  renderState.instanceBuffer = null
  renderState.instanceCapacity = 0
}

function createGridGpuPipeline(device: GPUDevice, format: GPUTextureFormat): GPURenderPipeline {
  const shader = device.createShaderModule({ code: GRID_GPU_SHADER })
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

function isGpuCanvasContext(value: unknown): value is GPUCanvasContext {
  return typeof value === 'object' && value !== null && 'configure' in value && 'getCurrentTexture' in value
}
