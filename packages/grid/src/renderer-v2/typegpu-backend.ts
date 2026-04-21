import typegpuCore, {
  d,
  type SampledFlag,
  type TgpuBindGroup,
  type TgpuBuffer,
  type TgpuFixedSampler,
  type TgpuRenderPipeline,
  type TgpuRoot,
  type TgpuTexture,
  type TgpuUniform,
  type VertexFlag,
} from 'typegpu'
import type { WgslArray } from 'typegpu/data'
import {
  noteTypeGpuAtlasUpload,
  noteTypeGpuBufferAllocation,
  noteTypeGpuBufferWrite,
  noteTypeGpuConfigure,
  noteTypeGpuUniformWrite,
} from '../renderer/grid-render-counters.js'

const UNIT_QUAD = new Float32Array([0, 0, 1, 0, 0, 1, 0, 1, 1, 0, 1, 1])
const UNIT_QUAD_VERTEX_COUNT = 6
const RECT_INSTANCE_FLOAT_COUNT = 20
const TEXT_INSTANCE_FLOAT_COUNT = 16

const surfaceUniformSchema = d.struct({
  origin: d.vec2f,
  viewportSize: d.vec2f,
  scrollOffset: d.vec2f,
  dpr: d.f32,
  _pad: d.f32,
})

const rectInstanceSchema = d.struct({
  rectOrigin: d.vec2f,
  rectSize: d.vec2f,
  rectColor: d.vec4f,
  borderColor: d.vec4f,
  cornerRadius: d.f32,
  borderThickness: d.f32,
  clipRect: d.vec4f,
})

const textInstanceSchema = d.struct({
  rectOrigin: d.vec2f,
  rectSize: d.vec2f,
  uv0: d.vec2f,
  uv1: d.vec2f,
  tint: d.vec4f,
  clipRect: d.vec4f,
})

const surfaceBindGroupLayout = typegpuCore
  .bindGroupLayout({
    surface: { uniform: surfaceUniformSchema },
  })
  .$idx(0)

const textBindGroupLayout = typegpuCore
  .bindGroupLayout({
    surface: { uniform: surfaceUniformSchema },
    atlasSampler: { sampler: 'filtering' },
    atlasTexture: { texture: d.texture2d() },
  })
  .$idx(0)

const rectVertex = typegpuCore.vertexFn({
  in: {
    quad: d.location(0, d.vec2f),
    rectOrigin: d.location(1, d.vec2f),
    rectSize: d.location(2, d.vec2f),
    rectColor: d.location(3, d.vec4f),
    borderColor: d.location(4, d.vec4f),
    cornerRadius: d.location(5, d.f32),
    borderThickness: d.location(6, d.f32),
    clipRect: d.location(7, d.vec4f),
  },
  out: {
    position: d.builtin.position,
    color: d.location(0, d.vec4f),
    panePixel: d.location(1, d.vec2f),
    clip: d.location(2, d.vec4f),
  },
})`{
  let localPixel = in.quad * in.rectSize;
  let panePixel = in.rectOrigin + surface.scrollOffset + localPixel;
  let screenPixel = surface.origin + panePixel;
  let ndc = vec2f(
    (screenPixel.x / surface.viewportSize.x) * 2.0 - 1.0,
    1.0 - (screenPixel.y / surface.viewportSize.y) * 2.0,
  );

  var color = in.rectColor;
  if (in.borderThickness > 0.0) {
    let maxPixel = in.rectSize - vec2f(in.borderThickness, in.borderThickness);
    if (
      localPixel.x < in.borderThickness ||
      localPixel.y < in.borderThickness ||
      localPixel.x >= maxPixel.x ||
      localPixel.y >= maxPixel.y
    ) {
      color = in.borderColor;
    }
  }

  return Out(
    vec4f(ndc, 0.0, 1.0),
    color,
    panePixel,
    in.clipRect,
  );
}`.$uses({
  surface: surfaceBindGroupLayout.bound.surface,
})

const rectFragment = typegpuCore.fragmentFn({
  in: {
    color: d.location(0, d.vec4f),
    panePixel: d.location(1, d.vec2f),
    clip: d.location(2, d.vec4f),
  },
  out: d.location(0, d.vec4f),
})`{
  if (
    in.panePixel.x < in.clip.x ||
    in.panePixel.y < in.clip.y ||
    in.panePixel.x > in.clip.z ||
    in.panePixel.y > in.clip.w
  ) {
    discard;
  }
  return vec4f(in.color.rgb * in.color.a, in.color.a);
}`

const textVertex = typegpuCore.vertexFn({
  in: {
    quad: d.location(0, d.vec2f),
    rectOrigin: d.location(1, d.vec2f),
    rectSize: d.location(2, d.vec2f),
    uv0: d.location(3, d.vec2f),
    uv1: d.location(4, d.vec2f),
    tint: d.location(5, d.vec4f),
    clipRect: d.location(6, d.vec4f),
  },
  out: {
    position: d.builtin.position,
    uv: d.location(0, d.vec2f),
    color: d.location(1, d.vec4f),
    clipSpacePixel: d.location(2, d.vec2f),
    clip: d.location(3, d.vec4f),
  },
})`{
  let clipSpacePixel = in.rectOrigin + in.quad * in.rectSize;
  let panePixel = in.rectOrigin + surface.scrollOffset + in.quad * in.rectSize;
  let screenPixel = surface.origin + panePixel;
  let ndc = vec2f(
    (screenPixel.x / surface.viewportSize.x) * 2.0 - 1.0,
    1.0 - (screenPixel.y / surface.viewportSize.y) * 2.0,
  );

  return Out(
    vec4f(ndc, 0.0, 1.0),
    vec2f(
      mix(in.uv0.x, in.uv1.x, in.quad.x),
      mix(in.uv0.y, in.uv1.y, in.quad.y),
    ),
    in.tint,
    clipSpacePixel,
    in.clipRect,
  );
}`.$uses({
  surface: textBindGroupLayout.bound.surface,
})

const textFragment = typegpuCore.fragmentFn({
  in: {
    uv: d.location(0, d.vec2f),
    color: d.location(1, d.vec4f),
    clipSpacePixel: d.location(2, d.vec2f),
    clip: d.location(3, d.vec4f),
  },
  out: d.location(0, d.vec4f),
})`{
  if (
    in.clipSpacePixel.x < in.clip.x ||
    in.clipSpacePixel.y < in.clip.y ||
    in.clipSpacePixel.x > in.clip.z ||
    in.clipSpacePixel.y > in.clip.w
  ) {
    discard;
  }

  let sampled = textureSample(atlasTexture, atlasSampler, in.uv);
  let alpha = in.color.a * sampled.a;
  return vec4f(in.color.rgb * alpha, alpha);
}`.$uses({
  atlasSampler: textBindGroupLayout.bound.atlasSampler,
  atlasTexture: textBindGroupLayout.bound.atlasTexture,
})

export const WORKBOOK_UNIT_QUAD_LAYOUT = typegpuCore.vertexLayout((count) => d.arrayOf(d.vec2f, count))
export const WORKBOOK_RECT_INSTANCE_LAYOUT = typegpuCore.vertexLayout((count) => d.arrayOf(rectInstanceSchema, count), 'instance')
export const WORKBOOK_TEXT_INSTANCE_LAYOUT = typegpuCore.vertexLayout((count) => d.arrayOf(textInstanceSchema, count), 'instance')

type UnitQuadSchema = ReturnType<(typeof WORKBOOK_UNIT_QUAD_LAYOUT)['schemaForCount']>
type RectInstanceSchemaArray = ReturnType<(typeof WORKBOOK_RECT_INSTANCE_LAYOUT)['schemaForCount']>
type TextInstanceSchemaArray = ReturnType<(typeof WORKBOOK_TEXT_INSTANCE_LAYOUT)['schemaForCount']>

export type TypeGpuVertexBuffer<TData extends WgslArray> = TgpuBuffer<TData> & VertexFlag
export type UnitQuadVertexBuffer = TypeGpuVertexBuffer<UnitQuadSchema>
export type RectInstanceVertexBuffer = TypeGpuVertexBuffer<RectInstanceSchemaArray>
export type TextInstanceVertexBuffer = TypeGpuVertexBuffer<TextInstanceSchemaArray>
export type SurfaceUniformBuffer = TgpuUniform<typeof surfaceUniformSchema>

type AtlasTexture = TgpuTexture & SampledFlag

export const MIN_TYPEGPU_RECT_VERTEX_CAPACITY = 2048
export const MIN_TYPEGPU_TEXT_VERTEX_CAPACITY = 4096

export interface TypeGpuRendererArtifacts {
  readonly root: TgpuRoot
  readonly device: GPUDevice
  readonly context: GPUCanvasContext
  readonly format: GPUTextureFormat
  readonly rectPipeline: TgpuRenderPipeline
  readonly textPipeline: TgpuRenderPipeline
  readonly quadBuffer: UnitQuadVertexBuffer
  readonly sampler: TgpuFixedSampler
  atlasTexture: AtlasTexture | null
  atlasVersion: number
  atlasWidth: number
  atlasHeight: number
}

export async function createTypeGpuRenderer(canvas: HTMLCanvasElement): Promise<TypeGpuRendererArtifacts | null> {
  if (typeof navigator === 'undefined' || !('gpu' in navigator)) {
    return null
  }

  const adapter = await navigator.gpu.requestAdapter()
  if (!adapter) {
    return null
  }

  const device = await adapter.requestDevice()
  const root = typegpuCore.initFromDevice({ device })
  const format = navigator.gpu.getPreferredCanvasFormat()
  const context = root.configureContext({
    alphaMode: 'premultiplied',
    canvas,
    format,
  })
  noteTypeGpuConfigure()

  const quadBuffer = root.createBuffer(WORKBOOK_UNIT_QUAD_LAYOUT.schemaForCount(UNIT_QUAD_VERTEX_COUNT)).$usage('vertex')
  noteTypeGpuBufferAllocation(UNIT_QUAD.byteLength, 'unit-quad')
  quadBuffer.write(UNIT_QUAD.buffer, { endOffset: UNIT_QUAD.byteLength })
  noteTypeGpuBufferWrite(UNIT_QUAD.byteLength, 'unit-quad')

  const sampler = root.createSampler({
    magFilter: 'linear',
    minFilter: 'linear',
  })

  const rectPipeline = root.createRenderPipeline({
    attribs: {
      quad: WORKBOOK_UNIT_QUAD_LAYOUT.attrib,
      rectOrigin: WORKBOOK_RECT_INSTANCE_LAYOUT.attrib.rectOrigin,
      rectSize: WORKBOOK_RECT_INSTANCE_LAYOUT.attrib.rectSize,
      rectColor: WORKBOOK_RECT_INSTANCE_LAYOUT.attrib.rectColor,
      borderColor: WORKBOOK_RECT_INSTANCE_LAYOUT.attrib.borderColor,
      cornerRadius: WORKBOOK_RECT_INSTANCE_LAYOUT.attrib.cornerRadius,
      borderThickness: WORKBOOK_RECT_INSTANCE_LAYOUT.attrib.borderThickness,
      clipRect: WORKBOOK_RECT_INSTANCE_LAYOUT.attrib.clipRect,
    },
    fragment: rectFragment,
    primitive: { topology: 'triangle-list' },
    targets: {
      blend: {
        alpha: {
          dstFactor: 'one-minus-src-alpha',
          operation: 'add',
          srcFactor: 'one',
        },
        color: {
          dstFactor: 'one-minus-src-alpha',
          operation: 'add',
          srcFactor: 'one',
        },
      },
      format,
    },
    vertex: rectVertex,
  })

  const textPipeline = root.createRenderPipeline({
    attribs: {
      quad: WORKBOOK_UNIT_QUAD_LAYOUT.attrib,
      rectOrigin: WORKBOOK_TEXT_INSTANCE_LAYOUT.attrib.rectOrigin,
      rectSize: WORKBOOK_TEXT_INSTANCE_LAYOUT.attrib.rectSize,
      uv0: WORKBOOK_TEXT_INSTANCE_LAYOUT.attrib.uv0,
      uv1: WORKBOOK_TEXT_INSTANCE_LAYOUT.attrib.uv1,
      tint: WORKBOOK_TEXT_INSTANCE_LAYOUT.attrib.tint,
      clipRect: WORKBOOK_TEXT_INSTANCE_LAYOUT.attrib.clipRect,
    },
    fragment: textFragment,
    primitive: { topology: 'triangle-list' },
    targets: {
      blend: {
        alpha: {
          dstFactor: 'one-minus-src-alpha',
          operation: 'add',
          srcFactor: 'one',
        },
        color: {
          dstFactor: 'one-minus-src-alpha',
          operation: 'add',
          srcFactor: 'one',
        },
      },
      format,
    },
    vertex: textVertex,
  })

  return {
    atlasHeight: 0,
    atlasTexture: null,
    atlasVersion: -1,
    atlasWidth: 0,
    context,
    device,
    format,
    quadBuffer,
    rectPipeline,
    root,
    sampler,
    textPipeline,
  }
}

export function destroyTypeGpuRenderer(artifacts: TypeGpuRendererArtifacts): void {
  artifacts.atlasTexture?.destroy()
  artifacts.root.destroy()
  artifacts.device.destroy()
}

export function ensureTypeGpuVertexBuffer(
  root: TgpuRoot,
  layout: typeof WORKBOOK_RECT_INSTANCE_LAYOUT,
  current: RectInstanceVertexBuffer | null,
  currentCapacity: number,
  nextCount: number,
): { buffer: RectInstanceVertexBuffer; capacity: number }
export function ensureTypeGpuVertexBuffer(
  root: TgpuRoot,
  layout: typeof WORKBOOK_TEXT_INSTANCE_LAYOUT,
  current: TextInstanceVertexBuffer | null,
  currentCapacity: number,
  nextCount: number,
): { buffer: TextInstanceVertexBuffer; capacity: number }
export function ensureTypeGpuVertexBuffer(
  root: TgpuRoot,
  layout: typeof WORKBOOK_RECT_INSTANCE_LAYOUT | typeof WORKBOOK_TEXT_INSTANCE_LAYOUT,
  current: RectInstanceVertexBuffer | TextInstanceVertexBuffer | null,
  currentCapacity: number,
  nextCount: number,
): { buffer: RectInstanceVertexBuffer | TextInstanceVertexBuffer; capacity: number } {
  const minCapacity = Math.max(1, nextCount)
  if (current && currentCapacity >= minCapacity) {
    return {
      buffer: current,
      capacity: currentCapacity,
    }
  }

  const nextCapacity = resolveTypeGpuVertexBufferCapacity({
    currentCapacity,
    minimumCapacity: layout === WORKBOOK_RECT_INSTANCE_LAYOUT ? MIN_TYPEGPU_RECT_VERTEX_CAPACITY : MIN_TYPEGPU_TEXT_VERTEX_CAPACITY,
    nextCount,
  })
  current?.destroy()

  if (layout === WORKBOOK_RECT_INSTANCE_LAYOUT) {
    noteTypeGpuBufferAllocation(nextCapacity * RECT_INSTANCE_FLOAT_COUNT * Float32Array.BYTES_PER_ELEMENT, 'rect-instances')
    return {
      buffer: root.createBuffer(WORKBOOK_RECT_INSTANCE_LAYOUT.schemaForCount(nextCapacity)).$usage('vertex') as RectInstanceVertexBuffer,
      capacity: nextCapacity,
    }
  }

  noteTypeGpuBufferAllocation(nextCapacity * TEXT_INSTANCE_FLOAT_COUNT * Float32Array.BYTES_PER_ELEMENT, 'text-instances')
  return {
    buffer: root.createBuffer(WORKBOOK_TEXT_INSTANCE_LAYOUT.schemaForCount(nextCapacity)).$usage('vertex') as TextInstanceVertexBuffer,
    capacity: nextCapacity,
  }
}

export function resolveTypeGpuVertexBufferCapacity(input: {
  readonly currentCapacity: number
  readonly minimumCapacity: number
  readonly nextCount: number
}): number {
  const required = Math.max(1, Math.ceil(input.nextCount))
  const currentCapacity = Math.max(0, Math.floor(input.currentCapacity))
  if (currentCapacity >= required) {
    return currentCapacity
  }

  let nextCapacity = Math.max(1, Math.ceil(input.minimumCapacity))
  while (nextCapacity < required) {
    nextCapacity *= 2
  }
  return nextCapacity
}

export function writeTypeGpuVertexBuffer<TData extends WgslArray>(
  buffer: TypeGpuVertexBuffer<TData>,
  floats: Float32Array,
  label = 'vertex',
): void {
  const source: ArrayBuffer =
    floats.byteOffset === 0 && floats.byteLength === floats.buffer.byteLength && floats.buffer instanceof ArrayBuffer
      ? floats.buffer
      : floats.slice().buffer
  buffer.write(source, {
    endOffset: floats.byteLength,
  })
  noteTypeGpuBufferWrite(floats.byteLength, label)
}

export function createTypeGpuSurfaceUniform(root: TgpuRoot): SurfaceUniformBuffer {
  return root.createUniform(surfaceUniformSchema, {
    origin: d.vec2f(0, 0),
    viewportSize: d.vec2f(1, 1),
    scrollOffset: d.vec2f(0, 0),
    dpr: 1,
    _pad: 0,
  })
}

export function createTypeGpuSurfaceBindGroup(root: TgpuRoot, surfaceUniform: SurfaceUniformBuffer): TgpuBindGroup {
  return root.createBindGroup(surfaceBindGroupLayout, {
    surface: surfaceUniform.buffer,
  })
}

export function createTypeGpuTextBindGroup(
  artifacts: TypeGpuRendererArtifacts,
  surfaceUniform: SurfaceUniformBuffer,
): TgpuBindGroup | null {
  if (!artifacts.atlasTexture) {
    return null
  }

  return artifacts.root.createBindGroup(textBindGroupLayout, {
    atlasSampler: artifacts.sampler,
    atlasTexture: artifacts.atlasTexture.createView(),
    surface: surfaceUniform.buffer,
  })
}

export function updateTypeGpuSurfaceUniform(
  surfaceUniform: SurfaceUniformBuffer,
  surface: { readonly width: number; readonly height: number; readonly dpr: number },
  origin: { readonly x: number; readonly y: number },
  scroll: { readonly x: number; readonly y: number },
): void {
  surfaceUniform.write({
    origin: d.vec2f(origin.x, origin.y),
    viewportSize: d.vec2f(surface.width, surface.height),
    scrollOffset: d.vec2f(scroll.x, scroll.y),
    dpr: surface.dpr,
    _pad: 0,
  })
  noteTypeGpuUniformWrite(32, 'surface')
}

export function syncTypeGpuAtlasResources(
  artifacts: TypeGpuRendererArtifacts,
  atlas: { getCanvas(): HTMLCanvasElement | OffscreenCanvas | null; getVersion(): number; getSize(): { width: number; height: number } },
): void {
  const atlasCanvas = atlas.getCanvas()
  if (!atlasCanvas) {
    return
  }

  const nextVersion = atlas.getVersion()
  if (artifacts.atlasVersion === nextVersion) {
    return
  }

  const nextSize = atlas.getSize()
  const needsTexture = !artifacts.atlasTexture || artifacts.atlasWidth !== nextSize.width || artifacts.atlasHeight !== nextSize.height

  if (needsTexture) {
    artifacts.atlasTexture?.destroy()
    artifacts.atlasTexture = createAtlasTexture(artifacts.root, nextSize.width, nextSize.height)
    artifacts.atlasWidth = nextSize.width
    artifacts.atlasHeight = nextSize.height
  }

  artifacts.atlasTexture?.write(atlasCanvas)
  noteTypeGpuAtlasUpload(nextSize.width * nextSize.height * 4)
  artifacts.atlasVersion = nextVersion
}

function createAtlasTexture(root: TgpuRoot, width: number, height: number): AtlasTexture {
  return root
    .createTexture({
      format: 'rgba8unorm',
      size: [width, height],
    })
    .$usage('sampled', 'render') as AtlasTexture
}
