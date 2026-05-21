import typegpuCore, {
  d,
  type TgpuBindGroup,
  type TgpuBuffer,
  type TgpuFixedSampler,
  type TgpuRenderPipeline,
  type TgpuRoot,
  type TgpuUniform,
  type VertexFlag,
} from 'typegpu'
import type { WgslArray } from 'typegpu/data'
import {
  noteTypeGpuAtlasUpload,
  noteTypeGpuAtlasDirtyPageUpload,
  noteTypeGpuBufferAllocation,
  noteTypeGpuBufferWrite,
  noteTypeGpuConfigure,
  noteTypeGpuUniformWrite,
} from '../grid-render-counters.js'
import type { GlyphAtlasDirtyPageUpload } from './typegpu-atlas-manager.js'

const UNIT_QUAD = new Float32Array([0, 0, 1, 0, 0, 1, 0, 1, 1, 0, 1, 1])
const UNIT_QUAD_VERTEX_COUNT = 6
const ATLAS_READBACK_SETTINGS: CanvasRenderingContext2DSettings = { willReadFrequently: true }
type AtlasReadbackContext = CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D
export const TYPEGPU_WORKBOOK_CANVAS_ALPHA_MODE = 'premultiplied' as const
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
    clipSpacePixel: d.location(1, d.vec2f),
    clip: d.location(2, d.vec4f),
  },
})`{
  let localPixel = in.quad * in.rectSize;
  let clipSpacePixel = in.rectOrigin + localPixel;
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
    clipSpacePixel,
    in.clipRect,
  );
}`.$uses({
  surface: surfaceBindGroupLayout.bound.surface,
})

const rectFragment = typegpuCore.fragmentFn({
  in: {
    color: d.location(0, d.vec4f),
    clipSpacePixel: d.location(1, d.vec2f),
    clip: d.location(2, d.vec4f),
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

type AtlasTexture = GPUTexture

export interface TypeGpuAtlasResourceArtifacts {
  readonly root: Pick<TgpuRoot, 'createTexture' | 'unwrap'>
  readonly device: Pick<GPUDevice, 'createTexture' | 'queue'>
  atlasTexture: AtlasTexture | null
  atlasVersion: number
  atlasWidth: number
  atlasHeight: number
}

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
    alphaMode: TYPEGPU_WORKBOOK_CANVAS_ALPHA_MODE,
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

export function writeTypeGpuVertexBufferSubrange<TData extends WgslArray>(input: {
  readonly buffer: Pick<TypeGpuVertexBuffer<TData>, 'write'>
  readonly floats: Float32Array
  readonly startFloat: number
  readonly floatCount: number
  readonly sourceStartFloat?: number | undefined
  readonly label?: string | undefined
}): void {
  const startFloat = Math.max(0, Math.floor(input.startFloat))
  const sourceStartFloat = Math.max(0, Math.floor(input.sourceStartFloat ?? input.startFloat))
  const sourceEndFloat = Math.max(
    sourceStartFloat,
    Math.min(input.floats.length, sourceStartFloat + Math.max(0, Math.floor(input.floatCount))),
  )
  if (sourceEndFloat <= sourceStartFloat) {
    return
  }
  const source = input.floats.subarray(sourceStartFloat, sourceEndFloat).slice().buffer
  const startOffset = startFloat * Float32Array.BYTES_PER_ELEMENT
  const endOffset = startOffset + source.byteLength
  input.buffer.write(source, {
    endOffset,
    startOffset,
  })
  noteTypeGpuBufferWrite(source.byteLength, input.label ?? 'vertex-subrange')
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
  artifacts: TypeGpuAtlasResourceArtifacts,
  atlas: {
    drainDirtyPages?: (() => readonly GlyphAtlasDirtyPageUpload[]) | undefined
    getCanvas(): HTMLCanvasElement | OffscreenCanvas | null
    getVersion(): number
    getSize(): { width: number; height: number }
  },
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
    artifacts.atlasTexture = createAtlasTexture(artifacts.device, nextSize.width, nextSize.height)
    artifacts.atlasWidth = nextSize.width
    artifacts.atlasHeight = nextSize.height
    atlas.drainDirtyPages?.()
    uploadFullAtlasTexture(artifacts, atlasCanvas, nextSize.width, nextSize.height)
    noteTypeGpuAtlasUpload(nextSize.width * nextSize.height * 4)
    artifacts.atlasVersion = nextVersion
    return
  }

  const dirtyPages = atlas.drainDirtyPages?.()
  if (!dirtyPages) {
    uploadFullAtlasTexture(artifacts, atlasCanvas, nextSize.width, nextSize.height)
    noteTypeGpuAtlasUpload(nextSize.width * nextSize.height * 4)
    artifacts.atlasVersion = nextVersion
    return
  }

  if (dirtyPages.length > 0 && artifacts.atlasTexture) {
    uploadAtlasDirtyPages(artifacts, atlasCanvas, dirtyPages)
  }
  artifacts.atlasVersion = nextVersion
}

const TEXT_ATLAS_TEXTURE_USAGE_COPY_DST = 0x02
const TEXT_ATLAS_TEXTURE_USAGE_BINDING = 0x04
const TEXT_ATLAS_TEXTURE_USAGE_RENDER_ATTACHMENT = 0x10

function createAtlasTexture(device: Pick<GPUDevice, 'createTexture'>, width: number, height: number): AtlasTexture {
  return device.createTexture({
    format: 'rgba8unorm',
    size: [width, height],
    usage: TEXT_ATLAS_TEXTURE_USAGE_COPY_DST | TEXT_ATLAS_TEXTURE_USAGE_BINDING | TEXT_ATLAS_TEXTURE_USAGE_RENDER_ATTACHMENT,
  })
}

function uploadFullAtlasTexture(
  artifacts: TypeGpuAtlasResourceArtifacts,
  atlasCanvas: HTMLCanvasElement | OffscreenCanvas,
  width: number,
  height: number,
): void {
  if (!artifacts.atlasTexture || width <= 0 || height <= 0) {
    return
  }
  writeAtlasCanvasRegionToTexture({
    artifacts,
    atlasCanvas,
    height,
    width,
    x: 0,
    y: 0,
  })
}

function uploadAtlasDirtyPages(
  artifacts: TypeGpuAtlasResourceArtifacts,
  atlasCanvas: HTMLCanvasElement | OffscreenCanvas,
  dirtyPages: readonly GlyphAtlasDirtyPageUpload[],
): void {
  if (!artifacts.atlasTexture) {
    return
  }
  let uploadBytes = 0
  for (const page of dirtyPages) {
    writeAtlasCanvasRegionToTexture({
      artifacts,
      atlasCanvas,
      height: page.height,
      width: page.width,
      x: page.x,
      y: page.y,
    })
    uploadBytes += page.byteSize
  }
  noteTypeGpuAtlasDirtyPageUpload(uploadBytes, dirtyPages.length)
  noteTypeGpuAtlasUpload(uploadBytes)
}

function writeAtlasCanvasRegionToTexture(input: {
  readonly artifacts: TypeGpuAtlasResourceArtifacts
  readonly atlasCanvas: HTMLCanvasElement | OffscreenCanvas
  readonly x: number
  readonly y: number
  readonly width: number
  readonly height: number
}): void {
  const context = getAtlasCanvasReadbackContext(input.atlasCanvas)
  if (!context || !input.artifacts.atlasTexture || input.width <= 0 || input.height <= 0) {
    return
  }
  const imageData = context.getImageData(input.x, input.y, input.width, input.height)
  input.artifacts.device.queue.writeTexture(
    {
      origin: { x: input.x, y: input.y },
      texture: input.artifacts.atlasTexture,
    },
    imageData.data,
    {
      bytesPerRow: input.width * 4,
      rowsPerImage: input.height,
    },
    {
      height: input.height,
      width: input.width,
    },
  )
}

function getAtlasCanvasReadbackContext(atlasCanvas: HTMLCanvasElement | OffscreenCanvas): AtlasReadbackContext | null {
  return atlasCanvas.getContext('2d', ATLAS_READBACK_SETTINGS) as AtlasReadbackContext | null
}
