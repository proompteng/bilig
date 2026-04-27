import type { TgpuBindGroup } from 'typegpu'
import type { GridHeaderPaneState } from '../gridHeaderPanes.js'
import { parseGpuColor } from '../gridGpuScene.js'
import { noteTypeGpuBufferAllocation } from '../renderer-v2/grid-render-counters.js'
import { buildTextDecorationRectsFromRuns, buildTextQuadsFromRuns, type TextDecorationRect } from '../renderer-v2/line-text-quad-buffer.js'
import type { createGlyphAtlas } from '../renderer-v2/typegpu-atlas-manager.js'
import {
  WORKBOOK_RECT_INSTANCE_LAYOUT,
  WORKBOOK_TEXT_INSTANCE_LAYOUT,
  createTypeGpuSurfaceBindGroup,
  createTypeGpuSurfaceUniform,
  createTypeGpuTextBindGroup,
  type RectInstanceVertexBuffer,
  type SurfaceUniformBuffer,
  type TextInstanceVertexBuffer,
  type TypeGpuRendererArtifacts,
  writeTypeGpuVertexBuffer,
} from '../renderer-v2/typegpu-backend.js'
import type { DynamicGridOverlayBatchV3 } from './dynamic-overlay-batch.js'
import { GpuBufferArenaV3, type GpuBufferHandleV3 } from './gpu-buffer-arena.js'
import { GRID_RECT_INSTANCE_FLOAT_COUNT_V3 } from './rect-instance-buffer.js'

export const WORKBOOK_DYNAMIC_OVERLAY_LAYER_KEY_V3 = 'overlay:v3'

const WORKBOOK_HEADER_LAYER_PREFIX_V3 = 'header:v3'
const RECT_INSTANCE_FLOAT_COUNT = GRID_RECT_INSTANCE_FLOAT_COUNT_V3
const TEXT_INSTANCE_FLOAT_COUNT = 16
const RECT_INSTANCE_BYTE_COUNT = RECT_INSTANCE_FLOAT_COUNT * Float32Array.BYTES_PER_ELEMENT
const TEXT_INSTANCE_BYTE_COUNT = TEXT_INSTANCE_FLOAT_COUNT * Float32Array.BYTES_PER_ELEMENT

export interface TypeGpuLayerResourceEntryV3 {
  rectHandle: GpuBufferHandleV3<RectInstanceVertexBuffer> | null
  rectCount: number
  rectSignature: string | null
  textHandle: GpuBufferHandleV3<TextInstanceVertexBuffer> | null
  textCount: number
  textSignature: string | null
  decorationRects: readonly TextDecorationRect[] | null
  surfaceUniform: SurfaceUniformBuffer | null
  surfaceBindGroup: TgpuBindGroup | null
  textBindGroup: TgpuBindGroup | null
  textBindGroupAtlasVersion: number
}

function createEmptyLayerEntry(): TypeGpuLayerResourceEntryV3 {
  return {
    decorationRects: null,
    rectCount: 0,
    rectHandle: null,
    rectSignature: null,
    surfaceBindGroup: null,
    surfaceUniform: null,
    textBindGroup: null,
    textBindGroupAtlasVersion: -1,
    textCount: 0,
    textHandle: null,
    textSignature: null,
  }
}

export class TypeGpuLayerResourceCacheV3 {
  private readonly entries = new Map<string, TypeGpuLayerResourceEntryV3>()
  private readonly rectArena: GpuBufferArenaV3<RectInstanceVertexBuffer>
  private readonly textArena: GpuBufferArenaV3<TextInstanceVertexBuffer>

  constructor(private readonly artifacts: TypeGpuRendererArtifacts | null = null) {
    this.rectArena = new GpuBufferArenaV3<RectInstanceVertexBuffer>(
      ({ capacityBytes }) => createRectBuffer(this.requireArtifacts(), capacityBytes),
      (buffer) => buffer.destroy(),
    )
    this.textArena = new GpuBufferArenaV3<TextInstanceVertexBuffer>(
      ({ capacityBytes }) => createTextBuffer(this.requireArtifacts(), capacityBytes),
      (buffer) => buffer.destroy(),
    )
  }

  get(key: string): TypeGpuLayerResourceEntryV3 {
    const existing = this.entries.get(key)
    if (existing) {
      return existing
    }
    const next = createEmptyLayerEntry()
    this.entries.set(key, next)
    return next
  }

  peek(key: string): TypeGpuLayerResourceEntryV3 | null {
    return this.entries.get(key) ?? null
  }

  pruneExcept(keys: ReadonlySet<string>): void {
    for (const key of this.entries.keys()) {
      if (!keys.has(key)) {
        const entry = this.entries.get(key)
        if (entry) {
          this.releaseEntry(entry)
          this.destroySurface(entry)
        }
        this.entries.delete(key)
      }
    }
  }

  acquireRectHandle(requiredCount: number): GpuBufferHandleV3<RectInstanceVertexBuffer> {
    return this.rectArena.acquire('rectInstances', Math.max(1, requiredCount) * RECT_INSTANCE_BYTE_COUNT)
  }

  releaseRectHandle(handle: GpuBufferHandleV3<RectInstanceVertexBuffer>): void {
    this.rectArena.release(handle)
  }

  acquireTextHandle(requiredCount: number): GpuBufferHandleV3<TextInstanceVertexBuffer> {
    return this.textArena.acquire('textRuns', Math.max(1, requiredCount) * TEXT_INSTANCE_BYTE_COUNT)
  }

  releaseTextHandle(handle: GpuBufferHandleV3<TextInstanceVertexBuffer>): void {
    this.textArena.release(handle)
  }

  dispose(): void {
    for (const entry of this.entries.values()) {
      this.destroyEntry(entry)
    }
    this.entries.clear()
    this.rectArena.trim(Number.MAX_SAFE_INTEGER)
    this.textArena.trim(Number.MAX_SAFE_INTEGER)
  }

  private releaseEntry(entry: TypeGpuLayerResourceEntryV3): void {
    if (entry.rectHandle) {
      this.rectArena.release(entry.rectHandle)
      entry.rectHandle = null
    }
    if (entry.textHandle) {
      this.textArena.release(entry.textHandle)
      entry.textHandle = null
    }
    entry.rectCount = 0
    entry.rectSignature = null
    entry.textCount = 0
    entry.textSignature = null
    entry.decorationRects = null
  }

  private destroyEntry(entry: TypeGpuLayerResourceEntryV3): void {
    entry.rectHandle?.buffer.destroy()
    entry.textHandle?.buffer.destroy()
    entry.rectHandle = null
    entry.textHandle = null
    this.destroySurface(entry)
  }

  private destroySurface(entry: TypeGpuLayerResourceEntryV3): void {
    entry.surfaceUniform?.buffer.destroy()
    entry.surfaceUniform = null
    entry.surfaceBindGroup = null
    entry.textBindGroup = null
    entry.textBindGroupAtlasVersion = -1
  }

  private requireArtifacts(): TypeGpuRendererArtifacts {
    if (!this.artifacts) {
      throw new Error('TypeGpuLayerResourceCacheV3 requires TypeGPU artifacts to allocate buffers')
    }
    return this.artifacts
  }
}

export function syncTypeGpuHeaderResourcesV3(input: {
  readonly artifacts: TypeGpuRendererArtifacts
  readonly atlas: ReturnType<typeof createGlyphAtlas>
  readonly layerResources: TypeGpuLayerResourceCacheV3
  readonly headerPanes: readonly GridHeaderPaneState[]
}): void {
  input.headerPanes.forEach((pane) => {
    const entry = input.layerResources.get(resolveWorkbookHeaderLayerKeyV3(pane))
    const textSignature = resolveHeaderTextSignatureV3(pane)
    if (entry.textSignature !== textSignature) {
      syncHeaderTextResource({
        atlas: input.atlas,
        entry,
        layerResources: input.layerResources,
        pane,
        textSignature,
      })
    }
    const rectSignature = resolveHeaderRectSignatureV3({
      decorationRects: entry.decorationRects ?? [],
      pane,
    })
    if (entry.rectSignature !== rectSignature) {
      syncHeaderRectResource({
        entry,
        layerResources: input.layerResources,
        pane,
        rectSignature,
      })
    }
  })
}

export function syncTypeGpuOverlayResourcesV3(input: {
  readonly layerResources: TypeGpuLayerResourceCacheV3
  readonly overlay: DynamicGridOverlayBatchV3 | null | undefined
}): void {
  if (!input.overlay) {
    return
  }
  const entry = input.layerResources.get(WORKBOOK_DYNAMIC_OVERLAY_LAYER_KEY_V3)
  entry.textCount = 0
  entry.textSignature = null
  releaseTextBuffer(input.layerResources, entry)
  if (input.overlay.rectCount === 0) {
    releaseRectBuffer(input.layerResources, entry)
    entry.rectCount = 0
    entry.rectSignature = input.overlay.rectSignature
    return
  }
  const rectSignature = resolveOverlayRectSignatureV3(input.overlay)
  if (entry.rectSignature === rectSignature) {
    return
  }
  const handle = prepareRectBuffer(input.layerResources, entry, input.overlay.rectCount)
  entry.rectHandle = handle
  entry.rectCount = input.overlay.rectCount
  writeTypeGpuVertexBuffer(handle.buffer, input.overlay.rectInstances, `overlay:${input.overlay.seq}`)
  entry.rectSignature = rectSignature
}

export function pruneTypeGpuLayerResourcesV3(input: {
  readonly layerResources: TypeGpuLayerResourceCacheV3
  readonly headerPanes: readonly GridHeaderPaneState[]
  readonly overlay: DynamicGridOverlayBatchV3 | null | undefined
}): void {
  const keys = new Set(input.headerPanes.map(resolveWorkbookHeaderLayerKeyV3))
  if (input.overlay) {
    keys.add(WORKBOOK_DYNAMIC_OVERLAY_LAYER_KEY_V3)
  }
  input.layerResources.pruneExcept(keys)
}

export function ensureLayerSurfaceBindingsV3(artifacts: TypeGpuRendererArtifacts, entry: TypeGpuLayerResourceEntryV3): void {
  if (!entry.surfaceUniform) {
    entry.surfaceUniform = createTypeGpuSurfaceUniform(artifacts.root)
  }
  if (!entry.surfaceBindGroup) {
    entry.surfaceBindGroup = createTypeGpuSurfaceBindGroup(artifacts.root, entry.surfaceUniform)
  }

  if (!artifacts.atlasTexture) {
    entry.textBindGroup = null
    entry.textBindGroupAtlasVersion = -1
    return
  }

  if (!entry.textBindGroup || entry.textBindGroupAtlasVersion !== artifacts.atlasVersion) {
    entry.textBindGroup = createTypeGpuTextBindGroup(artifacts, entry.surfaceUniform)
    entry.textBindGroupAtlasVersion = artifacts.atlasVersion
  }
}

export function resolveWorkbookHeaderLayerKeyV3(pane: Pick<GridHeaderPaneState, 'paneId'>): string {
  return `${WORKBOOK_HEADER_LAYER_PREFIX_V3}:${pane.paneId}`
}

function syncHeaderTextResource(input: {
  readonly atlas: ReturnType<typeof createGlyphAtlas>
  readonly pane: GridHeaderPaneState
  readonly layerResources: TypeGpuLayerResourceCacheV3
  readonly entry: TypeGpuLayerResourceEntryV3
  readonly textSignature: string
}): void {
  input.entry.decorationRects = buildTextDecorationRectsFromRuns(input.pane.textRuns, input.atlas)
  const textPayload = buildTextQuadsFromRuns(input.pane.textRuns, input.atlas)
  if (textPayload.quadCount === 0) {
    releaseTextBuffer(input.layerResources, input.entry)
    input.entry.textCount = 0
    input.entry.textSignature = input.textSignature
    return
  }
  const handle = prepareTextBuffer(input.layerResources, input.entry, textPayload.quadCount)
  input.entry.textHandle = handle
  input.entry.textCount = textPayload.quadCount
  writeTypeGpuVertexBuffer(handle.buffer, textPayload.floats, `header-text:${input.pane.paneId}`)
  input.entry.textSignature = input.textSignature
}

function syncHeaderRectResource(input: {
  readonly pane: GridHeaderPaneState
  readonly layerResources: TypeGpuLayerResourceCacheV3
  readonly entry: TypeGpuLayerResourceEntryV3
  readonly rectSignature: string
}): void {
  const rectPayload = buildRectInstanceDataFromHeader({
    decorationRects: input.entry.decorationRects ?? [],
    pane: input.pane,
  })
  if (rectPayload.count === 0) {
    releaseRectBuffer(input.layerResources, input.entry)
    input.entry.rectCount = 0
    input.entry.rectSignature = input.rectSignature
    return
  }
  const handle = prepareRectBuffer(input.layerResources, input.entry, rectPayload.count)
  input.entry.rectHandle = handle
  input.entry.rectCount = rectPayload.count
  writeTypeGpuVertexBuffer(handle.buffer, rectPayload.floats, `header-rect:${input.pane.paneId}`)
  input.entry.rectSignature = input.rectSignature
}

function prepareRectBuffer(
  layerResources: TypeGpuLayerResourceCacheV3,
  entry: TypeGpuLayerResourceEntryV3,
  requiredCount: number,
): GpuBufferHandleV3<RectInstanceVertexBuffer> {
  const requiredBytes = Math.max(1, requiredCount) * RECT_INSTANCE_BYTE_COUNT
  if (entry.rectHandle && entry.rectHandle.capacityBytes >= requiredBytes) {
    entry.rectHandle.usedBytes = requiredBytes
    return entry.rectHandle
  }
  releaseRectBuffer(layerResources, entry)
  return layerResources.acquireRectHandle(requiredCount)
}

function prepareTextBuffer(
  layerResources: TypeGpuLayerResourceCacheV3,
  entry: TypeGpuLayerResourceEntryV3,
  requiredCount: number,
): GpuBufferHandleV3<TextInstanceVertexBuffer> {
  const requiredBytes = Math.max(1, requiredCount) * TEXT_INSTANCE_BYTE_COUNT
  if (entry.textHandle && entry.textHandle.capacityBytes >= requiredBytes) {
    entry.textHandle.usedBytes = requiredBytes
    return entry.textHandle
  }
  releaseTextBuffer(layerResources, entry)
  return layerResources.acquireTextHandle(requiredCount)
}

function releaseRectBuffer(layerResources: TypeGpuLayerResourceCacheV3, entry: TypeGpuLayerResourceEntryV3): void {
  if (!entry.rectHandle) {
    return
  }
  layerResources.releaseRectHandle(entry.rectHandle)
  entry.rectHandle = null
}

function releaseTextBuffer(layerResources: TypeGpuLayerResourceCacheV3, entry: TypeGpuLayerResourceEntryV3): void {
  if (!entry.textHandle) {
    return
  }
  layerResources.releaseTextHandle(entry.textHandle)
  entry.textHandle = null
}

function createRectBuffer(artifacts: TypeGpuRendererArtifacts, capacityBytes: number): RectInstanceVertexBuffer {
  const count = Math.max(1, Math.ceil(capacityBytes / RECT_INSTANCE_BYTE_COUNT))
  const byteLength = count * RECT_INSTANCE_BYTE_COUNT
  noteTypeGpuBufferAllocation(byteLength, 'layer-v3-rect-instances')
  return artifacts.root.createBuffer(WORKBOOK_RECT_INSTANCE_LAYOUT.schemaForCount(count)).$usage('vertex') as RectInstanceVertexBuffer
}

function createTextBuffer(artifacts: TypeGpuRendererArtifacts, capacityBytes: number): TextInstanceVertexBuffer {
  const count = Math.max(1, Math.ceil(capacityBytes / TEXT_INSTANCE_BYTE_COUNT))
  const byteLength = count * TEXT_INSTANCE_BYTE_COUNT
  noteTypeGpuBufferAllocation(byteLength, 'layer-v3-text-instances')
  return artifacts.root.createBuffer(WORKBOOK_TEXT_INSTANCE_LAYOUT.schemaForCount(count)).$usage('vertex') as TextInstanceVertexBuffer
}

function resolveHeaderTextSignatureV3(pane: GridHeaderPaneState): string {
  return ['header-text-v3', pane.paneId, pane.textCount, pane.textSignature].join(':')
}

function resolveHeaderRectSignatureV3(input: {
  readonly pane: GridHeaderPaneState
  readonly decorationRects?: readonly TextDecorationRect[] | undefined
}): string {
  const decorationRects = input.decorationRects ?? []
  return [
    'header-rect-v3',
    input.pane.paneId,
    input.pane.rectCount,
    input.pane.fillRectCount,
    input.pane.borderRectCount,
    input.pane.rectSignature,
    input.pane.textSignature,
    input.pane.frame.width,
    input.pane.frame.height,
    decorationRects.length,
  ].join(':')
}

function resolveOverlayRectSignatureV3(overlay: DynamicGridOverlayBatchV3): string {
  return [
    'overlay-v3',
    overlay.sheetName,
    overlay.rectCount,
    overlay.fillRectCount,
    overlay.borderRectCount,
    overlay.surfaceSize.width,
    overlay.surfaceSize.height,
    overlay.rectSignature,
  ].join(':')
}

function buildRectInstanceDataFromHeader(input: {
  readonly pane: GridHeaderPaneState
  readonly decorationRects?: readonly TextDecorationRect[]
}): { readonly floats: Float32Array; readonly count: number } {
  const decorationRects = input.decorationRects ?? []
  const total = input.pane.rectCount + decorationRects.length
  if (decorationRects.length === 0) {
    return { count: total, floats: input.pane.rectInstances }
  }
  const floats = new Float32Array(Math.max(1, total) * RECT_INSTANCE_FLOAT_COUNT)
  const clipX = 0
  const clipY = 0
  const clipX1 = input.pane.surfaceSize.width
  const clipY1 = input.pane.surfaceSize.height
  const packedFloatCount = input.pane.rectCount * RECT_INSTANCE_FLOAT_COUNT
  floats.set(input.pane.rectInstances.subarray(0, packedFloatCount), 0)
  writeDecorationRects(floats, packedFloatCount, decorationRects, clipX, clipY, clipX1, clipY1)
  return { count: total, floats }
}

function writeDecorationRects(
  floats: Float32Array,
  offset: number,
  decorationRects: readonly TextDecorationRect[],
  clipX: number,
  clipY: number,
  clipX1: number,
  clipY1: number,
): number {
  let next = offset
  for (const rect of decorationRects) {
    next = writeDecorationRect(floats, next, rect, clipX, clipY, clipX1, clipY1)
  }
  return next
}

function writeDecorationRect(
  floats: Float32Array,
  offset: number,
  rect: TextDecorationRect,
  clipX: number,
  clipY: number,
  clipX1: number,
  clipY1: number,
): number {
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
  return offset + RECT_INSTANCE_FLOAT_COUNT
}
