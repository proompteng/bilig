import type { TgpuBindGroup } from 'typegpu'
import { parseGpuColor } from '../gridGpuScene.js'
import { noteTypeGpuBufferAllocation } from '../renderer-v2/grid-render-counters.js'
import { buildTextDecorationRectsFromRuns, buildTextQuadsFromRuns, type TextDecorationRect } from '../renderer-v2/line-text-quad-buffer.js'
import { GRID_SCENE_PACKET_V2_RECT_INSTANCE_FLOAT_COUNT } from '../renderer-v2/scene-packet-v2.js'
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
import { GpuBufferArenaV3, type GpuBufferHandleV3 } from './gpu-buffer-arena.js'
import type { GridRenderTile } from './render-tile-source.js'
import type { WorkbookRenderTilePaneState } from './render-tile-pane-state.js'

const RECT_INSTANCE_FLOAT_COUNT = GRID_SCENE_PACKET_V2_RECT_INSTANCE_FLOAT_COUNT
const TEXT_INSTANCE_FLOAT_COUNT = 16
const RECT_INSTANCE_BYTE_COUNT = RECT_INSTANCE_FLOAT_COUNT * Float32Array.BYTES_PER_ELEMENT
const TEXT_INSTANCE_BYTE_COUNT = TEXT_INSTANCE_FLOAT_COUNT * Float32Array.BYTES_PER_ELEMENT

export interface TypeGpuTileContentResourceEntryV3 {
  rectHandle: GpuBufferHandleV3<RectInstanceVertexBuffer> | null
  rectCount: number
  rectSignature: string | null
  textHandle: GpuBufferHandleV3<TextInstanceVertexBuffer> | null
  textCount: number
  textSignature: string | null
  decorationRects: readonly TextDecorationRect[] | null
}

export interface TypeGpuTilePlacementResourceEntryV3 {
  surfaceUniform: SurfaceUniformBuffer | null
  surfaceBindGroup: TgpuBindGroup | null
  textBindGroup: TgpuBindGroup | null
  textBindGroupAtlasVersion: number
}

function createEmptyContentEntry(): TypeGpuTileContentResourceEntryV3 {
  return {
    decorationRects: null,
    rectCount: 0,
    rectHandle: null,
    rectSignature: null,
    textCount: 0,
    textHandle: null,
    textSignature: null,
  }
}

function createEmptyPlacementEntry(): TypeGpuTilePlacementResourceEntryV3 {
  return {
    surfaceBindGroup: null,
    surfaceUniform: null,
    textBindGroup: null,
    textBindGroupAtlasVersion: -1,
  }
}

export class TypeGpuTileResourceCacheV3 {
  private readonly contentEntries = new Map<number, TypeGpuTileContentResourceEntryV3>()
  private readonly placementEntries = new Map<string, TypeGpuTilePlacementResourceEntryV3>()
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

  getContent(tileId: number): TypeGpuTileContentResourceEntryV3 {
    const existing = this.contentEntries.get(tileId)
    if (existing) {
      return existing
    }
    const next = createEmptyContentEntry()
    this.contentEntries.set(tileId, next)
    return next
  }

  peekContent(tileId: number): TypeGpuTileContentResourceEntryV3 | null {
    return this.contentEntries.get(tileId) ?? null
  }

  getPlacement(key: string): TypeGpuTilePlacementResourceEntryV3 {
    const existing = this.placementEntries.get(key)
    if (existing) {
      return existing
    }
    const next = createEmptyPlacementEntry()
    this.placementEntries.set(key, next)
    return next
  }

  peekPlacement(key: string): TypeGpuTilePlacementResourceEntryV3 | null {
    return this.placementEntries.get(key) ?? null
  }

  pruneExcept(input: { readonly contentKeys: ReadonlySet<number>; readonly placementKeys: ReadonlySet<string> }): void {
    for (const tileId of this.contentEntries.keys()) {
      if (!input.contentKeys.has(tileId)) {
        const entry = this.contentEntries.get(tileId)
        if (entry) {
          this.releaseContent(entry)
        }
        this.contentEntries.delete(tileId)
      }
    }
    for (const key of this.placementEntries.keys()) {
      if (!input.placementKeys.has(key)) {
        const entry = this.placementEntries.get(key)
        if (entry) {
          this.destroyPlacement(entry)
        }
        this.placementEntries.delete(key)
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

  trimFreeBuffers(bytesToFree: number): number {
    const rectFreed = this.rectArena.trim(bytesToFree)
    return rectFreed + this.textArena.trim(Math.max(0, bytesToFree - rectFreed))
  }

  dispose(): void {
    for (const entry of this.contentEntries.values()) {
      this.destroyContent(entry)
    }
    this.contentEntries.clear()
    for (const entry of this.placementEntries.values()) {
      this.destroyPlacement(entry)
    }
    this.placementEntries.clear()
    this.rectArena.trim(Number.MAX_SAFE_INTEGER)
    this.textArena.trim(Number.MAX_SAFE_INTEGER)
  }

  private releaseContent(entry: TypeGpuTileContentResourceEntryV3): void {
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

  private destroyContent(entry: TypeGpuTileContentResourceEntryV3): void {
    entry.rectHandle?.buffer.destroy()
    entry.textHandle?.buffer.destroy()
    entry.rectHandle = null
    entry.textHandle = null
    entry.rectCount = 0
    entry.textCount = 0
  }

  private destroyPlacement(entry: TypeGpuTilePlacementResourceEntryV3): void {
    entry.surfaceUniform?.buffer.destroy()
    entry.surfaceUniform = null
    entry.surfaceBindGroup = null
    entry.textBindGroup = null
    entry.textBindGroupAtlasVersion = -1
  }

  private requireArtifacts(): TypeGpuRendererArtifacts {
    if (!this.artifacts) {
      throw new Error('TypeGpuTileResourceCacheV3 requires TypeGPU artifacts to allocate buffers')
    }
    return this.artifacts
  }
}

export function syncTypeGpuTilePaneResourcesV3(input: {
  readonly artifacts: TypeGpuRendererArtifacts
  readonly atlas: ReturnType<typeof createGlyphAtlas>
  readonly tileResources: TypeGpuTileResourceCacheV3
  readonly panes: readonly WorkbookRenderTilePaneState[]
  readonly retainPanes?: readonly WorkbookRenderTilePaneState[] | undefined
}): void {
  const retainPanes = input.retainPanes ?? input.panes
  input.tileResources.pruneExcept({
    contentKeys: new Set(retainPanes.map(resolveWorkbookTileContentBufferKeyV3)),
    placementKeys: new Set(retainPanes.map(resolveWorkbookTilePlacementBufferKeyV3)),
  })

  input.panes.forEach((pane) => {
    const content = input.tileResources.getContent(resolveWorkbookTileContentBufferKeyV3(pane))
    const textSignature = resolveGridTextTileSignatureV3(pane.tile)
    if (content.textSignature !== textSignature) {
      syncTileTextResource({
        atlas: input.atlas,
        content,
        pane,
        textSignature,
        tileResources: input.tileResources,
      })
    }
    const rectSignature = resolveGridRectTileSignatureV3({
      decorationRects: content.decorationRects ?? [],
      tile: pane.tile,
    })
    if (content.rectSignature !== rectSignature) {
      syncTileRectResource({
        content,
        pane,
        rectSignature,
        tileResources: input.tileResources,
      })
    }
    input.tileResources.getPlacement(resolveWorkbookTilePlacementBufferKeyV3(pane))
  })
}

export function ensureTilePlacementSurfaceBindingsV3(
  artifacts: TypeGpuRendererArtifacts,
  placement: TypeGpuTilePlacementResourceEntryV3,
): void {
  if (!placement.surfaceUniform) {
    placement.surfaceUniform = createTypeGpuSurfaceUniform(artifacts.root)
  }
  if (!placement.surfaceBindGroup) {
    placement.surfaceBindGroup = createTypeGpuSurfaceBindGroup(artifacts.root, placement.surfaceUniform)
  }

  if (!artifacts.atlasTexture) {
    placement.textBindGroup = null
    placement.textBindGroupAtlasVersion = -1
    return
  }

  if (!placement.textBindGroup || placement.textBindGroupAtlasVersion !== artifacts.atlasVersion) {
    placement.textBindGroup = createTypeGpuTextBindGroup(artifacts, placement.surfaceUniform)
    placement.textBindGroupAtlasVersion = artifacts.atlasVersion
  }
}

export function resolveWorkbookTileContentBufferKeyV3(pane: Pick<WorkbookRenderTilePaneState, 'tile'>): number {
  return pane.tile.tileId
}

export function resolveWorkbookTilePlacementBufferKeyV3(pane: Pick<WorkbookRenderTilePaneState, 'paneId' | 'tile'>): string {
  return `tile-placement:v3:${pane.paneId}:${pane.tile.tileId}`
}

export function resolveGridTextTileSignatureV3(tile: GridRenderTile): string {
  return [
    'render-tile-v3:text',
    tile.tileId,
    tile.textCount,
    tile.version.values,
    tile.version.styles,
    tile.version.text,
    tile.version.axisX,
    tile.version.axisY,
    tile.lastBatchId,
  ].join(':')
}

export function resolveGridRectTileSignatureV3(input: {
  readonly tile: GridRenderTile
  readonly decorationRects?: readonly TextDecorationRect[] | undefined
}): string {
  const decorationRects = input.decorationRects ?? []
  return [
    'render-tile-v3:rect',
    input.tile.tileId,
    input.tile.rectCount,
    input.tile.version.values,
    input.tile.version.styles,
    input.tile.version.axisX,
    input.tile.version.axisY,
    input.tile.lastBatchId,
    decorationRects.length,
  ].join(':')
}

function syncTileTextResource(input: {
  readonly atlas: ReturnType<typeof createGlyphAtlas>
  readonly pane: WorkbookRenderTilePaneState
  readonly tileResources: TypeGpuTileResourceCacheV3
  readonly content: TypeGpuTileContentResourceEntryV3
  readonly textSignature: string
}): void {
  input.content.decorationRects = buildTextDecorationRectsFromRuns(input.pane.tile.textRuns, input.atlas)
  const textPayload = buildTextQuadsFromRuns(input.pane.tile.textRuns, input.atlas)
  if (textPayload.quadCount === 0) {
    releaseTextBuffer(input.tileResources, input.content)
    input.content.textCount = 0
    input.content.textSignature = input.textSignature
    return
  }
  const handle = prepareTextBuffer(input.tileResources, input.content, textPayload.quadCount)
  input.content.textHandle = handle
  input.content.textCount = textPayload.quadCount
  writeTypeGpuVertexBuffer(handle.buffer, textPayload.floats, `tile-text:${resolveWorkbookTileContentBufferKeyV3(input.pane)}`)
  input.content.textSignature = input.textSignature
}

function syncTileRectResource(input: {
  readonly pane: WorkbookRenderTilePaneState
  readonly tileResources: TypeGpuTileResourceCacheV3
  readonly content: TypeGpuTileContentResourceEntryV3
  readonly rectSignature: string
}): void {
  const decorationRects = input.content.decorationRects ?? []
  const rectPayload = buildRectInstanceDataFromTile({
    decorationRects,
    tile: input.pane.tile,
  })
  if (rectPayload.count === 0) {
    releaseRectBuffer(input.tileResources, input.content)
    input.content.rectCount = 0
    input.content.rectSignature = input.rectSignature
    return
  }
  const handle = prepareRectBuffer(input.tileResources, input.content, rectPayload.count)
  input.content.rectHandle = handle
  input.content.rectCount = rectPayload.count
  writeTypeGpuVertexBuffer(handle.buffer, rectPayload.floats, `tile-rect:${resolveWorkbookTileContentBufferKeyV3(input.pane)}`)
  input.content.rectSignature = input.rectSignature
}

function prepareRectBuffer(
  tileResources: TypeGpuTileResourceCacheV3,
  content: TypeGpuTileContentResourceEntryV3,
  requiredCount: number,
): GpuBufferHandleV3<RectInstanceVertexBuffer> {
  const requiredBytes = Math.max(1, requiredCount) * RECT_INSTANCE_BYTE_COUNT
  if (content.rectHandle && content.rectHandle.capacityBytes >= requiredBytes) {
    content.rectHandle.usedBytes = requiredBytes
    return content.rectHandle
  }
  releaseRectBuffer(tileResources, content)
  return tileResources.acquireRectHandle(requiredCount)
}

function prepareTextBuffer(
  tileResources: TypeGpuTileResourceCacheV3,
  content: TypeGpuTileContentResourceEntryV3,
  requiredCount: number,
): GpuBufferHandleV3<TextInstanceVertexBuffer> {
  const requiredBytes = Math.max(1, requiredCount) * TEXT_INSTANCE_BYTE_COUNT
  if (content.textHandle && content.textHandle.capacityBytes >= requiredBytes) {
    content.textHandle.usedBytes = requiredBytes
    return content.textHandle
  }
  releaseTextBuffer(tileResources, content)
  return tileResources.acquireTextHandle(requiredCount)
}

function releaseRectBuffer(tileResources: TypeGpuTileResourceCacheV3, content: TypeGpuTileContentResourceEntryV3): void {
  if (!content.rectHandle) {
    return
  }
  tileResources.releaseRectHandle(content.rectHandle)
  content.rectHandle = null
}

function releaseTextBuffer(tileResources: TypeGpuTileResourceCacheV3, content: TypeGpuTileContentResourceEntryV3): void {
  if (!content.textHandle) {
    return
  }
  tileResources.releaseTextHandle(content.textHandle)
  content.textHandle = null
}

function createRectBuffer(artifacts: TypeGpuRendererArtifacts, capacityBytes: number): RectInstanceVertexBuffer {
  const count = Math.max(1, Math.ceil(capacityBytes / RECT_INSTANCE_BYTE_COUNT))
  const byteLength = count * RECT_INSTANCE_BYTE_COUNT
  noteTypeGpuBufferAllocation(byteLength, 'tile-v3-rect-instances')
  return artifacts.root.createBuffer(WORKBOOK_RECT_INSTANCE_LAYOUT.schemaForCount(count)).$usage('vertex') as RectInstanceVertexBuffer
}

function createTextBuffer(artifacts: TypeGpuRendererArtifacts, capacityBytes: number): TextInstanceVertexBuffer {
  const count = Math.max(1, Math.ceil(capacityBytes / TEXT_INSTANCE_BYTE_COUNT))
  const byteLength = count * TEXT_INSTANCE_BYTE_COUNT
  noteTypeGpuBufferAllocation(byteLength, 'tile-v3-text-instances')
  return artifacts.root.createBuffer(WORKBOOK_TEXT_INSTANCE_LAYOUT.schemaForCount(count)).$usage('vertex') as TextInstanceVertexBuffer
}

function buildRectInstanceDataFromTile(input: {
  readonly tile: GridRenderTile
  readonly decorationRects?: readonly TextDecorationRect[]
}): { readonly floats: Float32Array; readonly count: number } {
  const decorationRects = input.decorationRects ?? []
  const total = input.tile.rectCount + decorationRects.length
  if (decorationRects.length === 0) {
    return { count: total, floats: input.tile.rectInstances }
  }
  const floats = new Float32Array(Math.max(1, total) * RECT_INSTANCE_FLOAT_COUNT)
  const clipX = 0
  const clipY = 0
  const clipX1 = Number.MAX_SAFE_INTEGER
  const clipY1 = Number.MAX_SAFE_INTEGER
  const packedFloatCount = input.tile.rectCount * RECT_INSTANCE_FLOAT_COUNT
  floats.set(input.tile.rectInstances.subarray(0, packedFloatCount), 0)
  const offset = packedFloatCount
  writeDecorationRects(floats, offset, decorationRects, clipX, clipY, clipX1, clipY1)
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
