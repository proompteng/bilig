import { parseGpuColor } from '../gridGpuScene.js'
import type { Rectangle } from '../gridTypes.js'
import {
  WORKBOOK_RECT_INSTANCE_LAYOUT,
  WORKBOOK_TEXT_INSTANCE_LAYOUT,
  createTypeGpuSurfaceBindGroup,
  createTypeGpuSurfaceUniform,
  createTypeGpuTextBindGroup,
  ensureTypeGpuVertexBuffer,
  type TypeGpuRendererArtifacts,
  writeTypeGpuVertexBuffer,
} from './typegpu-backend.js'
import { buildTextDecorationRectsFromRuns, buildTextQuadsFromRuns, type TextDecorationRect } from './line-text-quad-buffer.js'
import type { WorkbookPaneBufferEntry, WorkbookPaneBufferCache } from './pane-buffer-cache.js'
import type { createGlyphAtlas } from './typegpu-atlas-manager.js'
import type { WorkbookRenderPaneState } from './pane-scene-types.js'
import { GRID_SCENE_PACKET_V2_RECT_INSTANCE_FLOAT_COUNT, type GridScenePacketV2 } from './scene-packet-v2.js'
import { buildTileGpuCacheKey } from './tile-gpu-cache.js'
import type { GridHeaderPaneState } from '../gridHeaderPanes.js'
import type { DynamicGridOverlayBatchV3 } from '../renderer-v3/dynamic-overlay-batch.js'
import type { WorkbookRenderTilePaneState } from '../renderer-v3/render-tile-pane-state.js'
import type { GridRenderTile } from '../renderer-v3/render-tile-source.js'

const RECT_INSTANCE_FLOAT_COUNT = GRID_SCENE_PACKET_V2_RECT_INSTANCE_FLOAT_COUNT
export const WORKBOOK_DYNAMIC_OVERLAY_BUFFER_KEY = 'dynamic-overlay:v3'
const WORKBOOK_HEADER_BUFFER_PREFIX = 'header:v3'

export function syncTypeGpuPaneResources(input: {
  readonly artifacts: TypeGpuRendererArtifacts
  readonly atlas: ReturnType<typeof createGlyphAtlas>
  readonly paneBuffers: WorkbookPaneBufferCache
  readonly panes: readonly WorkbookRenderPaneState[]
  readonly retainPanes?: readonly WorkbookRenderPaneState[] | undefined
  readonly retainBufferKeys?: readonly string[] | undefined
}): void {
  const paneIds = new Set((input.retainPanes ?? input.panes).map(resolveWorkbookPaneBufferKey))
  for (const key of input.retainBufferKeys ?? []) {
    paneIds.add(key)
  }
  input.paneBuffers.pruneExcept(paneIds)

  input.panes.forEach((pane) => {
    const paneCache = input.paneBuffers.get(resolveWorkbookPaneBufferKey(pane))
    const textSignature = resolveGridTextPacketSignature(pane.packedScene)
    if (paneCache.textSignature !== textSignature) {
      syncTextResource({
        artifacts: input.artifacts,
        atlas: input.atlas,
        pane,
        paneBuffers: input.paneBuffers,
        paneCache,
        textSignature,
      })
    }
    const rectSignature = resolveGridRectPacketSignature({
      decorationRects: paneCache.decorationRects ?? [],
      frameHeight: pane.frame.height,
      frameWidth: pane.frame.width,
      packedScene: pane.packedScene,
    })
    if (paneCache.rectSignature !== rectSignature) {
      syncRectResource({
        artifacts: input.artifacts,
        pane,
        paneBuffers: input.paneBuffers,
        paneCache,
        rectSignature,
      })
    }
  })
}

export function syncTypeGpuTilePaneResources(input: {
  readonly artifacts: TypeGpuRendererArtifacts
  readonly atlas: ReturnType<typeof createGlyphAtlas>
  readonly paneBuffers: WorkbookPaneBufferCache
  readonly panes: readonly WorkbookRenderTilePaneState[]
  readonly retainPanes?: readonly WorkbookRenderTilePaneState[] | undefined
  readonly retainBufferKeys?: readonly string[] | undefined
}): void {
  const paneIds = new Set((input.retainPanes ?? input.panes).map(resolveWorkbookTilePaneBufferKey))
  for (const key of input.retainBufferKeys ?? []) {
    paneIds.add(key)
  }
  input.paneBuffers.pruneExcept(paneIds)

  input.panes.forEach((pane) => {
    const paneCache = input.paneBuffers.get(resolveWorkbookTilePaneBufferKey(pane))
    const textSignature = resolveGridTextTileSignature(pane.tile)
    if (paneCache.textSignature !== textSignature) {
      syncTileTextResource({
        artifacts: input.artifacts,
        atlas: input.atlas,
        pane,
        paneBuffers: input.paneBuffers,
        paneCache,
        textSignature,
      })
    }
    const rectSignature = resolveGridRectTileSignature({
      decorationRects: paneCache.decorationRects ?? [],
      tile: pane.tile,
    })
    if (paneCache.rectSignature !== rectSignature) {
      syncTileRectResource({
        artifacts: input.artifacts,
        pane,
        paneBuffers: input.paneBuffers,
        paneCache,
        rectSignature,
      })
    }
  })
}

export function syncTypeGpuHeaderResources(input: {
  readonly artifacts: TypeGpuRendererArtifacts
  readonly atlas: ReturnType<typeof createGlyphAtlas>
  readonly paneBuffers: WorkbookPaneBufferCache
  readonly headerPanes: readonly GridHeaderPaneState[]
}): void {
  input.headerPanes.forEach((pane) => {
    const paneCache = input.paneBuffers.get(resolveWorkbookHeaderBufferKey(pane))
    const textSignature = resolveHeaderTextSignature(pane)
    if (paneCache.textSignature !== textSignature) {
      syncHeaderTextResource({
        artifacts: input.artifacts,
        atlas: input.atlas,
        pane,
        paneBuffers: input.paneBuffers,
        paneCache,
        textSignature,
      })
    }
    const rectSignature = resolveHeaderRectSignature({
      decorationRects: paneCache.decorationRects ?? [],
      pane,
    })
    if (paneCache.rectSignature !== rectSignature) {
      syncHeaderRectResource({
        artifacts: input.artifacts,
        pane,
        paneBuffers: input.paneBuffers,
        paneCache,
        rectSignature,
      })
    }
  })
}

export function syncTypeGpuOverlayResources(input: {
  readonly artifacts: TypeGpuRendererArtifacts
  readonly paneBuffers: WorkbookPaneBufferCache
  readonly overlay: DynamicGridOverlayBatchV3 | null | undefined
}): void {
  if (!input.overlay) {
    return
  }
  const paneCache = input.paneBuffers.get(WORKBOOK_DYNAMIC_OVERLAY_BUFFER_KEY)
  paneCache.textCount = 0
  paneCache.textSignature = null
  if (input.overlay.rectCount === 0) {
    releaseRectBuffer(input.paneBuffers, paneCache)
    paneCache.rectCount = 0
    paneCache.rectSignature = input.overlay.rectSignature
    return
  }
  const rectSignature = resolveOverlayRectSignature(input.overlay)
  if (paneCache.rectSignature === rectSignature) {
    return
  }
  const reusable = prepareRectBuffer(input.paneBuffers, paneCache, input.overlay.rectCount)
  const rectBuffer = ensureTypeGpuVertexBuffer(
    input.artifacts.root,
    WORKBOOK_RECT_INSTANCE_LAYOUT,
    reusable.buffer,
    reusable.capacity,
    input.overlay.rectCount,
  )
  paneCache.rectBuffer = rectBuffer.buffer
  paneCache.rectCapacity = rectBuffer.capacity
  paneCache.rectCount = input.overlay.rectCount
  writeTypeGpuVertexBuffer(paneCache.rectBuffer, input.overlay.rectInstances, `overlay:${input.overlay.seq}`)
  paneCache.rectSignature = rectSignature
}

export function resolveWorkbookPaneBufferKey(pane: WorkbookRenderPaneState): string {
  return `${pane.paneId}:${buildTileGpuCacheKey(pane.packedScene)}`
}

export function resolveWorkbookTilePaneBufferKey(pane: Pick<WorkbookRenderTilePaneState, 'tile'>): string {
  return `tile:v3:${pane.tile.tileId}`
}

export function resolveWorkbookHeaderBufferKey(pane: Pick<GridHeaderPaneState, 'paneId'>): string {
  return `${WORKBOOK_HEADER_BUFFER_PREFIX}:${pane.paneId}`
}

export function ensurePaneSurfaceBindings(artifacts: TypeGpuRendererArtifacts, paneCache: WorkbookPaneBufferEntry): void {
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

function syncTextResource(input: {
  readonly artifacts: TypeGpuRendererArtifacts
  readonly atlas: ReturnType<typeof createGlyphAtlas>
  readonly pane: WorkbookRenderPaneState
  readonly paneBuffers: WorkbookPaneBufferCache
  readonly paneCache: WorkbookPaneBufferEntry
  readonly textSignature: string
}): void {
  input.paneCache.decorationRects = buildTextDecorationRectsFromRuns(input.pane.packedScene.textRuns, input.atlas)
  const textPayload = buildTextInstanceData({
    atlas: input.atlas,
    packet: input.pane.packedScene,
  })
  if (textPayload.quadCount === 0) {
    releaseTextBuffer(input.paneBuffers, input.paneCache)
    input.paneCache.textCount = 0
    input.paneCache.textSignature = input.textSignature
    return
  }
  const reusable = prepareTextBuffer(input.paneBuffers, input.paneCache, textPayload.quadCount)
  const textBuffer = ensureTypeGpuVertexBuffer(
    input.artifacts.root,
    WORKBOOK_TEXT_INSTANCE_LAYOUT,
    reusable.buffer,
    reusable.capacity,
    textPayload.quadCount,
  )
  input.paneCache.textBuffer = textBuffer.buffer
  input.paneCache.textCapacity = textBuffer.capacity
  input.paneCache.textCount = textPayload.quadCount
  writeTypeGpuVertexBuffer(input.paneCache.textBuffer, textPayload.floats, `text:${resolveWorkbookPaneBufferKey(input.pane)}`)
  input.paneCache.textSignature = input.textSignature
}

function syncTileTextResource(input: {
  readonly artifacts: TypeGpuRendererArtifacts
  readonly atlas: ReturnType<typeof createGlyphAtlas>
  readonly pane: WorkbookRenderTilePaneState
  readonly paneBuffers: WorkbookPaneBufferCache
  readonly paneCache: WorkbookPaneBufferEntry
  readonly textSignature: string
}): void {
  input.paneCache.decorationRects = buildTextDecorationRectsFromRuns(input.pane.tile.textRuns, input.atlas)
  const textPayload = buildTextQuadsFromRuns(input.pane.tile.textRuns, input.atlas)
  if (textPayload.quadCount === 0) {
    releaseTextBuffer(input.paneBuffers, input.paneCache)
    input.paneCache.textCount = 0
    input.paneCache.textSignature = input.textSignature
    return
  }
  const reusable = prepareTextBuffer(input.paneBuffers, input.paneCache, textPayload.quadCount)
  const textBuffer = ensureTypeGpuVertexBuffer(
    input.artifacts.root,
    WORKBOOK_TEXT_INSTANCE_LAYOUT,
    reusable.buffer,
    reusable.capacity,
    textPayload.quadCount,
  )
  input.paneCache.textBuffer = textBuffer.buffer
  input.paneCache.textCapacity = textBuffer.capacity
  input.paneCache.textCount = textPayload.quadCount
  writeTypeGpuVertexBuffer(input.paneCache.textBuffer, textPayload.floats, `tile-text:${resolveWorkbookTilePaneBufferKey(input.pane)}`)
  input.paneCache.textSignature = input.textSignature
}

function syncHeaderTextResource(input: {
  readonly artifacts: TypeGpuRendererArtifacts
  readonly atlas: ReturnType<typeof createGlyphAtlas>
  readonly pane: GridHeaderPaneState
  readonly paneBuffers: WorkbookPaneBufferCache
  readonly paneCache: WorkbookPaneBufferEntry
  readonly textSignature: string
}): void {
  input.paneCache.decorationRects = buildTextDecorationRectsFromRuns(input.pane.textRuns, input.atlas)
  const textPayload = buildTextQuadsFromRuns(input.pane.textRuns, input.atlas)
  if (textPayload.quadCount === 0) {
    releaseTextBuffer(input.paneBuffers, input.paneCache)
    input.paneCache.textCount = 0
    input.paneCache.textSignature = input.textSignature
    return
  }
  const reusable = prepareTextBuffer(input.paneBuffers, input.paneCache, textPayload.quadCount)
  const textBuffer = ensureTypeGpuVertexBuffer(
    input.artifacts.root,
    WORKBOOK_TEXT_INSTANCE_LAYOUT,
    reusable.buffer,
    reusable.capacity,
    textPayload.quadCount,
  )
  input.paneCache.textBuffer = textBuffer.buffer
  input.paneCache.textCapacity = textBuffer.capacity
  input.paneCache.textCount = textPayload.quadCount
  writeTypeGpuVertexBuffer(input.paneCache.textBuffer, textPayload.floats, `header-text:${input.pane.paneId}`)
  input.paneCache.textSignature = input.textSignature
}

function syncRectResource(input: {
  readonly artifacts: TypeGpuRendererArtifacts
  readonly pane: WorkbookRenderPaneState
  readonly paneBuffers: WorkbookPaneBufferCache
  readonly paneCache: WorkbookPaneBufferEntry
  readonly rectSignature: string
}): void {
  const decorationRects = input.paneCache.decorationRects ?? []
  const rectPayload = buildRectInstanceDataFromPacket({
    decorationRects,
    frame: input.pane.frame,
    packet: input.pane.packedScene,
  })
  if (rectPayload.count === 0) {
    releaseRectBuffer(input.paneBuffers, input.paneCache)
    input.paneCache.rectCount = 0
    input.paneCache.rectSignature = input.rectSignature
    return
  }
  const reusable = prepareRectBuffer(input.paneBuffers, input.paneCache, rectPayload.count)
  const rectBuffer = ensureTypeGpuVertexBuffer(
    input.artifacts.root,
    WORKBOOK_RECT_INSTANCE_LAYOUT,
    reusable.buffer,
    reusable.capacity,
    rectPayload.count,
  )
  input.paneCache.rectBuffer = rectBuffer.buffer
  input.paneCache.rectCapacity = rectBuffer.capacity
  input.paneCache.rectCount = rectPayload.count
  writeTypeGpuVertexBuffer(input.paneCache.rectBuffer, rectPayload.floats, `rect:${resolveWorkbookPaneBufferKey(input.pane)}`)
  input.paneCache.rectSignature = input.rectSignature
}

function syncTileRectResource(input: {
  readonly artifacts: TypeGpuRendererArtifacts
  readonly pane: WorkbookRenderTilePaneState
  readonly paneBuffers: WorkbookPaneBufferCache
  readonly paneCache: WorkbookPaneBufferEntry
  readonly rectSignature: string
}): void {
  const decorationRects = input.paneCache.decorationRects ?? []
  const rectPayload = buildRectInstanceDataFromTile({
    decorationRects,
    tile: input.pane.tile,
  })
  if (rectPayload.count === 0) {
    releaseRectBuffer(input.paneBuffers, input.paneCache)
    input.paneCache.rectCount = 0
    input.paneCache.rectSignature = input.rectSignature
    return
  }
  const reusable = prepareRectBuffer(input.paneBuffers, input.paneCache, rectPayload.count)
  const rectBuffer = ensureTypeGpuVertexBuffer(
    input.artifacts.root,
    WORKBOOK_RECT_INSTANCE_LAYOUT,
    reusable.buffer,
    reusable.capacity,
    rectPayload.count,
  )
  input.paneCache.rectBuffer = rectBuffer.buffer
  input.paneCache.rectCapacity = rectBuffer.capacity
  input.paneCache.rectCount = rectPayload.count
  writeTypeGpuVertexBuffer(input.paneCache.rectBuffer, rectPayload.floats, `tile-rect:${resolveWorkbookTilePaneBufferKey(input.pane)}`)
  input.paneCache.rectSignature = input.rectSignature
}

function syncHeaderRectResource(input: {
  readonly artifacts: TypeGpuRendererArtifacts
  readonly pane: GridHeaderPaneState
  readonly paneBuffers: WorkbookPaneBufferCache
  readonly paneCache: WorkbookPaneBufferEntry
  readonly rectSignature: string
}): void {
  const decorationRects = input.paneCache.decorationRects ?? []
  const rectPayload = buildRectInstanceDataFromHeader({
    decorationRects,
    pane: input.pane,
  })
  if (rectPayload.count === 0) {
    releaseRectBuffer(input.paneBuffers, input.paneCache)
    input.paneCache.rectCount = 0
    input.paneCache.rectSignature = input.rectSignature
    return
  }
  const reusable = prepareRectBuffer(input.paneBuffers, input.paneCache, rectPayload.count)
  const rectBuffer = ensureTypeGpuVertexBuffer(
    input.artifacts.root,
    WORKBOOK_RECT_INSTANCE_LAYOUT,
    reusable.buffer,
    reusable.capacity,
    rectPayload.count,
  )
  input.paneCache.rectBuffer = rectBuffer.buffer
  input.paneCache.rectCapacity = rectBuffer.capacity
  input.paneCache.rectCount = rectPayload.count
  writeTypeGpuVertexBuffer(input.paneCache.rectBuffer, rectPayload.floats, `header-rect:${input.pane.paneId}`)
  input.paneCache.rectSignature = input.rectSignature
}

function prepareRectBuffer(
  paneBuffers: WorkbookPaneBufferCache,
  paneCache: WorkbookPaneBufferEntry,
  requiredCount: number,
): {
  readonly buffer: WorkbookPaneBufferEntry['rectBuffer']
  readonly capacity: number
} {
  if (paneCache.rectBuffer && paneCache.rectCapacity >= requiredCount) {
    return { buffer: paneCache.rectBuffer, capacity: paneCache.rectCapacity }
  }
  releaseRectBuffer(paneBuffers, paneCache)
  const reused = paneBuffers.acquireRectBuffer(requiredCount)
  return {
    buffer: reused?.buffer ?? null,
    capacity: reused?.capacity ?? 0,
  }
}

function prepareTextBuffer(
  paneBuffers: WorkbookPaneBufferCache,
  paneCache: WorkbookPaneBufferEntry,
  requiredCount: number,
): {
  readonly buffer: WorkbookPaneBufferEntry['textBuffer']
  readonly capacity: number
} {
  if (paneCache.textBuffer && paneCache.textCapacity >= requiredCount) {
    return { buffer: paneCache.textBuffer, capacity: paneCache.textCapacity }
  }
  releaseTextBuffer(paneBuffers, paneCache)
  const reused = paneBuffers.acquireTextBuffer(requiredCount)
  return {
    buffer: reused?.buffer ?? null,
    capacity: reused?.capacity ?? 0,
  }
}

function releaseRectBuffer(paneBuffers: WorkbookPaneBufferCache, paneCache: WorkbookPaneBufferEntry): void {
  if (!paneCache.rectBuffer) {
    return
  }
  paneBuffers.releaseRectBuffer(paneCache.rectBuffer, paneCache.rectCapacity)
  paneCache.rectBuffer = null
  paneCache.rectCapacity = 0
}

function releaseTextBuffer(paneBuffers: WorkbookPaneBufferCache, paneCache: WorkbookPaneBufferEntry): void {
  if (!paneCache.textBuffer) {
    return
  }
  paneBuffers.releaseTextBuffer(paneCache.textBuffer, paneCache.textCapacity)
  paneCache.textBuffer = null
  paneCache.textCapacity = 0
}

export function resolveGridTextPacketSignature(packet: GridScenePacketV2): string {
  return [buildTileGpuCacheKey(packet), packet.textCount, packet.textSignature].join(':')
}

export function resolveGridTextTileSignature(tile: GridRenderTile): string {
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

export function resolveGridRectPacketSignature(input: {
  readonly frameWidth: number
  readonly frameHeight: number
  readonly packedScene: GridScenePacketV2
  readonly decorationRects?: readonly TextDecorationRect[] | undefined
}): string {
  const decorationRects = input.decorationRects ?? []
  return [
    buildTileGpuCacheKey(input.packedScene),
    input.packedScene.rectCount,
    input.packedScene.fillRectCount,
    input.packedScene.borderRectCount,
    input.packedScene.rectSignature,
    input.packedScene.textSignature,
    input.frameWidth,
    input.frameHeight,
    decorationRects.length,
  ].join(':')
}

export function resolveGridRectTileSignature(input: {
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

function resolveHeaderTextSignature(pane: GridHeaderPaneState): string {
  return ['header-text-v3', pane.paneId, pane.textCount, pane.textSignature].join(':')
}

function resolveHeaderRectSignature(input: {
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

function resolveOverlayRectSignature(overlay: DynamicGridOverlayBatchV3): string {
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

function buildTextInstanceData(input: { packet: GridScenePacketV2; atlas: ReturnType<typeof createGlyphAtlas> }): {
  floats: Float32Array
  quadCount: number
} {
  return buildTextQuadsFromRuns(input.packet.textRuns, input.atlas)
}

function buildRectInstanceDataFromPacket(input: {
  readonly frame: Rectangle
  readonly packet: GridScenePacketV2
  readonly decorationRects?: readonly TextDecorationRect[]
}): { readonly floats: Float32Array; readonly count: number } {
  const decorationRects = input.decorationRects ?? []
  const total = input.packet.rectCount + decorationRects.length
  if (decorationRects.length === 0) {
    return { count: total, floats: input.packet.rectInstances }
  }
  const floats = new Float32Array(Math.max(1, total) * RECT_INSTANCE_FLOAT_COUNT)
  const clipX = 0
  const clipY = 0
  const clipX1 = input.frame.width
  const clipY1 = input.frame.height
  const packedFloatCount = input.packet.rectCount * RECT_INSTANCE_FLOAT_COUNT
  floats.set(input.packet.rectInstances.subarray(0, packedFloatCount), 0)
  const offset = packedFloatCount
  writeDecorationRects(floats, offset, decorationRects, clipX, clipY, clipX1, clipY1)
  return { count: total, floats }
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
