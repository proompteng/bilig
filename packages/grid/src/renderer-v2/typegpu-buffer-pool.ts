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

const RECT_INSTANCE_FLOAT_COUNT = GRID_SCENE_PACKET_V2_RECT_INSTANCE_FLOAT_COUNT

export function syncTypeGpuPaneResources(input: {
  readonly artifacts: TypeGpuRendererArtifacts
  readonly atlas: ReturnType<typeof createGlyphAtlas>
  readonly paneBuffers: WorkbookPaneBufferCache
  readonly panes: readonly WorkbookRenderPaneState[]
  readonly retainPanes?: readonly WorkbookRenderPaneState[] | undefined
}): void {
  const paneIds = new Set((input.retainPanes ?? input.panes).map(resolveWorkbookPaneBufferKey))
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

export function resolveWorkbookPaneBufferKey(pane: WorkbookRenderPaneState): string {
  return `${pane.paneId}:${buildTileGpuCacheKey(pane.packedScene)}`
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
