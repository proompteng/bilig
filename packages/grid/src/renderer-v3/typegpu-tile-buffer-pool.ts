import { parseGpuColor } from '../gridGpuScene.js'
import type { WorkbookPaneBufferCache, WorkbookPaneBufferEntry } from '../renderer-v2/pane-buffer-cache.js'
import { GRID_SCENE_PACKET_V2_RECT_INSTANCE_FLOAT_COUNT } from '../renderer-v2/scene-packet-v2.js'
import type { createGlyphAtlas } from '../renderer-v2/typegpu-atlas-manager.js'
import {
  WORKBOOK_RECT_INSTANCE_LAYOUT,
  WORKBOOK_TEXT_INSTANCE_LAYOUT,
  ensureTypeGpuVertexBuffer,
  type TypeGpuRendererArtifacts,
  writeTypeGpuVertexBuffer,
} from '../renderer-v2/typegpu-backend.js'
import { buildTextDecorationRectsFromRuns, buildTextQuadsFromRuns, type TextDecorationRect } from '../renderer-v2/line-text-quad-buffer.js'
import type { GridRenderTile } from './render-tile-source.js'
import type { WorkbookRenderTilePaneState } from './render-tile-pane-state.js'

const RECT_INSTANCE_FLOAT_COUNT = GRID_SCENE_PACKET_V2_RECT_INSTANCE_FLOAT_COUNT

export function syncTypeGpuTilePaneResourcesV3(input: {
  readonly artifacts: TypeGpuRendererArtifacts
  readonly atlas: ReturnType<typeof createGlyphAtlas>
  readonly paneBuffers: WorkbookPaneBufferCache
  readonly panes: readonly WorkbookRenderTilePaneState[]
  readonly retainPanes?: readonly WorkbookRenderTilePaneState[] | undefined
  readonly retainBufferKeys?: readonly string[] | undefined
}): void {
  const paneIds = new Set((input.retainPanes ?? input.panes).map(resolveWorkbookTilePaneBufferKeyV3))
  for (const key of input.retainBufferKeys ?? []) {
    paneIds.add(key)
  }
  input.paneBuffers.pruneExcept(paneIds)

  input.panes.forEach((pane) => {
    const paneCache = input.paneBuffers.get(resolveWorkbookTilePaneBufferKeyV3(pane))
    const textSignature = resolveGridTextTileSignatureV3(pane.tile)
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
    const rectSignature = resolveGridRectTileSignatureV3({
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

export function resolveWorkbookTilePaneBufferKeyV3(pane: Pick<WorkbookRenderTilePaneState, 'tile'>): string {
  return `tile:v3:${pane.tile.tileId}`
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
  writeTypeGpuVertexBuffer(input.paneCache.textBuffer, textPayload.floats, `tile-text:${resolveWorkbookTilePaneBufferKeyV3(input.pane)}`)
  input.paneCache.textSignature = input.textSignature
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
  writeTypeGpuVertexBuffer(input.paneCache.rectBuffer, rectPayload.floats, `tile-rect:${resolveWorkbookTilePaneBufferKeyV3(input.pane)}`)
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
