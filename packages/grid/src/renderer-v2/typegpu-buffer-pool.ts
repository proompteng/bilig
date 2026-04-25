import { parseGpuColor } from '../gridGpuScene.js'
import type { GridTextScene } from '../gridTextScene.js'
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
import { buildTextDecorationRectsFromScene, buildTextQuadsFromScene, type TextDecorationRect } from './line-text-quad-buffer.js'
import type { WorkbookPaneBufferEntry, WorkbookPaneBufferCache } from './pane-buffer-cache.js'
import type { createGlyphAtlas } from './typegpu-atlas-manager.js'
import type { WorkbookRenderPaneState } from '../renderer/pane-scene-types.js'
import {
  GRID_SCENE_PACKET_V2_RECT_FLOAT_COUNT,
  GRID_SCENE_PACKET_V2_RECT_INSTANCE_FLOAT_COUNT,
  type GridScenePacketV2,
} from './scene-packet-v2.js'
import { buildTileGpuCacheKey } from './tile-gpu-cache.js'

const RECT_INSTANCE_FLOAT_COUNT = GRID_SCENE_PACKET_V2_RECT_INSTANCE_FLOAT_COUNT

export function syncTypeGpuPaneResources(input: {
  readonly artifacts: TypeGpuRendererArtifacts
  readonly atlas: ReturnType<typeof createGlyphAtlas>
  readonly paneBuffers: WorkbookPaneBufferCache
  readonly panes: readonly WorkbookRenderPaneState[]
  readonly retainPanes?: readonly WorkbookRenderPaneState[] | undefined
  readonly deferTextUploads?: boolean | undefined
}): void {
  const paneIds = new Set((input.retainPanes ?? input.panes).map(resolveWorkbookPaneBufferKey))
  input.paneBuffers.pruneExcept(paneIds)

  input.panes.forEach((pane) => {
    const paneCache = input.paneBuffers.get(resolveWorkbookPaneBufferKey(pane))
    const textSignature =
      paneCache.textScene === pane.textScene && paneCache.textSignature !== null
        ? paneCache.textSignature
        : resolveGridTextSceneSignature(pane.textScene)
    const textSceneChanged = paneCache.textSignature !== textSignature
    const deferTextUpload = shouldDeferPaneTextUpload({
      currentTextCount: paneCache.textCount,
      currentTextSignature: paneCache.textSignature,
      deferTextUploads: input.deferTextUploads === true,
      hasTextBuffer: paneCache.textBuffer !== null,
      nextTextItemCount: pane.textScene.items.length,
    })
    if (textSceneChanged && !deferTextUpload) {
      syncTextResource({
        artifacts: input.artifacts,
        atlas: input.atlas,
        pane,
        paneCache,
        textSignature,
      })
    }
    const rectSignature =
      paneCache.rectScene === pane.gpuScene && !textSceneChanged
        ? (paneCache.rectSignature ??
          resolveGridRectSceneSignature({
            decorationRects: paneCache.decorationRects ?? [],
            frame: pane.frame,
            packedScene: pane.packedScene,
          }))
        : resolveGridRectSceneSignature({
            decorationRects: paneCache.decorationRects ?? [],
            frame: pane.frame,
            packedScene: pane.packedScene,
          })
    if (paneCache.rectSignature !== rectSignature) {
      syncRectResource({
        artifacts: input.artifacts,
        pane,
        paneCache,
        rectSignature,
      })
    }
  })
}

export function shouldDeferPaneTextUpload(input: {
  readonly deferTextUploads: boolean
  readonly currentTextSignature: string | null
  readonly hasTextBuffer: boolean
  readonly currentTextCount: number
  readonly nextTextItemCount: number
}): boolean {
  if (!input.deferTextUploads) {
    return false
  }
  if (input.currentTextSignature === null) {
    return false
  }
  if (input.nextTextItemCount > 0 && (!input.hasTextBuffer || input.currentTextCount === 0)) {
    return false
  }
  return true
}

export function resolveWorkbookPaneBufferKey(pane: WorkbookRenderPaneState): string {
  return buildTileGpuCacheKey(pane.packedScene)
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
  readonly paneCache: WorkbookPaneBufferEntry
  readonly textSignature: string
}): void {
  input.paneCache.decorationRects = buildTextDecorationRectsFromScene(input.pane.textScene.items, input.atlas)
  const textPayload = buildTextInstanceData({
    atlas: input.atlas,
    textScene: input.pane.textScene,
  })
  if (textPayload.quadCount === 0) {
    input.paneCache.textCount = 0
    input.paneCache.textScene = input.pane.textScene
    input.paneCache.textSignature = input.textSignature
    return
  }
  const textBuffer = ensureTypeGpuVertexBuffer(
    input.artifacts.root,
    WORKBOOK_TEXT_INSTANCE_LAYOUT,
    input.paneCache.textBuffer,
    input.paneCache.textCapacity,
    textPayload.quadCount,
  )
  input.paneCache.textBuffer = textBuffer.buffer
  input.paneCache.textCapacity = textBuffer.capacity
  input.paneCache.textCount = textPayload.quadCount
  writeTypeGpuVertexBuffer(input.paneCache.textBuffer, textPayload.floats, `text:${resolveWorkbookPaneBufferKey(input.pane)}`)
  input.paneCache.textScene = input.pane.textScene
  input.paneCache.textSignature = input.textSignature
}

function syncRectResource(input: {
  readonly artifacts: TypeGpuRendererArtifacts
  readonly pane: WorkbookRenderPaneState
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
    input.paneCache.rectCount = 0
    input.paneCache.rectScene = input.pane.gpuScene
    input.paneCache.rectSignature = input.rectSignature
    return
  }
  const rectBuffer = ensureTypeGpuVertexBuffer(
    input.artifacts.root,
    WORKBOOK_RECT_INSTANCE_LAYOUT,
    input.paneCache.rectBuffer,
    input.paneCache.rectCapacity,
    rectPayload.count,
  )
  input.paneCache.rectBuffer = rectBuffer.buffer
  input.paneCache.rectCapacity = rectBuffer.capacity
  input.paneCache.rectCount = rectPayload.count
  writeTypeGpuVertexBuffer(input.paneCache.rectBuffer, rectPayload.floats, `rect:${resolveWorkbookPaneBufferKey(input.pane)}`)
  input.paneCache.rectScene = input.pane.gpuScene
  input.paneCache.rectSignature = input.rectSignature
}

export function resolveGridTextSceneSignature(scene: GridTextScene): string {
  let hash = createHash()
  hash = mixNumber(hash, scene.items.length)
  for (const item of scene.items) {
    hash = mixNumber(hash, item.x)
    hash = mixNumber(hash, item.y)
    hash = mixNumber(hash, item.width)
    hash = mixNumber(hash, item.height)
    hash = mixNumber(hash, item.clipInsetTop)
    hash = mixNumber(hash, item.clipInsetRight)
    hash = mixNumber(hash, item.clipInsetBottom)
    hash = mixNumber(hash, item.clipInsetLeft)
    hash = mixString(hash, item.text)
    hash = mixString(hash, item.align)
    hash = mixString(hash, item.color)
    hash = mixString(hash, item.font)
    hash = mixNumber(hash, item.fontSize)
    hash = mixNumber(hash, item.wrap ? 1 : 0)
    hash = mixNumber(hash, item.underline ? 1 : 0)
    hash = mixNumber(hash, item.strike ? 1 : 0)
  }
  return hash.toString(36)
}

export function resolveGridRectSceneSignature(input: {
  readonly frame: Rectangle
  readonly packedScene: GridScenePacketV2
  readonly decorationRects?: readonly TextDecorationRect[] | undefined
}): string {
  let hash = createHash()
  hash = mixNumber(hash, input.frame.width)
  hash = mixNumber(hash, input.frame.height)
  hash = mixNumber(hash, input.packedScene.rectCount)
  hash = mixNumber(hash, input.packedScene.fillRectCount)
  hash = mixNumber(hash, input.packedScene.borderRectCount)
  for (let index = 0; index < input.packedScene.rectCount * GRID_SCENE_PACKET_V2_RECT_FLOAT_COUNT; index += 1) {
    hash = mixNumber(hash, input.packedScene.rects[index] ?? 0)
  }
  const decorationRects = input.decorationRects ?? []
  hash = mixNumber(hash, decorationRects.length)
  for (const rect of decorationRects) {
    hash = mixNumber(hash, rect.x)
    hash = mixNumber(hash, rect.y)
    hash = mixNumber(hash, rect.width)
    hash = mixNumber(hash, rect.height)
    hash = mixString(hash, rect.color)
  }
  return hash.toString(36)
}

function createHash(): number {
  return 2_166_136_261
}

function mixString(hash: number, value: string): number {
  let next = hash
  for (let index = 0; index < value.length; index += 1) {
    next = mixInteger(next, value.charCodeAt(index))
  }
  return next
}

function mixNumber(hash: number, value: number): number {
  return mixInteger(hash, Math.round(value * 1_000))
}

function mixInteger(hash: number, value: number): number {
  return Math.imul((hash ^ value) >>> 0, 16_777_619) >>> 0
}

function buildTextInstanceData(input: { textScene: GridTextScene; atlas: ReturnType<typeof createGlyphAtlas> }): {
  floats: Float32Array
  quadCount: number
} {
  return buildTextQuadsFromScene(input.textScene.items, input.atlas)
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
