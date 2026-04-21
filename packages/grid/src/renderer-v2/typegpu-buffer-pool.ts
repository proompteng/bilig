import { parseGpuColor, type GridGpuScene } from '../gridGpuScene.js'
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
import { GRID_SCENE_PACKET_V2_RECT_FLOAT_COUNT, type GridScenePacketV2 } from './scene-packet-v2.js'
import { buildTileGpuCacheKey } from './tile-gpu-cache.js'

const RECT_INSTANCE_FLOAT_COUNT = 20

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
    const deferTextUpload = input.deferTextUploads === true && pane.packedScene !== undefined
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
            scene: pane.gpuScene,
          }))
        : resolveGridRectSceneSignature({
            decorationRects: paneCache.decorationRects ?? [],
            frame: pane.frame,
            packedScene: pane.packedScene,
            scene: pane.gpuScene,
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

export function resolveWorkbookPaneBufferKey(pane: WorkbookRenderPaneState): string {
  return pane.packedScene ? buildTileGpuCacheKey(pane.packedScene) : pane.paneId
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
  writeTypeGpuVertexBuffer(input.paneCache.textBuffer, textPayload.floats)
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
  const rectPayload = input.pane.packedScene
    ? buildRectInstanceDataFromPacket({
        decorationRects,
        frame: input.pane.frame,
        packet: input.pane.packedScene,
      })
    : {
        count: input.pane.gpuScene.fillRects.length + input.pane.gpuScene.borderRects.length + decorationRects.length,
        floats: buildRectInstanceData({
          frame: input.pane.frame,
          scene: input.pane.gpuScene,
          ...(decorationRects.length > 0 ? { decorationRects } : {}),
        }),
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
  writeTypeGpuVertexBuffer(input.paneCache.rectBuffer, rectPayload.floats)
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
  readonly scene: GridGpuScene
  readonly packedScene?: GridScenePacketV2 | undefined
  readonly decorationRects?: readonly TextDecorationRect[] | undefined
}): string {
  let hash = createHash()
  hash = mixNumber(hash, input.frame.width)
  hash = mixNumber(hash, input.frame.height)
  if (input.packedScene) {
    hash = mixNumber(hash, input.packedScene.generation)
    hash = mixNumber(hash, input.packedScene.rectCount)
    hash = mixNumber(hash, input.packedScene.fillRectCount)
    hash = mixNumber(hash, input.packedScene.borderRectCount)
    for (let index = 0; index < input.packedScene.rectCount * GRID_SCENE_PACKET_V2_RECT_FLOAT_COUNT; index += 1) {
      hash = mixNumber(hash, input.packedScene.rects[index] ?? 0)
    }
  } else {
    hash = mixNumber(hash, input.scene.fillRects.length)
    hash = mixNumber(hash, input.scene.borderRects.length)
    for (const rect of input.scene.fillRects) {
      hash = mixGpuRect(hash, rect)
    }
    for (const rect of input.scene.borderRects) {
      hash = mixGpuRect(hash, rect)
    }
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

function mixGpuRect(hash: number, rect: GridGpuScene['fillRects'][number]): number {
  hash = mixNumber(hash, rect.x)
  hash = mixNumber(hash, rect.y)
  hash = mixNumber(hash, rect.width)
  hash = mixNumber(hash, rect.height)
  hash = mixNumber(hash, rect.color.r)
  hash = mixNumber(hash, rect.color.g)
  hash = mixNumber(hash, rect.color.b)
  return mixNumber(hash, rect.color.a)
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
  const floats = new Float32Array(Math.max(1, total) * RECT_INSTANCE_FLOAT_COUNT)
  const clipX = 0
  const clipY = 0
  const clipX1 = input.frame.width
  const clipY1 = input.frame.height
  let offset = 0
  for (let index = 0; index < input.packet.rectCount; index += 1) {
    const source = index * GRID_SCENE_PACKET_V2_RECT_FLOAT_COUNT
    const isFill = index < input.packet.fillRectCount
    floats[offset + 0] = input.packet.rects[source + 0] ?? 0
    floats[offset + 1] = input.packet.rects[source + 1] ?? 0
    floats[offset + 2] = input.packet.rects[source + 2] ?? 0
    floats[offset + 3] = input.packet.rects[source + 3] ?? 0
    floats[offset + 4] = isFill ? (input.packet.rects[source + 4] ?? 0) : 0
    floats[offset + 5] = isFill ? (input.packet.rects[source + 5] ?? 0) : 0
    floats[offset + 6] = isFill ? (input.packet.rects[source + 6] ?? 0) : 0
    floats[offset + 7] = isFill ? (input.packet.rects[source + 7] ?? 0) : 0
    floats[offset + 8] = isFill ? 0 : (input.packet.rects[source + 4] ?? 0)
    floats[offset + 9] = isFill ? 0 : (input.packet.rects[source + 5] ?? 0)
    floats[offset + 10] = isFill ? 0 : (input.packet.rects[source + 6] ?? 0)
    floats[offset + 11] = isFill ? 0 : (input.packet.rects[source + 7] ?? 0)
    floats[offset + 12] = isFill && (input.packet.rects[source + 7] ?? 0) < 0.2 ? 2 : 0
    floats[offset + 13] = isFill ? 0 : 1
    floats[offset + 14] = 0
    floats[offset + 15] = 0
    floats[offset + 16] = clipX
    floats[offset + 17] = clipY
    floats[offset + 18] = clipX1
    floats[offset + 19] = clipY1
    offset += RECT_INSTANCE_FLOAT_COUNT
  }
  writeDecorationRects(floats, offset, decorationRects, clipX, clipY, clipX1, clipY1)
  return { count: total, floats }
}

function buildRectInstanceData(input: {
  frame: Rectangle
  scene: GridGpuScene
  decorationRects?: readonly TextDecorationRect[]
}): Float32Array {
  const rects = input.scene.fillRects
  const borders = input.scene.borderRects
  const decorationRects = input.decorationRects ?? []
  const total = rects.length + borders.length + decorationRects.length
  const floats = new Float32Array(Math.max(1, total) * RECT_INSTANCE_FLOAT_COUNT)

  const clipX = 0
  const clipY = 0
  const clipX1 = input.frame.width
  const clipY1 = input.frame.height

  let offset = 0
  for (const rect of rects) {
    floats[offset + 0] = rect.x
    floats[offset + 1] = rect.y
    floats[offset + 2] = rect.width
    floats[offset + 3] = rect.height
    floats[offset + 4] = rect.color.r
    floats[offset + 5] = rect.color.g
    floats[offset + 6] = rect.color.b
    floats[offset + 7] = rect.color.a
    floats[offset + 8] = 0
    floats[offset + 9] = 0
    floats[offset + 10] = 0
    floats[offset + 11] = 0
    floats[offset + 12] = rect.color.a < 0.2 ? 2 : 0
    floats[offset + 13] = 0
    floats[offset + 14] = 0
    floats[offset + 15] = 0
    floats[offset + 16] = clipX
    floats[offset + 17] = clipY
    floats[offset + 18] = clipX1
    floats[offset + 19] = clipY1
    offset += RECT_INSTANCE_FLOAT_COUNT
  }
  for (const rect of borders) {
    floats[offset + 0] = rect.x
    floats[offset + 1] = rect.y
    floats[offset + 2] = rect.width
    floats[offset + 3] = rect.height
    floats[offset + 4] = 0
    floats[offset + 5] = 0
    floats[offset + 6] = 0
    floats[offset + 7] = 0
    floats[offset + 8] = rect.color.r
    floats[offset + 9] = rect.color.g
    floats[offset + 10] = rect.color.b
    floats[offset + 11] = rect.color.a
    floats[offset + 12] = 0
    floats[offset + 13] = 1
    floats[offset + 14] = 0
    floats[offset + 15] = 0
    floats[offset + 16] = clipX
    floats[offset + 17] = clipY
    floats[offset + 18] = clipX1
    floats[offset + 19] = clipY1
    offset += RECT_INSTANCE_FLOAT_COUNT
  }

  for (const rect of decorationRects) {
    offset = writeDecorationRect(floats, offset, rect, clipX, clipY, clipX1, clipY1)
  }

  return floats
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
