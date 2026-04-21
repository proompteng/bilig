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
} from './typegpu-renderer.js'
import { buildTextDecorationRectsFromScene, buildTextQuadsFromScene, type TextDecorationRect } from './text-quad-buffer.js'
import type { WorkbookPaneBufferEntry, WorkbookPaneBufferCache } from './pane-buffer-cache.js'
import type { createGlyphAtlas } from './glyph-atlas.js'
import type { WorkbookRenderPaneState } from './pane-scene-types.js'

const RECT_INSTANCE_FLOAT_COUNT = 20

export function syncTypeGpuPaneResources(input: {
  readonly artifacts: TypeGpuRendererArtifacts
  readonly atlas: ReturnType<typeof createGlyphAtlas>
  readonly paneBuffers: WorkbookPaneBufferCache
  readonly panes: readonly WorkbookRenderPaneState[]
}): void {
  const paneIds = new Set(input.panes.map((pane) => pane.paneId))
  input.paneBuffers.pruneExcept(paneIds)

  input.panes.forEach((pane) => {
    const paneCache = input.paneBuffers.get(pane.paneId)
    const textSceneChanged = paneCache.textScene !== pane.textScene
    if (textSceneChanged) {
      syncTextResource({
        artifacts: input.artifacts,
        atlas: input.atlas,
        pane,
        paneCache,
      })
    }
    if (paneCache.rectScene !== pane.gpuScene || textSceneChanged) {
      syncRectResource({
        artifacts: input.artifacts,
        pane,
        paneCache,
      })
    }
  })
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
}

function syncRectResource(input: {
  readonly artifacts: TypeGpuRendererArtifacts
  readonly pane: WorkbookRenderPaneState
  readonly paneCache: WorkbookPaneBufferEntry
}): void {
  const decorationRects = input.paneCache.decorationRects ?? []
  const rectFloats = buildRectInstanceData({
    frame: input.pane.frame,
    scene: input.pane.gpuScene,
    ...(decorationRects.length > 0 ? { decorationRects } : {}),
  })
  const rectBuffer = ensureTypeGpuVertexBuffer(
    input.artifacts.root,
    WORKBOOK_RECT_INSTANCE_LAYOUT,
    input.paneCache.rectBuffer,
    input.paneCache.rectCapacity,
    input.pane.gpuScene.fillRects.length + input.pane.gpuScene.borderRects.length + decorationRects.length,
  )
  input.paneCache.rectBuffer = rectBuffer.buffer
  input.paneCache.rectCapacity = rectBuffer.capacity
  input.paneCache.rectCount = input.pane.gpuScene.fillRects.length + input.pane.gpuScene.borderRects.length + decorationRects.length
  writeTypeGpuVertexBuffer(input.paneCache.rectBuffer, rectFloats)
  input.paneCache.rectScene = input.pane.gpuScene
}

function buildTextInstanceData(input: { textScene: GridTextScene; atlas: ReturnType<typeof createGlyphAtlas> }): {
  floats: Float32Array
  quadCount: number
} {
  return buildTextQuadsFromScene(input.textScene.items, input.atlas)
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
    offset += RECT_INSTANCE_FLOAT_COUNT
  }

  return floats
}
