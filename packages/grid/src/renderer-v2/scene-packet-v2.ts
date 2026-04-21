import type { Viewport } from '@bilig/protocol'
import type { GridGpuScene } from '../gridGpuScene.js'
import type { GridTextScene } from '../gridTextScene.js'

export const GRID_SCENE_PACKET_V2_MAGIC = 'bilig.grid.scene.v2'
export const GRID_SCENE_PACKET_V2_VERSION = 1
export const GRID_SCENE_PACKET_V2_RECT_FLOAT_COUNT = 8
export const GRID_SCENE_PACKET_V2_TEXT_METRIC_FLOAT_COUNT = 8
export type GridScenePacketPaneId = 'body' | 'top' | 'left' | 'corner' | 'top-frozen' | 'top-body' | 'left-frozen' | 'left-body' | 'overlay'

export interface GridScenePacketV2 {
  readonly magic: typeof GRID_SCENE_PACKET_V2_MAGIC
  readonly version: typeof GRID_SCENE_PACKET_V2_VERSION
  readonly generation: number
  readonly sheetName: string
  readonly paneId: GridScenePacketPaneId
  readonly viewport: Viewport
  readonly surfaceSize: {
    readonly width: number
    readonly height: number
  }
  readonly rects: Float32Array
  readonly rectCount: number
  readonly fillRectCount: number
  readonly borderRectCount: number
  readonly textMetrics: Float32Array
  readonly textCount: number
}

export function packGridScenePacketV2(input: {
  readonly generation: number
  readonly sheetName: string
  readonly paneId: GridScenePacketPaneId
  readonly viewport: Viewport
  readonly surfaceSize: { readonly width: number; readonly height: number }
  readonly gpuScene: GridGpuScene
  readonly textScene: GridTextScene
}): GridScenePacketV2 {
  return {
    generation: input.generation,
    borderRectCount: input.gpuScene.borderRects.length,
    fillRectCount: input.gpuScene.fillRects.length,
    magic: GRID_SCENE_PACKET_V2_MAGIC,
    paneId: input.paneId,
    rectCount: input.gpuScene.fillRects.length + input.gpuScene.borderRects.length,
    rects: packRects(input.gpuScene),
    sheetName: input.sheetName,
    surfaceSize: input.surfaceSize,
    textCount: input.textScene.items.length,
    textMetrics: packTextMetrics(input.textScene),
    version: GRID_SCENE_PACKET_V2_VERSION,
    viewport: input.viewport,
  }
}

function packRects(scene: GridGpuScene): Float32Array {
  const rectCount = scene.fillRects.length + scene.borderRects.length
  const floats = new Float32Array(Math.max(1, rectCount) * GRID_SCENE_PACKET_V2_RECT_FLOAT_COUNT)
  let offset = 0
  for (const rect of scene.fillRects) {
    floats[offset + 0] = rect.x
    floats[offset + 1] = rect.y
    floats[offset + 2] = rect.width
    floats[offset + 3] = rect.height
    floats[offset + 4] = rect.color.r
    floats[offset + 5] = rect.color.g
    floats[offset + 6] = rect.color.b
    floats[offset + 7] = rect.color.a
    offset += GRID_SCENE_PACKET_V2_RECT_FLOAT_COUNT
  }
  for (const rect of scene.borderRects) {
    floats[offset + 0] = rect.x
    floats[offset + 1] = rect.y
    floats[offset + 2] = rect.width
    floats[offset + 3] = rect.height
    floats[offset + 4] = rect.color.r
    floats[offset + 5] = rect.color.g
    floats[offset + 6] = rect.color.b
    floats[offset + 7] = rect.color.a
    offset += GRID_SCENE_PACKET_V2_RECT_FLOAT_COUNT
  }
  return floats
}

function packTextMetrics(scene: GridTextScene): Float32Array {
  const floats = new Float32Array(Math.max(1, scene.items.length) * GRID_SCENE_PACKET_V2_TEXT_METRIC_FLOAT_COUNT)
  scene.items.forEach((item, index) => {
    const offset = index * GRID_SCENE_PACKET_V2_TEXT_METRIC_FLOAT_COUNT
    floats[offset + 0] = item.x
    floats[offset + 1] = item.y
    floats[offset + 2] = item.width
    floats[offset + 3] = item.height
    floats[offset + 4] = item.clipInsetTop
    floats[offset + 5] = item.clipInsetRight
    floats[offset + 6] = item.clipInsetBottom
    floats[offset + 7] = item.clipInsetLeft
  })
  return floats
}
