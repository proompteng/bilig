import type { GridGpuScene, GridTextScene, WorkbookPaneId } from './resident-pane-scene-types.js'
import type { Viewport } from '@bilig/protocol'

export const WORKER_PACKED_RECT_FLOAT_COUNT = 8
export const WORKER_PACKED_TEXT_ITEM_FLOAT_COUNT = 8

export interface WorkerPackedGridScenePacket {
  readonly generation: number
  readonly paneId: WorkbookPaneId
  readonly viewport: Viewport
  readonly rects: Float32Array
  readonly rectCount: number
  readonly textMetrics: Float32Array
  readonly textCount: number
}

export function packWorkerGridScenePacket(input: {
  readonly generation: number
  readonly paneId: WorkbookPaneId
  readonly viewport: Viewport
  readonly gpuScene: GridGpuScene
  readonly textScene: GridTextScene
}): WorkerPackedGridScenePacket {
  return {
    generation: input.generation,
    paneId: input.paneId,
    rectCount: input.gpuScene.fillRects.length + input.gpuScene.borderRects.length,
    rects: packRects(input.gpuScene),
    textCount: input.textScene.items.length,
    textMetrics: packTextMetrics(input.textScene),
    viewport: input.viewport,
  }
}

function packRects(scene: GridGpuScene): Float32Array {
  const rectCount = scene.fillRects.length + scene.borderRects.length
  const floats = new Float32Array(Math.max(1, rectCount) * WORKER_PACKED_RECT_FLOAT_COUNT)
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
    offset += WORKER_PACKED_RECT_FLOAT_COUNT
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
    offset += WORKER_PACKED_RECT_FLOAT_COUNT
  }
  return floats
}

function packTextMetrics(scene: GridTextScene): Float32Array {
  const floats = new Float32Array(Math.max(1, scene.items.length) * WORKER_PACKED_TEXT_ITEM_FLOAT_COUNT)
  scene.items.forEach((item, index) => {
    const offset = index * WORKER_PACKED_TEXT_ITEM_FLOAT_COUNT
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
