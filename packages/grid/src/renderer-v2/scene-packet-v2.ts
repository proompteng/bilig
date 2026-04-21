import type { Viewport } from '@bilig/protocol'
import type { WorkbookPaneId, WorkbookPaneSurfaceSize } from '../renderer/pane-scene-types.js'

export const GRID_SCENE_PACKET_V2_MAGIC = 'bilig.grid.scene.v2'
export const GRID_SCENE_PACKET_V2_VERSION = 1
export const GRID_SCENE_PACKET_V2_RECT_FLOAT_COUNT = 8
export const GRID_SCENE_PACKET_V2_TEXT_METRIC_FLOAT_COUNT = 8

export interface GridScenePacketV2 {
  readonly magic: typeof GRID_SCENE_PACKET_V2_MAGIC
  readonly version: typeof GRID_SCENE_PACKET_V2_VERSION
  readonly generation: number
  readonly sheetName: string
  readonly paneId: WorkbookPaneId
  readonly viewport: Viewport
  readonly surfaceSize: WorkbookPaneSurfaceSize
  readonly rects: Float32Array
  readonly rectCount: number
  readonly textMetrics: Float32Array
  readonly textCount: number
}
