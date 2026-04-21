import type { Viewport } from '@bilig/protocol'

export const GRID_SCENE_PACKET_V2_MAGIC = 'bilig.grid.scene.v2'
export const GRID_SCENE_PACKET_V2_VERSION = 1
export const GRID_SCENE_PACKET_V2_RECT_FLOAT_COUNT = 8
export const GRID_SCENE_PACKET_V2_TEXT_METRIC_FLOAT_COUNT = 8
export type GridScenePacketPaneId = 'body' | 'top' | 'left' | 'corner'

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
