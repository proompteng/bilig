import type { Viewport } from '@bilig/protocol'
import type { GridGpuScene } from '../gridGpuScene.js'
import type { GridTextScene } from '../gridTextScene.js'
import type { Rectangle } from '../gridTypes.js'
import type { GridScenePacketV2 } from '../renderer-v2/scene-packet-v2.js'

export type WorkbookPaneId = 'body' | 'top' | 'left' | 'corner'

export interface WorkbookPaneSurfaceSize {
  readonly width: number
  readonly height: number
}

export interface WorkbookPaneScenePacket {
  readonly generation: number
  readonly paneId: WorkbookPaneId
  readonly viewport: Viewport
  readonly surfaceSize: WorkbookPaneSurfaceSize
  readonly gpuScene: GridGpuScene
  readonly textScene: GridTextScene
  readonly packedScene?: GridScenePacketV2
}

export interface WorkbookPaneSceneRequest {
  readonly sheetName: string
  readonly residentViewport: Viewport
  readonly freezeRows: number
  readonly freezeCols: number
  readonly selectedCell: {
    readonly col: number
    readonly row: number
  }
  readonly selectionRange: Pick<Rectangle, 'x' | 'y' | 'width' | 'height'> | null
  readonly editingCell: {
    readonly col: number
    readonly row: number
  } | null
}

export interface WorkbookPaneRenderState extends WorkbookPaneScenePacket {
  readonly frame: Rectangle
  readonly contentOffset: {
    readonly x: number
    readonly y: number
  }
}

export interface WorkbookPaneScrollAxes {
  readonly x: boolean
  readonly y: boolean
}

export interface WorkbookRenderPaneState {
  readonly paneId: string
  readonly generation: number
  readonly viewport?: Viewport | undefined
  readonly packedScene?: GridScenePacketV2 | undefined
  readonly frame: Rectangle
  readonly surfaceSize: WorkbookPaneSurfaceSize
  readonly contentOffset: {
    readonly x: number
    readonly y: number
  }
  readonly scrollAxes: WorkbookPaneScrollAxes
  readonly gpuScene: GridGpuScene
  readonly textScene: GridTextScene
}
