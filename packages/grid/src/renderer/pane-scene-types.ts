import type { Viewport } from '@bilig/protocol'
import type { GridGpuScene } from '../gridGpuScene.js'
import type { GridTextScene } from '../gridTextScene.js'
import type { Rectangle } from '../gridTypes.js'

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
