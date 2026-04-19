import type { Viewport } from '@bilig/protocol'

export type WorkbookPaneId = 'body' | 'top' | 'left' | 'corner'

export interface WorkbookPaneSurfaceSize {
  readonly width: number
  readonly height: number
}

export interface GridGpuColor {
  readonly r: number
  readonly g: number
  readonly b: number
  readonly a: number
}

export interface GridGpuRect {
  readonly x: number
  readonly y: number
  readonly width: number
  readonly height: number
  readonly color: GridGpuColor
}

export interface GridGpuScene {
  readonly fillRects: readonly GridGpuRect[]
  readonly borderRects: readonly GridGpuRect[]
}

export interface GridTextItem {
  readonly x: number
  readonly y: number
  readonly width: number
  readonly height: number
  readonly clipInsetTop: number
  readonly clipInsetRight: number
  readonly clipInsetBottom: number
  readonly clipInsetLeft: number
  readonly text: string
  readonly align: 'left' | 'center' | 'right'
  readonly wrap: boolean
  readonly color: string
  readonly font: string
  readonly fontSize: number
  readonly underline: boolean
  readonly strike: boolean
}

export interface GridTextScene {
  readonly items: readonly GridTextItem[]
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
  readonly selectionRange: {
    readonly x: number
    readonly y: number
    readonly width: number
    readonly height: number
  } | null
  readonly editingCell: {
    readonly col: number
    readonly row: number
  } | null
}
