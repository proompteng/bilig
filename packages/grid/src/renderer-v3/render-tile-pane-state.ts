import type { Viewport } from '@bilig/protocol'
import type { Rectangle } from '../gridTypes.js'
import type { GridRenderTile } from './render-tile-source.js'

export interface WorkbookTilePaneSurfaceSize {
  readonly width: number
  readonly height: number
}

export interface WorkbookTilePaneScrollAxes {
  readonly x: boolean
  readonly y: boolean
}

export interface WorkbookRenderTilePaneState {
  readonly paneId: string
  readonly generation: number
  readonly viewport: Viewport
  readonly tile: GridRenderTile
  readonly frame: Rectangle
  readonly surfaceSize: WorkbookTilePaneSurfaceSize
  readonly contentOffset: {
    readonly x: number
    readonly y: number
  }
  readonly scrollAxes: WorkbookTilePaneScrollAxes
}
