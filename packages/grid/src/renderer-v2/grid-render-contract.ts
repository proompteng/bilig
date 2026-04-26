import type { Viewport } from '@bilig/protocol'
import type { GridGpuScene } from '../gridGpuScene.js'
import type { GridTextScene } from '../gridTextScene.js'
import type { Rectangle } from '../gridTypes.js'

export interface GridAxisSnapshot {
  readonly index: number
  readonly offset: number
  readonly size: number
}

export interface GridCameraSnapshot {
  readonly scrollLeft: number
  readonly scrollTop: number
  readonly tx: number
  readonly ty: number
  readonly visibleViewport: Viewport
  readonly residentViewport: Viewport
  readonly viewportWidth: number
  readonly viewportHeight: number
  readonly dpr: number
  readonly velocityX: number
  readonly velocityY: number
  readonly updatedAt: number
}

export interface GridRenderFrame {
  readonly camera: GridCameraSnapshot
  readonly inputAt: number
  readonly frameAt: number
}

export interface GridGpuCounters {
  readonly configureCount: number
  readonly submitCount: number
  readonly drawCalls: number
  readonly paneDraws: number
  readonly uniformWriteBytes: number
  readonly vertexUploadBytes: number
  readonly bufferAllocations: number
  readonly bufferAllocationBytes: number
  readonly atlasUploadBytes: number
  readonly surfaceResizes: number
  readonly tileMisses: number
  readonly scenePacketsApplied: number
}

export interface GridRenderStats {
  readonly inputToDrawMs: readonly number[]
  readonly gpu: GridGpuCounters
}

export interface GridResidentTileKey {
  readonly sheetName: string
  readonly paneId: string
  readonly rowStart: number
  readonly rowEnd: number
  readonly colStart: number
  readonly colEnd: number
}

export interface GridTextLayoutPacket {
  readonly version: number
  readonly floats: Float32Array
  readonly quadCount: number
}

export interface GridPaneRenderPacket {
  readonly key: GridResidentTileKey
  readonly generation: number
  readonly frame: Rectangle
  readonly surfaceSize: {
    readonly width: number
    readonly height: number
  }
  readonly contentOffset: {
    readonly x: number
    readonly y: number
  }
  readonly gpuScene: GridGpuScene
  readonly textScene: GridTextScene
}

export interface GridResidentScenePacket {
  readonly key: GridResidentTileKey
  readonly generation: number
  readonly panes: readonly GridPaneRenderPacket[]
}
