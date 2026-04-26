import type { Viewport } from '@bilig/protocol'
import type { WorkerPackedGridScenePacket } from './worker-runtime-render-packet.js'

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
  readonly packedScene: WorkerPackedGridScenePacket
}

export interface WorkbookPaneSceneRequest {
  readonly sheetName: string
  readonly residentViewport: Viewport
  readonly freezeRows: number
  readonly freezeCols: number
  readonly dprBucket?: number | undefined
  readonly requestSeq?: number | undefined
  readonly cameraSeq?: number | undefined
  readonly sceneRevision?: number | undefined
  readonly priority?: number | undefined
  readonly reason?: 'visible' | 'prefetch' | 'edit' | 'resize' | 'sheet-switch' | undefined
}
