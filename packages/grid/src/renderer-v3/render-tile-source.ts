import type { Viewport } from '@bilig/protocol'
import type { GridTilePacketV3 } from './tile-packet-v3.js'

export type GridRenderTilePaneKind = 'body' | 'frozenTop' | 'frozenLeft' | 'frozenCorner'

export interface GridRenderTileDeltaSubscription extends Viewport {
  readonly sheetId: number
  readonly sheetName: string
  readonly cameraSeq?: number | undefined
  readonly dprBucket?: number | undefined
  readonly initialDelta?: 'full' | 'none' | undefined
}

export interface GridRenderTileCoord {
  readonly sheetId: number
  readonly paneKind: GridRenderTilePaneKind
  readonly rowTile: number
  readonly colTile: number
  readonly dprBucket: number
}

export interface GridRenderTileVersion {
  readonly axisX: number
  readonly axisY: number
  readonly values: number
  readonly styles: number
  readonly text: number
  readonly freeze: number
}

export interface GridRenderTileTextRun {
  readonly text: string
  readonly x: number
  readonly y: number
  readonly width: number
  readonly height: number
  readonly clipX: number
  readonly clipY: number
  readonly clipWidth: number
  readonly clipHeight: number
  readonly align?: 'left' | 'center' | 'right'
  readonly wrap?: boolean
  readonly font: string
  readonly fontSize: number
  readonly color: string
  readonly underline: boolean
  readonly strike: boolean
}

export interface GridRenderTile {
  readonly tileId: number
  readonly packet?: GridTilePacketV3 | undefined
  readonly coord: GridRenderTileCoord
  readonly version: GridRenderTileVersion
  readonly bounds: Viewport
  readonly rectInstances: Float32Array
  readonly rectCount: number
  readonly textMetrics: Float32Array
  readonly textRuns: readonly GridRenderTileTextRun[]
  readonly textCount: number
  readonly lastBatchId: number
  readonly lastCameraSeq: number
}

export interface GridRenderTileSceneChange {
  readonly batchId: number
  readonly cameraSeq: number
  readonly changedTileIds: readonly number[]
  readonly invalidatedTileIds: readonly number[]
  readonly structural: boolean
}

export interface GridRenderTileSource {
  subscribeRenderTileDeltas(
    subscription: GridRenderTileDeltaSubscription,
    listener: (change: GridRenderTileSceneChange) => void,
  ): () => void
  peekRenderTile(tileId: number): GridRenderTile | null
}
