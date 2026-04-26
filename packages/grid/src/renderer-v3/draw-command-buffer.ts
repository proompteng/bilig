import type { TileKey53 } from './tile-key.js'

export type PanePlacementKindV3 = 'body' | 'frozenRows' | 'frozenCols' | 'frozenCorner'

export interface PanePlacementV3 {
  readonly tileKey: TileKey53
  readonly pane: PanePlacementKindV3
  readonly clipX: number
  readonly clipY: number
  readonly clipWidth: number
  readonly clipHeight: number
  readonly translateX: number
  readonly translateY: number
  readonly z: number
}

export interface DrawCommandBufferSnapshotV3 {
  readonly frameSeq: number
  readonly placements: readonly PanePlacementV3[]
}

export class DrawCommandBufferV3 {
  private frameSeq = 0
  private readonly placements: PanePlacementV3[] = []

  beginFrame(): void {
    this.frameSeq += 1
    this.placements.length = 0
  }

  addPlacement(placement: PanePlacementV3): void {
    this.placements.push({
      ...placement,
      clipWidth: Math.max(0, placement.clipWidth),
      clipHeight: Math.max(0, placement.clipHeight),
    })
  }

  addBodyPlacement(input: {
    readonly tileKey: TileKey53
    readonly clipWidth: number
    readonly clipHeight: number
    readonly translateX: number
    readonly translateY: number
    readonly z?: number | undefined
  }): void {
    this.addPlacement({
      tileKey: input.tileKey,
      pane: 'body',
      clipX: 0,
      clipY: 0,
      clipWidth: input.clipWidth,
      clipHeight: input.clipHeight,
      translateX: input.translateX,
      translateY: input.translateY,
      z: input.z ?? 0,
    })
  }

  snapshot(): DrawCommandBufferSnapshotV3 {
    return {
      frameSeq: this.frameSeq,
      placements: this.placements.map((placement) => ({ ...placement })),
    }
  }
}
