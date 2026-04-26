import { DirtyTileIndexV3, type WorkbookDeltaBatchLikeV3 } from '../renderer-v3/tile-damage-index.js'
import { TileResidencyV3, type TileEntryV3, type TileUpsertInputV3 } from '../renderer-v3/tile-residency.js'
import { unpackTileKey53, type TileKey53 } from '../renderer-v3/tile-key.js'

export type GridTileInterestReasonV3 = 'scroll' | 'sheetSwitch' | 'mutation' | 'viewportRestore' | 'prefetch'

export interface GridTileInterestBatchV3 {
  readonly seq: number
  readonly sheetId: number
  readonly sheetOrdinal: number
  readonly cameraSeq: number
  readonly axisSeqX: number
  readonly axisSeqY: number
  readonly freezeSeq: number
  readonly visibleTileKeys: readonly TileKey53[]
  readonly warmTileKeys: readonly TileKey53[]
  readonly pinnedTileKeys: readonly TileKey53[]
  readonly reason: GridTileInterestReasonV3
}

export interface GridTileReadinessSnapshotV3 {
  readonly exactHits: readonly TileKey53[]
  readonly staleHits: readonly TileKey53[]
  readonly misses: readonly TileKey53[]
  readonly visibleDirtyTileKeys: readonly TileKey53[]
  readonly warmDirtyTileKeys: readonly TileKey53[]
}

export class GridTileCoordinator<Packet = unknown, Resources = unknown> {
  readonly residency = new TileResidencyV3<Packet, Resources>()
  readonly dirtyTiles = new DirtyTileIndexV3()
  private nextInterestSeq = 1

  buildInterest(input: {
    readonly sheetId: number
    readonly sheetOrdinal: number
    readonly cameraSeq: number
    readonly axisSeqX: number
    readonly axisSeqY: number
    readonly freezeSeq: number
    readonly visibleTileKeys: Iterable<TileKey53>
    readonly warmTileKeys?: Iterable<TileKey53> | undefined
    readonly pinnedTileKeys?: Iterable<TileKey53> | undefined
    readonly reason: GridTileInterestReasonV3
  }): GridTileInterestBatchV3 {
    return {
      seq: this.nextInterestSeq++,
      sheetId: input.sheetId,
      sheetOrdinal: input.sheetOrdinal,
      cameraSeq: input.cameraSeq,
      axisSeqX: input.axisSeqX,
      axisSeqY: input.axisSeqY,
      freezeSeq: input.freezeSeq,
      visibleTileKeys: [...input.visibleTileKeys],
      warmTileKeys: [...(input.warmTileKeys ?? [])],
      pinnedTileKeys: [...(input.pinnedTileKeys ?? [])],
      reason: input.reason,
    }
  }

  applyWorkbookDelta(batch: WorkbookDeltaBatchLikeV3, options: { readonly dprBucket: number }): void {
    this.dirtyTiles.applyWorkbookDelta(batch, options)
  }

  upsertTile(input: TileUpsertInputV3<Packet, Resources>): TileEntryV3<Packet, Resources> {
    return this.residency.upsert(input)
  }

  reconcileInterest(input: GridTileInterestBatchV3): GridTileReadinessSnapshotV3 {
    this.residency.markVisible(input.visibleTileKeys)
    input.pinnedTileKeys.forEach((key) => {
      this.residency.pin(key, 2)
    })

    const exactHits: number[] = []
    const staleHits: number[] = []
    const misses: number[] = []
    const visibleDirtyTileKeys = this.dirtyTiles.consumeVisible(input.visibleTileKeys)
    const visibleDirty = new Set(visibleDirtyTileKeys)

    input.visibleTileKeys.forEach((key) => {
      const exact = this.residency.getExact(key)
      if (exact && exact.state === 'ready' && !visibleDirty.has(key)) {
        exactHits.push(key)
        return
      }
      if (this.findStaleCompatible(key, input)) {
        staleHits.push(key)
        return
      }
      misses.push(key)
    })

    return {
      exactHits,
      staleHits,
      misses,
      visibleDirtyTileKeys,
      warmDirtyTileKeys: this.dirtyTiles.peekWarm(input.warmTileKeys),
    }
  }

  reset(): void {
    this.nextInterestSeq = 1
    this.residency.clear()
    this.dirtyTiles.clear()
  }

  private findStaleCompatible(key: TileKey53, input: GridTileInterestBatchV3): TileEntryV3<Packet, Resources> | null {
    return this.residency.findStaleCompatible({
      ...unpackTileKey53(key),
      axisSeqX: input.axisSeqX,
      axisSeqY: input.axisSeqY,
      freezeSeq: input.freezeSeq,
      excludeKey: key,
    })
  }
}
