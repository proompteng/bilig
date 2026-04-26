import type { WorkbookDeltaBatchV3 } from '@bilig/worker-transport'
import { DirtyTileIndexV3, markWorkbookDeltaDirtyTilesV3 } from '../../../packages/grid/src/renderer-v3/tile-damage-index.js'
import type { TileKey53 } from '../../../packages/grid/src/renderer-v3/tile-key.js'

export interface ProjectedDamageBusApplyResult {
  readonly applied: boolean
  readonly seq: number
}

export class ProjectedDamageBus {
  private readonly dirtyTiles = new DirtyTileIndexV3()
  private readonly lastSeqBySheetOrdinal = new Map<number, number>()

  applyWorkbookDelta(batch: WorkbookDeltaBatchV3, options: { readonly dprBucket: number }): ProjectedDamageBusApplyResult {
    const lastSeq = this.lastSeqBySheetOrdinal.get(batch.sheetOrdinal) ?? -1
    if (batch.seq <= lastSeq) {
      return {
        applied: false,
        seq: batch.seq,
      }
    }
    this.lastSeqBySheetOrdinal.set(batch.sheetOrdinal, batch.seq)
    markWorkbookDeltaDirtyTilesV3(this.dirtyTiles, batch, options)
    return {
      applied: true,
      seq: batch.seq,
    }
  }

  getMask(key: TileKey53): number {
    return this.dirtyTiles.getMask(key)
  }

  consumeVisible(keys: Iterable<TileKey53>): TileKey53[] {
    return this.dirtyTiles.consumeVisible(keys)
  }

  peekWarm(keys: Iterable<TileKey53>): TileKey53[] {
    return this.dirtyTiles.peekWarm(keys)
  }

  reset(): void {
    this.dirtyTiles.clear()
    this.lastSeqBySheetOrdinal.clear()
  }
}
