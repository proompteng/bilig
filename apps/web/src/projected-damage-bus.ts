import { decodeWorkbookDeltaBatchV3, type WorkbookDeltaBatchV3, type WorkerEngineClient } from '@bilig/worker-transport'
import { DirtyTileIndexV3, markWorkbookDeltaDirtyTilesV3 } from '../../../packages/grid/src/renderer-v3/tile-damage-index.js'
import type { TileKey53 } from '../../../packages/grid/src/renderer-v3/tile-key.js'

export interface ProjectedDamageBusApplyResult {
  readonly applied: boolean
  readonly seq: number
}

export class ProjectedDamageBus {
  private readonly dirtyTiles = new DirtyTileIndexV3()
  private readonly lastSeqBySheetOrdinal = new Map<number, number>()

  constructor(private readonly client?: Pick<WorkerEngineClient, 'subscribeWorkbookDeltas'>) {}

  subscribeWorkbookDeltas(
    options: { readonly dprBucket: number },
    listener: (result: ProjectedDamageBusApplyResult, batch: WorkbookDeltaBatchV3) => void = () => undefined,
  ): () => void {
    if (!this.client) {
      throw new Error('Projected damage bus subscriptions require a worker engine client')
    }
    return this.client.subscribeWorkbookDeltas((bytes) => {
      const batch = decodeWorkbookDeltaBatchV3(bytes)
      const result = this.applyWorkbookDelta(batch, options)
      listener(result, batch)
    })
  }

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
