import {
  decodeRenderTileDeltaBatch,
  type RenderTileCellRun,
  type RenderTileCoord,
  type RenderTileDeltaBatch,
  type RenderTileDeltaSubscription,
  type RenderTileDirtySpans,
  type RenderTileTextRun,
  type RenderTileVersion,
  type WorkerEngineClient,
} from '@bilig/worker-transport'
import type { Viewport } from '@bilig/protocol'
import { getWorkbookScrollPerfCollector } from './perf/workbook-scroll-perf.js'

export interface ProjectedRenderTile {
  readonly tileId: number
  readonly coord: RenderTileCoord
  readonly version: RenderTileVersion
  readonly bounds: Viewport
  readonly rectInstances: Float32Array
  readonly rectCount: number
  readonly textMetrics: Float32Array
  readonly glyphRefs: Uint32Array
  readonly textRuns: readonly RenderTileTextRun[]
  readonly textCount: number
  readonly dirty: RenderTileDirtySpans
  readonly lastCellRuns: readonly RenderTileCellRun[]
  readonly lastBatchId: number
  readonly lastCameraSeq: number
}

export interface ProjectedTileSceneChange {
  readonly batchId: number
  readonly cameraSeq: number
  readonly changedTileIds: readonly number[]
  readonly invalidatedTileIds: readonly number[]
  readonly structural: boolean
}

export class ProjectedTileSceneStore {
  private readonly tiles = new Map<number, ProjectedRenderTile>()
  private readonly sheetIdsByName = new Map<string, number>()
  private lastBatchId = 0
  private lastCameraSeq = 0
  private overlayRevision = 0

  constructor(private readonly client?: Pick<WorkerEngineClient, 'subscribeRenderTileDeltas'>) {}

  subscribe(subscription: RenderTileDeltaSubscription, listener: (change: ProjectedTileSceneChange) => void): () => void {
    if (!this.client) {
      throw new Error('Worker render tile subscriptions require a worker engine client')
    }
    this.sheetIdsByName.set(subscription.sheetName, subscription.sheetId)
    return this.client.subscribeRenderTileDeltas(subscription, (bytes: Uint8Array) => {
      listener(this.applyDelta(decodeRenderTileDeltaBatch(bytes)))
    })
  }

  applyDelta(batch: RenderTileDeltaBatch): ProjectedTileSceneChange {
    const startedAt = nowMs()
    if (batch.batchId < this.lastBatchId) {
      const staleChange = {
        batchId: batch.batchId,
        cameraSeq: batch.cameraSeq,
        changedTileIds: [],
        invalidatedTileIds: [],
        structural: false,
      }
      this.noteRendererDeltaApply(batch, staleChange, startedAt)
      return staleChange
    }

    const changedTileIds = new Set<number>()
    const invalidatedTileIds = new Set<number>()
    const structural = batch.mutations.some((mutation) => mutation.kind === 'axis' || mutation.kind === 'freeze')
    if (structural) {
      for (const [tileId, tile] of this.tiles) {
        if (tile.coord.sheetId === batch.sheetId) {
          this.tiles.delete(tileId)
          invalidatedTileIds.add(tileId)
        }
      }
    }

    for (const mutation of batch.mutations) {
      switch (mutation.kind) {
        case 'tileReplace':
          this.tiles.set(mutation.tileId, {
            tileId: mutation.tileId,
            coord: mutation.coord,
            version: mutation.version,
            bounds: mutation.bounds,
            rectInstances: mutation.rectInstances,
            rectCount: mutation.rectCount,
            textMetrics: mutation.textMetrics,
            glyphRefs: mutation.glyphRefs,
            textRuns: mutation.textRuns,
            textCount: mutation.textCount,
            dirty: mutation.dirty,
            lastCellRuns: [],
            lastBatchId: batch.batchId,
            lastCameraSeq: batch.cameraSeq,
          })
          changedTileIds.add(mutation.tileId)
          invalidatedTileIds.delete(mutation.tileId)
          break
        case 'cellRuns': {
          const current = this.tiles.get(mutation.tileId)
          if (!current) {
            break
          }
          this.tiles.set(mutation.tileId, {
            ...current,
            version: mutation.version,
            lastCellRuns: mutation.runs,
            lastBatchId: batch.batchId,
            lastCameraSeq: batch.cameraSeq,
          })
          changedTileIds.add(mutation.tileId)
          break
        }
        case 'invalidate':
          this.tiles.delete(mutation.tileId)
          changedTileIds.delete(mutation.tileId)
          invalidatedTileIds.add(mutation.tileId)
          break
        case 'overlay':
          this.overlayRevision = Math.max(this.overlayRevision, mutation.overlayRevision)
          break
        case 'axis':
        case 'freeze':
          break
      }
    }

    this.lastBatchId = Math.max(this.lastBatchId, batch.batchId)
    this.lastCameraSeq = Math.max(this.lastCameraSeq, batch.cameraSeq)
    const change = {
      batchId: batch.batchId,
      cameraSeq: batch.cameraSeq,
      changedTileIds: [...changedTileIds],
      invalidatedTileIds: [...invalidatedTileIds],
      structural,
    }
    this.noteRendererDeltaApply(batch, change, startedAt)
    return change
  }

  peekTile(tileId: number): ProjectedRenderTile | null {
    return this.tiles.get(tileId) ?? null
  }

  getLastBatchId(): number {
    return this.lastBatchId
  }

  getLastCameraSeq(): number {
    return this.lastCameraSeq
  }

  getOverlayRevision(): number {
    return this.overlayRevision
  }

  dropSheets(sheetNames: readonly string[]): void {
    const sheetIds = new Set<number>()
    sheetNames.forEach((sheetName) => {
      const sheetId = this.sheetIdsByName.get(sheetName)
      if (sheetId !== undefined) {
        sheetIds.add(sheetId)
      }
      this.sheetIdsByName.delete(sheetName)
    })
    if (sheetIds.size === 0) {
      return
    }
    for (const [tileId, tile] of this.tiles) {
      if (sheetIds.has(tile.coord.sheetId)) {
        this.tiles.delete(tileId)
      }
    }
  }

  reset(): void {
    this.tiles.clear()
    this.sheetIdsByName.clear()
    this.lastBatchId = 0
    this.lastCameraSeq = 0
    this.overlayRevision = 0
  }

  private noteRendererDeltaApply(
    batch: RenderTileDeltaBatch,
    change: Pick<ProjectedTileSceneChange, 'changedTileIds' | 'invalidatedTileIds'>,
    startedAt: number,
  ): void {
    getWorkbookScrollPerfCollector()?.noteRendererDeltaApply({
      dirtyTileCount: new Set([...change.changedTileIds, ...change.invalidatedTileIds]).size,
      durationMs: Math.max(0, nowMs() - startedAt),
      mutationCount: batch.mutations.length,
    })
  }
}

function nowMs(): number {
  return typeof performance === 'undefined' ? Date.now() : performance.now()
}
