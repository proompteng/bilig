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
import { DirtyMaskV3, DirtyTileIndexV3 } from '../../../packages/grid/src/renderer-v3/tile-damage-index.js'
import { TileResidencyV3 } from '../../../packages/grid/src/renderer-v3/tile-residency.js'
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
  readonly dirtyLocalRows?: Uint32Array | undefined
  readonly dirtyLocalCols?: Uint32Array | undefined
  readonly dirtyMasks?: Uint32Array | undefined
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
  private readonly residency = new TileResidencyV3<ProjectedRenderTile>()
  private readonly dirtyTiles = new DirtyTileIndexV3()
  private readonly sheetIdsByName = new Map<string, number>()
  private lastBatchId = 0
  private lastCameraSeq = 0

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
      for (const entry of this.residency.entries()) {
        const tile = entry.packet
        if (tile?.coord.sheetId === batch.sheetId) {
          this.residency.delete(entry.key)
          invalidatedTileIds.add(entry.key)
        }
      }
    }

    for (const mutation of batch.mutations) {
      switch (mutation.kind) {
        case 'tileReplace': {
          const tile = {
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
            dirtyLocalCols: mutation.dirtyLocalCols,
            dirtyLocalRows: mutation.dirtyLocalRows,
            dirtyMasks: mutation.dirtyMasks,
            lastCellRuns: [],
            lastBatchId: batch.batchId,
            lastCameraSeq: batch.cameraSeq,
          }
          this.upsertTile(tile)
          changedTileIds.add(mutation.tileId)
          invalidatedTileIds.delete(mutation.tileId)
          break
        }
        case 'cellRuns': {
          const current = this.residency.getExact(mutation.tileId)?.packet
          if (!current) {
            break
          }
          this.upsertTile({
            ...current,
            version: mutation.version,
            lastCellRuns: mutation.runs,
            lastBatchId: batch.batchId,
            lastCameraSeq: batch.cameraSeq,
          })
          this.dirtyTiles.markTile(mutation.tileId, DirtyMaskV3.Value | DirtyMaskV3.Text | DirtyMaskV3.Rect)
          changedTileIds.add(mutation.tileId)
          break
        }
        case 'invalidate':
          this.residency.delete(mutation.tileId)
          changedTileIds.delete(mutation.tileId)
          invalidatedTileIds.add(mutation.tileId)
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
    return this.residency.getExact(tileId)?.packet ?? null
  }

  getLastBatchId(): number {
    return this.lastBatchId
  }

  getLastCameraSeq(): number {
    return this.lastCameraSeq
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
    for (const entry of this.residency.entries()) {
      const tile = entry.packet
      if (tile && sheetIds.has(tile.coord.sheetId)) {
        this.residency.delete(entry.key)
      }
    }
  }

  reset(): void {
    this.residency.clear()
    this.dirtyTiles.clear()
    this.sheetIdsByName.clear()
    this.lastBatchId = 0
    this.lastCameraSeq = 0
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

  private upsertTile(tile: ProjectedRenderTile): void {
    this.residency.upsert({
      axisSeqX: tile.version.axisX,
      axisSeqY: tile.version.axisY,
      byteSizeCpu: estimateTileBytes(tile),
      byteSizeGpu: tile.rectInstances.byteLength + tile.textMetrics.byteLength + tile.glyphRefs.byteLength,
      colTile: tile.coord.colTile,
      dprBucket: tile.coord.dprBucket,
      freezeSeq: tile.version.freeze,
      key: tile.tileId,
      packet: tile,
      rectSeq: tile.version.styles,
      rowTile: tile.coord.rowTile,
      sheetOrdinal: tile.coord.sheetId,
      state: 'ready',
      styleSeq: tile.version.styles,
      textSeq: tile.version.text,
      valueSeq: tile.version.values,
    })
  }
}

function nowMs(): number {
  return typeof performance === 'undefined' ? Date.now() : performance.now()
}

function estimateTileBytes(tile: ProjectedRenderTile): number {
  let stringBytes = 0
  for (const run of tile.textRuns) {
    stringBytes += run.text.length * 2
    stringBytes += run.font.length * 2
    stringBytes += run.color.length * 2
  }
  return tile.rectInstances.byteLength + tile.textMetrics.byteLength + tile.glyphRefs.byteLength + stringBytes
}
