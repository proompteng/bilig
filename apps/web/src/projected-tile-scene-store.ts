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
  readonly rectSignature?: string | undefined
  readonly textMetrics: Float32Array
  readonly glyphRefs: Uint32Array
  readonly textRuns: readonly RenderTileTextRun[]
  readonly textCount: number
  readonly textSignature?: string | undefined
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
  private readonly sheetIdentityByName = new Map<string, { readonly sheetId: number; readonly sheetOrdinal: number }>()
  private readonly lastSequenceBySheetIdentity = new Map<string, { readonly batchId: number; readonly cameraSeq: number }>()
  private lastBatchId = 0
  private lastCameraSeq = 0

  constructor(private readonly client?: Pick<WorkerEngineClient, 'subscribeRenderTileDeltas'>) {}

  subscribe(subscription: RenderTileDeltaSubscription, listener: (change: ProjectedTileSceneChange) => void): () => void {
    if (!this.client) {
      throw new Error('Worker render tile subscriptions require a worker engine client')
    }
    this.sheetIdentityByName.set(subscription.sheetName, {
      sheetId: subscription.sheetId,
      sheetOrdinal: subscription.sheetOrdinal ?? subscription.tileInterest?.sheetOrdinal ?? subscription.sheetId,
    })
    return this.client.subscribeRenderTileDeltas(subscription, (bytes: Uint8Array) => {
      listener(this.applyDelta(decodeRenderTileDeltaBatch(bytes)))
    })
  }

  applyDelta(batch: RenderTileDeltaBatch): ProjectedTileSceneChange {
    const startedAt = nowMs()
    if (this.isStaleBatchForSheet(batch)) {
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
        if (tile && matchesProjectedSheetIdentity(tile.coord, batch)) {
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
            rectSignature: mutation.rectSignature,
            textMetrics: mutation.textMetrics,
            glyphRefs: mutation.glyphRefs,
            textRuns: mutation.textRuns,
            textCount: mutation.textCount,
            textSignature: mutation.textSignature,
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
          this.residency.delete(mutation.tileId)
          this.dirtyTiles.markTile(mutation.tileId, DirtyMaskV3.Value | DirtyMaskV3.Text | DirtyMaskV3.Rect)
          changedTileIds.delete(mutation.tileId)
          invalidatedTileIds.add(mutation.tileId)
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

    this.noteAcceptedBatch(batch)
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
    const sheetOrdinals = new Set<number>()
    const sheetIdentities = new Set<string>()
    sheetNames.forEach((sheetName) => {
      const identity = this.sheetIdentityByName.get(sheetName)
      if (identity) {
        sheetIds.add(identity.sheetId)
        sheetOrdinals.add(identity.sheetOrdinal)
        const sheetIdentity = projectedSheetIdentityKey(identity)
        sheetIdentities.add(sheetIdentity)
        this.lastSequenceBySheetIdentity.delete(sheetIdentity)
      }
      this.sheetIdentityByName.delete(sheetName)
    })
    if (sheetIds.size === 0 && sheetOrdinals.size === 0) {
      return
    }
    for (const entry of this.residency.entries()) {
      const tile = entry.packet
      if (
        tile &&
        (sheetIdentities.has(projectedSheetIdentityKey(tile.coord)) ||
          (sheetIds.has(tile.coord.sheetId) && sheetOrdinals.size === 0) ||
          (sheetIds.size === 0 && sheetOrdinals.has(tile.coord.sheetOrdinal)))
      ) {
        this.residency.delete(entry.key)
      }
    }
  }

  reset(): void {
    this.residency.clear()
    this.dirtyTiles.clear()
    this.sheetIdentityByName.clear()
    this.lastSequenceBySheetIdentity.clear()
    this.lastBatchId = 0
    this.lastCameraSeq = 0
  }

  private isStaleBatchForSheet(batch: RenderTileDeltaBatch): boolean {
    const key = projectedSheetIdentityKey(batch)
    const lastSequence = this.lastSequenceBySheetIdentity.get(key)
    return (
      lastSequence !== undefined &&
      (batch.batchId < lastSequence.batchId || (batch.batchId === lastSequence.batchId && batch.cameraSeq < lastSequence.cameraSeq))
    )
  }

  private noteAcceptedBatch(batch: RenderTileDeltaBatch): void {
    const key = projectedSheetIdentityKey(batch)
    const lastSequence = this.lastSequenceBySheetIdentity.get(key)
    if (
      lastSequence === undefined ||
      batch.batchId > lastSequence.batchId ||
      (batch.batchId === lastSequence.batchId && batch.cameraSeq > lastSequence.cameraSeq)
    ) {
      this.lastSequenceBySheetIdentity.set(key, {
        batchId: batch.batchId,
        cameraSeq: batch.cameraSeq,
      })
    }
    this.lastBatchId = Math.max(this.lastBatchId, batch.batchId)
    this.lastCameraSeq = Math.max(this.lastCameraSeq, batch.cameraSeq)
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
      sheetOrdinal: tile.coord.sheetOrdinal,
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

function matchesProjectedSheetIdentity(
  tile: { readonly sheetId?: number | undefined; readonly sheetOrdinal?: number | undefined },
  expected: { readonly sheetId?: number | undefined; readonly sheetOrdinal?: number | undefined },
): boolean {
  if (expected.sheetId !== undefined && expected.sheetOrdinal !== undefined) {
    return tile.sheetId === expected.sheetId && tile.sheetOrdinal === expected.sheetOrdinal
  }
  if (expected.sheetId !== undefined) {
    return tile.sheetId === expected.sheetId
  }
  if (expected.sheetOrdinal !== undefined) {
    return tile.sheetOrdinal === expected.sheetOrdinal
  }
  return false
}

function projectedSheetIdentityKey(identity: { readonly sheetId: number; readonly sheetOrdinal: number }): string {
  return `${identity.sheetId}:${identity.sheetOrdinal}`
}
