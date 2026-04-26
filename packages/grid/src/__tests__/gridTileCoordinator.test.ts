import { describe, expect, it } from 'vitest'
import { DirtyMaskV3 } from '../renderer-v3/tile-damage-index.js'
import { packTileKey53 } from '../renderer-v3/tile-key.js'
import { GridTileCoordinator } from '../runtime/gridTileCoordinator.js'

function key(rowTile: number, colTile: number): number {
  return packTileKey53({ sheetOrdinal: 1, rowTile, colTile, dprBucket: 1 })
}

function tile(rowTile: number, colTile: number, overrides: Partial<Parameters<GridTileCoordinator['upsertTile']>[0]> = {}) {
  return {
    sheetOrdinal: 1,
    rowTile,
    colTile,
    dprBucket: 1,
    axisSeqX: 2,
    axisSeqY: 3,
    freezeSeq: 4,
    valueSeq: 5,
    styleSeq: 6,
    textSeq: 7,
    rectSeq: 8,
    key: key(rowTile, colTile),
    ...overrides,
  }
}

describe('GridTileCoordinator', () => {
  it('builds one-frame tile interest batches', () => {
    const coordinator = new GridTileCoordinator()
    const visible = [key(0, 0)]
    const warm = [key(0, 1)]
    const pinned = [key(0, 2)]

    expect(
      coordinator.buildInterest({
        sheetId: 10,
        sheetOrdinal: 1,
        cameraSeq: 2,
        axisSeqX: 3,
        axisSeqY: 4,
        freezeSeq: 5,
        visibleTileKeys: visible,
        warmTileKeys: warm,
        pinnedTileKeys: pinned,
        reason: 'scroll',
      }),
    ).toEqual({
      seq: 1,
      sheetId: 10,
      sheetOrdinal: 1,
      cameraSeq: 2,
      axisSeqX: 3,
      axisSeqY: 4,
      freezeSeq: 5,
      visibleTileKeys: visible,
      warmTileKeys: warm,
      pinnedTileKeys: pinned,
      reason: 'scroll',
    })
  })

  it('classifies exact, stale-compatible, and missing visible tiles', () => {
    const coordinator = new GridTileCoordinator()
    const exact = coordinator.upsertTile(tile(0, 0))
    coordinator.upsertTile(tile(0, 3, { key: key(0, 3), valueSeq: 1 }))
    const dirty = coordinator.upsertTile(tile(0, 1))
    const missing = key(0, 2)
    coordinator.applyWorkbookDelta(
      {
        sheetOrdinal: 1,
        dirty: {
          axisX: new Uint32Array(),
          axisY: new Uint32Array(),
          cellRanges: Uint32Array.from([0, 0, 128, 128, DirtyMaskV3.Value]),
        },
      },
      { dprBucket: 1 },
    )

    const interest = coordinator.buildInterest({
      sheetId: 10,
      sheetOrdinal: 1,
      cameraSeq: 2,
      axisSeqX: 2,
      axisSeqY: 3,
      freezeSeq: 4,
      visibleTileKeys: [exact.key, dirty.key, missing],
      warmTileKeys: [dirty.key],
      reason: 'mutation',
    })

    expect(coordinator.reconcileInterest(interest)).toEqual({
      exactHits: [exact.key],
      staleHits: [dirty.key, missing],
      misses: [],
      visibleDirtyTileKeys: [dirty.key],
      warmDirtyTileKeys: [],
    })
  })
})
