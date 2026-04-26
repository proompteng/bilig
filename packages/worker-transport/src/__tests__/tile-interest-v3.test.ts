import { describe, expect, it } from 'vitest'
import { decodeTileInterestBatchV3, encodeTileInterestBatchV3, type TileInterestBatchV3 } from '../index.js'

function createBatch(): TileInterestBatchV3 {
  return {
    magic: 'bilig.tile.interest.v3',
    version: 1,
    seq: 12,
    sheetId: 7,
    sheetOrdinal: 3,
    cameraSeq: 14,
    axisSeqX: 15,
    axisSeqY: 16,
    freezeSeq: 17,
    visibleTileKeys: [2 ** 40 + 1, 2 ** 40 + 2],
    warmTileKeys: [2 ** 40 + 3],
    pinnedTileKeys: [2 ** 40 + 4],
    reason: 'scroll',
  }
}

describe('tile interest v3 codec', () => {
  it('round-trips safe-integer tile keys without u32 truncation', () => {
    const batch = createBatch()
    const bytes = encodeTileInterestBatchV3(batch)
    const decoded = decodeTileInterestBatchV3(bytes)

    expect(bytes[0]).not.toBe('{'.charCodeAt(0))
    expect(decoded).toEqual(batch)
  })
})
