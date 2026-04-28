import { describe, expect, it } from 'vitest'
import { packTileKey53 } from '../renderer-v3/tile-key.js'
import { TileResidencyV3 } from '../renderer-v3/tile-residency.js'

function entry(rowTile: number, colTile: number, overrides: Partial<Parameters<TileResidencyV3['upsert']>[0]> = {}) {
  return {
    sheetOrdinal: 1,
    rowTile,
    colTile,
    dprBucket: 1,
    axisSeqX: 1,
    axisSeqY: 1,
    freezeSeq: 1,
    valueSeq: 1,
    styleSeq: 1,
    textSeq: 1,
    rectSeq: 1,
    key: packTileKey53({ sheetOrdinal: 1, rowTile, colTile, dprBucket: 1 }),
    byteSizeCpu: 10,
    byteSizeGpu: 20,
    ...overrides,
  }
}

describe('TileResidencyV3', () => {
  it('uses exact numeric lookup and compatibility buckets for stale reuse', () => {
    const residency = new TileResidencyV3()
    const first = residency.upsert(entry(0, 0, { valueSeq: 1 }))
    const second = residency.upsert(entry(0, 1, { valueSeq: 2 }))

    expect(residency.getExact(first.key)).toBe(first)
    expect(residency.findStaleCompatible({ ...entry(0, 2), excludeKey: second.key })).toBe(first)
    expect(residency.getLastStaleScanCount()).toBe(2)

    residency.upsert(entry(1, 0, { axisSeqX: 2 }))
    expect(residency.findStaleCompatible({ ...entry(0, 3), axisSeqX: 1 })).not.toMatchObject({ axisSeqX: 2 })
  })

  it('marks only provided visible tiles for the current generation', () => {
    const residency = new TileResidencyV3()
    const first = residency.upsert(entry(0, 0))
    const second = residency.upsert(entry(0, 1))

    expect(residency.markVisible([first.key])).toBe(1)
    expect(residency.isVisible(first)).toBe(true)
    expect(residency.isVisible(second)).toBe(false)

    residency.markVisible([second.key])
    expect(residency.isVisible(first)).toBe(false)
    expect(residency.isVisible(second)).toBe(true)
  })

  it('evicts by byte budget without removing visible or pinned entries', () => {
    const residency = new TileResidencyV3()
    const first = residency.upsert(entry(0, 0))
    const second = residency.upsert(entry(0, 1))
    const third = residency.upsert(entry(0, 2))
    residency.markVisible([third.key])
    residency.pin(second.key, 3)

    const evicted: number[] = []
    expect(residency.evictToBudgets({ maxCpuBytes: 20, maxGpuBytes: 40, onEvict: (victim) => evicted.push(victim.key) })).toBe(1)
    expect(evicted).toEqual([first.key])
    expect(residency.getExact(first.key)).toBeNull()
    expect(residency.getExact(second.key)).toBe(second)
    expect(residency.getExact(third.key)).toBe(third)
  })

  it('evicts by entry count from the LRU tail', () => {
    const residency = new TileResidencyV3()
    const first = residency.upsert(entry(0, 0))
    const second = residency.upsert(entry(0, 1))
    const third = residency.upsert(entry(0, 2))
    residency.getExact(first.key)

    expect(residency.evictToSize(2)).toBe(1)
    expect(residency.getExact(first.key)).toBe(first)
    expect(residency.getExact(second.key)).toBeNull()
    expect(residency.getExact(third.key)).toBe(third)
  })

  it('updates revision fields in place when an existing tile is refreshed', () => {
    const residency = new TileResidencyV3()
    const original = residency.upsert(entry(0, 0, { axisSeqX: 1, valueSeq: 1 }))
    const refreshed = residency.upsert(entry(0, 0, { axisSeqX: 3, valueSeq: 7, textSeq: 8 }))

    expect(refreshed).toBe(original)
    expect(refreshed).toMatchObject({
      axisSeqX: 3,
      axisSeqY: 1,
      freezeSeq: 1,
      valueSeq: 7,
      styleSeq: 1,
      textSeq: 8,
      rectSeq: 1,
    })
    expect(residency.findStaleCompatible({ ...entry(0, 1), axisSeqX: 1 })).toBeNull()
    expect(residency.findStaleCompatible({ ...entry(0, 1), axisSeqX: 3 })).toBe(refreshed)
  })
})
