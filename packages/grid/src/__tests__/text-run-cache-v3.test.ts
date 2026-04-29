import { describe, expect, it } from 'vitest'
import { TextRunCacheV3, type TextRunKeyV3 } from '../renderer-v3/text-run-cache.js'

function key(overrides: Partial<TextRunKeyV3> = {}): TextRunKeyV3 {
  return {
    clipWidthBucket: 20,
    colorId: 3,
    dprBucket: 1,
    fontInternId: 2,
    horizontalAlign: 0,
    textInternId: 1,
    verticalAlign: 0,
    wrapMode: 'clip',
    ...overrides,
  }
}

describe('TextRunCacheV3', () => {
  it('reuses text runs by interned text/style/clip key', () => {
    const cache = new TextRunCacheV3()
    const first = cache.getOrCreate({
      create: () => ({ glyphIds: [1, 2], payload: Uint32Array.from([10, 20]) }),
      key: key(),
    })
    const second = cache.getOrCreate({
      create: () => ({ glyphIds: [3], payload: Uint32Array.from([30]) }),
      key: key(),
    })

    expect(second).toBe(first)
    expect(cache.stats()).toMatchObject({ byteSize: 8, runCount: 1 })
    expect(cache.getById(first.runId)).toBe(first)
  })

  it('treats clip width and dpr as cache key inputs', () => {
    const cache = new TextRunCacheV3()

    const narrow = cache.put({ glyphIds: [1], key: key({ clipWidthBucket: 8 }), payload: Uint32Array.from([1]) })
    const wide = cache.put({ glyphIds: [1], key: key({ clipWidthBucket: 64 }), payload: Uint32Array.from([1]) })
    const retina = cache.put({ glyphIds: [1], key: key({ clipWidthBucket: 64, dprBucket: 2 }), payload: Uint32Array.from([1]) })

    expect(new Set([narrow.runId, wide.runId, retina.runId]).size).toBe(3)
    expect(cache.stats().runCount).toBe(3)
  })

  it('updates an existing run without changing its ID', () => {
    const cache = new TextRunCacheV3()
    const first = cache.put({ glyphIds: [1], key: key(), payload: Uint32Array.from([1]) })
    const updated = cache.put({ glyphIds: [1, 2], key: key(), payload: Uint32Array.from([1, 2, 3]) })

    expect(updated.runId).toBe(first.runId)
    expect(updated.byteSize).toBe(12)
    expect(updated.glyphIds).toEqual([1, 2])
    expect(cache.stats()).toMatchObject({ byteSize: 12, runCount: 1 })
  })

  it('evicts least recently used runs by byte budget', () => {
    const cache = new TextRunCacheV3()
    const first = cache.put({ glyphIds: [1], key: key({ textInternId: 1 }), payload: Uint32Array.from([1, 2]) })
    const second = cache.put({ glyphIds: [2], key: key({ textInternId: 2 }), payload: Uint32Array.from([3, 4]) })
    const third = cache.put({ glyphIds: [3], key: key({ textInternId: 3 }), payload: Uint32Array.from([5, 6]) })
    cache.getById(first.runId)

    const evicted: number[] = []
    expect(cache.evictToBudget(16, (record) => evicted.push(record.runId))).toBe(1)

    expect(evicted).toEqual([second.runId])
    expect(cache.getById(first.runId)).toBe(first)
    expect(cache.getById(second.runId)).toBeNull()
    expect(cache.getById(third.runId)).toBe(third)
  })

  it('evicts repeatedly from the LRU head without disturbing recently touched runs', () => {
    const cache = new TextRunCacheV3()
    const first = cache.put({ glyphIds: [1], key: key({ textInternId: 1 }), payload: Uint32Array.from([1, 2]) })
    const second = cache.put({ glyphIds: [2], key: key({ textInternId: 2 }), payload: Uint32Array.from([3, 4]) })
    const third = cache.put({ glyphIds: [3], key: key({ textInternId: 3 }), payload: Uint32Array.from([5, 6]) })
    const fourth = cache.put({ glyphIds: [4], key: key({ textInternId: 4 }), payload: Uint32Array.from([7, 8]) })

    expect(cache.getById(second.runId)).toBe(second)
    expect(cache.evictToBudget(8)).toBe(3)

    expect(cache.getById(first.runId)).toBeNull()
    expect(cache.getById(third.runId)).toBeNull()
    expect(cache.getById(fourth.runId)).toBeNull()
    expect(cache.getById(second.runId)).toBe(second)
  })
})
