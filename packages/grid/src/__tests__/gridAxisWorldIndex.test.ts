import { describe, expect, test } from 'vitest'
import { createGridAxisWorldIndex, createGridAxisWorldIndexFromRecords } from '../gridAxisWorldIndex.js'

describe('gridAxisWorldIndex', () => {
  test('resolves default offsets, anchors, and total size directly', () => {
    const axis = createGridAxisWorldIndex({ axisLength: 1_048_576, defaultSize: 22, version: 7 })

    expect(axis.version).toBe(7)
    expect(axis.totalSize).toBe(23_068_672)
    expect(axis.offsetOf(900_000)).toBe(19_800_000)
    expect(axis.endOffsetOf(900_000)).toBe(19_800_022)
    expect(axis.anchorAt(19_800_010)).toEqual({ index: 900_000, intraOffset: 10, offset: 19_800_000, size: 22 })
    expect(axis.hitTest(19_800_021)).toBe(900_000)
  })

  test('supports variable sizes and hidden entries as rendered zero-size axes', () => {
    const axis = createGridAxisWorldIndex({
      axisLength: 8,
      defaultSize: 10,
      overrides: [
        { index: 1, size: 20 },
        { hidden: true, index: 2 },
        { index: 4, size: 5 },
      ],
    })

    expect(axis.totalSize).toBe(75)
    expect(axis.offsetOf(3)).toBe(30)
    expect(axis.sizeOf(2)).toBe(0)
    expect(axis.isHidden(2)).toBe(true)
    expect(axis.anchorAt(30)).toEqual({ index: 3, intraOffset: 0, offset: 30, size: 10 })
    expect(axis.hitTest(29)).toBe(1)
    expect(axis.hitTest(30)).toBe(3)
  })

  test('merges size and hidden records with hidden state winning', () => {
    const axis = createGridAxisWorldIndexFromRecords({
      axisLength: 6,
      defaultSize: 10,
      hidden: { 1: true },
      sizes: { 1: 30, 3: 25 },
    })

    expect(axis.sizeOf(1)).toBe(0)
    expect(axis.offsetOf(2)).toBe(10)
    expect(axis.offsetOf(4)).toBe(45)
    expect(axis.hitTest(10)).toBe(2)
  })

  test('returns visible ranges and counts without including hidden entries', () => {
    const axis = createGridAxisWorldIndex({
      axisLength: 8,
      defaultSize: 10,
      overrides: [
        { hidden: true, index: 2 },
        { hidden: true, index: 3 },
      ],
    })

    expect(axis.visibleRangeForWorldRect(10, 35)).toEqual({ count: 4, endIndexExclusive: 7, startIndex: 1 })
    expect(axis.visibleCountFrom(1, 35)).toBe(4)
  })
})
