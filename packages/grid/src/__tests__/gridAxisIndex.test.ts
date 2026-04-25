import { describe, expect, test } from 'vitest'
import { createAxisIndex } from '../gridAxisIndex.js'

function linearAnchor(input: {
  readonly axisLength: number
  readonly defaultSize: number
  readonly overrides?: Readonly<Record<number, number>>
  readonly scrollOffset: number
}) {
  let consumed = 0
  for (let index = 0; index < input.axisLength; index += 1) {
    const size = input.overrides?.[index] ?? input.defaultSize
    if (consumed + size > input.scrollOffset) {
      return { index, offset: input.scrollOffset - consumed }
    }
    consumed += size
  }
  return { index: input.axisLength - 1, offset: 0 }
}

describe('gridAxisIndex', () => {
  test('resolves default-size offsets and anchors directly', () => {
    const axis = createAxisIndex({ axisLength: 1_048_576, defaultSize: 22 })

    expect(axis.resolveOffset(900_000)).toBe(19_800_000)
    expect(axis.resolveAnchor(19_800_010)).toEqual({ index: 900_000, offset: 10 })
    expect(axis.resolveVisibleCount(900_000, 220, 22)).toBe(11)
  })

  test('resolves sparse variable sizes', () => {
    const axis = createAxisIndex({
      axisLength: 10,
      defaultSize: 100,
      overrides: [
        { index: 1, size: 150 },
        { index: 4, size: 40 },
        { index: 6, size: 125 },
      ],
    })

    expect(axis.resolveOffset(5)).toBe(490)
    expect(axis.resolveSpan(1, 5)).toBe(390)
    expect(axis.resolveSize(4)).toBe(40)
    expect(axis.resolveAnchor(510)).toEqual({ index: 5, offset: 20 })
  })

  test('skips hidden zero-size entries when resolving anchors', () => {
    const axis = createAxisIndex({
      axisLength: 8,
      defaultSize: 22,
      overrides: [
        { index: 1, size: 0 },
        { index: 2, size: 0 },
      ],
    })

    expect(axis.resolveOffset(3)).toBe(22)
    expect(axis.resolveAnchor(22)).toEqual({ index: 3, offset: 0 })
  })

  test('matches a linear oracle for small sparse axes', () => {
    const overrides = {
      0: 12,
      2: 0,
      5: 33,
      7: 4,
    }
    const axis = createAxisIndex({
      axisLength: 12,
      defaultSize: 10,
      overrides: Object.entries(overrides).map(([index, size]) => ({ index: Number(index), size })),
    })

    for (let scrollOffset = 0; scrollOffset < 140; scrollOffset += 1) {
      expect(axis.resolveAnchor(scrollOffset)).toEqual(
        linearAnchor({
          axisLength: 12,
          defaultSize: 10,
          overrides,
          scrollOffset,
        }),
      )
    }
  })
})
