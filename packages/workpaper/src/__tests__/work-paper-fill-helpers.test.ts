import { describe, expect, it } from 'vitest'
import { buildWorkPaperFillRangeData, buildWorkPaperNullMatrixForRange } from '../work-paper-fill-helpers.js'
import type { WorkPaperCellRange } from '../work-paper-types.js'

const range = (startRow: number, startCol: number, endRow: number, endCol: number): WorkPaperCellRange => ({
  start: { sheet: 0, row: startRow, col: startCol },
  end: { sheet: 0, row: endRow, col: endCol },
})

describe('work paper fill helpers', () => {
  it('builds null matrices sized to the requested range', () => {
    expect(buildWorkPaperNullMatrixForRange(range(2, 3, 3, 5))).toEqual([
      [null, null, null],
      [null, null, null],
    ])
  })

  it('tiles serialized source values across the target range', () => {
    expect(
      buildWorkPaperFillRangeData({
        source: range(0, 0, 1, 1),
        target: range(0, 0, 2, 2),
        sourceSerialized: [
          [1, 2],
          [3, 4],
        ],
        offsetsFromTarget: false,
      }),
    ).toEqual([
      [1, 2, 1],
      [3, 4, 3],
      [1, 2, 1],
    ])
  })

  it('relocates tiled formulas from their source cell to each target cell', () => {
    expect(
      buildWorkPaperFillRangeData({
        source: range(0, 0, 0, 0),
        target: range(1, 1, 2, 1),
        sourceSerialized: [['=A1']],
        offsetsFromTarget: false,
      }),
    ).toEqual([['=B2'], ['=B3']])
  })

  it('relocates formula prefixes with leading whitespace across rows and columns', () => {
    expect(
      buildWorkPaperFillRangeData({
        source: range(0, 0, 0, 0),
        target: range(1, 1, 2, 2),
        sourceSerialized: [[' \t=A1']],
        offsetsFromTarget: false,
      }),
    ).toEqual([
      [' \t=B2', ' \t=C2'],
      [' \t=B3', ' \t=C3'],
    ])
  })

  it('can offset formula tiling from the target origin', () => {
    expect(
      buildWorkPaperFillRangeData({
        source: range(10, 10, 10, 10),
        target: range(20, 20, 21, 20),
        sourceSerialized: [['=K11']],
        offsetsFromTarget: true,
      }),
    ).toEqual([['=U21'], ['=U22']])
  })
})
