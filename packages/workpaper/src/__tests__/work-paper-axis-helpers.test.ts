import { describe, expect, it } from 'vitest'
import {
  applyWorkPaperAxisOrder,
  applyWorkPaperAxisSwapMappings,
  normalizeAxisIntervals,
  normalizeAxisSwapMappings,
} from '../work-paper-axis-helpers.js'
import { WorkPaperInvalidArgumentsError } from '../work-paper-errors.js'

describe('work paper axis helpers', () => {
  it('normalizes numeric and tuple intervals', () => {
    expect(normalizeAxisIntervals(4)).toEqual([[4, 1]])
    expect(normalizeAxisIntervals(4, 3)).toEqual([[4, 3]])
    expect(normalizeAxisIntervals([1], [3, 2], [[7]])).toEqual([
      [1, 1],
      [3, 2],
      [7, 1],
    ])
  })

  it('rejects ambiguous interval overloads', () => {
    expect(() => normalizeAxisIntervals(1, [2, 3])).toThrow(WorkPaperInvalidArgumentsError)
    expect(() => normalizeAxisIntervals([1, 2], 3)).toThrow(WorkPaperInvalidArgumentsError)
  })

  it('normalizes row and column swap mappings', () => {
    const mappings: readonly [number, number][] = [
      [1, 2],
      [3, 4],
    ]

    expect(normalizeAxisSwapMappings('row', 1, 2)).toEqual([[1, 2]])
    expect(normalizeAxisSwapMappings('column', mappings)).toEqual(mappings)
    expect(normalizeAxisSwapMappings('column', mappings)).not.toBe(mappings)
  })

  it('requires both indexes for numeric swap overloads', () => {
    expect(() => normalizeAxisSwapMappings('row', 1)).toThrow(WorkPaperInvalidArgumentsError)
    expect(() => normalizeAxisSwapMappings('column', 1)).toThrow(WorkPaperInvalidArgumentsError)
  })

  it('applies axis swap mappings through the same move choreography used by rows and columns', () => {
    const moves: Array<[start: number, count: number, target: number]> = []

    applyWorkPaperAxisSwapMappings(
      [
        [1, 3],
        [4, 2],
        [5, 5],
      ],
      (start, count, target) => {
        moves.push([start, count, target])
      },
    )

    expect(moves).toEqual([
      [1, 1, 3],
      [2, 1, 1],
      [4, 1, 2],
      [3, 1, 4],
    ])
  })

  it('applies target axis order with stable one-item moves', () => {
    const moves: Array<[start: number, count: number, target: number]> = []

    applyWorkPaperAxisOrder([2, 3, 0, 1], (start, count, target) => {
      moves.push([start, count, target])
    })

    expect(moves).toEqual([
      [2, 1, 0],
      [3, 1, 1],
    ])
  })
})
