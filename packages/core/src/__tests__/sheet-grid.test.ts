import { describe, expect, it } from 'vitest'
import { SheetGrid } from '../sheet-grid.js'

describe('SheetGrid', () => {
  it('stores, retrieves, and clears sparse cells across blocks', () => {
    const grid = new SheetGrid()
    grid.set(0, 0, 4)
    grid.set(130, 33, 9)

    expect(grid.get(0, 0)).toBe(4)
    expect(grid.get(130, 33)).toBe(9)
    expect(grid.get(1, 1)).toBe(-1)

    grid.clear(0, 0)
    expect(grid.get(0, 0)).toBe(-1)
  })

  it('iterates visible cells by range and across all occupied blocks', () => {
    const grid = new SheetGrid()
    grid.set(0, 0, 1)
    grid.set(0, 2, 2)
    grid.set(200, 40, 3)

    const inRange: number[] = []
    grid.forEachInRange(0, 0, 1, 2, (cellIndex) => inRange.push(cellIndex))
    expect(inRange).toEqual([1, 2])

    const all: number[] = []
    grid.forEachCell((cellIndex) => all.push(cellIndex))
    expect(all.toSorted((left, right) => left - right)).toEqual([1, 2, 3])
  })

  it('remaps only cells inside the requested axis scope', () => {
    const grid = new SheetGrid()
    grid.set(0, 0, 1)
    grid.set(129, 0, 2)
    grid.set(260, 0, 3)

    const changed = grid.remapAxis('row', (index) => index + 1, { start: 128, end: 256 })

    expect(changed).toEqual([
      {
        cellIndex: 2,
        row: 129,
        col: 0,
        nextRow: 130,
        nextCol: 0,
      },
    ])
    expect(grid.get(0, 0)).toBe(1)
    expect(grid.get(129, 0)).toBe(-1)
    expect(grid.get(130, 0)).toBe(2)
    expect(grid.get(260, 0)).toBe(3)
  })

  it('returns stable remap entries for row deletes without touching unrelated cells', () => {
    const grid = new SheetGrid()
    grid.set(0, 0, 1)
    grid.set(10, 2, 2)
    grid.set(11, 2, 3)

    const changed = grid.remapAxis(
      'row',
      (row) => {
        if (row === 10) {
          return undefined
        }
        if (row > 10) {
          return row - 1
        }
        return row
      },
      { start: 10 },
    )

    expect(changed).toEqual([
      {
        cellIndex: 2,
        row: 10,
        col: 2,
        nextRow: undefined,
        nextCol: 2,
      },
      {
        cellIndex: 3,
        row: 11,
        col: 2,
        nextRow: 10,
        nextCol: 2,
      },
    ])
    expect(grid.get(0, 0)).toBe(1)
    expect(grid.get(10, 2)).toBe(3)
    expect(grid.get(11, 2)).toBe(-1)
  })

  it('remaps only cells inside the requested column scope', () => {
    const grid = new SheetGrid()
    grid.set(0, 0, 1)
    grid.set(40, 33, 2)
    grid.set(90, 65, 3)

    const changed = grid.remapAxis('column', (col) => col + 1, { start: 32, end: 64 })

    expect(changed).toEqual([
      {
        cellIndex: 2,
        row: 40,
        col: 33,
        nextRow: 40,
        nextCol: 34,
      },
    ])
    expect(grid.get(0, 0)).toBe(1)
    expect(grid.get(40, 33)).toBe(-1)
    expect(grid.get(40, 34)).toBe(2)
    expect(grid.get(90, 65)).toBe(3)
  })
})
