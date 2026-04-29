import { describe, expect, it, vi } from 'vitest'
import { ValueTag, type CellSnapshot } from '@bilig/protocol'
import { ProjectedViewportCellCache } from '../projected-viewport-cell-cache.js'

function countSheetCells(cache: ProjectedViewportCellCache, sheetName: string): number {
  let count = 0
  cache.getSheet(sheetName)?.grid.forEachCellEntry(() => {
    count += 1
  })
  return count
}

function snapshot(address: string, value: number | string): CellSnapshot {
  return {
    sheetName: 'Sheet1',
    address,
    value: typeof value === 'number' ? { tag: ValueTag.Number, value } : { tag: ValueTag.String, value, stringId: 1 },
    flags: 0,
    version: 1,
  }
}

function resetEmptySnapshot(address: string): CellSnapshot {
  return {
    sheetName: 'Sheet1',
    address,
    value: { tag: ValueTag.Empty },
    flags: 0,
    version: 0,
  }
}

function columnLabel(columnIndex: number): string {
  let index = columnIndex + 1
  let label = ''
  while (index > 0) {
    const remainder = (index - 1) % 26
    label = String.fromCharCode(65 + remainder) + label
    index = Math.floor((index - 1) / 26)
  }
  return label
}

describe('ProjectedViewportCellCache', () => {
  it('does not publish absent version-zero empty selection snapshots', () => {
    const cache = new ProjectedViewportCellCache()
    const listener = vi.fn()
    const globalListener = vi.fn()
    cache.subscribeCells('Sheet1', ['C21'], listener)
    cache.subscribe(globalListener)

    expect(cache.setCellSnapshot(resetEmptySnapshot('C21'))).toBe(false)

    expect(cache.peekCell('Sheet1', 'C21')).toBeUndefined()
    expect(cache.getCell('Sheet1', 'C21')).toEqual(resetEmptySnapshot('C21'))
    expect(listener).not.toHaveBeenCalled()
    expect(globalListener).not.toHaveBeenCalled()
  })

  it('keeps existing cached values when a reset-empty selection snapshot is older', () => {
    const cache = new ProjectedViewportCellCache()
    const listener = vi.fn()
    cache.subscribeCells('Sheet1', ['C21'], listener)
    cache.setCellSnapshot(snapshot('C21', 'stale'))
    listener.mockClear()

    expect(cache.setCellSnapshot(resetEmptySnapshot('C21'))).toBe(false)

    expect(cache.peekCell('Sheet1', 'C21')?.input).toBeUndefined()
    expect(cache.peekCell('Sheet1', 'C21')?.value).toEqual({ tag: ValueTag.String, value: 'stale', stringId: 1 })
    expect(listener).not.toHaveBeenCalled()
  })

  it('tracks cell subscriptions and exposes sheet grid entries', () => {
    const cache = new ProjectedViewportCellCache()
    const listener = vi.fn()
    cache.subscribeCells('Sheet1', ['B2'], listener)

    cache.setCellSnapshot(snapshot('B2', 'left'))

    expect(listener).toHaveBeenCalledTimes(1)
    expect(cache.getCell('Sheet1', 'B2')).toMatchObject({
      value: { tag: ValueTag.String, value: 'left' },
    })
    expect(countSheetCells(cache, 'Sheet1')).toBe(1)
  })

  it('prunes to the cache cap after the last viewport unsubscribes while keeping pinned cells', () => {
    const cache = new ProjectedViewportCellCache({ maxCachedCellsPerSheet: 6000 })
    const untrackViewport = cache.trackViewport('Sheet1', {
      rowStart: 0,
      rowEnd: 600,
      colStart: 0,
      colEnd: 9,
    })
    const unsubscribeCell = cache.subscribeCells('Sheet1', ['A1'], () => undefined)

    for (let row = 0; row <= 600; row += 1) {
      for (let col = 0; col < 10; col += 1) {
        cache.setCellSnapshot(snapshot(`${columnLabel(col)}${row + 1}`, row * 10 + col))
      }
    }

    expect(countSheetCells(cache, 'Sheet1')).toBe(6010)

    untrackViewport()

    expect(countSheetCells(cache, 'Sheet1')).toBe(6000)
    expect(cache.peekCell('Sheet1', 'A1')).toBeDefined()

    unsubscribeCell()
  })
})
