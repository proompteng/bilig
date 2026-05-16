import { ValueTag } from '@bilig/protocol'
import { describe, expect, it, vi } from 'vitest'

import { WorkPaper, type WorkPaperCellAddress, type WorkPaperChange } from '../index.js'

function cell(sheet: number, row: number, col: number): WorkPaperCellAddress {
  return { sheet, row, col }
}

function hasCaptureVisibilitySnapshot(value: unknown): value is WorkPaper & { captureVisibilitySnapshot: () => unknown } {
  return typeof Reflect.get(value, 'captureVisibilitySnapshot') === 'function'
}

interface TestSheetDimensionCache {
  updateAfterCellMutationRefs(...args: unknown[]): unknown
}

interface TestEngineRuntimeSupport {
  clearOwnedSpillNow(...args: unknown[]): unknown
}

function hasSheetDimensionCacheUpdater(value: unknown): value is TestSheetDimensionCache {
  return typeof value === 'object' && value !== null && typeof Reflect.get(value, 'updateAfterCellMutationRefs') === 'function'
}

function hasClearOwnedSpill(value: unknown): value is TestEngineRuntimeSupport {
  return typeof value === 'object' && value !== null && typeof Reflect.get(value, 'clearOwnedSpillNow') === 'function'
}

function trackSheetDimensionCacheUpdates(workbook: WorkPaper): { readonly count: number; restore: () => void } {
  const cache: unknown = Reflect.get(workbook, 'sheetDimensionCache')
  if (!hasSheetDimensionCacheUpdater(cache)) {
    throw new Error('Expected WorkPaper to expose a sheet dimension cache in tests')
  }
  const spy = vi.spyOn(cache, 'updateAfterCellMutationRefs')
  return {
    get count() {
      return spy.mock.calls.length
    },
    restore: () => {
      spy.mockRestore()
    },
  }
}

function trackCoreSpillOwnerClears(workbook: WorkPaper): { readonly count: number; restore: () => void } {
  const engine: unknown = Reflect.get(workbook, 'engine')
  const runtime: unknown = typeof engine === 'object' && engine !== null ? Reflect.get(engine, 'runtime') : undefined
  const support: unknown = typeof runtime === 'object' && runtime !== null ? Reflect.get(runtime, 'support') : undefined
  if (!hasClearOwnedSpill(support)) {
    throw new Error('Expected WorkPaper to expose core spill-owner cleanup in tests')
  }
  const spy = vi.spyOn(support, 'clearOwnedSpillNow')
  return {
    get count() {
      return spy.mock.calls.length
    },
    restore: () => {
      spy.mockRestore()
    },
  }
}

describe('work paper batched structural fast path', () => {
  it('keeps appended formula rows on the tracked batch path', () => {
    const workbook = WorkPaper.buildFromSheets({
      Data: [
        [1, 2, '=SUM(A1:B1)'],
        [3, 4, '=SUM(A2:B2)'],
      ],
    })
    const sheetId = workbook.getSheetId('Data')!
    expect(hasCaptureVisibilitySnapshot(workbook)).toBe(true)
    if (!hasCaptureVisibilitySnapshot(workbook)) {
      throw new Error('Expected WorkPaper to expose captureVisibilitySnapshot in tests')
    }
    const captureVisibilitySnapshot = vi.spyOn(workbook, 'captureVisibilitySnapshot').mockImplementation(() => {
      throw new Error('batched append formulas should not rebuild visibility snapshots')
    })
    const dimensionUpdates = trackSheetDimensionCacheUpdates(workbook)
    const spillOwnerClears = trackCoreSpillOwnerClears(workbook)

    let changes: WorkPaperChange[] = []
    try {
      changes = workbook.batch(() => {
        expect(workbook.addRows(sheetId, 2, 2)).toEqual([])
        expect(
          workbook.setCellContents(cell(sheetId, 2, 0), [
            [5, 6, '=SUM(A3:B3)'],
            [7, 8, '=SUM(A4:B4)'],
          ]),
        ).toEqual([])
      })
      expect(dimensionUpdates.count).toBe(1)
      expect(spillOwnerClears.count).toBe(0)
    } finally {
      spillOwnerClears.restore()
      dimensionUpdates.restore()
      captureVisibilitySnapshot.mockRestore()
    }

    expect(changes.length).toBeGreaterThan(0)
    expect(workbook.getCellValue(cell(sheetId, 2, 2))).toEqual({ tag: ValueTag.Number, value: 11 })
    expect(workbook.getCellValue(cell(sheetId, 3, 2))).toEqual({ tag: ValueTag.Number, value: 15 })

    workbook.setCellContents(cell(sheetId, 2, 0), 10)
    expect(workbook.getCellValue(cell(sheetId, 2, 2))).toEqual({ tag: ValueTag.Number, value: 16 })

    workbook.undo()
    expect(workbook.getCellValue(cell(sheetId, 2, 2))).toEqual({ tag: ValueTag.Number, value: 11 })
    workbook.undo()
    expect(workbook.getSheetDimensions(sheetId).height).toBe(2)
  })
})
