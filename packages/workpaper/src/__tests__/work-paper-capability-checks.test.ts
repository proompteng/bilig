import { describe, expect, it } from 'vitest'
import {
  isWorkPaperAxisIntervalEditPossible,
  isWorkPaperAxisOrderPossible,
  isWorkPaperAxisSwapPossible,
  isWorkPaperMoveAxisPossible,
  isWorkPaperMoveCellsPossible,
  isWorkPaperSetCellContentsPossible,
  isWorkPaperSheetContentReplaceable,
  isWorkPaperSheetNameAvailable,
} from '../work-paper-capability-checks.js'
import { WorkPaperInvalidArgumentsError, WorkPaperNoSheetWithIdError } from '../work-paper-errors.js'
import type { WorkPaperSheet } from '../work-paper-types.js'

const context = {
  config: { maxRows: 3, maxColumns: 4 },
  requireSheet: (sheetId: number): void => {
    if (sheetId !== 1) {
      throw new WorkPaperNoSheetWithIdError(sheetId)
    }
  },
  doesSheetExist: (sheetName: string): boolean => sheetName === 'Sheet1',
  getSheetIdByName: (sheetName: string): number | undefined => (sheetName === 'Sheet1' ? 1 : undefined),
}

describe('work paper capability checks', () => {
  it('validates cell content bounds and matrix shape', () => {
    const invalidMatrix: WorkPaperSheet = [[1]]
    Reflect.set(invalidMatrix, 0, 1)

    expect(isWorkPaperSetCellContentsPossible(context, { sheet: 1, row: 2, col: 3 }, 5)).toBe(true)
    expect(isWorkPaperSetCellContentsPossible(context, { sheet: 1, row: 3, col: 0 }, 5)).toBe(false)
    expect(isWorkPaperSetCellContentsPossible(context, { sheet: 1, row: 2, col: 3 }, [[1, 2]])).toBe(false)
    expect(() => isWorkPaperSetCellContentsPossible(context, { sheet: 2, row: 0, col: 0 }, 5)).toThrow(WorkPaperNoSheetWithIdError)
    expect(() => isWorkPaperSetCellContentsPossible(context, { sheet: 1, row: 0, col: 0 }, invalidMatrix)).toThrow(
      WorkPaperInvalidArgumentsError,
    )
  })

  it('validates axis swaps, intervals, moves, and explicit orders', () => {
    expect(isWorkPaperAxisSwapPossible(context, 'row', 1, 0, 2)).toBe(true)
    expect(isWorkPaperAxisSwapPossible(context, 'row', 1, 0, 3)).toBe(false)
    expect(isWorkPaperAxisOrderPossible(context, 'column', 1, [0, 2, 3])).toBe(true)
    expect(isWorkPaperAxisOrderPossible(context, 'column', 1, [0, 2, 2])).toBe(false)
    expect(isWorkPaperAxisIntervalEditPossible(context, 'column', 1, 2, 2, [])).toBe(true)
    expect(isWorkPaperAxisIntervalEditPossible(context, 'column', 1, 3, 2, [])).toBe(false)
    expect(isWorkPaperMoveAxisPossible(context, 'row', 1, 0, 2, 1)).toBe(true)
    expect(isWorkPaperMoveAxisPossible(context, 'row', 1, 0, 2, 2)).toBe(false)
  })

  it('validates moved ranges and sheet names', () => {
    const source = { start: { sheet: 1, row: 0, col: 0 }, end: { sheet: 1, row: 1, col: 1 } }
    expect(isWorkPaperMoveCellsPossible(context, source, { sheet: 1, row: 1, col: 2 })).toBe(true)
    expect(isWorkPaperMoveCellsPossible(context, source, { sheet: 1, row: 2, col: 2 })).toBe(false)
    expect(isWorkPaperMoveCellsPossible(context, source, { sheet: 2, row: 0, col: 0 })).toBe(false)

    expect(isWorkPaperSheetNameAvailable(context, 'Archive')).toBe(true)
    expect(isWorkPaperSheetNameAvailable(context, 'Sheet1')).toBe(false)
    expect(isWorkPaperSheetNameAvailable(context, 'Sheet1', 1)).toBe(true)
    expect(() => isWorkPaperSheetNameAvailable(context, '   ')).toThrow(WorkPaperInvalidArgumentsError)
  })

  it('validates replacement sheet shape and dimensions', () => {
    const invalidSheet: WorkPaperSheet = [[1]]
    Reflect.set(invalidSheet, 0, 1)

    expect(
      isWorkPaperSheetContentReplaceable(context, 1, [
        [1, 2],
        [3, 4],
      ]),
    ).toBe(true)
    expect(isWorkPaperSheetContentReplaceable(context, 1, [[1, 2, 3, 4, 5]])).toBe(false)
    expect(() => isWorkPaperSheetContentReplaceable(context, 1, invalidSheet)).toThrow(WorkPaperInvalidArgumentsError)
  })
})
