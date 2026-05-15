import { describe, expect, it, vi } from 'vitest'
import * as formula from '@bilig/formula'
import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import {
  cellHasFormulaPrefix,
  classifyWorkPaperCell,
  compareSheetNames,
  doesWorkPaperCellHaveSimpleValue,
  inspectRuntimeSnapshotSheetDimensionsWithinLimits,
  inspectSheetDimensionsWithinLimits,
  inspectSheetWithinLimits,
  isWorkPaperCellPartOfArray,
  runtimeSnapshotMatchesSheetEntries,
  validateSheetWithinLimits,
  workPaperCellIsInsideSpillRange,
  workPaperCellValueDetailedType,
  workPaperCellValueType,
} from '../work-paper-sheet-inspection.js'
import { WorkPaperSheetSizeLimitExceededError, WorkPaperUnableToParseError } from '../work-paper-errors.js'
import type { WorkPaperSheet } from '../work-paper-types.js'

function requireSheetIdForInspectionTest(sheetName: string): number {
  if (sheetName === 'Sheet1') {
    return 1
  }
  if (sheetName === 'Other') {
    return 2
  }
  throw new Error('Unknown sheet')
}

describe('work paper sheet inspection', () => {
  it('inspects materialized dimensions and formula counts', () => {
    const sheet: WorkPaperSheet = [[null, 1, '=A1'], [], [null, null, ' text ']]

    expect(inspectSheetDimensionsWithinLimits('Sheet1', sheet, {})).toEqual({ width: 3, height: 3 })
    expect(inspectSheetWithinLimits('Sheet1', sheet, {})).toEqual({
      hasFormula: true,
      hasDynamicSpillFormula: false,
      dimensions: { width: 3, height: 3 },
      materializedCellCount: 3,
      maxColumnCount: 3,
      formulaCellCount: 1,
    })
  })

  it('does not compile definite scalar formulas just to inspect spill-resizing dimensions', () => {
    const compileSpy = vi.spyOn(formula, 'compileFormula')
    try {
      const sheet: WorkPaperSheet = [
        [1, 2, '=a1 + b1', '= sum ( A1:A1 ) + 1'],
        [3, 4, '=A2*B2+5', '=COUNTIFS(A1:A2,">0")'],
        [5, 6, '=ABS(A3)', '=MAX(A1:A3)'],
      ]

      expect(inspectSheetWithinLimits('Sheet1', sheet, {})).toEqual({
        hasFormula: true,
        hasDynamicSpillFormula: false,
        dimensions: { width: 4, height: 3 },
        materializedCellCount: 12,
        maxColumnCount: 4,
        formulaCellCount: 6,
      })
      expect(compileSpy).not.toHaveBeenCalled()
    } finally {
      compileSpy.mockRestore()
    }
  })

  it('keeps dynamic array formulas on compiler-backed spill inspection', () => {
    const compileSpy = vi.spyOn(formula, 'compileFormula')
    try {
      const filterInspection = inspectSheetWithinLimits('Sheet1', [[1], [2], ['=FILTER(A1:A2,A1:A2>1)']], {})
      const rangeExpressionInspection = inspectSheetWithinLimits('Sheet1', [[1], [2], ['=A1:A2>1']], {})

      expect(filterInspection).toMatchObject({
        hasFormula: true,
        hasDynamicSpillFormula: true,
        formulaCellCount: 1,
      })
      expect(rangeExpressionInspection).toMatchObject({
        hasFormula: true,
        hasDynamicSpillFormula: true,
        formulaCellCount: 1,
      })
      expect(compileSpy).toHaveBeenCalledTimes(2)
    } finally {
      compileSpy.mockRestore()
    }
  })

  it('rejects invalid rows and sheets over configured limits', () => {
    const invalidSheet: WorkPaperSheet = []
    Reflect.set(invalidSheet, 0, 1)

    expect(() => inspectSheetWithinLimits('Bad', invalidSheet, {})).toThrow(WorkPaperUnableToParseError)
    expect(() => validateSheetWithinLimits('Wide', [[1, 2]], { maxColumns: 1 })).toThrow(WorkPaperSheetSizeLimitExceededError)
    expect(() => validateSheetWithinLimits('Tall', [[1], [2]], { maxRows: 1 })).toThrow(WorkPaperSheetSizeLimitExceededError)
  })

  it('uses runtime dimensions, coords, or snapshot cells for snapshot dimensions', () => {
    const snapshotSheet = {
      name: 'Sheet1',
      order: 0,
      cells: [{ address: 'C4', value: 1 }],
    }

    expect(
      inspectRuntimeSnapshotSheetDimensionsWithinLimits({
        sheetName: 'Sheet1',
        snapshotSheet,
        runtimeSheetCells: { dimensions: { width: 5, height: 6 } },
        config: {},
      }),
    ).toEqual({ width: 5, height: 6 })
    expect(
      inspectRuntimeSnapshotSheetDimensionsWithinLimits({
        sheetName: 'Sheet1',
        snapshotSheet,
        runtimeSheetCells: { coords: [{ row: 2, col: 8 }] },
        config: {},
      }),
    ).toEqual({ width: 9, height: 3 })
    expect(
      inspectRuntimeSnapshotSheetDimensionsWithinLimits({
        sheetName: 'Sheet1',
        snapshotSheet,
        config: {},
      }),
    ).toEqual({ width: 3, height: 4 })
  })

  it('matches runtime snapshot sheets by unique names', () => {
    const sheets: readonly (readonly [string, WorkPaperSheet])[] = [
      ['A', []],
      ['B', []],
    ]

    expect(runtimeSnapshotMatchesSheetEntries(sheets, { sheets: [{ name: 'B' }, { name: 'A' }] })).toBe(true)
    expect(runtimeSnapshotMatchesSheetEntries(sheets, { sheets: [{ name: 'A' }, { name: 'A' }] })).toBe(false)
    expect(runtimeSnapshotMatchesSheetEntries(sheets, { sheets: [{ name: 'A' }] })).toBe(false)
  })

  it('detects formula prefixes and compares sheet names', () => {
    expect(cellHasFormulaPrefix('=A1')).toBe(true)
    expect(cellHasFormulaPrefix('  =A1')).toBe(true)
    expect(cellHasFormulaPrefix(' A1')).toBe(false)
    expect(['z', 'a'].toSorted(compareSheetNames)).toEqual(['a', 'z'])
  })

  it('classifies cell shape and simple value state', () => {
    expect(classifyWorkPaperCell({ hasFormula: true, isEmpty: true, isPartOfArray: true })).toBe('EMPTY')
    expect(classifyWorkPaperCell({ hasFormula: true, isEmpty: false, isPartOfArray: true })).toBe('ARRAY')
    expect(classifyWorkPaperCell({ hasFormula: true, isEmpty: false, isPartOfArray: false })).toBe('FORMULA')
    expect(classifyWorkPaperCell({ hasFormula: false, isEmpty: false, isPartOfArray: false })).toBe('VALUE')
    expect(doesWorkPaperCellHaveSimpleValue({ hasFormula: false, isEmpty: false })).toBe(true)
    expect(doesWorkPaperCellHaveSimpleValue({ hasFormula: true, isEmpty: false })).toBe(false)
    expect(doesWorkPaperCellHaveSimpleValue({ hasFormula: false, isEmpty: true })).toBe(false)
  })

  it('classifies value tags and date/time number formats', () => {
    const numberValue: CellValue = { tag: ValueTag.Number, value: 1 }
    expect(workPaperCellValueType({ tag: ValueTag.Empty })).toBe('EMPTY')
    expect(workPaperCellValueType(numberValue)).toBe('NUMBER')
    expect(workPaperCellValueType({ tag: ValueTag.String, value: 'x', stringId: 1 })).toBe('STRING')
    expect(workPaperCellValueType({ tag: ValueTag.Boolean, value: true })).toBe('BOOLEAN')
    expect(workPaperCellValueType({ tag: ValueTag.Error, code: ErrorCode.Value })).toBe('ERROR')
    expect(workPaperCellValueDetailedType({ value: numberValue })).toBe('NUMBER')
    expect(workPaperCellValueDetailedType({ value: numberValue, format: 'yyyy-mm-dd' })).toBe('DATE')
    expect(workPaperCellValueDetailedType({ value: numberValue, format: 'hh:mm:ss' })).toBe('TIME')
    expect(workPaperCellValueDetailedType({ value: numberValue, format: 'yyyy-mm-dd hh:mm' })).toBe('DATETIME')
    expect(workPaperCellValueDetailedType({ value: { tag: ValueTag.String, value: 'x', stringId: 1 }, format: 'yyyy' })).toBe('STRING')
  })

  it('matches cells inside spill ranges by owning sheet and zero-based coordinates', () => {
    const spill = { sheetName: 'Sheet1', address: 'B2', rows: 2, cols: 3 }

    expect(
      workPaperCellIsInsideSpillRange({ address: { sheet: 1, row: 1, col: 1 }, spill, requireSheetId: requireSheetIdForInspectionTest }),
    ).toBe(true)
    expect(
      workPaperCellIsInsideSpillRange({ address: { sheet: 1, row: 2, col: 3 }, spill, requireSheetId: requireSheetIdForInspectionTest }),
    ).toBe(true)
    expect(
      workPaperCellIsInsideSpillRange({ address: { sheet: 1, row: 3, col: 3 }, spill, requireSheetId: requireSheetIdForInspectionTest }),
    ).toBe(false)
    expect(
      workPaperCellIsInsideSpillRange({ address: { sheet: 2, row: 1, col: 1 }, spill, requireSheetId: requireSheetIdForInspectionTest }),
    ).toBe(false)
    expect(
      isWorkPaperCellPartOfArray({
        address: { sheet: 1, row: 1, col: 2 },
        spillRanges: [spill],
        requireSheetId: requireSheetIdForInspectionTest,
      }),
    ).toBe(true)
  })
})
