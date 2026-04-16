import { describe, expect, test } from 'vitest'
import {
  clampSelectionRange,
  createColumnSliceSelection,
  createGridSelection,
  createRangeSelection,
  createRowSliceSelection,
  createSheetSelection,
  formatSelectionSummary,
  isSheetSelection,
  rectangleToAddresses,
  selectionToAddresses,
} from '../gridSelection.js'

describe('gridSelection', () => {
  test('formats single-cell and rectangular selections', () => {
    expect(formatSelectionSummary(createGridSelection(2, 4), 'A1')).toBe('C5')

    const range = createRangeSelection(createGridSelection(1, 1), [1, 1], [3, 4])
    expect(formatSelectionSummary(range, 'A1')).toBe('B2:D5')
  })

  test('formats row and column slice selections', () => {
    expect(formatSelectionSummary(createColumnSliceSelection(1, 3, 0), 'A1')).toBe('B:D')
    expect(formatSelectionSummary(createRowSliceSelection(0, 1, 3), 'A1')).toBe('2:4')
  })

  test('detects full sheet selections', () => {
    expect(isSheetSelection(createSheetSelection())).toBe(true)
    expect(isSheetSelection(createGridSelection(0, 0))).toBe(false)
  })

  test('clamps oversized ranges and converts them to addresses', () => {
    const clamped = clampSelectionRange({ x: -10, y: -20, width: 5, height: 8 })
    expect(clamped).toEqual({ x: 0, y: 0, width: 5, height: 8 })
    expect(rectangleToAddresses({ x: 1, y: 2, width: 3, height: 2 })).toEqual({
      startAddress: 'B3',
      endAddress: 'D4',
    })
  })

  test('derives authoritative address bounds for cell, range, row, column, and sheet selections', () => {
    expect(selectionToAddresses(createGridSelection(2, 4), 'C5')).toEqual({
      startAddress: 'C5',
      endAddress: 'C5',
    })
    expect(selectionToAddresses(createRangeSelection(createGridSelection(1, 1), [1, 1], [3, 4]), 'B2')).toEqual({
      startAddress: 'B2',
      endAddress: 'D5',
    })
    expect(selectionToAddresses(createColumnSliceSelection(1, 3, 0), 'B1')).toEqual({
      startAddress: 'B1',
      endAddress: 'D1048576',
    })
    expect(selectionToAddresses(createRowSliceSelection(0, 1, 3), 'A2')).toEqual({
      startAddress: 'A2',
      endAddress: 'XFD4',
    })
    expect(selectionToAddresses(createSheetSelection(), 'A1')).toEqual({
      startAddress: 'A1',
      endAddress: 'XFD1048576',
    })
  })
})
