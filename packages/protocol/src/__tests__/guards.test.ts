import { describe, expect, it } from 'vitest'
import { ErrorCode, ValueTag } from '../enums.js'
import { isCellRangeRef, isCellSnapshot, isLiteralInput, isWorkbookSnapshot } from '../guards.js'

describe('protocol guards', () => {
  it('accepts workbook snapshots with the shipped shape', () => {
    expect(
      isWorkbookSnapshot({
        version: 1,
        workbook: { name: 'guarded' },
        sheets: [
          {
            id: 1,
            name: 'Sheet1',
            order: 0,
            cells: [
              { address: 'A1', row: 0, col: 0, value: 'ready' },
              { address: 'B2', formula: 'A1', format: 'text' },
            ],
          },
        ],
      }),
    ).toBe(true)
  })

  it('rejects workbook snapshots without a workbook name', () => {
    expect(
      isWorkbookSnapshot({
        version: 1,
        workbook: {},
        sheets: [],
      }),
    ).toBe(false)
  })

  it('rejects workbook snapshots with malformed sheets or cells', () => {
    const base = {
      version: 1,
      workbook: { name: 'guarded' },
      sheets: [
        {
          id: 1,
          name: 'Sheet1',
          order: 0,
          cells: [{ address: 'A1', value: 1 }],
        },
      ],
    }

    expect(isWorkbookSnapshot({ ...base, sheets: [{ ...base.sheets[0], order: 1.5 }] })).toBe(false)
    expect(isWorkbookSnapshot({ ...base, sheets: [{ ...base.sheets[0], cells: [{ address: 'A1', value: Number.NaN }] }] })).toBe(false)
    expect(isWorkbookSnapshot({ ...base, sheets: [{ ...base.sheets[0], cells: [{ address: 'A1', row: -1, value: 1 }] }] })).toBe(false)
    expect(isWorkbookSnapshot({ ...base, sheets: [{ ...base.sheets[0], cells: [{ value: 1 }] }] })).toBe(false)
  })

  it('accepts cell snapshots with a valid value tag', () => {
    expect(
      isCellSnapshot({
        sheetName: 'Sheet1',
        address: 'A1',
        value: { tag: ValueTag.Number, value: 7 },
        flags: 0,
        version: 1,
      }),
    ).toBe(true)
    expect(
      isCellSnapshot({
        sheetName: 'Sheet1',
        address: 'B2',
        value: { tag: ValueTag.String, value: 'ready', stringId: 0 },
        flags: 0,
        version: 1,
      }),
    ).toBe(true)
    expect(
      isCellSnapshot({
        sheetName: 'Sheet1',
        address: 'C3',
        value: { tag: ValueTag.Error, code: ErrorCode.Ref },
        flags: 0,
        version: 1,
      }),
    ).toBe(true)
  })

  it('rejects cell snapshots with malformed values', () => {
    expect(
      isCellSnapshot({
        sheetName: 'Sheet1',
        address: 'A1',
        value: { tag: 99, value: 7 },
        flags: 0,
        version: 1,
      }),
    ).toBe(false)
    expect(
      isCellSnapshot({
        sheetName: 'Sheet1',
        address: 'A1',
        value: { tag: ValueTag.Number, value: Number.NaN },
        flags: 0,
        version: 1,
      }),
    ).toBe(false)
    expect(
      isCellSnapshot({
        sheetName: 'Sheet1',
        address: 'A1',
        value: { tag: ValueTag.String, value: 'missing-id' },
        flags: 0,
        version: 1,
      }),
    ).toBe(false)
    expect(
      isCellSnapshot({
        sheetName: 'Sheet1',
        address: 'A1',
        value: { tag: ValueTag.Error, code: 99 },
        flags: 0,
        version: 1,
      }),
    ).toBe(false)
  })

  it('rejects cell snapshots with unsafe structural metadata', () => {
    const baseSnapshot = {
      sheetName: 'Sheet1',
      address: 'A1',
      value: { tag: ValueTag.Number, value: 7 },
      flags: 0,
      version: 1,
    }

    expect(isCellSnapshot({ ...baseSnapshot, flags: -1 })).toBe(false)
    expect(isCellSnapshot({ ...baseSnapshot, flags: 1.5 })).toBe(false)
    expect(isCellSnapshot({ ...baseSnapshot, flags: Number.MAX_SAFE_INTEGER + 1 })).toBe(false)
    expect(isCellSnapshot({ ...baseSnapshot, version: Number.NaN })).toBe(false)
    expect(isCellSnapshot({ ...baseSnapshot, version: Number.POSITIVE_INFINITY })).toBe(false)
    expect(isCellSnapshot({ ...baseSnapshot, version: Number.MAX_SAFE_INTEGER + 1 })).toBe(false)
  })

  it('accepts literal inputs and cell range refs with the shipped shapes', () => {
    expect(isLiteralInput(null)).toBe(true)
    expect(isLiteralInput('text')).toBe(true)
    expect(isLiteralInput(12.5)).toBe(true)
    expect(isLiteralInput(true)).toBe(true)
    expect(
      isCellRangeRef({
        sheetName: 'Sheet1',
        startAddress: 'A1',
        endAddress: 'B2',
      }),
    ).toBe(true)
  })

  it('rejects malformed literal inputs', () => {
    expect(isLiteralInput(Number.NaN)).toBe(false)
    expect(isLiteralInput(Number.POSITIVE_INFINITY)).toBe(false)
    expect(isLiteralInput({ value: 1 })).toBe(false)
    expect(isLiteralInput(undefined)).toBe(false)
  })

  it('rejects malformed cell range refs', () => {
    expect(
      isCellRangeRef({
        sheetName: 'Sheet1',
        startAddress: 'A1',
      }),
    ).toBe(false)
  })
})
