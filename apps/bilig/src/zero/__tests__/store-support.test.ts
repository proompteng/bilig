import { ErrorCode, ValueTag } from '@bilig/protocol'
import { describe, expect, it } from 'vitest'
import {
  createEmptyWorkbookSnapshot,
  eventRequiresRecalc,
  isDirtyRegion,
  normalizeRangeBounds,
  parseCellEvalValue,
  parseCellStyleRecord,
  parseCheckpointPayload,
  parseNonNegativeInteger,
  parseNullableInteger,
  parsePositiveInteger,
} from '../store-support.js'

describe('store support helpers', () => {
  it('falls back to an empty workbook snapshot when checkpoint payloads are invalid', () => {
    expect(parseCheckpointPayload(null, 'book-1')).toEqual(createEmptyWorkbookSnapshot('book-1'))
  })

  it('normalizes reversed range bounds', () => {
    expect(
      normalizeRangeBounds({
        sheetName: 'Sheet1',
        startAddress: 'D5',
        endAddress: 'B2',
      }),
    ).toEqual({
      sheetName: 'Sheet1',
      rowStart: 1,
      rowEnd: 4,
      colStart: 1,
      colEnd: 3,
    })
  })

  it('rejects unsafe dirty region bounds', () => {
    expect(
      isDirtyRegion({
        sheetName: 'Sheet1',
        rowStart: 0,
        rowEnd: 1,
        colStart: 0,
        colEnd: 1,
      }),
    ).toBe(true)
    expect(
      isDirtyRegion({
        sheetName: 'Sheet1',
        rowStart: 1.5,
        rowEnd: 2,
        colStart: 0,
        colEnd: 1,
      }),
    ).toBe(false)
    expect(
      isDirtyRegion({
        sheetName: 'Sheet1',
        rowStart: 2,
        rowEnd: 1,
        colStart: 0,
        colEnd: 1,
      }),
    ).toBe(false)
    expect(
      isDirtyRegion({
        sheetName: 'Sheet1',
        rowStart: 0,
        rowEnd: Number.MAX_SAFE_INTEGER + 1,
        colStart: 0,
        colEnd: 1,
      }),
    ).toBe(false)
  })

  it('rejects malformed persisted cell values', () => {
    expect(parseCellEvalValue({ tag: ValueTag.Number, value: 42.5 })).toEqual({ tag: ValueTag.Number, value: 42.5 })
    expect(parseCellEvalValue({ tag: ValueTag.String, value: 'ready', stringId: 0 })).toEqual({
      tag: ValueTag.String,
      value: 'ready',
      stringId: 0,
    })
    expect(parseCellEvalValue({ tag: ValueTag.Error, code: ErrorCode.Value })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })

    expect(parseCellEvalValue({ tag: ValueTag.Number, value: Number.NaN })).toEqual({ tag: ValueTag.Empty })
    expect(parseCellEvalValue({ tag: ValueTag.Number, value: Number.POSITIVE_INFINITY })).toEqual({ tag: ValueTag.Empty })
    expect(parseCellEvalValue({ tag: ValueTag.String, value: 'missing-id' })).toEqual({ tag: ValueTag.Empty })
    expect(parseCellEvalValue({ tag: ValueTag.String, value: 'unsafe-id', stringId: Number.MAX_SAFE_INTEGER + 1 })).toEqual({
      tag: ValueTag.Empty,
    })
    expect(parseCellEvalValue({ tag: ValueTag.Error, code: 99 })).toEqual({ tag: ValueTag.Empty })
    expect(parseCellEvalValue({ tag: ValueTag.Error, code: 1.5 })).toEqual({ tag: ValueTag.Empty })
  })

  it('keeps style records but drops invalid nested fields', () => {
    expect(
      parseCellStyleRecord({
        id: 'style-1',
        font: { family: 'Aptos', size: 12, bold: true, color: 17 },
        alignment: { horizontal: 'center', wrap: true, indent: 'x', textRotation: Number.POSITIVE_INFINITY },
        protection: { locked: true, hidden: false, mode: 'ignored' },
        borders: {
          top: { style: 'solid', weight: 'thin', color: '#111' },
          left: { style: 'invalid', weight: 'thin', color: '#222' },
        },
      }),
    ).toEqual({
      id: 'style-1',
      font: { family: 'Aptos', size: 12, bold: true },
      alignment: { horizontal: 'center', wrap: true },
      borders: {
        top: { style: 'solid', weight: 'thin', color: '#111' },
      },
      protection: { locked: true, hidden: false },
    })

    expect(
      parseCellStyleRecord({
        id: 'style-2',
        font: { size: Number.NaN, italic: true },
        alignment: { indent: 2, readingOrder: Number.NEGATIVE_INFINITY, textRotation: 45 },
      }),
    ).toEqual({
      id: 'style-2',
      font: { italic: true },
      alignment: { indent: 2, textRotation: 45 },
    })
  })

  it('treats formatting-only mutations as no-recalc events', () => {
    expect(
      eventRequiresRecalc({
        kind: 'setRangeStyle',
        range: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B2' },
        patch: { font: { bold: true } },
      }),
    ).toBe(false)

    expect(
      eventRequiresRecalc({
        kind: 'setCellValue',
        sheetName: 'Sheet1',
        address: 'A1',
        value: 123,
      }),
    ).toBe(true)
  })

  it('parses nullable integer fields strictly', () => {
    expect(parseNullableInteger(42)).toBe(42)
    expect(parseNullableInteger(' 42 ')).toBe(42)
    expect(parseNullableInteger(-3)).toBe(-3)
    expect(parseNullableInteger(42.5)).toBeNull()
    expect(parseNullableInteger('12abc')).toBeNull()
    expect(parseNullableInteger('1.5')).toBeNull()
    expect(parseNullableInteger(String(Number.MAX_SAFE_INTEGER + 1))).toBeNull()
  })

  it('parses positive and non-negative integer fields by domain', () => {
    expect(parsePositiveInteger(1)).toBe(1)
    expect(parsePositiveInteger('7')).toBe(7)
    expect(parsePositiveInteger(0)).toBeNull()
    expect(parsePositiveInteger(-1)).toBeNull()
    expect(parsePositiveInteger('1.5')).toBeNull()
    expect(parseNonNegativeInteger(0)).toBe(0)
    expect(parseNonNegativeInteger('9')).toBe(9)
    expect(parseNonNegativeInteger(-1)).toBeNull()
    expect(parseNonNegativeInteger(Number.MAX_SAFE_INTEGER + 1)).toBeNull()
  })
})
