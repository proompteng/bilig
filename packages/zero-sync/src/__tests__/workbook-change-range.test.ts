import { describe, expect, it } from 'vitest'
import { isWorkbookChangeRange, normalizeWorkbookChangeRange } from '../workbook-change-range.js'

describe('workbook change range guards', () => {
  it('normalizes legacy cell ranges without explicit scope', () => {
    expect(
      normalizeWorkbookChangeRange({
        sheetName: 'Sheet1',
        startAddress: 'A1',
        endAddress: 'B2',
      }),
    ).toEqual({
      sheetName: 'Sheet1',
      startAddress: 'A1',
      endAddress: 'B2',
    })
  })

  it('preserves valid structural scope', () => {
    expect(
      normalizeWorkbookChangeRange({
        sheetName: 'Sheet1',
        startAddress: 'A3',
        endAddress: 'A4',
        scope: 'rows',
      }),
    ).toEqual({
      sheetName: 'Sheet1',
      startAddress: 'A3',
      endAddress: 'A4',
      scope: 'rows',
    })
  })

  it('canonicalizes reversed persisted ranges before they become history authority', () => {
    expect(
      normalizeWorkbookChangeRange({
        sheetName: 'Sheet1',
        startAddress: 'D5',
        endAddress: 'B2',
        scope: 'cells',
      }),
    ).toEqual({
      sheetName: 'Sheet1',
      startAddress: 'B2',
      endAddress: 'D5',
    })
  })

  it('rejects malformed scope instead of silently downgrading to a cell range', () => {
    const malformed = {
      sheetName: 'Sheet1',
      startAddress: 'A3',
      endAddress: 'A4',
      scope: 'row-band',
    }

    expect(normalizeWorkbookChangeRange(malformed)).toBeNull()
    expect(isWorkbookChangeRange(malformed)).toBe(false)
  })

  it('rejects missing range authority coordinates', () => {
    for (const malformed of [
      { sheetName: '', startAddress: 'A1', endAddress: 'A1' },
      { sheetName: 'Sheet1', startAddress: '', endAddress: 'A1' },
      { sheetName: 'Sheet1', startAddress: 'A1', endAddress: '' },
    ]) {
      expect(normalizeWorkbookChangeRange(malformed)).toBeNull()
      expect(isWorkbookChangeRange(malformed)).toBe(false)
    }
  })

  it('rejects unparseable persisted cell addresses', () => {
    for (const malformed of [
      { sheetName: 'Sheet1', startAddress: '1A', endAddress: 'A1' },
      { sheetName: 'Sheet1', startAddress: 'A0', endAddress: 'A1' },
      { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'ZZZ999' },
    ]) {
      expect(normalizeWorkbookChangeRange(malformed)).toBeNull()
      expect(isWorkbookChangeRange(malformed)).toBe(false)
    }
  })
})
