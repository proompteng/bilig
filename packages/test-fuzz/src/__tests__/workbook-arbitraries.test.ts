import { describe, expect, it } from 'vitest'
import * as fc from 'fast-check'

import {
  cloneJsonValue,
  corruptRecord,
  formatAddress,
  fuzzCellRangeRefArbitrary,
  fuzzCellSnapshotArbitrary,
  fuzzWorkbookSnapshotArbitrary,
} from '../workbook-arbitraries.js'
import { runProperty } from '../index.js'

describe('workbook fuzz arbitraries', () => {
  it('should produce 1-based A1 addresses', async () => {
    await runProperty({
      suite: 'test-fuzz/arbitraries/a1-addresses',
      arbitrary: fc.record({
        row: fc.integer({ min: 0, max: 25 }),
        col: fc.integer({ min: 0, max: 25 }),
      }),
      predicate: async ({ row, col }) => {
        expect(formatAddress(row, col)).toMatch(/^[A-Z]+[1-9]\d*$/u)
      },
      parameters: { numRuns: 80 },
    })
  })

  it('should generate protocol-shaped range refs', async () => {
    await runProperty({
      suite: 'test-fuzz/arbitraries/range-refs',
      arbitrary: fuzzCellRangeRefArbitrary,
      predicate: async (range) => {
        expect(range.sheetName).toBeTruthy()
        expect(range.startAddress).toMatch(/^[A-Z]+[1-9]\d*$/u)
        expect(range.endAddress).toMatch(/^[A-Z]+[1-9]\d*$/u)
      },
      parameters: { numRuns: 80 },
    })
  })

  it('should generate JSON-cloneable workbook payloads', async () => {
    await runProperty({
      suite: 'test-fuzz/arbitraries/json-cloneable-workbook-payloads',
      arbitrary: fc.record({
        workbook: fuzzWorkbookSnapshotArbitrary,
        cell: fuzzCellSnapshotArbitrary,
      }),
      predicate: async ({ workbook, cell }) => {
        expect(cloneJsonValue(workbook)).toEqual(workbook)
        expect(cloneJsonValue(cell)).toEqual(cell)
      },
      parameters: { numRuns: 80 },
    })
  })

  it('should remove fields when corrupting records without an override value', () => {
    expect(corruptRecord({ id: 'a', value: 1 }, 'id')).toEqual({ value: 1 })
    expect(corruptRecord({ id: 'a', value: 1 }, 'value', 'bad')).toEqual({ id: 'a', value: 'bad' })
  })
})
