import { describe, expect, it } from 'vitest'
import * as fc from 'fast-check'
import {
  cloneJsonValue,
  corruptRecord,
  fuzzCellRangeRefArbitrary,
  fuzzCellSnapshotArbitrary,
  fuzzLiteralInputArbitrary,
  fuzzWorkbookSnapshotArbitrary,
  runProperty,
} from '@bilig/test-fuzz'
import { isCellRangeRef, isCellSnapshot, isLiteralInput, isWorkbookSnapshot } from '../guards.js'

describe('protocol guard fuzz', () => {
  it('should accept generated valid workbook protocol payloads', async () => {
    await runProperty({
      suite: 'protocol/guards/valid-generated-payloads',
      arbitrary: fc.record({
        literal: fuzzLiteralInputArbitrary,
        range: fuzzCellRangeRefArbitrary,
        workbook: fuzzWorkbookSnapshotArbitrary,
        cell: fuzzCellSnapshotArbitrary,
      }),
      predicate: async ({ literal, range, workbook, cell }) => {
        expect(isLiteralInput(cloneJsonValue(literal))).toBe(true)
        expect(isCellRangeRef(cloneJsonValue(range))).toBe(true)
        expect(isWorkbookSnapshot(cloneJsonValue(workbook))).toBe(true)
        expect(isCellSnapshot(cloneJsonValue(cell))).toBe(true)
      },
      parameters: { numRuns: 120 },
    })
  })

  it('should reject required-field corruptions for generated protocol payloads', async () => {
    await runProperty({
      suite: 'protocol/guards/reject-required-field-corruption',
      arbitrary: fc.record({
        range: fuzzCellRangeRefArbitrary,
        workbook: fuzzWorkbookSnapshotArbitrary,
        cell: fuzzCellSnapshotArbitrary,
      }),
      predicate: async ({ range, workbook, cell }) => {
        expect(isCellRangeRef(corruptRecord(range, 'startAddress'))).toBe(false)
        expect(isWorkbookSnapshot(corruptRecord(workbook, 'workbook'))).toBe(false)
        expect(isCellSnapshot(corruptRecord(cell, 'value'))).toBe(false)
      },
      parameters: { numRuns: 120 },
    })
  })
})
