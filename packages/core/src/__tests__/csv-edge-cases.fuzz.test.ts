import { describe, expect, it } from 'vitest'
import * as fc from 'fast-check'
import { runProperty } from '@bilig/test-fuzz'
import { parseCsv, parseCsvCellInput, serializeCsv } from '../csv.js'

describe('csv edge-case fuzz', () => {
  it('should roundtrip generated quoted cells through serialize and parse', async () => {
    await runProperty({
      suite: 'core/csv/quoted-cell-roundtrip',
      arbitrary: fc
        .array(fc.array(csvCellTextArbitrary, { minLength: 1, maxLength: 5 }), { minLength: 1, maxLength: 8 })
        .filter((rows) => rows.some((row) => row.length > 1 || row.some((value) => value !== ''))),
      predicate: async (rows) => {
        expect(parseCsv(serializeCsv(rows))).toEqual(rows)
      },
      parameters: { numRuns: 120 },
    })
  })

  it('should parse generated quoted cells with explicit delimiters', async () => {
    await runProperty({
      suite: 'core/csv/explicit-delimiter-quoted-cell-roundtrip',
      arbitrary: fc
        .record({
          delimiter: fc.constantFrom(',' as const, ';' as const, '\t' as const),
          rows: fc.array(fc.array(csvCellTextArbitrary, { minLength: 1, maxLength: 5 }), { minLength: 1, maxLength: 8 }),
        })
        .filter(({ rows }) => rows.some((row) => row.length > 1 || row.some((value) => value !== ''))),
      predicate: async ({ delimiter, rows }) => {
        expect(parseCsv(serializeDelimitedCsv(rows, delimiter), { delimiter })).toEqual(rows)
      },
      parameters: { numRuns: 120 },
    })
  })

  it('should keep formula-like text, leading zero ids, booleans, and accounting numbers distinct', async () => {
    await runProperty({
      suite: 'core/csv/cell-input-edge-semantics',
      arbitrary: fc.oneof(
        fc.constant('=SUM(A1:A2)'),
        fc.constant('00123'),
        fc.constantFrom('TRUE', 'FALSE'),
        fc.constantFrom('$1,234.56', '(987.65)', '12.5%'),
      ),
      predicate: async (raw) => {
        const parsed = parseCsvCellInput(raw)
        if (raw.startsWith('=')) {
          expect(parsed).toEqual({ formula: raw.slice(1) })
        } else if (/^0\d+/u.test(raw)) {
          expect(parsed).toEqual({ value: raw })
        } else {
          expect(parsed?.value).not.toBeUndefined()
        }
      },
      parameters: { numRuns: 80 },
    })
  })
})

// Helpers

const csvCellTextArbitrary = fc.oneof(
  fc.string({ maxLength: 24 }),
  fc.constantFrom('contains,comma', 'contains;semicolon', 'quote"inside', 'line\nbreak', 'crlf\r\nbreak', '\tindent', '00123', '=A1+1'),
)

function serializeDelimitedCsv(rows: readonly (readonly string[])[], delimiter: ',' | ';' | '\t'): string {
  return rows.map((row) => row.map((value) => escapeDelimitedCsvValue(value, delimiter)).join(delimiter)).join('\n')
}

function escapeDelimitedCsvValue(value: string, delimiter: ',' | ';' | '\t'): string {
  return value.includes(delimiter) || /[",;\t\n\r]/u.test(value) ? `"${value.replaceAll('"', '""')}"` : value
}
