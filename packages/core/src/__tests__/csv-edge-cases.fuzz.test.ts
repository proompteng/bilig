import { describe, expect, it } from 'vitest'
import * as fc from 'fast-check'
import { runProperty } from '@bilig/test-fuzz'
import { parseCsv, parseCsvCellInput, serializeCsv } from '../csv.js'

describe('csv edge-case fuzz', () => {
  it('should roundtrip generated quoted cells through serialize and parse', async () => {
    await runProperty({
      suite: 'core/csv/quoted-cell-roundtrip',
      arbitrary: fc.array(fc.array(csvCellTextArbitrary, { minLength: 1, maxLength: 5 }), { minLength: 1, maxLength: 8 }),
      predicate: async (rows) => {
        expect(parseCsv(serializeCsv(rows))).toEqual(rows)
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
