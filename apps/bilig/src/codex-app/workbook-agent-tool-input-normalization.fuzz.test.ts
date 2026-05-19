import { describe, expect, it } from 'vitest'
import * as fc from 'fast-check'
import { fuzzLiteralInputArbitrary, runProperty } from '@bilig/test-fuzz'
import {
  normalizeWorkbookAgentToolNumberFormatInput,
  normalizeWorkbookAgentWriteCellInput,
} from './workbook-agent-tool-input-normalization.js'

describe('workbook agent tool input normalization fuzz', () => {
  it('should normalize generated supported write cell inputs into protocol-safe values', async () => {
    await runProperty({
      suite: 'bilig/codex-app/tool-input/write-cell-normalization',
      arbitrary: writeCellInputArbitrary,
      predicate: async (input) => {
        const normalized = normalizeWorkbookAgentWriteCellInput(input)

        if (typeof normalized === 'number') {
          expect(Number.isFinite(normalized)).toBe(true)
        } else if (normalized && typeof normalized === 'object') {
          expect('formula' in normalized).toBe(true)
          expect(String(normalized.formula)).toMatch(/^=/u)
        } else {
          expect(normalized === null || typeof normalized === 'boolean' || typeof normalized === 'string').toBe(true)
        }
      },
      parameters: { numRuns: 120 },
    })
  })

  it('should reject generated malformed typed write cell inputs with stable errors', async () => {
    await runProperty({
      suite: 'bilig/codex-app/tool-input/reject-malformed-write-cells',
      arbitrary: fc.oneof(
        fc.record({ type: fc.constant('number'), value: fc.constantFrom('001', '12px', Number.NaN, Number.POSITIVE_INFINITY) }),
        fc.record({ type: fc.constant('date'), value: fc.constantFrom('2025-02-30', 'not-a-date', {}, []) }),
        fc.record({ type: fc.constant('boolean'), value: fc.constantFrom('maybe', 1, null) }),
      ),
      predicate: async (input) => {
        expect(() => normalizeWorkbookAgentWriteCellInput(input)).toThrow()
      },
      parameters: { numRuns: 100 },
    })
  })

  it('should normalize generated number format inputs without dropping required preset fields', async () => {
    await runProperty({
      suite: 'bilig/codex-app/tool-input/number-format-normalization',
      arbitrary: numberFormatInputArbitrary,
      predicate: async (input) => {
        const normalized = normalizeWorkbookAgentToolNumberFormatInput(input)
        if (typeof input === 'string') {
          expect(normalized).toBe(input)
        } else {
          expect(typeof normalized).toBe('object')
          expect(normalized.kind).toBe(input.kind)
        }
      },
      parameters: { numRuns: 100 },
    })
  })
})

// Helpers

const formulaTextArbitrary = fc.constantFrom('=A1+1', 'SUM(A1:A3)', '=IF(A1>0,TRUE,FALSE)')

const writeCellInputArbitrary = fc.oneof(
  fuzzLiteralInputArbitrary,
  formulaTextArbitrary,
  fc.record({ type: fc.constant('blank') }),
  fc.record({ type: fc.constant('formula'), formula: formulaTextArbitrary }),
  fc.record({ type: fc.constant('text'), value: fc.string({ maxLength: 40 }) }),
  fc.record({
    type: fc.constant('number'),
    value: fc.oneof(fc.integer({ min: -1_000_000, max: 1_000_000 }), fc.constantFrom('0', '12.5', '-7', '1e3')),
  }),
  fc.record({ type: fc.constant('date'), value: fc.constantFrom('1970-01-01', '2026-05-18', 45_000) }),
  fc.record({ type: fc.constant('boolean'), value: fc.constantFrom(true, false, 'TRUE', 'false') }),
  fc.record({ value: fuzzLiteralInputArbitrary }),
)

const numberFormatInputArbitrary = fc.oneof(
  fc.constantFrom('0.00', '0%', '@', 'yyyy-mm-dd'),
  fc.record(
    {
      kind: fc.constantFrom('number', 'currency', 'percent', 'date', 'text'),
      currency: fc.option(fc.constantFrom('usd', 'EUR'), { nil: undefined }),
      decimals: fc.option(fc.integer({ min: 0, max: 6 }), { nil: undefined }),
      useGrouping: fc.option(fc.boolean(), { nil: undefined }),
      negativeStyle: fc.option(fc.constantFrom('minus', 'parentheses'), { nil: undefined }),
      zeroStyle: fc.option(fc.constantFrom('zero', 'dash'), { nil: undefined }),
      dateStyle: fc.option(fc.constantFrom('short', 'iso'), { nil: undefined }),
    },
    { requiredKeys: ['kind'] },
  ),
)
