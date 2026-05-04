import { describe, expect, it } from 'vitest'
import {
  normalizeWorkbookAgentToolNumberFormatInput,
  normalizeWorkbookAgentWriteCellInput,
} from './workbook-agent-tool-input-normalization.js'

describe('workbook agent tool input normalization', () => {
  it('normalizes workbook agent number format presets with defaults', () => {
    expect(normalizeWorkbookAgentToolNumberFormatInput({ kind: 'currency', currency: 'usd' })).toEqual({
      kind: 'currency',
      currency: 'USD',
      decimals: 2,
      useGrouping: true,
      negativeStyle: 'minus',
      zeroStyle: 'zero',
    })
  })

  it('coerces workbook agent write cell inputs into protocol values', () => {
    expect(normalizeWorkbookAgentWriteCellInput('1200')).toBe(1200)
    expect(normalizeWorkbookAgentWriteCellInput({ type: 'number', value: '12' })).toBe(12)
    expect(normalizeWorkbookAgentWriteCellInput({ type: 'formula', formula: 'SUM(A1:A1)' })).toEqual({
      formula: '=SUM(A1:A1)',
    })
    expect(normalizeWorkbookAgentWriteCellInput({ type: 'date', value: '1970-01-01' })).toBe(25569)
    expect(normalizeWorkbookAgentWriteCellInput({ type: 'boolean', value: 'FALSE' })).toBe(false)
    expect(normalizeWorkbookAgentWriteCellInput({ type: 'blank' })).toBeNull()
  })

  it('rejects invalid workbook agent write cell inputs', () => {
    expect(() => normalizeWorkbookAgentWriteCellInput({ type: 'boolean', value: 'maybe' })).toThrow(
      'Boolean value must be true or false, received maybe',
    )
    expect(() => normalizeWorkbookAgentWriteCellInput({ type: 'date', value: '2025-02-30' })).toThrow('Invalid date value 2025-02-30')
  })
})
