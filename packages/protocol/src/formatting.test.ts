import { describe, expect, it } from 'vitest'
import { ValueTag } from './enums.js'
import { formatCellDisplayValue, formatGeneralNumberValue } from './formatting.js'

describe('formatCellDisplayValue', () => {
  it('renders Excel date serials without local timezone drift', () => {
    expect(formatCellDisplayValue({ tag: ValueTag.Number, value: 46023 }, 'date:short')).toBe('01/01/2026')
    expect(formatCellDisplayValue({ tag: ValueTag.Number, value: 46357 }, 'date:short')).toBe('12/01/2026')
  })

  it('uses General-format text for numeric display coercion', () => {
    expect(formatGeneralNumberValue(1989)).toBe('1989')
    expect(formatGeneralNumberValue(1)).toBe('1')
    expect(formatGeneralNumberValue(0.1 + 0.2)).toBe('0.3')
    expect(formatCellDisplayValue({ tag: ValueTag.Number, value: 0.1 + 0.2 }, 'general')).toBe('0.3')
  })
})
