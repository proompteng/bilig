import { describe, expect, it } from 'vitest'
import { ValueTag } from './enums.js'
import { formatCellDisplayValue } from './formatting.js'

describe('formatCellDisplayValue', () => {
  it('renders Excel date serials without local timezone drift', () => {
    expect(formatCellDisplayValue({ tag: ValueTag.Number, value: 46023 }, 'date:short')).toBe('01/01/2026')
    expect(formatCellDisplayValue({ tag: ValueTag.Number, value: 46357 }, 'date:short')).toBe('12/01/2026')
  })
})
