import { describe, expect, it } from 'vitest'
import { createFormulaInstanceTable } from '../formula/formula-instance-table.js'

describe('FormulaInstanceTable', () => {
  it('hydrates formula instance snapshots into an empty table', () => {
    const table = createFormulaInstanceTable()

    table.hydrate([
      { cellIndex: 2, sheetName: 'Sheet1', row: 0, col: 1, source: 'A1+1' },
      { cellIndex: 1, sheetName: 'Sheet1', row: 0, col: 0, source: '1+1' },
    ])

    expect(table.get(2)).toMatchObject({ cellIndex: 2, source: 'A1+1' })
    expect(table.list()).toEqual([
      { cellIndex: 1, sheetName: 'Sheet1', row: 0, col: 0, source: '1+1' },
      { cellIndex: 2, sheetName: 'Sheet1', row: 0, col: 1, source: 'A1+1' },
    ])
  })
})
