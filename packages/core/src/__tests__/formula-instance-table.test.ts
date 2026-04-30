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

  it('lists sparse upserts and deletes in sheet order', () => {
    const table = createFormulaInstanceTable()

    table.upsert({ cellIndex: 8, sheetName: 'Sheet2', row: 0, col: 0, source: 'A1+1' })
    table.upsert({ cellIndex: 2, sheetName: 'Sheet1', row: 1, col: 0, source: 'A2+1' })
    table.upsert({ cellIndex: 5, sheetName: 'Sheet1', row: 0, col: 0, source: 'A1+1' })

    expect(table.delete(2)).toBe(true)
    expect(table.delete(2)).toBe(false)
    expect(table.get(2)).toBeUndefined()
    expect(table.list()).toEqual([
      { cellIndex: 5, sheetName: 'Sheet1', row: 0, col: 0, source: 'A1+1' },
      { cellIndex: 8, sheetName: 'Sheet2', row: 0, col: 0, source: 'A1+1' },
    ])
  })
})
