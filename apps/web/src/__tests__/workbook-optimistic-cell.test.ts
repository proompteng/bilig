import { describe, expect, it } from 'vitest'
import { ValueTag, type CellSnapshot } from '@bilig/protocol'
import { createOptimisticCellSnapshot, evaluateOptimisticFormula } from '../workbook-optimistic-cell.js'

function cell(address: string, value: CellSnapshot['value'], version = 1): CellSnapshot {
  return {
    sheetName: 'Sheet1',
    address,
    value,
    flags: 0,
    version,
  }
}

describe('workbook optimistic cell snapshots', () => {
  it('evaluates formula commits from projected visible cells before worker readback arrives', () => {
    const cells = new Map<string, CellSnapshot>([['Sheet1:A1', cell('A1', { tag: ValueTag.String, value: 'hello', stringId: 0 })]])
    const current = cell('A2', { tag: ValueTag.Empty })

    const optimistic = createOptimisticCellSnapshot({
      sheetName: 'Sheet1',
      address: 'A2',
      current,
      parsed: { kind: 'formula', formula: 'A1="HELLO"' },
      evaluateFormula: (formula) =>
        evaluateOptimisticFormula({
          sheetName: 'Sheet1',
          address: 'A2',
          formula,
          getCell: (sheetName, address) => cells.get(`${sheetName}:${address}`) ?? cell(address, { tag: ValueTag.Empty }, 0),
        }),
    })

    expect(optimistic).toMatchObject({
      address: 'A2',
      formula: 'A1="HELLO"',
      value: {
        tag: ValueTag.Boolean,
        value: true,
      },
      version: 2,
    })
  })

  it('keeps formula source visible when optimistic evaluation cannot be resolved safely', () => {
    const current = cell('A2', { tag: ValueTag.Empty })

    const optimistic = createOptimisticCellSnapshot({
      sheetName: 'Sheet1',
      address: 'A2',
      current,
      parsed: { kind: 'formula', formula: 'SUM(1:1)' },
      evaluateFormula: () => null,
    })

    expect(optimistic).toMatchObject({
      formula: 'SUM(1:1)',
      value: {
        tag: ValueTag.String,
        value: '=SUM(1:1)',
      },
    })
  })
})
