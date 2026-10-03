import { describe, expect, it } from 'vitest'
import { ValueTag } from '@bilig/protocol'
import { orderWorkPaperCellChanges } from '../change-order.js'
import type { WorkPaperChange } from '../work-paper-types.js'

function cellChange(sheet: number, row: number, col: number): WorkPaperChange {
  return {
    kind: 'cell',
    address: { sheet, row, col },
    sheetName: `Sheet${sheet}`,
    a1: `${row}:${col}`,
    newValue: { tag: ValueTag.Number, value: row * 10 + col },
  }
}

describe('orderWorkPaperCellChanges', () => {
  it('orders single-sheet changes by row and column', () => {
    const changes = [cellChange(1, 2, 0), cellChange(1, 0, 1), cellChange(1, 0, 0)]

    const ordered = orderWorkPaperCellChanges(changes, [{ id: 1, order: 0 }])

    expect(ordered.map((change) => (change.kind === 'cell' ? `${change.address.row}:${change.address.col}` : 'other'))).toEqual([
      '0:0',
      '0:1',
      '2:0',
    ])
  })

  it('reverses descending single-sheet batches without falling back to a full sort', () => {
    const changes = [cellChange(1, 3, 0), cellChange(1, 2, 0), cellChange(1, 1, 0)]

    const ordered = orderWorkPaperCellChanges(changes, [{ id: 1, order: 0 }])

    expect(ordered).not.toBe(changes)
    expect(ordered.map((change) => (change.kind === 'cell' ? `${change.address.row}:${change.address.col}` : 'other'))).toEqual([
      '1:0',
      '2:0',
      '3:0',
    ])
  })
})
