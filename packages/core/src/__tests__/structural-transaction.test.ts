import { describe, expect, it } from 'vitest'
import { buildStructuralTransaction, structuralScopeForTransform } from '../engine/structural-transaction.js'

describe('StructuralTransaction', () => {
  it('builds a bounded move scope that covers both source and target windows', () => {
    expect(
      structuralScopeForTransform({
        axis: 'row',
        kind: 'move',
        start: 10,
        count: 3,
        target: 4,
      }),
    ).toEqual({ start: 4, end: 13 })
  })

  it('records deleted cells separately from remapped survivors', () => {
    const transaction = buildStructuralTransaction({
      sheetName: 'Sheet1',
      sheetId: 1,
      transform: {
        axis: 'column',
        kind: 'delete',
        start: 2,
        count: 1,
      },
      remappedCells: [
        { cellIndex: 10, fromRow: 0, fromCol: 4, toRow: 0, toCol: 3 },
        { cellIndex: 11, fromRow: 5, fromCol: 2, toRow: undefined, toCol: undefined },
      ],
    })

    expect(transaction.removedCellIndices).toEqual([11])
    expect(transaction.remappedCells).toHaveLength(2)
  })
})
