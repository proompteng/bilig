import { describe, expect, it } from 'vitest'
import { ValueTag } from '@bilig/protocol'
import { WorkPaper, type WorkPaperCellAddress, type WorkPaperChange } from '../index.js'

function cell(sheet: number, row: number, col: number): WorkPaperCellAddress {
  return { sheet, row, col }
}

function trackComputeCellChangesFromTrackedEvents(workbook: WorkPaper): { readonly count: number; restore: () => void } {
  const original = Reflect.get(workbook, 'computeCellChangesFromTrackedEvents')
  if (typeof original !== 'function') {
    throw new Error('Expected WorkPaper to expose computeCellChangesFromTrackedEvents in tests')
  }
  let count = 0
  Reflect.set(workbook, 'computeCellChangesFromTrackedEvents', (...args: unknown[]) => {
    count += 1
    return Reflect.apply(original, workbook, args)
  })
  return {
    get count() {
      return count
    },
    restore: () => {
      Reflect.set(workbook, 'computeCellChangesFromTrackedEvents', original)
    },
  }
}

describe('WorkPaper literal fanout fast path', () => {
  it('emits eager listener payloads without the generic tracked reducer', () => {
    const workbook = WorkPaper.buildFromSheets({
      Bench: [['old', '=A1', '=A1', '=A1', '=A1']],
    })
    const sheetId = workbook.getSheetId('Bench')!
    const reducerTracker = trackComputeCellChangesFromTrackedEvents(workbook)
    const events: WorkPaperChange[][] = []
    workbook.on('valuesUpdated', (changes) => {
      events.push(changes)
    })

    try {
      const changes = workbook.setCellContents(cell(sheetId, 0, 0), 'new')

      expect(events).toEqual([changes])
      expect(changes.map((change) => (change.kind === 'cell' ? `${change.sheetName}!${change.a1}` : ''))).toEqual([
        'Bench!A1',
        'Bench!B1',
        'Bench!C1',
        'Bench!D1',
        'Bench!E1',
      ])
      expect(changes.every((change) => change.kind === 'cell' && change.newValue.tag === ValueTag.String)).toBe(true)
      expect(workbook.getCellValue(cell(sheetId, 0, 4))).toEqual({ tag: ValueTag.String, value: 'new', stringId: 2 })
      expect(reducerTracker.count).toBe(0)
    } finally {
      reducerTracker.restore()
    }
  })
})
