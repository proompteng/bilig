import { describe, expect, it, vi } from 'vitest'
import { WorkPaper } from './index.js'
import { createWorkPaperToolHandlers } from './ai-sdk.js'

describe('@bilig/workpaper AI SDK handlers', () => {
  it('edits a WorkPaper input and returns computed restore proof', () => {
    const workpaper = WorkPaper.buildFromSheets({
      Inputs: [
        ['Metric', 'Value'],
        ['Qualified opportunities', 20],
        ['Win rate', 0.25],
        ['Average ARR', 12000],
      ],
      Summary: [
        ['Metric', 'Value'],
        ['Expected customers', '=Inputs!B2*Inputs!B3'],
        ['Expected ARR', '=B2*Inputs!B4'],
      ],
    })
    const handlers = createWorkPaperToolHandlers({
      workpaper,
      defaultReadRange: 'Summary!A1:B3',
      proofRange: 'Summary!A1:B3',
      writableSheets: ['Inputs'],
    })

    const before = handlers.readWorkPaperSummary()
    const write = handlers.setWorkPaperInputCell({
      sheetName: 'Inputs',
      address: 'B3',
      value: 0.4,
    })

    expect(readNumber(before.values, 2, 1)).toBe(60000)
    expect(write.editedCell).toBe('Inputs!B3')
    expect(readNumber(write.after.values, 2, 1)).toBe(96000)
    expect(readNumber(write.restored.values, 2, 1)).toBe(96000)
    expect(write.checks).toMatchObject({
      previousValue: 0.25,
      newValue: 0.4,
      formulasPersisted: true,
      restoredMatchesAfter: true,
      proofRangeChanged: true,
    })
    expect(write.checks.serializedBytes).toBeGreaterThan(100)
  })

  it('blocks writes outside the configured input sheet boundary', () => {
    const workpaper = WorkPaper.buildFromSheets({
      Inputs: [['Metric', 'Value']],
      Summary: [['Metric', '=1+1']],
    })
    const handlers = createWorkPaperToolHandlers({
      workpaper,
      defaultReadRange: 'Summary!A1:B1',
      proofRange: 'Summary!A1:B1',
      writableSheets: ['Inputs'],
    })

    expect(() =>
      handlers.setWorkPaperInputCell({
        sheetName: 'Summary',
        address: 'B1',
        value: 3,
      }),
    ).toThrow('Sheet "Summary" is not writable')
  })

  it('disposes the temporary restore workbook after a write proof', () => {
    const workpaper = WorkPaper.buildFromSheets({
      Inputs: [
        ['Metric', 'Value'],
        ['Units', 20],
      ],
      Summary: [
        ['Metric', 'Value'],
        ['Units', '=Inputs!B2'],
      ],
    })
    const handlers = createWorkPaperToolHandlers({
      workpaper,
      proofRange: 'Summary!A1:B2',
      writableSheets: ['Inputs'],
    })
    const dispose = vi.spyOn(WorkPaper.prototype, 'dispose')

    try {
      handlers.setWorkPaperInputCell({
        sheetName: 'Inputs',
        address: 'B2',
        value: 24,
      })

      expect(dispose).toHaveBeenCalledTimes(1)
    } finally {
      dispose.mockRestore()
      workpaper.dispose()
    }
  })
})

function readNumber(values: unknown[][], row: number, col: number): number {
  const cell = values[row]?.[col]
  if (!cell || typeof cell !== 'object' || !('value' in cell) || typeof cell.value !== 'number') {
    throw new Error(`Expected numeric cell at row ${row}, col ${col}; received ${JSON.stringify(cell)}`)
  }
  return Math.round(cell.value * 100) / 100
}
