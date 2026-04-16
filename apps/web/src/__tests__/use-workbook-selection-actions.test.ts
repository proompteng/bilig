import { describe, expect, it } from 'vitest'
import { buildPasteCommitOps, createSheetScopedRangePair } from '../use-workbook-selection-actions.js'

describe('use workbook selection action helpers', () => {
  it('builds paste commit ops for formulas, clears, booleans, numbers, and strings', () => {
    expect(
      buildPasteCommitOps('Sheet1', 'B2', [
        ['=SUM(A1:A2)', ''],
        ['TRUE', '42', 'text'],
      ]),
    ).toEqual([
      { kind: 'upsertCell', sheetName: 'Sheet1', addr: 'B2', formula: 'SUM(A1:A2)' },
      { kind: 'deleteCell', sheetName: 'Sheet1', addr: 'C2' },
      { kind: 'upsertCell', sheetName: 'Sheet1', addr: 'B3', value: true },
      { kind: 'upsertCell', sheetName: 'Sheet1', addr: 'C3', value: 42 },
      { kind: 'upsertCell', sheetName: 'Sheet1', addr: 'D3', value: 'text' },
    ])
  })

  it('creates source and target ranges scoped to one sheet', () => {
    expect(createSheetScopedRangePair('Sheet1', 'A1', 'B2', 'C3', 'D4')).toEqual({
      source: {
        sheetName: 'Sheet1',
        startAddress: 'A1',
        endAddress: 'B2',
      },
      target: {
        sheetName: 'Sheet1',
        startAddress: 'C3',
        endAddress: 'D4',
      },
    })
  })
})
