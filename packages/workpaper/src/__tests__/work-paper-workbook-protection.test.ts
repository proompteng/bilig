import { ValueTag, type WorkbookSnapshot } from '@bilig/protocol'
import { describe, expect, it } from 'vitest'

import { WorkPaper, type WorkPaperCellAddress } from '../index.js'

describe('WorkPaper workbook structure protection', () => {
  it('blocks sheet topology mutations while allowing cell edits', () => {
    const workbook = WorkPaper.buildFromSnapshot(protectedWorkbookSnapshot())
    try {
      const dataSheet = workbook.getSheetId('Data')!
      const reportSheet = workbook.getSheetId('Report')!

      expect(workbook.isItPossibleToAddSheet('Added')).toBe(false)
      expect(workbook.isItPossibleToRemoveSheet(dataSheet)).toBe(false)
      expect(workbook.isItPossibleToRenameSheet(dataSheet, 'Source')).toBe(false)
      expect(workbook.isItPossibleToClearSheet(dataSheet)).toBe(true)

      expect(() => workbook.addSheet('Added')).toThrow(/Workbook structure is protected/)
      expect(() => workbook.moveSheet(reportSheet, 0)).toThrow(/Workbook structure is protected/)
      expect(() => workbook.removeSheet(dataSheet)).toThrow(/Workbook structure is protected/)
      expect(() => workbook.renameSheet(dataSheet, 'Source')).toThrow(/Workbook structure is protected/)
      expect(workbook.getSheetNames()).toEqual(['Data', 'Report'])

      workbook.setCellContents(cell(dataSheet, 0, 0), 42)

      expect(workbook.getCellValue(cell(dataSheet, 0, 0))).toEqual({ tag: ValueTag.Number, value: 42 })
      expect(workbook.exportSnapshot().workbook.metadata?.workbookProtection).toEqual(workbookStructureProtection)
    } finally {
      workbook.dispose()
    }
  })
})

const workbookStructureProtection = {
  lockStructure: true,
  xmlAttributes: [{ name: 'lockStructure', value: '1' }],
} as const

function protectedWorkbookSnapshot(): WorkbookSnapshot {
  const workbook = WorkPaper.buildFromSheets({
    Data: [[1]],
    Report: [[2]],
  })
  try {
    const snapshot = workbook.exportSnapshot()
    return {
      ...snapshot,
      workbook: {
        ...snapshot.workbook,
        metadata: {
          ...snapshot.workbook.metadata,
          workbookProtection: workbookStructureProtection,
        },
      },
    }
  } finally {
    workbook.dispose()
  }
}

function cell(sheet: number, row: number, col: number): WorkPaperCellAddress {
  return { sheet, row, col }
}
