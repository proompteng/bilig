import { describe, expect, it } from 'vitest'
import * as fc from 'fast-check'
import { runProperty } from '@bilig/test-fuzz'
import {
  createWorkPaperFromDocument,
  exportWorkPaperDocument,
  parseWorkPaperDocument,
  serializeWorkPaperDocument,
  WorkPaper,
} from '../index.js'

describe('WorkPaper action fuzz', () => {
  it('should preserve generated cell edits through save/load and undo/redo', async () => {
    await runProperty({
      suite: 'headless/work-paper/action-sequence-save-load',
      arbitrary: fc.array(workPaperActionArbitrary, { minLength: 1, maxLength: 16 }),
      predicate: async (actions) => {
        const workbook = WorkPaper.buildEmpty({
          parseDateTime: () => undefined,
          functionPlugins: [],
        })
        workbook.addSheet('Sheet1')
        const sheetId = workbook.getSheetId('Sheet1')
        if (sheetId === undefined) {
          throw new Error('Expected generated WorkPaper sheet to exist')
        }

        for (const action of actions) {
          workbook.setCellContents({ sheet: sheetId, row: action.row, col: action.col }, action.value)
        }

        const document = exportWorkPaperDocument(workbook, { includeConfig: true })
        const restored = createWorkPaperFromDocument(parseWorkPaperDocument(serializeWorkPaperDocument(document)))
        const restoredDocument = exportWorkPaperDocument(restored, { includeConfig: true })
        const secondRestored = createWorkPaperFromDocument(parseWorkPaperDocument(serializeWorkPaperDocument(restoredDocument)))
        expect(exportWorkPaperDocument(secondRestored, { includeConfig: true })).toEqual(restoredDocument)

        const beforeUndo = exportWorkPaperDocument(workbook, { includeConfig: true })
        if (workbook.undo().length > 0) {
          expect(exportWorkPaperDocument(workbook, { includeConfig: true })).not.toEqual(beforeUndo)
          workbook.redo()
          expect(exportWorkPaperDocument(workbook, { includeConfig: true })).toEqual(beforeUndo)
        }
      },
      parameters: { numRuns: 80 },
    })
  })
})

// Helpers

const workPaperActionArbitrary = fc.record({
  row: fc.integer({ min: 0, max: 8 }),
  col: fc.integer({ min: 0, max: 6 }),
  value: fc.oneof(
    fc.constant(null),
    fc.boolean(),
    fc.integer({ min: -1_000, max: 1_000 }),
    fc.string({ maxLength: 24 }),
    fc.constantFrom('=1+1', '=SUM(A1:A3)', '=IF(A1>0,1,0)'),
  ),
})
