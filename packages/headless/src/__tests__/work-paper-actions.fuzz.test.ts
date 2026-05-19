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
  it('does not record undo history for repeated identical cell writes', () => {
    const workbook = WorkPaper.buildFromSheets(
      {
        Sheet1: [],
      },
      {
        parseDateTime: () => undefined,
        functionPlugins: [],
      },
    )
    const sheetId = workbook.getSheetId('Sheet1')
    if (sheetId === undefined) {
      throw new Error('Expected generated WorkPaper sheet to exist')
    }

    const address = { sheet: sheetId, row: 2, col: 0 }
    expect(workbook.setCellContents(address, '=SUM(A1:A3)')).toHaveLength(1)
    const afterFirstWrite = exportWorkPaperDocument(workbook, { includeConfig: true })

    expect(workbook.setCellContents(address, '=SUM(A1:A3)')).toEqual([])
    expect(exportWorkPaperDocument(workbook, { includeConfig: true })).toEqual(afterFirstWrite)

    expect(workbook.undo()).toHaveLength(1)
    expect(exportWorkPaperDocument(workbook, { includeConfig: true })).not.toEqual(afterFirstWrite)
    expect(workbook.isThereSomethingToUndo()).toBe(false)
  })

  it('persists explicit zero writes after clearing a cell', () => {
    const workbook = WorkPaper.buildFromSheets(
      {
        Sheet1: [],
      },
      {
        parseDateTime: () => undefined,
        functionPlugins: [],
      },
    )
    const sheetId = workbook.getSheetId('Sheet1')
    if (sheetId === undefined) {
      throw new Error('Expected generated WorkPaper sheet to exist')
    }

    const address = { sheet: sheetId, row: 7, col: 5 }
    expect(workbook.setCellContents(address, '=1+1')).toHaveLength(1)
    expect(workbook.setCellContents(address, null)).toHaveLength(1)
    expect(workbook.setCellContents(address, 0)).toHaveLength(1)
    expect(workbook.getCellSerialized(address)).toBe(0)

    const beforeUndo = exportWorkPaperDocument(workbook, { includeConfig: true })
    expect(beforeUndo.sheets[0]?.content[7]?.[5]).toBe(0)
    const restored = createWorkPaperFromDocument(parseWorkPaperDocument(serializeWorkPaperDocument(beforeUndo)))
    const restoredSheetId = restored.getSheetId('Sheet1')
    if (restoredSheetId === undefined) {
      throw new Error('Expected restored WorkPaper sheet to exist')
    }
    expect(restored.getCellSerialized({ sheet: restoredSheetId, row: 7, col: 5 })).toBe(0)

    expect(workbook.undo()).toHaveLength(1)
    expect(workbook.getCellSerialized(address)).toBeNull()
    expect(exportWorkPaperDocument(workbook, { includeConfig: true })).not.toEqual(beforeUndo)

    expect(workbook.redo()).toHaveLength(1)
    expect(workbook.getCellSerialized(address)).toBe(0)
    expect(exportWorkPaperDocument(workbook, { includeConfig: true })).toEqual(beforeUndo)
  })

  it('redoes clears after a pruned cell is restored by undo', () => {
    const workbook = WorkPaper.buildFromSheets(
      {
        Sheet1: [],
      },
      {
        parseDateTime: () => undefined,
        functionPlugins: [],
      },
    )
    const sheetId = workbook.getSheetId('Sheet1')
    if (sheetId === undefined) {
      throw new Error('Expected generated WorkPaper sheet to exist')
    }

    const address = { sheet: sheetId, row: 4, col: 6 }
    expect(workbook.setCellContents(address, '=1+1')).toHaveLength(1)
    expect(workbook.setCellContents(address, null)).toHaveLength(1)
    const clearedDocument = exportWorkPaperDocument(workbook, { includeConfig: true })
    expect(clearedDocument.sheets[0]?.content).toEqual([])

    expect(workbook.undo()).toHaveLength(1)
    expect(workbook.getCellSerialized(address)).toBe('=1+1')

    expect(workbook.redo()).toHaveLength(1)
    expect(workbook.getCellSerialized(address)).toBeNull()
    expect(exportWorkPaperDocument(workbook, { includeConfig: true })).toEqual(clearedDocument)
  })

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
