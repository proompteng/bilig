import { describe, expect, it } from 'vitest'
import fc from 'fast-check'
import { runProperty } from '@bilig/test-fuzz'
import {
  createWorkPaperFromDocument,
  exportWorkPaperDocument,
  parseWorkPaperDocument,
  serializeWorkPaperDocument,
  WorkPaper,
} from '../index.js'

type WorkPaperCell = number | boolean | null | string

describe('WorkPaper persistence fuzz', () => {
  it('should roundtrip sheets, named expressions, and persistable config through the document format', async () => {
    await runProperty({
      suite: 'headless/work-paper/persistence-roundtrip',
      arbitrary: fc.record({
        useColumnIndex: fc.boolean(),
        decimalSeparator: fc.constantFrom('.', ','),
        sheets: fc.uniqueArray(
          fc.record({
            name: fc.constantFrom('SheetA', 'SheetB', 'SheetC'),
            content: fc.array(fc.array(workPaperCellArbitrary, { minLength: 1, maxLength: 4 }), { minLength: 1, maxLength: 4 }),
          }),
          {
            minLength: 1,
            maxLength: 2,
            selector: (sheet) => sheet.name,
          },
        ),
      }),
      predicate: async ({ useColumnIndex, decimalSeparator, sheets }) => {
        const workbook = WorkPaper.buildEmpty({
          useColumnIndex,
          decimalSeparator,
          parseDateTime: () => undefined,
          functionPlugins: [],
        })

        sheets.forEach((sheet) => {
          workbook.addSheet(sheet.name)
        })
        workbook.addNamedExpression('GlobalRate', '=5')
        const scopedSheetId = workbook.getSheetId(sheets[0].name)
        if (scopedSheetId !== undefined) {
          workbook.addNamedExpression('LocalRate', '=7', scopedSheetId)
        }
        sheets.forEach((sheet) => {
          const sheetId = workbook.getSheetId(sheet.name)
          if (sheetId === undefined) {
            throw new Error(`Missing sheet id for ${sheet.name}`)
          }
          workbook.setSheetContent(sheetId, sheet.content)
        })

        const document = exportWorkPaperDocument(workbook)
        const restored = createWorkPaperFromDocument(parseWorkPaperDocument(serializeWorkPaperDocument(document)))

        expect(exportWorkPaperDocument(restored)).toEqual(document)
      },
    })
  })
})

// Helpers

const workPaperCellArbitrary = fc.oneof<WorkPaperCell>(
  fc.integer({ min: -200, max: 200 }),
  fc.boolean(),
  fc.constant(null),
  fc.constantFrom('north', 'south', '=1+2', '=SUM(1,2)', '=A1'),
)
