import { describe, expect, it } from 'vitest'
import { ValueTag } from '@bilig/protocol'
import {
  createWorkPaperFromDocument,
  exportWorkPaperDocument,
  isPersistedWorkPaperDocument,
  parseWorkPaperDocument,
  pickPersistableWorkPaperConfig,
  serializeWorkPaperDocument,
  WORK_PAPER_DOCUMENT_FORMAT,
  WorkPaper,
  WorkPaperPersistenceError,
} from '../index.js'

function restoreThroughSerializedDocument(workbook: WorkPaper): WorkPaper {
  return createWorkPaperFromDocument(parseWorkPaperDocument(serializeWorkPaperDocument(exportWorkPaperDocument(workbook))))
}

describe('WorkPaper persistence helpers', () => {
  it('roundtrips sheets, named expressions, config, and sheet order through the persisted document format', () => {
    const workbook = WorkPaper.buildEmpty({
      useColumnIndex: true,
      decimalSeparator: '.',
      calculationSettings: { iterate: true, iterateCount: 64, iterateDelta: '0.001', calcOnSave: true, calcCompleted: false },
      parseDateTime: () => undefined,
      functionPlugins: [],
    })
    workbook.addSheet('10')
    workbook.addSheet('2')

    const tenId = workbook.getSheetId('10')
    const twoId = workbook.getSheetId('2')
    if (tenId === undefined || twoId === undefined) {
      throw new Error('Expected persisted sheets to exist')
    }

    workbook.addNamedExpression('GlobalRate', '=5')
    workbook.addNamedExpression('LocalRate', '=7', twoId)
    workbook.setSheetContent(tenId, [[1, '=GlobalRate+A1']])
    workbook.setSheetContent(twoId, [[2, '=LocalRate+A1']])

    const document = exportWorkPaperDocument(workbook)
    const restored = createWorkPaperFromDocument(parseWorkPaperDocument(serializeWorkPaperDocument(document)))

    expect(document.format).toBe(WORK_PAPER_DOCUMENT_FORMAT)
    expect(isPersistedWorkPaperDocument(document)).toBe(true)
    expect(document.config).toMatchObject({
      calculationSettings: {
        iterate: true,
        iterateCount: 64,
        iterateDelta: '0.001',
        calcOnSave: true,
        calcCompleted: false,
      },
      useColumnIndex: true,
      decimalSeparator: '.',
    })
    expect(document.config).not.toHaveProperty('parseDateTime')
    expect(document.config).not.toHaveProperty('functionPlugins')
    expect(restored.getSheetNames()).toEqual(['10', '2'])
    expect(restored.getAllSheetsSerialized()).toEqual(workbook.getAllSheetsSerialized())
    expect(restored.getAllNamedExpressionsSerialized()).toEqual(workbook.getAllNamedExpressionsSerialized())
    expect(restored.getCellValue({ sheet: restored.getSheetId('10')!, row: 0, col: 1 })).toEqual({
      tag: ValueTag.Number,
      value: 6,
    })
    expect(restored.getCellValue({ sheet: restored.getSheetId('2')!, row: 0, col: 1 })).toEqual({
      tag: ValueTag.Number,
      value: 9,
    })
  })

  it('keeps restored cross-sheet formula aggregate dependencies live after edits', () => {
    const workbook = WorkPaper.buildFromSheets(
      {
        Plan: [
          ['Month', 'Bookings', 'Churn', 'Net MRR'],
          ['January', 12000, 800, '=B2-C2'],
          ['February', 15000, 900, '=B3-C3'],
          ['March', 18000, 1200, '=B4-C4'],
        ],
        Summary: [
          ['Metric', 'Value'],
          ['Quarter net MRR', '=SUM(Plan!D2:D4)'],
          ['Annualized run rate', '=B2*12'],
          ['Expansion-adjusted ARR', null],
        ],
      },
      { maxRows: 1000, maxColumns: 64, useColumnIndex: true },
    )
    workbook.addNamedExpression('ExpansionRatePercent', 8)
    workbook.setCellContents({ sheet: workbook.getSheetId('Summary')!, row: 3, col: 1 }, '=B3*(100+ExpansionRatePercent)/100')

    const restored = createWorkPaperFromDocument(parseWorkPaperDocument(serializeWorkPaperDocument(exportWorkPaperDocument(workbook))))
    restored.setCellContents({ sheet: restored.getSheetId('Plan')!, row: 3, col: 1 }, 21000)

    expect(restored.getCellValue({ sheet: restored.getSheetId('Plan')!, row: 3, col: 3 })).toEqual({
      tag: ValueTag.Number,
      value: 19800,
    })
    expect(restored.getCellValue({ sheet: restored.getSheetId('Summary')!, row: 1, col: 1 })).toEqual({
      tag: ValueTag.Number,
      value: 45100,
    })
    expect(restored.getCellValue({ sheet: restored.getSheetId('Summary')!, row: 2, col: 1 })).toEqual({
      tag: ValueTag.Number,
      value: 541200,
    })
    expect(restored.getCellValue({ sheet: restored.getSheetId('Summary')!, row: 3, col: 1 })).toEqual({
      tag: ValueTag.Number,
      value: 584496,
    })
  })

  it('keeps computed display readback after JSON document restore', () => {
    const workbook = WorkPaper.buildFromSheets({
      Inputs: [
        ['Metric', 'Value'],
        ['Win rate', 0.25],
      ],
      Summary: [
        ['Metric', 'Value'],
        ['Expected customers', '=Inputs!B2*20'],
      ],
    })
    const inputId = workbook.getSheetId('Inputs')!
    const summaryId = workbook.getSheetId('Summary')!
    const summaryValue = { sheet: summaryId, row: 1, col: 1 }

    workbook.setCellContents({ sheet: inputId, row: 1, col: 1 }, 0.4)

    expect(workbook.getCellValue(summaryValue)).toEqual({ tag: ValueTag.Number, value: 8 })
    expect(workbook.getCellDisplayValue(summaryValue)).toBe('8')

    const restored = restoreThroughSerializedDocument(workbook)
    const restoredSummaryValue = { sheet: restored.getSheetId('Summary')!, row: 1, col: 1 }

    expect(restored.getCellSerialized(restoredSummaryValue)).toBe('=Inputs!B2*20')
    expect(restored.getCellValue(restoredSummaryValue)).toEqual(workbook.getCellValue(summaryValue))
    expect(restored.getCellDisplayValue(restoredSummaryValue)).toBe(workbook.getCellDisplayValue(summaryValue))
  })

  it('keeps changed named expressions driving formulas after document restore', () => {
    const workbook = WorkPaper.buildFromSheets(
      {
        Inputs: [
          ['Metric', 'Value'],
          ['Base customers', 100],
        ],
        Summary: [
          ['Metric', 'Value'],
          ['Projected customers', '=Inputs!B2*GrowthRate'],
        ],
      },
      { useColumnIndex: true },
      [{ name: 'GrowthRate', expression: '=1.1' }],
    )
    const summaryValue = { sheet: workbook.getSheetId('Summary')!, row: 1, col: 1 }

    workbook.changeNamedExpression('GrowthRate', '=1.25')

    expect(workbook.getNamedExpressionFormula('GrowthRate')).toBe('=1.25')
    expect(workbook.getNamedExpressionValue('GrowthRate')).toEqual({ tag: ValueTag.Number, value: 1.25 })
    expect(workbook.getCellValue(summaryValue)).toEqual({ tag: ValueTag.Number, value: 125 })

    const restored = restoreThroughSerializedDocument(workbook)
    const restoredSummaryValue = { sheet: restored.getSheetId('Summary')!, row: 1, col: 1 }

    expect(restored.getNamedExpressionFormula('GrowthRate')).toBe('=1.25')
    expect(restored.getNamedExpressionValue('GrowthRate')).toEqual({ tag: ValueTag.Number, value: 1.25 })
    expect(restored.getCellValue(restoredSummaryValue)).toEqual({ tag: ValueTag.Number, value: 125 })
  })

  it('preserves imported sheet names that end with spaces during document restore', () => {
    const workbook = WorkPaper.buildFromSheets(
      {
        'DFFH ': [[1, '=LocalRate+A1']],
      },
      { useColumnIndex: true },
      [{ name: 'LocalRate', expression: '=7', scope: 1 }],
    )

    const document = exportWorkPaperDocument(workbook)
    const restored = createWorkPaperFromDocument(parseWorkPaperDocument(serializeWorkPaperDocument(document)))
    const restoredSheetId = restored.getSheetId('DFFH ')

    expect(restored.getSheetNames()).toEqual(['DFFH '])
    expect(restoredSheetId).toBe(1)
    expect(restored.getSheetId('DFFH')).toBeUndefined()
    expect(restored.getAllNamedExpressionsSerialized()).toEqual(workbook.getAllNamedExpressionsSerialized())
    expect(restored.getCellValue({ sheet: restoredSheetId!, row: 0, col: 1 })).toEqual({
      tag: ValueTag.Number,
      value: 8,
    })
  })

  it('exports config only from the documented JSON-safe subset', () => {
    const config = pickPersistableWorkPaperConfig({
      calculationSettings: { iterate: true, iterateCount: 50, iterateDelta: '0.001', calcOnSave: true, calcCompleted: false },
      useColumnIndex: true,
      context: { requestId: 'ctx-1', featureFlags: ['alpha'] },
      stringifyDateTime: () => undefined,
      functionPlugins: [],
    })

    expect(config).toEqual({
      calculationSettings: { iterate: true, iterateCount: 50, iterateDelta: '0.001', calcOnSave: true, calcCompleted: false },
      useColumnIndex: true,
      context: { requestId: 'ctx-1', featureFlags: ['alpha'] },
    })
  })

  it('supports documents without persisted config', () => {
    const workbook = WorkPaper.buildFromSheets({ Sheet1: [[1, '=A1+1']] })
    const document = exportWorkPaperDocument(workbook, { includeConfig: false })
    const restored = createWorkPaperFromDocument(document)

    expect(document.config).toBeUndefined()
    expect(restored.getAllSheetsSerialized()).toEqual(workbook.getAllSheetsSerialized())
  })

  it('rejects invalid persisted WorkPaper documents', () => {
    expect(isPersistedWorkPaperDocument({})).toBe(false)
    expect(() => parseWorkPaperDocument('{}')).toThrow(WorkPaperPersistenceError)
    expect(() => parseWorkPaperDocument('{')).toThrow(WorkPaperPersistenceError)
    expect(
      isPersistedWorkPaperDocument({
        format: WORK_PAPER_DOCUMENT_FORMAT,
        sheets: [{ name: 'Sheet1', content: [[1]] }],
        namedExpressions: [{ name: 'Rate', expression: '=1', scopeSheetName: 1 }],
      }),
    ).toBe(false)
  })

  it('rejects non-json and structurally unsafe persisted document values', () => {
    const sparseRow: unknown[] = [null]
    Reflect.deleteProperty(sparseRow, '0')
    const sparseSheets = [{ name: 'Sheet1', content: [sparseRow] }]

    expect(
      isPersistedWorkPaperDocument({
        format: WORK_PAPER_DOCUMENT_FORMAT,
        sheets: [{ name: 'Sheet1', content: [[Number.NaN]] }],
        namedExpressions: [],
      }),
    ).toBe(false)
    expect(
      isPersistedWorkPaperDocument({
        format: WORK_PAPER_DOCUMENT_FORMAT,
        sheets: sparseSheets,
        namedExpressions: [],
      }),
    ).toBe(false)
    expect(
      isPersistedWorkPaperDocument({
        format: WORK_PAPER_DOCUMENT_FORMAT,
        sheets: [
          { name: 'Sheet1', content: [[1]] },
          { name: 'Sheet1', content: [[2]] },
        ],
        namedExpressions: [],
      }),
    ).toBe(false)
    expect(
      isPersistedWorkPaperDocument({
        format: WORK_PAPER_DOCUMENT_FORMAT,
        sheets: [{ name: 'Sheet1', content: [[1]] }],
        namedExpressions: [{ name: 'Scoped', expression: '=1', scopeSheetName: 'Missing' }],
      }),
    ).toBe(false)
    expect(
      isPersistedWorkPaperDocument({
        format: WORK_PAPER_DOCUMENT_FORMAT,
        sheets: [{ name: 'Sheet1', content: [[1]] }],
        namedExpressions: [],
        config: { maxRows: 0 },
      }),
    ).toBe(false)
  })
})
