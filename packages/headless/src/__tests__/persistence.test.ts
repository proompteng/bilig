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
} from '../index.js'

describe('WorkPaper persistence helpers', () => {
  it('roundtrips sheets, named expressions, config, and sheet order through the persisted document format', () => {
    const workbook = WorkPaper.buildEmpty({
      useColumnIndex: true,
      decimalSeparator: '.',
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

  it('exports config only from the documented JSON-safe subset', () => {
    const config = pickPersistableWorkPaperConfig({
      useColumnIndex: true,
      context: { requestId: 'ctx-1', featureFlags: ['alpha'] },
      stringifyDateTime: () => undefined,
      functionPlugins: [],
    })

    expect(config).toEqual({
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
    expect(() => parseWorkPaperDocument('{}')).toThrow('Invalid persisted WorkPaper document')
    expect(
      isPersistedWorkPaperDocument({
        format: WORK_PAPER_DOCUMENT_FORMAT,
        sheets: [{ name: 'Sheet1', content: [[1]] }],
        namedExpressions: [{ name: 'Rate', expression: '=1', scopeSheetName: 1 }],
      }),
    ).toBe(false)
  })
})
