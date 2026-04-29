import { compileFormulaAst, parseFormula } from '@bilig/formula'
import { ValueTag, type WorkbookSnapshot } from '@bilig/protocol'
import { describe, expect, it } from 'vitest'
import { restoreWorkbookFromRuntimeImage } from '../snapshot/runtime-image.js'
import { StringPool } from '../string-pool.js'
import { WorkbookStore } from '../workbook-store.js'

describe('restoreWorkbookFromRuntimeImage', () => {
  it('prefers prepared template initialization when runtime image instances carry template ids', () => {
    const workbook = new WorkbookStore('runtime-image-test')
    const preparedCalls: Array<{ col: number; row: number; sheetId: number; source: string; templateId?: number }> = []
    const plainCalls: Array<{ mutation: { col: number; formula: string; kind: string; row: number }; sheetId: number }> = []
    const compiled = compileFormulaAst('A1+1', parseFormula('A1+1'))
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: { name: 'runtime-image-test' },
      sheets: [
        {
          id: 1,
          name: 'Sheet1',
          order: 0,
          cells: [
            { address: 'A1', value: 1 },
            { address: 'B1', formula: 'A1+1' },
          ],
        },
      ],
    }

    restoreWorkbookFromRuntimeImage({
      snapshot,
      runtimeImage: {
        version: 1,
        templateBank: [
          {
            id: 7,
            templateKey: 'template:A1+1',
            baseSource: 'A1+1',
            baseRow: 0,
            baseCol: 1,
            compiled,
          },
        ],
        formulaInstances: [
          {
            cellIndex: 2,
            sheetName: 'Sheet1',
            row: 0,
            col: 1,
            source: 'A1+1',
            templateId: 7,
          },
        ],
        formulaValues: [
          {
            sheetName: 'Sheet1',
            row: 0,
            col: 1,
            value: { tag: ValueTag.Number, value: 2 },
          },
        ],
      },
      workbook,
      strings: new StringPool(),
      resetWorkbook: () => {},
      hydrateTemplateBank: () => {},
      resolveTemplateById: (templateId, source, row, col) =>
        templateId === 7
          ? {
              templateId,
              templateKey: 'template:A1+1',
              baseSource: source,
              compiled,
              translated: false,
              rowDelta: row,
              colDelta: col - 1,
            }
          : undefined,
      initializePreparedCellFormulasAt: (refs) => {
        preparedCalls.push(...refs)
      },
      initializeCellFormulasAt: (refs) => {
        plainCalls.push(...refs)
      },
    })

    const sheet = workbook.getSheet('Sheet1')
    expect(sheet).toBeDefined()
    const literalCellIndex = sheet!.grid.get(0, 0)
    expect(literalCellIndex).toBeGreaterThanOrEqual(0)
    expect(workbook.cellStore.tags[literalCellIndex]).toBe(ValueTag.Number)
    expect(workbook.cellStore.numbers[literalCellIndex]).toBe(1)
    expect(preparedCalls).toHaveLength(1)
    expect(preparedCalls[0]).toMatchObject({ sheetId: 1, row: 0, col: 1, source: 'A1+1', templateId: 7 })
    expect(plainCalls).toHaveLength(0)
  })

  it('prefers runtime-image coordinates over reparsing snapshot addresses during restore', () => {
    const workbook = new WorkbookStore('runtime-image-coords')
    const plainCalls: Array<{ mutation: { col: number; formula: string; kind: string; row: number }; sheetId: number }> = []
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: { name: 'runtime-image-coords' },
      sheets: [
        {
          id: 1,
          name: 'Sheet1',
          order: 0,
          cells: [
            { address: '__ignored_literal__', value: 1 },
            { address: '__ignored_formula__', formula: 'A1+1' },
          ],
        },
      ],
    }

    restoreWorkbookFromRuntimeImage({
      snapshot,
      runtimeImage: {
        version: 1,
        templateBank: [],
        formulaInstances: [],
        formulaValues: [],
        sheetCells: [
          {
            sheetName: 'Sheet1',
            coords: [
              { row: 0, col: 0 },
              { row: 0, col: 1 },
            ],
          },
        ],
      },
      workbook,
      strings: new StringPool(),
      resetWorkbook: () => {},
      hydrateTemplateBank: () => {},
      initializeCellFormulasAt: (refs) => {
        plainCalls.push(...refs)
      },
    })

    const sheet = workbook.getSheet('Sheet1')
    expect(sheet).toBeDefined()
    expect(sheet!.grid.get(0, 0)).toBeGreaterThanOrEqual(0)
    expect(sheet!.grid.get(0, 1)).toBeGreaterThanOrEqual(0)
    const formulaCellIndex = sheet!.grid.get(0, 1)
    expect(plainCalls).toEqual([
      {
        sheetId: 1,
        cellIndex: formulaCellIndex,
        mutation: {
          kind: 'setCellFormula',
          row: 0,
          col: 1,
          formula: 'A1+1',
        },
      },
    ])
  })

  it('falls back to address-matched runtime formula values when value order is not aligned', () => {
    const workbook = new WorkbookStore('runtime-image-misaligned-values')
    const hydratedCalls: Array<{ col: number; row: number; sheetId: number; source: string; value: unknown }> = []
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: { name: 'runtime-image-misaligned-values' },
      sheets: [
        {
          id: 1,
          name: 'Sheet1',
          order: 0,
          cells: [
            { address: 'A1', value: 1 },
            { address: 'B1', formula: 'A1+1' },
            { address: 'A2', value: 2 },
            { address: 'B2', formula: 'A2+1' },
          ],
        },
      ],
    }

    restoreWorkbookFromRuntimeImage({
      snapshot,
      runtimeImage: {
        version: 1,
        templateBank: [],
        formulaInstances: [
          {
            cellIndex: 2,
            sheetName: 'Sheet1',
            row: 0,
            col: 1,
            source: 'A1+1',
            templateId: 7,
          },
          {
            cellIndex: 4,
            sheetName: 'Sheet1',
            row: 1,
            col: 1,
            source: 'A2+1',
            templateId: 7,
          },
        ],
        formulaValues: [
          {
            sheetName: 'Sheet1',
            row: 1,
            col: 1,
            value: { tag: ValueTag.Number, value: 3 },
          },
          {
            sheetName: 'Sheet1',
            row: 0,
            col: 1,
            value: { tag: ValueTag.Number, value: 2 },
          },
        ],
      },
      workbook,
      strings: new StringPool(),
      resetWorkbook: () => {},
      hydrateTemplateBank: () => {},
      resolveTemplateById: (templateId, source, row, col) => ({
        templateId,
        templateKey: 'template:A1+1',
        baseSource: source,
        compiled: compileFormulaAst(source, parseFormula(source)),
        translated: false,
        rowDelta: row,
        colDelta: col,
      }),
      initializeHydratedPreparedCellFormulasAt: (refs) => {
        hydratedCalls.push(...refs)
      },
      initializeCellFormulasAt: () => {},
    })

    expect(hydratedCalls.map((call) => [call.row, call.col, call.value])).toEqual([
      [0, 1, { tag: ValueTag.Number, value: 2 }],
      [1, 1, { tag: ValueTag.Number, value: 3 }],
    ])
  })
})
