import { describe, expect, it } from 'vitest'
import { ValueTag, type WorkbookSnapshot } from '@bilig/protocol'
import { SpreadsheetEngine } from '../index.js'

describe('engine INDIRECT defined-name references', () => {
  it('dereferences dynamically constructed cell names after scalar implicit intersection', async () => {
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: {
        name: 'indirect-defined-name-reference',
        metadata: {
          definedNames: [
            { name: 'i', value: { kind: 'range-ref', sheetName: 'CubicSpline3', startAddress: 'E40', endAddress: 'E40' } },
            { name: 't', value: { kind: 'range-ref', sheetName: 'CubicSpline3', startAddress: 'F40', endAddress: 'F40' } },
            { name: 'a1_', value: { kind: 'cell-ref', sheetName: 'CubicSpline3', address: 'T25' } },
            { name: 'b1_', value: { kind: 'cell-ref', sheetName: 'CubicSpline3', address: 'T26' } },
            { name: '_c1', value: { kind: 'cell-ref', sheetName: 'CubicSpline3', address: 'T27' } },
            { name: 'd1_', value: { kind: 'cell-ref', sheetName: 'CubicSpline3', address: 'T28' } },
          ],
        },
      },
      sheets: [
        {
          id: 1,
          name: 'CubicSpline3',
          order: 0,
          cells: [
            { address: 'E40', value: 1 },
            { address: 'F40', value: 2 },
            { address: 'T25', value: 1 },
            { address: 'T26', value: 2 },
            { address: 'T27', value: 3 },
            { address: 'T28', value: 4 },
            {
              address: 'G40',
              formula: 'INDIRECT("a" & i & "_") * t^3 + INDIRECT("b" & i & "_") * t^2 + INDIRECT("_c" & i) * t^1 + INDIRECT("d" & i & "_")',
            },
          ],
        },
      ],
    }

    const engine = new SpreadsheetEngine({ workbookName: 'indirect-defined-name-reference' })
    await engine.ready()
    engine.importSnapshot(snapshot)

    expect(engine.getCellValue('CubicSpline3', 'G40')).toEqual({ tag: ValueTag.Number, value: 26 })

    engine.setCellValue('CubicSpline3', 'T25', 10)
    expect(engine.getCellValue('CubicSpline3', 'G40')).toEqual({ tag: ValueTag.Number, value: 98 })
  })

  it('returns range-valued defined names from INDIRECT to aggregate callers', async () => {
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: {
        name: 'indirect-defined-range-reference',
        metadata: {
          definedNames: [{ name: 'SalesRange', value: { kind: 'range-ref', sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A3' } }],
        },
      },
      sheets: [
        {
          id: 1,
          name: 'Sheet1',
          order: 0,
          cells: [
            { address: 'A1', value: 2 },
            { address: 'A2', value: 3 },
            { address: 'A3', value: 4 },
            { address: 'B1', formula: 'SUM(INDIRECT("SalesRange"))' },
          ],
        },
      ],
    }

    const engine = new SpreadsheetEngine({ workbookName: 'indirect-defined-range-reference' })
    await engine.ready()
    engine.importSnapshot(snapshot)

    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 9 })

    engine.setCellValue('Sheet1', 'A2', 5)
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 11 })
  })

  it('recalculates volatile INDIRECT names after initialization materializes target spills', async () => {
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: {
        name: 'indirect-defined-name-spill-initialization',
        metadata: {
          definedNames: [
            { name: 'i', value: { kind: 'cell-ref', sheetName: 'Sheet1', address: 'E1' } },
            { name: 'a1_', value: { kind: 'cell-ref', sheetName: 'Sheet1', address: 'T1' } },
            { name: 'b1_', value: { kind: 'cell-ref', sheetName: 'Sheet1', address: 'T2' } },
            { name: '_c1', value: { kind: 'cell-ref', sheetName: 'Sheet1', address: 'T3' } },
            { name: 'd1_', value: { kind: 'cell-ref', sheetName: 'Sheet1', address: 'T4' } },
          ],
        },
      },
      sheets: [
        {
          id: 1,
          name: 'Sheet1',
          order: 0,
          cells: [
            { address: 'E1', value: 1 },
            {
              address: 'G1',
              formula: 'INDIRECT("a"&i&"_")+INDIRECT("b"&i&"_")+INDIRECT("_c"&i)+INDIRECT("d"&i&"_")',
            },
            { address: 'T1', formula: 'SEQUENCE(4,1,1,1)' },
          ],
        },
      ],
    }

    const engine = new SpreadsheetEngine({ workbookName: 'indirect-defined-name-spill-initialization' })
    await engine.ready()
    engine.importSnapshot(snapshot)

    expect(engine.getCellValue('Sheet1', 'G1')).toEqual({ tag: ValueTag.Number, value: 10 })
  })
})
