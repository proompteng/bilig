import { describe, expect, it, vi } from 'vitest'
import type { WorkbookSnapshot } from '@bilig/protocol'
import { FormulaMode, ValueTag } from '@bilig/protocol'
import { SpreadsheetEngine } from '../index.js'

type RuntimeFormulaWithDirectCriteria = {
  directCriteria: {
    aggregateKind: string
    aggregateRange:
      | {
          sheetName: string
          rowStart: number
          rowEnd: number
          col: number
          length: number
        }
      | undefined
    criteriaPairs: Array<{
      range: { sheetName: string; rowStart: number; rowEnd: number; col: number; length: number }
      criterion: { kind: string }
    }>
  }
}

function readRuntimeFormula(engine: SpreadsheetEngine, cellIndex: number): unknown {
  const formulas = Reflect.get(engine, 'formulas')
  if (typeof formulas !== 'object' || formulas === null || typeof Reflect.get(formulas, 'get') !== 'function') {
    throw new TypeError('Expected internal formulas store')
  }
  return Reflect.get(formulas, 'get').call(formulas, cellIndex)
}

function isRuntimeFormulaWithDirectCriteria(value: unknown): value is RuntimeFormulaWithDirectCriteria {
  if (typeof value !== 'object' || value === null) {
    return false
  }
  const directCriteria = Reflect.get(value, 'directCriteria')
  return typeof directCriteria === 'object' && directCriteria !== null && Array.isArray(Reflect.get(directCriteria, 'criteriaPairs'))
}

describe('engine snapshot metadata formula restore', () => {
  it('binds defined names before evaluating imported snapshot formulas', async () => {
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: {
        name: 'Imported Financial Model',
        metadata: {
          definedNames: [
            { name: 'Currency', value: { kind: 'cell-ref', sheetName: 'Constants', address: 'F7' } },
            { name: 'Thousands', value: { kind: 'cell-ref', sheetName: 'Constants', address: 'F9' } },
          ],
        },
      },
      sheets: [
        {
          id: 1,
          name: 'Constants',
          order: 0,
          cells: [
            { address: 'F7', value: 'USD' },
            { address: 'F9', formula: 'Currency & "  000s"' },
            { address: 'C12', formula: 'Thousands' },
          ],
        },
      ],
    }

    const engine = new SpreadsheetEngine({ workbookName: 'restored-defined-names' })
    await engine.ready()
    engine.importSnapshot(snapshot)

    expect(engine.getCellValue('Constants', 'F9')).toMatchObject({ tag: ValueTag.String, value: 'USD  000s' })
    expect(engine.getCellValue('Constants', 'C12')).toMatchObject({ tag: ValueTag.String, value: 'USD  000s' })
    expect(engine.explainCell('Constants', 'F9').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Constants', 'C12').mode).toBe(FormulaMode.JsOnly)
    expect(engine.explainCell('Constants', 'F9').directPrecedents).toEqual(['Constants!F7'])
    expect(engine.explainCell('Constants', 'C12').directPrecedents).toEqual(['Constants!F9'])
  })

  it('binds imported INDEX MATCH criteria arrays as direct first-match lookups', async () => {
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: {
        name: 'Imported NYSE Model',
        metadata: {
          definedNames: [
            { name: 'Symbol', value: { kind: 'range-ref', sheetName: 'Projectdata_NYSE', startAddress: 'A1', endAddress: 'A4' } },
            { name: 'Year_num', value: { kind: 'range-ref', sheetName: 'Projectdata_NYSE', startAddress: 'B1', endAddress: 'B4' } },
            { name: 'Total_Revenue', value: { kind: 'range-ref', sheetName: 'Projectdata_NYSE', startAddress: 'C1', endAddress: 'C4' } },
          ],
        },
      },
      sheets: [
        {
          id: 1,
          name: 'Projectdata_NYSE',
          order: 0,
          cells: [
            { address: 'A1', value: 'AAA' },
            { address: 'A2', value: 'BBB' },
            { address: 'A3', value: 'AAA' },
            { address: 'A4', value: 'AAA' },
            { address: 'B1', value: 2024 },
            { address: 'B2', value: 2024 },
            { address: 'B3', value: 2025 },
            { address: 'B4', value: 2026 },
            { address: 'C1', value: 100 },
            { address: 'C2', value: 200 },
            { address: 'C3', value: 300 },
            { address: 'C4', value: 400 },
          ],
        },
        {
          id: 2,
          name: 'Forecasting Model',
          order: 1,
          cells: [
            { address: 'B2', value: 'AAA' },
            { address: 'C1', value: 2025 },
            { address: 'C5', formula: 'IFERROR(INDEX(Total_Revenue,MATCH(1,($B$2=Symbol)*(C$1=Year_num),0)),0)' },
          ],
        },
      ],
    }

    const engine = new SpreadsheetEngine({ workbookName: 'restored-index-match-criteria' })
    await engine.ready()
    engine.importSnapshot(snapshot)

    expect(engine.getCellValue('Forecasting Model', 'C5')).toEqual({ tag: ValueTag.Number, value: 300 })

    const formulaCellIndex = engine.workbook.getCellIndex('Forecasting Model', 'C5')
    if (formulaCellIndex === undefined) {
      throw new Error('expected imported criteria lookup formula to be materialized')
    }
    const runtimeFormula = readRuntimeFormula(engine, formulaCellIndex)
    if (!isRuntimeFormulaWithDirectCriteria(runtimeFormula)) {
      throw new Error('expected imported INDEX/MATCH criteria array to expose direct criteria metadata')
    }
    expect(runtimeFormula.directCriteria.aggregateKind).toBe('first')
    expect(runtimeFormula.directCriteria.aggregateRange).toMatchObject({
      sheetName: 'Projectdata_NYSE',
      rowStart: 0,
      rowEnd: 3,
      col: 2,
      length: 4,
    })
    expect(runtimeFormula.directCriteria.criteriaPairs).toHaveLength(2)

    expect(engine.setCellValue('Forecasting Model', 'C1', 2026)).toEqual({ tag: ValueTag.Number, value: 2026 })
    expect(engine.getCellValue('Forecasting Model', 'C5')).toEqual({ tag: ValueTag.Number, value: 400 })

    expect(engine.setCellValue('Forecasting Model', 'B2', 'ZZZ')).toMatchObject({ tag: ValueTag.String, value: 'ZZZ' })
    expect(engine.getCellValue('Forecasting Model', 'C5')).toEqual({ tag: ValueTag.Number, value: 0 })
  })

  it('restores imported style ranges through the bulk range path', async () => {
    const styleRanges = Array.from({ length: 200 }, (_value, index) => ({
      range: {
        sheetName: 'Sheet1',
        startAddress: `A${index + 1}`,
        endAddress: `A${index + 1}`,
      },
      styleId: 'xlsx-style-1',
    }))
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: {
        name: 'Imported Styled Model',
        metadata: {
          styles: [{ id: 'xlsx-style-1', font: { bold: true } }],
        },
      },
      sheets: [
        {
          id: 1,
          name: 'Sheet1',
          order: 0,
          cells: styleRanges.map((range, index) => ({ address: range.range.startAddress, value: index + 1 })),
          metadata: {
            styleRanges,
          },
        },
      ],
    }

    const engine = new SpreadsheetEngine({ workbookName: 'restored-style-ranges' })
    await engine.ready()
    const setStyleRange = vi.spyOn(engine.workbook, 'setStyleRange')
    const setStyleRanges = vi.spyOn(engine.workbook, 'setStyleRanges')

    engine.importSnapshot(snapshot)

    expect(setStyleRange).not.toHaveBeenCalled()
    expect(setStyleRanges).toHaveBeenCalledTimes(1)
    expect(engine.workbook.listStyleRanges('Sheet1')).toHaveLength(styleRanges.length)
    expect(engine.getCell('Sheet1', 'A200').styleId).toBe('xlsx-style-1')
  })

  it('does not mark lazy rolling INDEX branches as static cycles', async () => {
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: {
        name: 'Imported Rolling Debt Model',
        metadata: {
          definedNames: [{ name: 'Debt_Opening', value: { kind: 'cell-ref', sheetName: 'Constants', address: 'B12' } }],
        },
      },
      sheets: [
        {
          id: 1,
          name: 'Constants',
          order: 0,
          cells: [{ address: 'B12', value: 50_000 }],
        },
        {
          id: 2,
          name: 'Calcs',
          order: 1,
          cells: [
            { address: 'S7', value: 1 },
            { address: 'U7', value: 0 },
            { address: 'T7', formula: 'IF(S7=1, Debt_Opening, INDEX(V7:V8,S7-1))' },
            { address: 'V7', formula: 'T7+U7' },
            { address: 'S8', value: 2 },
            { address: 'U8', value: -500 },
            { address: 'T8', formula: 'IF(S8=1, Debt_Opening, INDEX(V7:V8,S8-1))' },
            { address: 'V8', formula: 'T8+U8' },
          ],
        },
      ],
    }

    const engine = new SpreadsheetEngine({ workbookName: 'restored-rolling-index' })
    await engine.ready()
    engine.importSnapshot(snapshot)

    expect(engine.getCellValue('Calcs', 'T7')).toEqual({ tag: ValueTag.Number, value: 50_000 })
    expect(engine.getCellValue('Calcs', 'V7')).toEqual({ tag: ValueTag.Number, value: 50_000 })
    expect(engine.getCellValue('Calcs', 'T8')).toEqual({ tag: ValueTag.Number, value: 50_000 })
    expect(engine.getCellValue('Calcs', 'V8')).toEqual({ tag: ValueTag.Number, value: 49_500 })
    expect(engine.explainCell('Calcs', 'T7').inCycle).toBe(false)
    expect(engine.explainCell('Calcs', 'T8').directPrecedents).toContain('Calcs!V7')
    expect(engine.explainCell('Calcs', 'T8').directPrecedents).not.toContain('Calcs!V8')
  })

  it('does not mark rolling SUMIFS previous-period aggregates as static cycles', async () => {
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: { name: 'Imported Rolling SUMIFS Model' },
      sheets: [
        {
          id: 1,
          name: 'Calcs',
          order: 0,
          cells: [
            { address: 'AE7', value: 1 },
            { address: 'AF7', formula: 'SUMIFS(AG7:AG8,AE7:AE8,AE7-1)' },
            { address: 'AG7', formula: 'AF7+1' },
            { address: 'AE8', value: 2 },
            { address: 'AF8', formula: 'SUMIFS(AG7:AG8,AE7:AE8,AE8-1)' },
            { address: 'AG8', formula: 'AF8+1' },
          ],
        },
      ],
    }

    const engine = new SpreadsheetEngine({ workbookName: 'restored-rolling-sumifs' })
    await engine.ready()
    engine.importSnapshot(snapshot)

    expect(engine.getCellValue('Calcs', 'AF7')).toEqual({ tag: ValueTag.Number, value: 0 })
    expect(engine.getCellValue('Calcs', 'AG7')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Calcs', 'AF8')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Calcs', 'AG8')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.explainCell('Calcs', 'AF7').inCycle).toBe(false)
    expect(engine.explainCell('Calcs', 'AF8').directPrecedents).toContain('Calcs!AG7')
    expect(engine.explainCell('Calcs', 'AF8').directPrecedents).not.toContain('Calcs!AG8')
  })

  it('keeps normal range formulas dirty when compacted criteria aggregates share the same range', async () => {
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: { name: 'Shared Range Criteria Aggregate Model' },
      sheets: [
        {
          id: 1,
          name: 'Deals',
          order: 0,
          cells: [
            { address: 'A1', value: 'Region' },
            { address: 'B1', value: 'Segment' },
            { address: 'C1', value: 'Customers' },
            { address: 'D1', value: 'ARPA' },
            { address: 'E1', value: 'Revenue' },
            { address: 'A2', value: 'West' },
            { address: 'B2', value: 'Enterprise' },
            { address: 'C2', value: 12 },
            { address: 'D2', value: 1200 },
            { address: 'E2', formula: 'C2*D2' },
            { address: 'A3', value: 'East' },
            { address: 'B3', value: 'SMB' },
            { address: 'C3', value: 30 },
            { address: 'D3', value: 250 },
            { address: 'E3', formula: 'C3*D3' },
            { address: 'A4', value: 'West' },
            { address: 'B4', value: 'SMB' },
            { address: 'C4', value: 18 },
            { address: 'D4', value: 300 },
            { address: 'E4', formula: 'C4*D4' },
          ],
        },
        {
          id: 2,
          name: 'Summary',
          order: 1,
          cells: [
            { address: 'A1', value: 'Metric' },
            { address: 'B1', value: 'Value' },
            { address: 'A2', value: 'Total revenue' },
            { address: 'B2', formula: 'SUM(Deals!E2:E4)' },
            { address: 'A3', value: 'West customers' },
            { address: 'B3', formula: 'SUMIF(Deals!A2:A4,"West",Deals!C2:C4)' },
            { address: 'A6', value: 'Qualified customer counts' },
            { address: 'B6', formula: 'FILTER(Deals!C2:C4,Deals!C2:C4>=18)' },
          ],
        },
      ],
    }

    const engine = new SpreadsheetEngine({ workbookName: 'shared-range-criteria-dirtying' })
    await engine.ready()
    engine.importSnapshot(snapshot)

    expect(engine.getCellValue('Summary', 'B3')).toEqual({ tag: ValueTag.Number, value: 30 })
    expect(engine.getCellValue('Summary', 'B6')).toEqual({ tag: ValueTag.Number, value: 30 })
    expect(engine.getCellValue('Summary', 'B7')).toEqual({ tag: ValueTag.Number, value: 18 })

    expect(engine.setCellValue('Deals', 'C2', 20)).toEqual({ tag: ValueTag.Number, value: 20 })

    expect(engine.getCellValue('Summary', 'B3')).toEqual({ tag: ValueTag.Number, value: 38 })
    expect(engine.getCellValue('Summary', 'B6')).toEqual({ tag: ValueTag.Number, value: 20 })
    expect(engine.getCellValue('Summary', 'B7')).toEqual({ tag: ValueTag.Number, value: 30 })
    expect(engine.getCellValue('Summary', 'B8')).toEqual({ tag: ValueTag.Number, value: 18 })
  })

  it('resolves imported data-model GETPIVOTDATA formulas from visible pivot output', async () => {
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: {
        name: 'Imported Data Model Pivot Workbook',
        metadata: {
          definedNames: [
            { name: 'Start_Year', value: { kind: 'cell-ref', sheetName: 'Constants', address: 'A1' } },
            { name: 'Periods', value: { kind: 'cell-ref', sheetName: 'Constants', address: 'A2' } },
          ],
        },
      },
      sheets: [
        {
          id: 1,
          name: 'Constants',
          order: 0,
          cells: [
            { address: 'A1', value: 2020 },
            { address: 'A2', value: 2 },
          ],
        },
        {
          id: 2,
          name: 'Cash Flow',
          order: 1,
          cells: [
            { address: 'B5', value: 'Year' },
            { address: 'A6', value: 'Cash Flow' },
            { address: 'B6', value: 2020 },
            { address: 'C6', value: 2021 },
            { address: 'A24', value: 'Closing Cash Balance' },
            { address: 'B24', value: 229_912.44379265167 },
            { address: 'C24', value: 279_173.6002754987 },
          ],
        },
        {
          id: 3,
          name: 'CHKs',
          order: 2,
          cells: [
            {
              address: 'B16',
              formula:
                'GETPIVOTDATA("[Measures].[Sum of Closing Cash Balance]",\'Cash Flow\'!$A$5,"[tblCal].[Year]","[tblCal].[Year].&[" & Start_Year + Periods - 1 & "]")',
            },
          ],
        },
      ],
    }

    const engine = new SpreadsheetEngine({ workbookName: 'restored-visible-pivot-getpivotdata' })
    await engine.ready()
    engine.importSnapshot(snapshot)

    expect(engine.getCellValue('CHKs', 'B16')).toEqual({ tag: ValueTag.Number, value: 279_173.6002754987 })
  })
})
