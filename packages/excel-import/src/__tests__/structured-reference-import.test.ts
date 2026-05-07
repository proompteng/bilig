import { describe, expect, it } from 'vitest'
import { SpreadsheetEngine } from '@bilig/core'
import type { WorkbookSnapshot } from '@bilig/protocol'
import { ValueTag } from '@bilig/protocol'
import { exportXlsx, importXlsx } from '../index.js'

describe('structured reference XLSX import', () => {
  it('translates Excel table sections and this-row references into executable formulas', async () => {
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: {
        name: 'Structured Financial Model',
        metadata: {
          definedNames: [
            { name: 'Currency', value: { kind: 'cell-ref', sheetName: 'Constants', address: 'F7' } },
            { name: 'Start_Year', value: { kind: 'cell-ref', sheetName: 'Constants', address: 'B10' } },
          ],
          tables: [
            {
              name: 'tblPFS',
              sheetName: 'Constants',
              startAddress: 'H6',
              endAddress: 'O7',
              columnNames: ['Period', 'Revenue', 'CoGS', 'Gross Profit', 'Opex', 'EBITDA', 'Tax', 'Cash'],
              headerRow: true,
              totalsRow: false,
            },
            {
              name: 'tblActuals',
              sheetName: 'Imports',
              startAddress: 'A6',
              endAddress: 'D8',
              columnNames: ['Account', 'Value', 'Year', 'Period'],
              headerRow: true,
              totalsRow: false,
            },
          ],
        },
      },
      sheets: [
        {
          id: 1,
          name: 'Constants',
          order: 0,
          cells: [
            { address: 'B10', value: 2012 },
            { address: 'F7', value: 'USD' },
            { address: 'F9', formula: 'Currency & "  000s"' },
            { address: 'H6', value: 'Period' },
            { address: 'I6', value: 'Revenue' },
            { address: 'J6', value: 'CoGS' },
            { address: 'K6', value: 'Gross Profit' },
            { address: 'L6', value: 'Opex' },
            { address: 'M6', value: 'EBITDA' },
            { address: 'N6', value: 'Tax' },
            { address: 'O6', value: 'Cash' },
            { address: 'H7', formula: 'ROW()-ROW(tblPFS[#Headers])' },
          ],
        },
        {
          id: 2,
          name: 'Imports',
          order: 1,
          cells: [
            { address: 'A6', value: 'Account' },
            { address: 'B6', value: 'Value' },
            { address: 'C6', value: 'Year' },
            { address: 'D6', value: 'Period' },
            { address: 'A7', value: 'Revenue' },
            { address: 'B7', value: 100 },
            { address: 'C7', value: 2011 },
            { address: 'D7', formula: 'tblActuals[[#This Row],[Year]]-Start_Year+1' },
            { address: 'A8', value: 'Revenue' },
            { address: 'B8', value: 125 },
            { address: 'C8', value: 2012 },
            { address: 'D8', formula: 'tblActuals[[#This Row],[Year]]-Start_Year+1' },
            { address: 'F10', formula: 'SUM(tblActuals[Value])' },
          ],
        },
      ],
    }

    const imported = importXlsx(exportXlsx(snapshot), 'structured-financial-model.xlsx')
    const constants = imported.snapshot.sheets.find((sheet) => sheet.name === 'Constants')
    const imports = imported.snapshot.sheets.find((sheet) => sheet.name === 'Imports')

    expect(constants?.cells.find((cell) => cell.address === 'H7')?.formula).not.toContain('tblPFS[#Headers]')
    expect(imports?.cells.find((cell) => cell.address === 'D7')?.formula).not.toContain('[[#This Row],[Year]]')
    expect(imports?.cells.find((cell) => cell.address === 'F10')?.formula).not.toContain('tblActuals[Value]')

    const engine = new SpreadsheetEngine({ workbookName: 'structured-import' })
    await engine.ready()
    engine.importSnapshot(imported.snapshot)

    expect(engine.getCellValue('Constants', 'F9')).toMatchObject({ tag: ValueTag.String, value: 'USD  000s' })
    expect(engine.getCellValue('Constants', 'H7')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Imports', 'D7')).toEqual({ tag: ValueTag.Number, value: 0 })
    expect(engine.getCellValue('Imports', 'D8')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Imports', 'F10')).toEqual({ tag: ValueTag.Number, value: 225 })
  })

  it('preserves lowercase defined names and named ranges used by simulation formulas', async () => {
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: {
        name: 'Simulation Model',
        metadata: {
          definedNames: [
            { name: 'distribution', value: { kind: 'range-ref', sheetName: 'newsvendor', startAddress: 'H35', endAddress: 'K36' } },
            { name: 'mean', value: { kind: 'cell-ref', sheetName: 'newsvendor', address: 'D9' } },
            { name: 'std', value: { kind: 'cell-ref', sheetName: 'newsvendor', address: 'D10' } },
          ],
        },
      },
      sheets: [
        {
          id: 1,
          name: 'newsvendor',
          order: 0,
          cells: [
            { address: 'D9', value: 5000 },
            { address: 'D10', value: 250 },
            { address: 'H35', value: 0.25 },
            { address: 'I35', value: 125 },
            { address: 'J35', value: 'low' },
            { address: 'K35', value: true },
            { address: 'H36', value: 0.5 },
            { address: 'I36', value: 225 },
            { address: 'J36', value: 'high' },
            { address: 'K36', value: false },
            { address: 'I17', value: 0.975 },
            { address: 'I18', formula: 'mean+(NORM.S.INV(I17)*std)' },
            { address: 'I19', formula: 'VLOOKUP(0.5,distribution,2,FALSE)' },
          ],
        },
      ],
    }

    const imported = importXlsx(exportXlsx(snapshot), 'newsvendor-simulation.xlsm')
    const definedNames = imported.snapshot.workbook.metadata?.definedNames ?? []

    expect(definedNames.find((name) => name.name === 'mean')?.value).toEqual({
      kind: 'cell-ref',
      sheetName: 'newsvendor',
      address: 'D9',
    })
    expect(definedNames.find((name) => name.name === 'distribution')?.value).toEqual({
      kind: 'range-ref',
      sheetName: 'newsvendor',
      startAddress: 'H35',
      endAddress: 'K36',
    })

    const engine = new SpreadsheetEngine({ workbookName: 'newsvendor-import' })
    await engine.ready()
    engine.importSnapshot(imported.snapshot)

    const i18 = engine.getCellValue('newsvendor', 'I18')
    expect(i18.tag).toBe(ValueTag.Number)
    expect(i18.tag === ValueTag.Number ? i18.value : Number.NaN).toBeCloseTo(5489.990996, 5)
    expect(engine.getCellValue('newsvendor', 'I19')).toEqual({ tag: ValueTag.Number, value: 225 })
  })
})
