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
})
