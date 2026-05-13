import { describe, expect, it } from 'vitest'
import { strFromU8, unzipSync } from 'fflate'
import * as XLSX from 'xlsx'

import type { WorkbookSnapshot } from '@bilig/protocol'
import { exportXlsx, importXlsx } from '../index.js'

describe('XLSX sheet-name whitespace export', () => {
  it('preserves trailing spaces in raw workbook sheet names and formulas', () => {
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: { name: 'country-erp-whitespace' },
      sheets: [
        {
          id: 'country-erp',
          name: 'Country ERP ',
          order: 0,
          cells: [
            { address: 'A5', value: 'United States' },
            { address: 'F196', value: 4.25 },
          ],
          metadata: {
            merges: [{ sheetName: 'Country ERP ', startAddress: 'A226', endAddress: 'C226' }],
          },
        },
        {
          id: 'inputs',
          name: 'Inputs',
          order: 1,
          cells: [{ address: 'B15', formula: "VLOOKUP(B8,'Country ERP '!A5:F196,6,FALSE)" }],
        },
      ],
    }

    const exported = exportXlsx(snapshot)
    const rawWorkbook = XLSX.read(exported, { type: 'array', cellFormula: true })
    const workbookXml = strFromU8(unzipSync(exported)['xl/workbook.xml'] ?? new Uint8Array())
    const roundTripped = importXlsx(exported, 'country-erp-whitespace.xlsx')

    expect(rawWorkbook.SheetNames).toContain('Country ERP ')
    expect(rawWorkbook.SheetNames).not.toContain('Country ERP')
    expect(workbookXml).toContain('<sheet name="Country ERP "')
    expect(roundTripped.snapshot.sheets.map((sheet) => sheet.name)).toContain('Country ERP ')
    expect(roundTripped.snapshot.sheets.map((sheet) => sheet.name)).not.toContain('Country ERP')

    const inputSheet = roundTripped.snapshot.sheets.find((sheet) => sheet.name === 'Inputs')
    expect(inputSheet?.cells).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          address: 'B15',
          formula: "VLOOKUP(B8,'Country ERP '!A5:F196,6,FALSE)",
        }),
      ]),
    )

    const countrySheet = roundTripped.snapshot.sheets.find((sheet) => sheet.name === 'Country ERP ')
    expect(countrySheet?.metadata?.merges).toEqual([{ sheetName: 'Country ERP ', startAddress: 'A226', endAddress: 'C226' }])
  })

  it('preserves whitespace-only sheet names across export round trips', () => {
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: { name: 'whitespace-only-sheet' },
      sheets: [
        {
          id: 'blank-named-sheet',
          name: ' ',
          order: 0,
          cells: [{ address: 'A1', value: 'Visible data' }],
          metadata: {
            merges: [{ sheetName: ' ', startAddress: 'A3', endAddress: 'B3' }],
          },
        },
      ],
    }

    const exported = exportXlsx(snapshot)
    const rawWorkbook = XLSX.read(exported, { type: 'array', cellFormula: true })
    const workbookXml = strFromU8(unzipSync(exported)['xl/workbook.xml'] ?? new Uint8Array())
    const roundTripped = importXlsx(exported, 'whitespace-only-sheet.xlsx')

    expect(rawWorkbook.SheetNames).toEqual([' '])
    expect(rawWorkbook.SheetNames).not.toContain('Sheet1')
    expect(workbookXml).toContain('<sheet name=" "')
    expect(roundTripped.snapshot.sheets[0]?.name).toBe(' ')
    expect(roundTripped.snapshot.sheets[0]?.cells).toEqual([{ address: 'A1', value: 'Visible data' }])
    expect(roundTripped.snapshot.sheets[0]?.metadata?.merges).toEqual([{ sheetName: ' ', startAddress: 'A3', endAddress: 'B3' }])
  })
})
