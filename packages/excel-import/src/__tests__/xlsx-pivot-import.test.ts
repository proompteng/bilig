import { describe, expect, it } from 'vitest'
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'

import type { WorkbookSnapshot } from '@bilig/protocol'
import { exportXlsx, importXlsx } from '../index.js'

describe('xlsx pivot import', () => {
  it('imports pivot data fields that rely on the OOXML default subtotal', () => {
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: {
        name: 'pivot-default-subtotal',
        metadata: {
          pivots: [
            {
              name: 'SalesByRegion',
              sheetName: 'Pivot',
              address: 'A1',
              source: { sheetName: 'Data', startAddress: 'A1', endAddress: 'B4' },
              groupBy: ['Region'],
              values: [{ sourceColumn: 'Sales', summarizeBy: 'sum', outputLabel: 'Sales Total' }],
              rows: 4,
              cols: 2,
            },
          ],
        },
      },
      sheets: [
        {
          id: 1,
          name: 'Data',
          order: 0,
          cells: [
            { address: 'A1', value: 'Region' },
            { address: 'B1', value: 'Sales' },
            { address: 'A2', value: 'East' },
            { address: 'B2', value: 10 },
            { address: 'A3', value: 'West' },
            { address: 'B3', value: 7 },
            { address: 'A4', value: 'East' },
            { address: 'B4', value: 5 },
          ],
        },
        {
          id: 2,
          name: 'Pivot',
          order: 1,
          cells: [],
        },
      ],
    }
    const zip = unzipSync(exportXlsx(snapshot))
    const pivotPath = 'xl/pivotTables/pivotTable1.xml'
    const pivotXml = strFromU8(zip[pivotPath] ?? new Uint8Array())

    expect(pivotXml).toContain('subtotal="sum"')
    zip[pivotPath] = strToU8(pivotXml.replace(' subtotal="sum"', ''))

    const imported = importXlsx(zipSync(zip), 'pivot-default-subtotal.xlsx')

    expect(imported.snapshot.workbook.metadata?.pivots).toEqual([
      {
        name: 'SalesByRegion',
        sheetName: 'Pivot',
        address: 'A1',
        source: { sheetName: 'Data', startAddress: 'A1', endAddress: 'B4' },
        groupBy: ['Region'],
        values: [{ sourceColumn: 'Sales', summarizeBy: 'sum', outputLabel: 'Sales Total' }],
        rows: 4,
        cols: 2,
      },
    ])
  })
})
