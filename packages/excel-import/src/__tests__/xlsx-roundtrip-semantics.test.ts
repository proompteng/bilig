import { describe, expect, it } from 'vitest'
import { strFromU8, unzipSync } from 'fflate'

import type { WorkbookSnapshot } from '@bilig/protocol'
import { exportXlsx, importXlsx } from '../index.js'

describe('XLSX round-trip semantics', () => {
  it('preserves sort keys that point at a header row outside the sorted data range', () => {
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: { name: 'header-sort-key' },
      sheets: [
        {
          id: 1,
          name: 'Data',
          order: 0,
          cells: [
            { address: 'A1', value: 'Status' },
            { address: 'A2', value: 'Closed' },
            { address: 'A3', value: 'Open' },
            { address: 'A4', value: 'Pending' },
          ],
          metadata: {
            sorts: [
              {
                range: { sheetName: 'Data', startAddress: 'A2', endAddress: 'A4' },
                keys: [{ keyAddress: 'A1', direction: 'asc' }],
              },
            ],
          },
        },
      ],
    }

    const exported = exportXlsx(snapshot)
    const sheetXml = strFromU8(unzipSync(exported)['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
    const imported = importXlsx(exported, 'header-sort-key.xlsx')

    expect(sheetXml).toContain('<sortState ref="A2:A4">')
    expect(sheetXml).toContain('<sortCondition ref="A1:A4"/>')
    expect(imported.snapshot.sheets[0]?.metadata?.sorts).toEqual(snapshot.sheets[0]?.metadata?.sorts)
  })

  it('preserves custom row metadata on rows that have no cells', () => {
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: { name: 'blank-row-metadata' },
      sheets: [
        {
          id: 1,
          name: 'Rows',
          order: 0,
          cells: [{ address: 'A3', value: 'visible data' }],
          metadata: {
            rows: [
              { id: 'row:0', index: 0, size: 44 },
              { id: 'row:1', index: 1, hidden: true },
              { id: 'row:2', index: 2, size: 30 },
            ],
          },
        },
      ],
    }

    const exported = exportXlsx(snapshot)
    const sheetXml = strFromU8(unzipSync(exported)['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
    const importedRows = importXlsx(exported, 'blank-row-metadata.xlsx').snapshot.sheets[0]?.metadata?.rows

    expect(sheetXml).toContain('<row r="1" ht="44" customHeight="1"/>')
    expect(sheetXml).toContain('<row r="2" hidden="1"/>')
    expect(importedRows).toEqual(snapshot.sheets[0]?.metadata?.rows)
  })
})
