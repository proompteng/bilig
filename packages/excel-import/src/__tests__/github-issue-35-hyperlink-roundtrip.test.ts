import { describe, expect, it } from 'vitest'
import * as XLSX from 'xlsx'

import { exportXlsx, importXlsx } from '../index.js'

describe('GitHub issue #35 hyperlink roundtrip', () => {
  it('preserves external and internal cell hyperlinks through import and export', () => {
    const imported = importXlsx(buildHyperlinkWorkbookBytes(), 'hyperlinks.xlsx')

    expect(imported.snapshot.sheets[0]?.metadata?.hyperlinks).toEqual([
      {
        sheetName: 'Inputs',
        address: 'A1',
        target: 'https://example.com/report',
        tooltip: 'Open report',
        display: 'Open report',
      },
      {
        sheetName: 'Inputs',
        address: 'B2',
        target: '#Summary!A1',
        tooltip: 'Jump to summary',
        display: 'Summary',
      },
    ])

    const reimported = XLSX.read(exportXlsx(imported.snapshot), { type: 'array' })
    const sheet = reimported.Sheets['Inputs']

    expect(sheet?.['A1']?.l).toMatchObject({
      Target: 'https://example.com/report',
      Tooltip: 'Open report',
      display: 'Open report',
    })
    expect(sheet?.['B2']?.l).toMatchObject({
      Target: '#Summary!A1',
      Tooltip: 'Jump to summary',
      display: 'Summary',
    })
  })
})

function buildHyperlinkWorkbookBytes(): Uint8Array {
  const workbook = XLSX.utils.book_new()
  const inputs = XLSX.utils.aoa_to_sheet([['Open report'], ['', 'Summary']])
  inputs['A1'] = {
    ...inputs['A1'],
    l: {
      Target: 'https://example.com/report',
      Tooltip: 'Open report',
    },
  }
  inputs['B2'] = {
    ...inputs['B2'],
    l: {
      Target: '#Summary!A1',
      Tooltip: 'Jump to summary',
    },
  }
  XLSX.utils.book_append_sheet(workbook, inputs, 'Inputs')
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([['Destination']]), 'Summary')
  return XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' })
}
