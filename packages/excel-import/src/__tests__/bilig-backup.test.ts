import { describe, expect, it } from 'vitest'
import type { WorkbookSnapshot } from '@bilig/protocol'
import { importWorkbookFile } from '../index.js'
import { importWorkbookFile as importBrowserWorkbookFile } from '../browser.js'

const contentType = 'application/vnd.bilig.workbook+json'
const snapshot: WorkbookSnapshot = {
  version: 1,
  workbook: { name: 'Quarterly budget', metadata: { definedNames: [{ name: 'TaxRate', value: 0.2 }] } },
  sheets: [
    {
      id: 1,
      name: 'Inputs',
      order: 0,
      cells: [
        { address: 'A1', value: 10 },
        { address: 'B1', formula: 'A1*TaxRate', format: '0.00' },
      ],
    },
  ],
}

describe('native Bilig backup import', () => {
  for (const [door, importFile] of [
    ['server', importWorkbookFile],
    ['browser', importBrowserWorkbookFile],
  ] as const) {
    it(`preserves names, formulas, formats, and metadata through the ${door} importer`, () => {
      const bytes = new TextEncoder().encode(JSON.stringify(snapshot))
      const result = importFile(bytes, 'budget.bilig.json', contentType)
      expect(result.snapshot).toEqual(snapshot)
      expect(result.workbookName).toBe('Quarterly budget')
      expect(result.sheetNames).toEqual(['Inputs'])
      expect(result.preview.sheets[0]?.previewRows[0]).toEqual(['10', '=A1*TaxRate'])
    })
    it(`rejects malformed snapshots through the ${door} importer`, () => {
      expect(() => importFile(new TextEncoder().encode('{"version":2}'), 'bad.bilig.json', contentType)).toThrow(
        'Invalid Bilig workbook backup',
      )
    })
  }
})
