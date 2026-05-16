import { describe, expect, it } from 'vitest'
import * as XLSX from 'xlsx'
import { workbookSheetPathEntries } from '../xlsx-workbook-sheet-paths.js'

function createWorkbookWithPackageParts(parts: {
  readonly Directory: { readonly sheets: readonly string[] }
  readonly files: Record<string, { readonly content: string }>
}): XLSX.WorkBook {
  return Object.assign(XLSX.utils.book_new(), parts)
}

describe('xlsx workbook sheet paths', () => {
  it('uses workbook relationships before directory-order fallback paths', () => {
    const workbook = createWorkbookWithPackageParts({
      Directory: {
        sheets: ['xl/worksheets/fallback1.xml', 'xl/worksheets/fallback2.xml'],
      },
      files: {
        'xl/workbook.xml': {
          content: [
            '<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
            '<sheets>',
            '<sheet name="Actuals" sheetId="1" r:id="rIdActuals"/>',
            '<sheet name="Archive" sheetId="2" r:id="rIdArchive"/>',
            '</sheets>',
            '</workbook>',
          ].join(''),
        },
        'xl/_rels/workbook.xml.rels': {
          content: [
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
            '<Relationship Id="rIdActuals" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/actuals.xml"/>',
            '<Relationship Id="rIdArchive" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/archive.xml"/>',
            '</Relationships>',
          ].join(''),
        },
      },
    })

    expect(workbookSheetPathEntries(workbook, ['Actuals', 'Archive'])).toEqual([
      { name: 'Actuals', index: 0, path: 'xl/worksheets/actuals.xml' },
      { name: 'Archive', index: 1, path: 'xl/worksheets/archive.xml' },
    ])
  })

  it('falls back to workbook directory sheet order when workbook relationships are absent', () => {
    const workbook = createWorkbookWithPackageParts({
      Directory: {
        sheets: ['xl/worksheets/sheet1.xml', 'xl/worksheets/sheet2.xml'],
      },
      files: {},
    })

    expect(workbookSheetPathEntries(workbook, ['Sheet1', 'Sheet2', 'Missing'])).toEqual([
      { name: 'Sheet1', index: 0, path: 'xl/worksheets/sheet1.xml' },
      { name: 'Sheet2', index: 1, path: 'xl/worksheets/sheet2.xml' },
    ])
  })
})
