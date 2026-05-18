import { describe, expect, it } from 'vitest'
import { strToU8, zipSync } from 'fflate'
import * as XLSX from 'xlsx'
import { readImportedWorksheetTextValues } from '../xlsx-worksheet-text-values.js'
import { workbookSheetPathEntries } from '../xlsx-workbook-sheet-paths.js'

const worksheetRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet'
const chartSheetRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet'

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

  it('does not align chartsheets to worksheet fallback paths by index', () => {
    const workbook = createWorkbookWithPackageParts({
      Directory: {
        sheets: ['xl/worksheets/sheet1.xml', 'xl/worksheets/sheet2.xml'],
      },
      files: {
        'xl/workbook.xml': {
          content: [
            '<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
            '<sheets>',
            '<sheet name="Chart1" sheetId="1" r:id="rIdChart"/>',
            '<sheet name="Sheet1" sheetId="2" r:id="rIdSheet1"/>',
            '<sheet name="Source" sheetId="3" r:id="rIdSource"/>',
            '</sheets>',
            '</workbook>',
          ].join(''),
        },
        'xl/_rels/workbook.xml.rels': {
          content: [
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
            `<Relationship Id="rIdChart" Type="${chartSheetRelationshipType}" Target="chartsheets/sheet1.xml"/>`,
            `<Relationship Id="rIdSheet1" Type="${worksheetRelationshipType}" Target="worksheets/sheet1.xml"/>`,
            `<Relationship Id="rIdSource" Type="${worksheetRelationshipType}" Target="worksheets/sheet2.xml"/>`,
            '</Relationships>',
          ].join(''),
        },
      },
    })

    expect(workbookSheetPathEntries(workbook, ['Chart1', 'Sheet1', 'Source'])).toEqual([
      { name: 'Sheet1', index: 1, path: 'xl/worksheets/sheet1.xml' },
      { name: 'Source', index: 2, path: 'xl/worksheets/sheet2.xml' },
    ])
  })

  it('does not read worksheet text values into unmapped chartsheets', () => {
    const source = zipSync({
      'xl/worksheets/sheet1.xml': strToU8(
        [
          '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
          '<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Actual text</t></is></c></row></sheetData>',
          '</worksheet>',
        ].join(''),
      ),
    })
    const values = readImportedWorksheetTextValues(source, ['Chart1', 'Sheet1'], new Map([['Sheet1', 'xl/worksheets/sheet1.xml']]), [
      'xl/worksheets/sheet1.xml',
    ])

    expect(values.has('Chart1')).toBe(false)
    expect(values.get('Sheet1')?.get('A1')).toBe('Actual text')
  })
})
