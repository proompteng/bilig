import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { describe, expect, it } from 'vitest'

import type { WorkbookSnapshot } from '@bilig/protocol'
import { exportXlsx, importXlsx } from '../index.js'

const targetedIgnoredErrorsXml =
  '<ignoredErrors><ignoredError sqref="N8:N17 J6:J8 J12 J15 J17 N31:N36 F8:F19 J31:J35" formula="1"/><ignoredError sqref="K73" xmlns:x16r3="http://schemas.microsoft.com/office/spreadsheetml/2018/08/main" x16r3:misleadingFormat="1"/></ignoredErrors>'
const rootNamespacedIgnoredErrorsXml = '<ignoredErrors><ignoredError sqref="K73" x16r3:misleadingFormat="1"/></ignoredErrors>'
const x16r3NamespaceDeclaration = 'xmlns:x16r3="http://schemas.microsoft.com/office/spreadsheetml/2018/08/main"'

describe('worksheet ignored error metadata import/export', () => {
  it('preserves targeted ignoredErrors metadata without adding broad number text suppressions', () => {
    const imported = importXlsx(buildWorkbookWithIgnoredErrors(), 'ignored-errors.xlsx')

    const exportedSheetXml = worksheetXml(exportXlsx(imported.snapshot), 1)

    expect(exportedSheetXml).toContain(targetedIgnoredErrorsXml)
    expect(exportedSheetXml).not.toContain('numberStoredAsText')
    expect(exportedSheetXml).not.toContain('sqref="A1:XFD')
  })

  it('does not add broad ignoredErrors metadata when exporting text values that look numeric', () => {
    const exportedSheetXml = worksheetXml(exportXlsx(buildNumericTextWorkbook()), 1)

    expect(exportedSheetXml).not.toContain('<ignoredErrors')
    expect(exportedSheetXml).not.toContain('numberStoredAsText')
    expect(exportedSheetXml).not.toContain('sqref="A1:XFD')
  })

  it('keeps namespace declarations needed by preserved ignoredErrors attributes', () => {
    const imported = importXlsx(buildWorkbookWithRootNamespacedIgnoredErrors(), 'ignored-errors-namespace.xlsx')

    const exportedSheetXml = worksheetXml(exportXlsx(imported.snapshot), 1)

    expect(exportedSheetXml).toContain(
      rootNamespacedIgnoredErrorsXml.replace('<ignoredErrors', `<ignoredErrors ${x16r3NamespaceDeclaration}`),
    )
    expect(exportedSheetXml).not.toContain('numberStoredAsText')
  })
})

function buildWorkbookWithIgnoredErrors(): Uint8Array {
  const zip = unzipSync(exportXlsx(buildNumericTextWorkbook()))
  const sheetPath = 'xl/worksheets/sheet1.xml'
  const sheetXml = strFromU8(zip[sheetPath] ?? new Uint8Array())
  zip[sheetPath] = strToU8(sheetXml.replace('</worksheet>', `${targetedIgnoredErrorsXml}</worksheet>`))
  return zipSync(zip)
}

function buildWorkbookWithRootNamespacedIgnoredErrors(): Uint8Array {
  const zip = unzipSync(exportXlsx(buildNumericTextWorkbook()))
  const sheetPath = 'xl/worksheets/sheet1.xml'
  const sheetXml = strFromU8(zip[sheetPath] ?? new Uint8Array())
  zip[sheetPath] = strToU8(
    sheetXml
      .replace('<worksheet ', `<worksheet ${x16r3NamespaceDeclaration} `)
      .replace('</worksheet>', `${rootNamespacedIgnoredErrorsXml}</worksheet>`),
  )
  return zipSync(zip)
}

function buildNumericTextWorkbook(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: {
      name: 'ignored-errors',
    },
    sheets: [
      {
        id: 1,
        name: 'Review',
        order: 0,
        cells: [
          { address: 'A1', value: 'Account' },
          { address: 'B1', value: 'Code' },
          { address: 'A2', value: 'Cash' },
          { address: 'B2', value: '00125' },
          { address: 'J6', formula: 'A2', value: 'Cash' },
          { address: 'K73', value: 1200 },
          { address: 'N8', formula: 'B2', value: '00125' },
        ],
      },
    ],
  }
}

function worksheetXml(bytes: Uint8Array, sheetIndex: number): string {
  return strFromU8(unzipSync(bytes)[`xl/worksheets/sheet${String(sheetIndex)}.xml`] ?? new Uint8Array())
}
