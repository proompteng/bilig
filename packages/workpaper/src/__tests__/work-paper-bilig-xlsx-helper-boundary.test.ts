import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { describe, expect, it, vi } from 'vitest'
import type * as BiligXlsx from '@bilig/xlsx'

import { exportWorkPaperXlsx, exportXlsx } from '../xlsx.js'

vi.mock('@bilig/excel-import', async () => {
  const biligXlsx = await vi.importActual<typeof BiligXlsx>('@bilig/xlsx')
  return {
    createFileImportedXlsxSourceReader: () => {
      throw new Error('file source reader should not be created in this test')
    },
    createTempFileImportedXlsxSourceReader: () => {
      throw new Error('temp source reader should not be created in this test')
    },
    exportXlsx: () => {
      throw new Error('Generic XLSX writer should not be called for source-preserving WorkPaper XLSX export')
    },
    exportXlsxToFile: () => {
      throw new Error('Generic XLSX file writer should not be called in this test')
    },
    exportXlsxSourceLiteralPatches: biligXlsx.exportXlsxSourceLiteralPatches,
    exportXlsxSourceLiteralPatchesToFile: biligXlsx.exportXlsxSourceLiteralPatchesToFile,
    exportXlsxSourceLiteralPatchesToFileAsync: biligXlsx.exportXlsxSourceLiteralPatchesToFileAsync,
    importXlsx: () => {
      throw new Error('XLSX import should not be called in this test')
    },
    importXlsxFromZipByteSource: () => {
      throw new Error('XLSX byte-source import should not be called in this test')
    },
  }
})

const importedXlsxSourceBytes = Symbol.for('bilig.importedXlsxSourceBytes')
const importedXlsxSourceCellPatches = Symbol.for('bilig.importedXlsxSourceCellPatches')

describe('WorkPaper @bilig/xlsx helper boundary', () => {
  it('exports WorkPaper inputs through source-preserving scalar patches without loading the generic writer', () => {
    const snapshot = sourcePreservingSnapshot(9)

    const exported = exportXlsx({
      exportSourcePreservingXlsxSnapshot: () => snapshot,
      exportSnapshot: () => {
        throw new Error('Full WorkPaper snapshot should not be exported for direct source-preserving XLSX export')
      },
    })

    const exportedZip = unzipSync(exported)
    const sheetXml = strFromU8(exportedZip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
    expect(cellXml(sheetXml, 'A1')).toContain('<v>9</v>')
    expect(strFromU8(exportedZip['customXml/item1.xml'] ?? new Uint8Array())).toBe('<keep source="true"/>')
  })

  it('exports source-preserving scalar patches without loading the SheetJS writer', () => {
    const snapshot = sourcePreservingSnapshot(9)

    const exported = exportWorkPaperXlsx({
      exportSourcePreservingXlsxSnapshot: () => snapshot,
      exportSnapshot: () => {
        throw new Error('Full WorkPaper snapshot should not be exported for source-preserving scalar patches')
      },
    })

    const exportedZip = unzipSync(exported)
    const sheetXml = strFromU8(exportedZip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
    expect(cellXml(sheetXml, 'A1')).toContain('<v>9</v>')
    expect(strFromU8(exportedZip['customXml/item1.xml'] ?? new Uint8Array())).toBe('<keep source="true"/>')
  })
})

function sourcePreservingSnapshot(value: number): {
  readonly version: 1
  readonly workbook: { readonly name: string }
  readonly sheets: readonly [{ readonly id: 1; readonly name: 'Data'; readonly order: 0; readonly cells: readonly [] }]
} {
  const snapshot = {
    version: 1 as const,
    workbook: { name: 'bilig-xlsx-helper-boundary' },
    sheets: [{ id: 1 as const, name: 'Data' as const, order: 0 as const, cells: [] as const }],
  }
  Object.defineProperty(snapshot, importedXlsxSourceBytes, {
    configurable: true,
    enumerable: false,
    value: sourceWorkbookBytes(),
  })
  Object.defineProperty(snapshot, importedXlsxSourceCellPatches, {
    configurable: true,
    enumerable: false,
    value: [{ kind: 'literal', sheetName: 'Data', address: 'A1', value }],
  })
  return snapshot
}

function sourceWorkbookBytes(): Uint8Array {
  return zipSync({
    '[Content_Types].xml': strToU8(
      [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
        '<Default Extension="xml" ContentType="application/xml"/>',
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>',
        '</Types>',
      ].join(''),
    ),
    '_rels/.rels': strToU8(
      [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>',
        '</Relationships>',
      ].join(''),
    ),
    'xl/workbook.xml': strToU8(
      [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
        '<sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>',
        '</workbook>',
      ].join(''),
    ),
    'xl/_rels/workbook.xml.rels': strToU8(
      [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>',
        '</Relationships>',
      ].join(''),
    ),
    'xl/worksheets/sheet1.xml': strToU8(
      [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        '<dimension ref="A1"/>',
        '<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>',
        '</worksheet>',
      ].join(''),
    ),
    'customXml/item1.xml': strToU8('<keep source="true"/>'),
  })
}

function cellXml(sheetXml: string, address: string): string {
  return new RegExp(`<c\\b(?=[^>]*\\br="${address}")(?:[^>"']|"[^"]*"|'[^']*')*(?:/>|>[\\s\\S]*?</c>)`, 'u').exec(sheetXml)?.[0] ?? ''
}
