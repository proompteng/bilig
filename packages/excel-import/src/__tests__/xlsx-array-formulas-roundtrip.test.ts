import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { describe, expect, it } from 'vitest'

import type { WorkbookSnapshot } from '@bilig/protocol'
import { exportXlsx, importXlsx } from '../index.js'

const sheetMetadataRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata'
const sheetMetadataContentType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml'

describe('xlsx array formulas roundtrip', () => {
  it('preserves array-formula ranges and related cell metadata references', () => {
    const source = buildWorkbookWithArrayFormulas()

    const imported = importXlsx(source, 'array-formula-roundtrip.xlsx')
    const exported = exportXlsx(imported.snapshot)

    expect(imported.snapshot.sheets[0]?.metadata?.arrayFormulas).toEqual({
      formulas: [
        {
          address: 'C2',
          formulaXml: '<f t="array" ref="C2:C4">TRANSPOSE(A2:C2)</f>',
        },
        {
          address: 'E2',
          formulaXml: '<f t="array" aca="1" ref="E2:F3" ca="1">_xlfn.STOCKHISTORY(&quot;MSFT&quot;,&quot;12/1/2020&quot;,,2)</f>',
        },
      ],
    })
    expect(arrayFormulaXml(exported)).toEqual(arrayFormulaXml(source))
    expect(cellOpenTag(exported, 'C2')).toContain('cm="1"')
    expect(cellOpenTag(exported, 'C2')).toContain('vm="1"')
    expect(cellXml(exported, 'C2')).toContain('<v>11</v>')
    expect(readZipText(exported, 'xl/metadata.xml')).toContain('XLDAPR')
  })
})

function buildWorkbookWithArrayFormulas(): Uint8Array {
  const zip = unzipSync(exportXlsx(buildWorkbook()))
  addWorkbookCellMetadata(zip)
  replaceWorksheetCell(zip, 'C2', '<c r="C2" cm="1" vm="1"><f t="array" ref="C2:C4">TRANSPOSE(A2:C2)</f><v>11</v></c>')
  replaceWorksheetCell(
    zip,
    'E2',
    '<c r="E2"><f t="array" aca="1" ref="E2:F3" ca="1">_xlfn.STOCKHISTORY(&quot;MSFT&quot;,&quot;12/1/2020&quot;,,2)</f><v>100</v></c>',
  )
  return zipSync(zip)
}

function buildWorkbook(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: { name: 'Array formulas' },
    sheets: [
      {
        id: 1,
        name: 'Forecast',
        order: 0,
        cells: [
          { address: 'A2', value: 10 },
          { address: 'B2', value: 20 },
          { address: 'C2', value: 11 },
          { address: 'C3', value: 12 },
          { address: 'C4', value: 13 },
          { address: 'E2', value: 100 },
          { address: 'F2', value: 101 },
          { address: 'E3', value: 102 },
          { address: 'F3', value: 103 },
        ],
      },
    ],
  }
}

function addWorkbookCellMetadata(zip: Record<string, Uint8Array>): void {
  zip['xl/metadata.xml'] = strToU8(buildMetadataXml())

  const workbookRelsXml = readZipTextFromZip(zip, 'xl/_rels/workbook.xml.rels')
  zip['xl/_rels/workbook.xml.rels'] = strToU8(
    workbookRelsXml.replace(
      '</Relationships>',
      `<Relationship Id="rIdArrayMetadata" Type="${sheetMetadataRelationshipType}" Target="metadata.xml"/></Relationships>`,
    ),
  )

  const contentTypesXml = readZipTextFromZip(zip, '[Content_Types].xml')
  zip['[Content_Types].xml'] = strToU8(
    contentTypesXml.replace('</Types>', `<Override PartName="/xl/metadata.xml" ContentType="${sheetMetadataContentType}"/></Types>`),
  )
}

function replaceWorksheetCell(zip: Record<string, Uint8Array>, address: string, nextCellXml: string): void {
  const sheetPath = 'xl/worksheets/sheet1.xml'
  const sheetXml = readZipTextFromZip(zip, sheetPath)
  const cellPattern = new RegExp(`<c\\b[^>]*\\br=(["'])${address}\\1[^>]*>[\\s\\S]*?<\\/c>`, 'u')
  if (!cellPattern.test(sheetXml)) {
    throw new Error(`Missing fixture cell ${address}`)
  }
  zip[sheetPath] = strToU8(sheetXml.replace(cellPattern, nextCellXml))
}

function buildMetadataXml(): string {
  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray">',
    '<metadataTypes count="1">',
    '<metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1" cellMeta="1"/>',
    '</metadataTypes>',
    '<futureMetadata name="XLDAPR" count="1"><bk><extLst><ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}"><xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk></futureMetadata>',
    '<cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata>',
    '<valueMetadata count="1"><bk><rc t="1" v="0"/></bk></valueMetadata>',
    '</metadata>',
  ].join('')
}

function arrayFormulaXml(bytes: Uint8Array): string[] {
  return [...readZipText(bytes, 'xl/worksheets/sheet1.xml').matchAll(/<f\b[^>]*\bt=(["'])array\1[^>]*(?:\/>|>[\s\S]*?<\/f>)/gu)].map(
    (match) => match[0],
  )
}

function cellOpenTag(bytes: Uint8Array, address: string): string {
  return readZipText(bytes, 'xl/worksheets/sheet1.xml').match(new RegExp(`<c[^>]* r="${address}"[^>]*>`, 'u'))?.[0] ?? ''
}

function cellXml(bytes: Uint8Array, address: string): string {
  return readZipText(bytes, 'xl/worksheets/sheet1.xml').match(new RegExp(`<c[^>]* r="${address}"[^>]*>[\\s\\S]*?<\\/c>`, 'u'))?.[0] ?? ''
}

function readZipText(bytes: Uint8Array, path: string): string {
  return readZipTextFromZip(unzipSync(bytes), path)
}

function readZipTextFromZip(zip: Record<string, Uint8Array>, path: string): string {
  const bytes = zip[path]
  if (!bytes) {
    throw new Error(`Missing XLSX part: ${path}`)
  }
  return strFromU8(bytes)
}
