import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { describe, expect, it } from 'vitest'
import * as XLSX from 'xlsx'

import { exportXlsx, importXlsx } from '../index.js'

const sheetMetadataRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata'
const sheetMetadataContentType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml'

describe('cell metadata roundtrip', () => {
  it('preserves workbook rich-value metadata and worksheet cm/vm references on no-op roundtrip', () => {
    const source = buildCellMetadataWorkbookBytes()
    const sourceMetadata = readCellMetadataSummary(source)

    const imported = importXlsx(source, 'cell-metadata.xlsx')
    const exported = exportXlsx(imported.snapshot)
    const exportedMetadata = readCellMetadataSummary(exported)

    expect(exportedMetadata).toEqual(sourceMetadata)
    expect(exportedMetadata.futureMetadataNames).toEqual(['XLDAPR', 'XLRICHVALUE'])
    expect(exportedMetadata.cellMetadataRefs).toEqual([
      { sheetPath: 'xl/worksheets/sheet1.xml', address: 'A2', cm: '1', vm: '1' },
      { sheetPath: 'xl/worksheets/sheet1.xml', address: 'B2', cm: '2' },
    ])
  })

  it('does not restore stale cell metadata references after the referenced cell changes', () => {
    const imported = importXlsx(buildCellMetadataWorkbookBytes(), 'cell-metadata.xlsx')
    const richValueSheet = imported.snapshot.sheets[0]
    const editedCell = richValueSheet?.cells.find((cell) => cell.address === 'A2')
    if (!editedCell) {
      throw new Error('Fixture import did not produce the expected A2 cell.')
    }
    editedCell.value = 'AAPL'

    const exportedMetadata = readCellMetadataSummary(exportXlsx(imported.snapshot))

    expect(exportedMetadata.futureMetadataNames).toEqual(['XLDAPR', 'XLRICHVALUE'])
    expect(exportedMetadata.cellMetadataRefs).toEqual([{ sheetPath: 'xl/worksheets/sheet1.xml', address: 'B2', cm: '2' }])
  })
})

interface CellMetadataSummary {
  readonly hasMetadataPart: boolean
  readonly hasWorkbookRelationship: boolean
  readonly hasContentTypeOverride: boolean
  readonly futureMetadataNames: readonly string[]
  readonly cellMetadataRefs: readonly CellMetadataRefSummary[]
}

interface CellMetadataRefSummary {
  readonly sheetPath: string
  readonly address: string
  readonly cm?: string
  readonly vm?: string
}

function buildCellMetadataWorkbookBytes(): Uint8Array {
  const workbook = XLSX.utils.book_new()
  const worksheet = XLSX.utils.aoa_to_sheet([
    ['Ticker', 'Price'],
    ['MSFT', 415.32],
  ])
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Rich Values')
  const zip = unzipSync(XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' }))
  addWorkbookCellMetadata(zip)
  addWorksheetCellMetadataRefs(zip)
  return zipSync(zip)
}

function addWorkbookCellMetadata(zip: Record<string, Uint8Array>): void {
  zip['xl/metadata.xml'] = strToU8(buildMetadataXml())

  const workbookRelsXml = strFromU8(zip['xl/_rels/workbook.xml.rels'] ?? new Uint8Array())
  zip['xl/_rels/workbook.xml.rels'] = strToU8(
    workbookRelsXml.replace(
      '</Relationships>',
      `<Relationship Id="rIdCellMetadata" Type="${sheetMetadataRelationshipType}" Target="metadata.xml"/></Relationships>`,
    ),
  )

  const contentTypesXml = strFromU8(zip['[Content_Types].xml'] ?? new Uint8Array())
  zip['[Content_Types].xml'] = strToU8(
    contentTypesXml.replace('</Types>', `<Override PartName="/xl/metadata.xml" ContentType="${sheetMetadataContentType}"/></Types>`),
  )
}

function addWorksheetCellMetadataRefs(zip: Record<string, Uint8Array>): void {
  const sheetPath = 'xl/worksheets/sheet1.xml'
  const sheetXml = strFromU8(zip[sheetPath] ?? new Uint8Array())
  zip[sheetPath] = strToU8(sheetXml.replace('<c r="A2"', '<c r="A2" cm="1" vm="1"').replace('<c r="B2"', '<c r="B2" cm="2"'))
}

function buildMetadataXml(): string {
  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray" xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">',
    '<metadataTypes count="2">',
    '<metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1" cellMeta="1"/>',
    '<metadataType name="XLRICHVALUE" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1"/>',
    '</metadataTypes>',
    '<futureMetadata name="XLDAPR" count="1"><bk><extLst><ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}"><xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk></futureMetadata>',
    '<futureMetadata name="XLRICHVALUE" count="1"><bk><extLst><ext uri="{3e2802c4-a4d2-4d8b-9148-e3be6c30e623}"><xlrd:rvb i="0"/></ext></extLst></bk></futureMetadata>',
    '<cellMetadata count="2"><bk><rc t="1" v="0"/></bk><bk><rc t="1" v="0"/></bk></cellMetadata>',
    '<valueMetadata count="1"><bk><rc t="2" v="0"/></bk></valueMetadata>',
    '</metadata>',
  ].join('')
}

function readCellMetadataSummary(bytes: Uint8Array): CellMetadataSummary {
  const zip = unzipSync(bytes)
  const metadataXml = strFromU8(zip['xl/metadata.xml'] ?? new Uint8Array())
  const workbookRelsXml = strFromU8(zip['xl/_rels/workbook.xml.rels'] ?? new Uint8Array())
  const contentTypesXml = strFromU8(zip['[Content_Types].xml'] ?? new Uint8Array())
  return {
    hasMetadataPart: Boolean(zip['xl/metadata.xml']),
    hasWorkbookRelationship:
      workbookRelsXml.includes(`Type="${sheetMetadataRelationshipType}"`) && workbookRelsXml.includes('Target="metadata.xml"'),
    hasContentTypeOverride:
      contentTypesXml.includes(`PartName="/xl/metadata.xml"`) && contentTypesXml.includes(`ContentType="${sheetMetadataContentType}"`),
    futureMetadataNames: [...metadataXml.matchAll(/<futureMetadata\b[^>]*\bname=(["'])([\s\S]*?)\1/gu)].map((match) => match[2] ?? ''),
    cellMetadataRefs: readWorksheetCellMetadataRefs(zip),
  }
}

function readWorksheetCellMetadataRefs(zip: Record<string, Uint8Array>): CellMetadataRefSummary[] {
  const refs: CellMetadataRefSummary[] = []
  for (const sheetPath of Object.keys(zip)
    .filter((path) => /^xl\/worksheets\/sheet\d+\.xml$/u.test(path))
    .toSorted()) {
    const sheetXml = strFromU8(zip[sheetPath] ?? new Uint8Array())
    for (const cellTag of sheetXml.match(/<c\b[^>]*>/gu) ?? []) {
      const cm = readAttribute(cellTag, 'cm')
      const vm = readAttribute(cellTag, 'vm')
      if (!cm && !vm) {
        continue
      }
      const address = readAttribute(cellTag, 'r')
      if (!address) {
        continue
      }
      refs.push({
        sheetPath,
        address,
        ...(cm ? { cm } : {}),
        ...(vm ? { vm } : {}),
      })
    }
  }
  return refs
}

function readAttribute(xml: string, attributeName: string): string | null {
  const match = new RegExp(`\\s${attributeName}=(["'])([\\s\\S]*?)\\1`, 'u').exec(xml)
  return match?.[2] ?? null
}
