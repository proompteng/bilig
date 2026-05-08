import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { describe, expect, it } from 'vitest'
import * as XLSX from 'xlsx'

import { exportXlsx, importXlsx } from '../index.js'

const printerSettingsRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings'
const printerSettingsContentType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings'
const relationshipNamespace = 'http://schemas.openxmlformats.org/package/2006/relationships'
const officeRelationshipNamespace = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

describe('printer settings roundtrip', () => {
  it('preserves binary printerSettings parts, worksheet relationships, and pageSetup links', () => {
    const source = buildPrinterSettingsWorkbookBytes()
    const imported = importXlsx(source, 'printer-settings.xlsx')

    expect(imported.snapshot.sheets[0]?.metadata?.printerSettings?.[0]).toMatchObject({
      relationshipTarget: '../printerSettings/printerSettings1.bin',
      storage: 'base64',
      byteLength: 5,
    })
    expect(imported.snapshot.sheets[1]?.metadata?.printerSettings?.[0]).toMatchObject({
      relationshipTarget: '../printerSettings/printerSettings2.bin',
      storage: 'base64',
      byteLength: 4,
    })

    const exported = exportXlsx(imported.snapshot)
    expect(readZipBytes(exported, 'xl/printerSettings/printerSettings1.bin')).toEqual([1, 3, 5, 7, 9])
    expect(readZipBytes(exported, 'xl/printerSettings/printerSettings2.bin')).toEqual([2, 4, 6, 8])

    expect(readPrinterSettingsRelationship(exported, 1, '../printerSettings/printerSettings1.bin')).toBeTruthy()
    expect(readPrinterSettingsRelationship(exported, 2, '../printerSettings/printerSettings2.bin')).toBeTruthy()
    expect(readPageSetup(exported, 1)).toMatchObject({
      paperSize: '9',
      orientation: 'landscape',
      relationshipId: readPrinterSettingsRelationship(exported, 1, '../printerSettings/printerSettings1.bin'),
    })
    expect(readPageSetup(exported, 2)).toMatchObject({
      paperSize: '1',
      orientation: 'portrait',
      relationshipId: readPrinterSettingsRelationship(exported, 2, '../printerSettings/printerSettings2.bin'),
    })
    expect(strFromU8(unzipSync(exported)['[Content_Types].xml'] ?? new Uint8Array())).toContain(printerSettingsContentType)
  })
})

function buildPrinterSettingsWorkbookBytes(): Uint8Array {
  const workbook = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([['sheet one']]), 'Print One')
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([['sheet two']]), 'Print Two')

  const zip = unzipSync(XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' }))
  addPrinterSettings(zip, {
    sheetIndex: 1,
    relationshipId: 'rIdPrinterA',
    target: '../printerSettings/printerSettings1.bin',
    bytes: [1, 3, 5, 7, 9],
    pageSetupXml: '<pageSetup paperSize="9" orientation="landscape" r:id="rIdPrinterA"/>',
  })
  addPrinterSettings(zip, {
    sheetIndex: 2,
    relationshipId: 'rIdPrinterB',
    target: '../printerSettings/printerSettings2.bin',
    bytes: [2, 4, 6, 8],
    pageSetupXml: '<pageSetup paperSize="1" orientation="portrait" r:id="rIdPrinterB"/>',
  })
  addPrinterSettingsContentType(zip, '/xl/printerSettings/printerSettings1.bin')
  addPrinterSettingsContentType(zip, '/xl/printerSettings/printerSettings2.bin')
  return zipSync(zip)
}

function addPrinterSettings(
  zip: Record<string, Uint8Array>,
  input: {
    readonly sheetIndex: number
    readonly relationshipId: string
    readonly target: string
    readonly bytes: readonly number[]
    readonly pageSetupXml: string
  },
): void {
  const sheetPath = `xl/worksheets/sheet${String(input.sheetIndex)}.xml`
  const relsPath = `xl/worksheets/_rels/sheet${String(input.sheetIndex)}.xml.rels`
  const partPath = `xl/printerSettings/printerSettings${String(input.sheetIndex)}.bin`
  zip[partPath] = Uint8Array.from(input.bytes)
  zip[relsPath] = strToU8(
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${relationshipNamespace}"><Relationship Id="${input.relationshipId}" Type="${printerSettingsRelationshipType}" Target="${input.target}"/></Relationships>`,
  )
  const sheetXml = ensureOfficeRelationshipNamespace(strFromU8(zip[sheetPath] ?? new Uint8Array()))
  zip[sheetPath] = strToU8(sheetXml.replace('</worksheet>', `${input.pageSetupXml}</worksheet>`))
}

function ensureOfficeRelationshipNamespace(sheetXml: string): string {
  return /\sxmlns:r=(["'])[\s\S]*?\1/u.test(sheetXml)
    ? sheetXml
    : sheetXml.replace(/<worksheet\b([^>]*)>/u, `<worksheet$1 xmlns:r="${officeRelationshipNamespace}">`)
}

function addPrinterSettingsContentType(zip: Record<string, Uint8Array>, partName: string): void {
  const contentTypesXml = strFromU8(zip['[Content_Types].xml'] ?? new Uint8Array())
  zip['[Content_Types].xml'] = strToU8(
    contentTypesXml.replace('</Types>', `<Override PartName="${partName}" ContentType="${printerSettingsContentType}"/></Types>`),
  )
}

function readZipBytes(bytes: Uint8Array, path: string): number[] {
  return [...(unzipSync(bytes)[path] ?? new Uint8Array())]
}

function readPrinterSettingsRelationship(bytes: Uint8Array, sheetIndex: number, target: string): string | null {
  const relsXml = strFromU8(unzipSync(bytes)[`xl/worksheets/_rels/sheet${String(sheetIndex)}.xml.rels`] ?? new Uint8Array())
  const relationship = [...relsXml.matchAll(/<Relationship\b([^>]*)\/?>/gu)]
    .map((match) => match[1] ?? '')
    .find(
      (attributes) =>
        readAttribute(attributes, 'Type') === printerSettingsRelationshipType && readAttribute(attributes, 'Target') === target,
    )
  return relationship ? readAttribute(relationship, 'Id') : null
}

function readPageSetup(
  bytes: Uint8Array,
  sheetIndex: number,
): {
  readonly paperSize: string | null
  readonly orientation: string | null
  readonly relationshipId: string | null
} {
  const sheetXml = strFromU8(unzipSync(bytes)[`xl/worksheets/sheet${String(sheetIndex)}.xml`] ?? new Uint8Array())
  const pageSetupXml = /<pageSetup\b[^>]*(?:\/>|>[\s\S]*?<\/pageSetup>)/u.exec(sheetXml)?.[0] ?? ''
  return {
    paperSize: readAttribute(pageSetupXml, 'paperSize'),
    orientation: readAttribute(pageSetupXml, 'orientation'),
    relationshipId: readAttribute(pageSetupXml, 'r:id'),
  }
}

function readAttribute(xml: string, attributeName: string): string | null {
  const match = new RegExp(`\\s${attributeName}=(["'])([\\s\\S]*?)\\1`, 'u').exec(xml)
  return match?.[2] ?? null
}
