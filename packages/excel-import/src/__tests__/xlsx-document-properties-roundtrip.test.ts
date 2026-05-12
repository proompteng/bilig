import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { describe, expect, it } from 'vitest'

import type { WorkbookSnapshot } from '@bilig/protocol'
import { exportXlsx, importXlsx } from '../index.js'

const officeDocumentRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument'
const corePropertiesRelationshipType = 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties'
const extendedPropertiesRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties'
const customPropertiesRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties'
const relationshipNamespace = 'http://schemas.openxmlformats.org/package/2006/relationships'
const corePropertiesContentType = 'application/vnd.openxmlformats-package.core-properties+xml'
const extendedPropertiesContentType = 'application/vnd.openxmlformats-officedocument.extended-properties+xml'
const customPropertiesContentType = 'application/vnd.openxmlformats-officedocument.custom-properties+xml'

describe('xlsx document properties roundtrip', () => {
  it('preserves document property XML parts and package declarations across round trips', () => {
    const source = buildWorkbookWithDocumentProperties()

    const imported = importXlsx(source, 'document-properties.xlsx')
    const exported = exportXlsx(imported.snapshot)

    expect(imported.snapshot.workbook.metadata?.documentPropertyArtifacts?.core?.path).toBe('docProps/core.xml')
    expect(imported.snapshot.workbook.metadata?.documentPropertyArtifacts?.app?.path).toBe('docProps/app.xml')
    expect(imported.snapshot.workbook.metadata?.documentPropertyArtifacts?.custom?.path).toBe('docProps/custom.xml')
    expect(imported.snapshot.workbook.metadata?.properties).toEqual([
      { key: 'AuditScore', value: 98.5 },
      { key: 'ReviewStatus', value: 'Approved' },
    ])
    expect(readDocumentProperties(exported)).toEqual(readDocumentProperties(source))
    expect(readRootRelationship(exported, corePropertiesRelationshipType)).toMatchObject({
      target: 'docProps/core.xml',
      type: corePropertiesRelationshipType,
    })
    expect(readRootRelationship(exported, extendedPropertiesRelationshipType)).toMatchObject({
      target: 'docProps/app.xml',
      type: extendedPropertiesRelationshipType,
    })
    expect(readRootRelationship(exported, customPropertiesRelationshipType)).toMatchObject({
      target: 'docProps/custom.xml',
      type: customPropertiesRelationshipType,
    })
    expect(readContentTypeOverride(exported, '/docProps/core.xml')).toBe(corePropertiesContentType)
    expect(readContentTypeOverride(exported, '/docProps/app.xml')).toBe(extendedPropertiesContentType)
    expect(readContentTypeOverride(exported, '/docProps/custom.xml')).toBe(customPropertiesContentType)
  })
})

function buildWorkbookWithDocumentProperties(): Uint8Array {
  const zip = unzipSync(exportXlsx(buildWorkbook()))
  zip['_rels/.rels'] = strToU8(
    [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      `<Relationships xmlns="${relationshipNamespace}">`,
      `<Relationship Id="rId1" Type="${officeDocumentRelationshipType}" Target="xl/workbook.xml"/>`,
      `<Relationship Id="rId2" Type="${corePropertiesRelationshipType}" Target="docProps/core.xml"/>`,
      `<Relationship Id="rId3" Type="${extendedPropertiesRelationshipType}" Target="docProps/app.xml"/>`,
      `<Relationship Id="rId4" Type="${customPropertiesRelationshipType}" Target="docProps/custom.xml"/>`,
      '</Relationships>',
    ].join(''),
  )
  zip['docProps/core.xml'] = strToU8(corePropertiesXml)
  zip['docProps/app.xml'] = strToU8(appPropertiesXml)
  zip['docProps/custom.xml'] = strToU8(customPropertiesXml)
  zip['[Content_Types].xml'] = strToU8(
    upsertContentTypeOverride(
      upsertContentTypeOverride(
        upsertContentTypeOverride(readZipText(zip, '[Content_Types].xml'), '/docProps/core.xml', corePropertiesContentType),
        '/docProps/app.xml',
        extendedPropertiesContentType,
      ),
      '/docProps/custom.xml',
      customPropertiesContentType,
    ),
  )
  return zipSync(zip)
}

function buildWorkbook(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: { name: 'Document properties' },
    sheets: [
      {
        id: 1,
        name: 'Model',
        order: 0,
        cells: [
          { address: 'A1', value: 'Metric' },
          { address: 'B1', value: 'Value' },
        ],
      },
    ],
  }
}

function readDocumentProperties(bytes: Uint8Array): { app: string; core: string; custom: string } {
  const zip = unzipSync(bytes)
  return {
    app: readZipText(zip, 'docProps/app.xml'),
    core: readZipText(zip, 'docProps/core.xml'),
    custom: readZipText(zip, 'docProps/custom.xml'),
  }
}

function readRootRelationship(
  bytes: Uint8Array,
  relationshipType: string,
): { target: string; type: string; targetMode?: string } | undefined {
  const relationshipsXml = readZipText(unzipSync(bytes), '_rels/.rels')
  for (const match of relationshipsXml.matchAll(/<Relationship\b([^>]*)\/?>/gu)) {
    const attributes = match[1] ?? ''
    if (readXmlAttribute(attributes, 'Type') === relationshipType) {
      return {
        target: readXmlAttribute(attributes, 'Target') ?? '',
        type: readXmlAttribute(attributes, 'Type') ?? '',
        ...(readXmlAttribute(attributes, 'TargetMode') ? { targetMode: readXmlAttribute(attributes, 'TargetMode') ?? undefined } : {}),
      }
    }
  }
  return undefined
}

function readContentTypeOverride(bytes: Uint8Array, partName: string): string | undefined {
  const contentTypesXml = readZipText(unzipSync(bytes), '[Content_Types].xml')
  for (const match of contentTypesXml.matchAll(/<Override\b([^>]*)\/?>/gu)) {
    const attributes = match[1] ?? ''
    if (readXmlAttribute(attributes, 'PartName') === partName) {
      return readXmlAttribute(attributes, 'ContentType') ?? undefined
    }
  }
  return undefined
}

function upsertContentTypeOverride(contentTypesXml: string, partName: string, contentType: string): string {
  const escapedPartName = escapeXml(partName)
  const escapedContentType = escapeXml(contentType)
  const overridePattern = /<Override\b([^>]*)\/?>/gu
  let replaced = false
  const nextXml = contentTypesXml.replace(overridePattern, (match: string, attributes: string) => {
    if (readXmlAttribute(attributes, 'PartName') !== partName) {
      return match
    }
    replaced = true
    return `<Override PartName="${escapedPartName}" ContentType="${escapedContentType}"/>`
  })
  return replaced
    ? nextXml
    : nextXml.replace('</Types>', `<Override PartName="${escapedPartName}" ContentType="${escapedContentType}"/></Types>`)
}

function readZipText(zip: Record<string, Uint8Array>, path: string): string {
  const bytes = zip[path]
  if (!bytes) {
    throw new Error(`Missing XLSX part: ${path}`)
  }
  return strFromU8(bytes)
}

function readXmlAttribute(attributes: string, name: string): string | null {
  return new RegExp(`\\b${name}=("|')([\\s\\S]*?)\\1`, 'u').exec(attributes)?.[2] ?? null
}

function escapeXml(value: string): string {
  return value.replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;').replaceAll('"', '&quot;').replaceAll("'", '&apos;')
}

const corePropertiesXml = [
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
  '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ',
  'xmlns:dc="http://purl.org/dc/elements/1.1/" ',
  'xmlns:dcterms="http://purl.org/dc/terms/" ',
  'xmlns:dcmitype="http://purl.org/dc/dcmitype/" ',
  'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">',
  '<dc:creator>Career Principles</dc:creator>',
  '<cp:lastModifiedBy>Nathan Ayers</cp:lastModifiedBy>',
  '<dc:title>Board Model</dc:title>',
  '<dc:description>Reviewed finance workbook</dc:description>',
  '<dcterms:created xsi:type="dcterms:W3CDTF">2024-01-02T03:04:05Z</dcterms:created>',
  '<dcterms:modified xsi:type="dcterms:W3CDTF">2024-02-03T04:05:06Z</dcterms:modified>',
  '</cp:coreProperties>',
].join('')

const appPropertiesXml = [
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
  '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" ',
  'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">',
  '<Application>Microsoft Excel</Application>',
  '<AppVersion>16.0300</AppVersion>',
  '<Company>Bilig Capital</Company>',
  '<Manager>Finance Ops</Manager>',
  '<DocSecurity>0</DocSecurity>',
  '<ScaleCrop>false</ScaleCrop>',
  '<LinksUpToDate>false</LinksUpToDate>',
  '<SharedDoc>false</SharedDoc>',
  '<HyperlinksChanged>false</HyperlinksChanged>',
  '<TotalTime>42</TotalTime>',
  '</Properties>',
].join('')

const customPropertiesXml = [
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
  '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" ',
  'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">',
  '<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="ReviewStatus"><vt:lpwstr>Approved</vt:lpwstr></property>',
  '<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="3" name="AuditScore"><vt:r8>98.5</vt:r8></property>',
  '</Properties>',
].join('')
