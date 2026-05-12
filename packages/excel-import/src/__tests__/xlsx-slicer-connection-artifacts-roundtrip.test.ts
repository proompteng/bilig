import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { describe, expect, it } from 'vitest'

import type { WorkbookSnapshot } from '@bilig/protocol'
import { exportXlsx, importXlsx } from '../index.js'

const relationshipNamespace = 'http://schemas.openxmlformats.org/package/2006/relationships'
const officeRelationshipNamespace = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const connectionsRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/connections'
const slicerCacheRelationshipType = 'http://schemas.microsoft.com/office/2007/relationships/slicerCache'
const slicerRelationshipType = 'http://schemas.microsoft.com/office/2007/relationships/slicer'
const tableRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/table'
const connectionsContentType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml'
const slicerCacheContentType = 'application/vnd.ms-excel.slicerCache+xml'
const slicerContentType = 'application/vnd.ms-excel.slicer+xml'

describe('xlsx slicer and connection artifacts roundtrip', () => {
  it('preserves slicer caches, slicers, and workbook connections as package artifacts', () => {
    const source = buildWorkbookWithSlicerAndConnectionArtifacts()

    const imported = importXlsx(source, 'slicer-connections.xlsx')
    const exported = exportXlsx(imported.snapshot)

    expect(imported.snapshot.workbook.metadata?.slicerConnectionArtifacts?.parts.map((part) => part.path).toSorted()).toEqual([
      'xl/connections.xml',
      'xl/slicerCaches/_rels/slicerCache1.xml.rels',
      'xl/slicerCaches/slicerCache1.xml',
      'xl/slicers/slicer1.xml',
    ])
    expect(imported.snapshot.workbook.metadata?.slicerConnectionArtifacts?.workbookSlicerCachesExtXml).toBe(workbookSlicerCachesExtXml)
    expect(imported.snapshot.workbook.metadata?.slicerConnectionArtifacts?.workbookRelationships).toEqual([
      { id: 'rId80', type: slicerCacheRelationshipType, target: 'slicerCaches/slicerCache1.xml' },
      { id: 'rId81', type: connectionsRelationshipType, target: 'connections.xml' },
    ])
    expect(imported.snapshot.workbook.metadata?.slicerConnectionArtifacts?.sheetArtifacts).toEqual([
      {
        sheetName: 'Revenue',
        sheetSlicerListExtXml,
        relationships: [{ id: 'rId20', type: slicerRelationshipType, target: '../slicers/slicer1.xml' }],
      },
    ])
    expect(slicerConnectionMetrics(exported)).toEqual(slicerConnectionMetrics(source))
    expect(readZipText(exported, 'xl/connections.xml')).toBe(connectionsXml)
    expect(readZipText(exported, 'xl/slicerCaches/slicerCache1.xml')).toBe(slicerCacheXml)
    expect(readZipText(exported, 'xl/slicers/slicer1.xml')).toBe(slicerXml)
    expect(readContentTypeOverride(exported, '/xl/connections.xml')).toBe(connectionsContentType)
    expect(readContentTypeOverride(exported, '/xl/slicerCaches/slicerCache1.xml')).toBe(slicerCacheContentType)
    expect(readContentTypeOverride(exported, '/xl/slicers/slicer1.xml')).toBe(slicerContentType)
  })
})

function buildWorkbookWithSlicerAndConnectionArtifacts(): Uint8Array {
  const zip = unzipSync(exportXlsx(buildWorkbook()))
  const tablePath = Object.keys(zip).find((path) => /^xl\/tables\/table[1-9][0-9]*\.xml$/u.test(path))
  if (!tablePath) {
    throw new Error('Expected exported workbook to include a table part')
  }
  const tableTarget = `../tables/${tablePath.slice(tablePath.lastIndexOf('/') + 1)}`
  zip['xl/workbook.xml'] = strToU8(
    ensureRelationshipNamespace(readZipTextFromZip(zip, 'xl/workbook.xml')).replace(
      '</workbook>',
      `<extLst>${workbookSlicerCachesExtXml}</extLst></workbook>`,
    ),
  )
  zip['xl/_rels/workbook.xml.rels'] = strToU8(
    readZipTextFromZip(zip, 'xl/_rels/workbook.xml.rels').replace(
      '</Relationships>',
      [
        `<Relationship Id="rId80" Type="${slicerCacheRelationshipType}" Target="slicerCaches/slicerCache1.xml"/>`,
        `<Relationship Id="rId81" Type="${connectionsRelationshipType}" Target="connections.xml"/>`,
        '</Relationships>',
      ].join(''),
    ),
  )
  zip['xl/worksheets/sheet1.xml'] = strToU8(
    ensureRelationshipNamespace(readZipTextFromZip(zip, 'xl/worksheets/sheet1.xml')).replace(
      '</worksheet>',
      `<extLst>${sheetSlicerListExtXml}</extLst></worksheet>`,
    ),
  )
  zip['xl/worksheets/_rels/sheet1.xml.rels'] = strToU8(
    appendRelationship(
      readZipTextFromZip(zip, 'xl/worksheets/_rels/sheet1.xml.rels'),
      `<Relationship Id="rId20" Type="${slicerRelationshipType}" Target="../slicers/slicer1.xml"/>`,
    ),
  )
  zip['xl/connections.xml'] = strToU8(connectionsXml)
  zip['xl/slicerCaches/slicerCache1.xml'] = strToU8(slicerCacheXml)
  zip['xl/slicerCaches/_rels/slicerCache1.xml.rels'] = strToU8(slicerCacheRelationshipsXml(tableTarget))
  zip['xl/slicers/slicer1.xml'] = strToU8(slicerXml)
  zip['[Content_Types].xml'] = strToU8(
    [
      { partName: '/xl/connections.xml', contentType: connectionsContentType },
      { partName: '/xl/slicerCaches/slicerCache1.xml', contentType: slicerCacheContentType },
      { partName: '/xl/slicers/slicer1.xml', contentType: slicerContentType },
    ].reduce((xml, entry) => upsertContentTypeOverride(xml, entry), readZipTextFromZip(zip, '[Content_Types].xml')),
  )
  return zipSync(zip)
}

function buildWorkbook(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: {
      name: 'Slicer connection artifacts',
      metadata: {
        tables: [
          {
            name: 'RevenueTable',
            sheetName: 'Revenue',
            startAddress: 'A1',
            endAddress: 'B4',
            columnNames: ['Region', 'Amount'],
            headerRow: true,
            totalsRow: false,
          },
        ],
      },
    },
    sheets: [
      {
        id: 1,
        name: 'Revenue',
        order: 0,
        cells: [
          { address: 'A1', value: 'Region' },
          { address: 'B1', value: 'Amount' },
          { address: 'A2', value: 'North' },
          { address: 'B2', value: 1200 },
          { address: 'A3', value: 'South' },
          { address: 'B3', value: 900 },
          { address: 'A4', value: 'North' },
          { address: 'B4', value: 300 },
        ],
      },
    ],
  }
}

function slicerConnectionMetrics(bytes: Uint8Array): {
  packageParts: string[]
  sheetSlicerRelationships: number
  sheetSlicerRefs: number
  workbookConnectionsRelationships: number
  workbookSlicerCacheRelationships: number
  workbookSlicerCacheRefs: number
} {
  const zip = unzipSync(bytes)
  const workbookXml = readZipTextFromZip(zip, 'xl/workbook.xml')
  const workbookRelationshipsXml = readZipTextFromZip(zip, 'xl/_rels/workbook.xml.rels')
  const sheetXml = readZipTextFromZip(zip, 'xl/worksheets/sheet1.xml')
  const sheetRelationshipsXml = readZipTextFromZip(zip, 'xl/worksheets/_rels/sheet1.xml.rels')
  return {
    packageParts: Object.keys(zip)
      .filter((path) => path === 'xl/connections.xml' || path.startsWith('xl/slicerCaches/') || path.startsWith('xl/slicers/'))
      .toSorted(),
    sheetSlicerRelationships: relationshipsWithType(sheetRelationshipsXml, slicerRelationshipType).length,
    sheetSlicerRefs: [...sheetXml.matchAll(/<x14:slicer\b/gu)].length,
    workbookConnectionsRelationships: relationshipsWithType(workbookRelationshipsXml, connectionsRelationshipType).length,
    workbookSlicerCacheRelationships: relationshipsWithType(workbookRelationshipsXml, slicerCacheRelationshipType).length,
    workbookSlicerCacheRefs: [...workbookXml.matchAll(/<x15:slicerCache\b/gu)].length,
  }
}

function relationshipsWithType(relationshipsXml: string, relationshipType: string): string[] {
  return [...relationshipsXml.matchAll(/<Relationship\b([^>]*)\/?>/gu)].flatMap((match) => {
    const attributes = match[1] ?? ''
    return readXmlAttribute(attributes, 'Type') === relationshipType ? [match[0]] : []
  })
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

function readContentTypeOverride(bytes: Uint8Array, partName: string): string | undefined {
  const contentTypesXml = readZipText(bytes, '[Content_Types].xml')
  for (const match of contentTypesXml.matchAll(/<Override\b([^>]*)\/?>/gu)) {
    const attributes = match[1] ?? ''
    if (readXmlAttribute(attributes, 'PartName') === partName) {
      return readXmlAttribute(attributes, 'ContentType') ?? undefined
    }
  }
  return undefined
}

function ensureRelationshipNamespace(xml: string): string {
  if (/xmlns:r=/u.test(xml)) {
    return xml
  }
  return xml.replace(/<([A-Za-z0-9:]+)\b([^>]*)>/u, `<$1$2 xmlns:r="${officeRelationshipNamespace}">`)
}

function appendRelationship(relationshipsXml: string, relationshipXml: string): string {
  return relationshipsXml.replace('</Relationships>', `${relationshipXml}</Relationships>`)
}

function upsertContentTypeOverride(
  contentTypesXml: string,
  input: {
    readonly partName: string
    readonly contentType: string
  },
): string {
  if (contentTypesXml.includes(`PartName="${input.partName}"`)) {
    return contentTypesXml
  }
  return contentTypesXml.replace('</Types>', `<Override PartName="${input.partName}" ContentType="${input.contentType}"/></Types>`)
}

function readXmlAttribute(attributes: string, name: string): string | null {
  return new RegExp(`\\b${name}=("|')([\\s\\S]*?)\\1`, 'u').exec(attributes)?.[2] ?? null
}

const workbookSlicerCachesExtXml = [
  '<ext uri="{A8765BA9-456A-4DAB-B4F3-ACF838C121DE}" ',
  'xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">',
  '<x15:slicerCaches><x15:slicerCache r:id="rId80"/></x15:slicerCaches>',
  '</ext>',
].join('')

const sheetSlicerListExtXml = [
  '<ext uri="{A8765BA9-456A-4DAB-B4F3-ACF838C121DE}" ',
  'xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">',
  '<x14:slicerList><x14:slicer r:id="rId20"/></x14:slicerList>',
  '</ext>',
].join('')

const connectionsXml = [
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
  '<connections xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1">',
  '<connection id="1" name="Revenue connection" type="5" refreshedVersion="8" background="1">',
  '<dbPr connection="Provider=Microsoft.ACE.OLEDB.12.0;Data Source=revenue.xlsx" command="SELECT * FROM Revenue" commandType="2"/>',
  '</connection>',
  '</connections>',
].join('')

const slicerCacheXml = [
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
  '<slicerCacheDefinition xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" ',
  `xmlns:r="${officeRelationshipNamespace}" name="Slicer_Region" sourceName="Region" cache="Slicer_Region" r:id="rId1">`,
  '<tableSlicerCache tableId="1" column="1">',
  '<items count="2"><i x="0" s="1"/><i x="1"/></items>',
  '</tableSlicerCache>',
  '</slicerCacheDefinition>',
].join('')

function slicerCacheRelationshipsXml(tableTarget: string): string {
  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    `<Relationships xmlns="${relationshipNamespace}">`,
    `<Relationship Id="rId1" Type="${tableRelationshipType}" Target="${tableTarget}"/>`,
    '</Relationships>',
  ].join('')
}

const slicerXml = [
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
  '<slicer xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" ',
  'name="Slicer_Region" cache="Slicer_Region" caption="Region" startItem="0" columnCount="1"/>',
].join('')
