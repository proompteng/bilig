import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { describe, expect, it } from 'vitest'
import * as XLSX from 'xlsx'

import type { WorkbookSnapshot } from '@bilig/protocol'
import { exportXlsx, importXlsx } from '../index.js'

const relationshipNamespace = 'http://schemas.openxmlformats.org/package/2006/relationships'
const officeRelationshipNamespace = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const chartSheetRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet'
const drawingRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing'
const chartRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart'
const chartSheetContentType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml'
const drawingContentType = 'application/vnd.openxmlformats-officedocument.drawing+xml'
const chartContentType = 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml'

describe('xlsx chart artifacts roundtrip', () => {
  it('preserves chart sheets, chart parts, series, and relationships across XLSX round trips', () => {
    const source = buildWorkbookWithChartSheet()

    const imported = importXlsx(source, 'chart-sheet.xlsx')
    const exported = exportXlsx(imported.snapshot)

    expect(imported.snapshot.workbook.metadata?.chartSheetArtifacts).toEqual([
      { name: 'Revenue Chart', relationshipTarget: 'chartsheets/sheet1.xml', sheetId: 2 },
    ])
    expect(imported.snapshot.workbook.metadata?.chartArtifacts?.parts.map((part) => part.path).toSorted()).toEqual([
      'xl/charts/chart1.xml',
      'xl/chartsheets/_rels/sheet1.xml.rels',
      'xl/chartsheets/sheet1.xml',
      'xl/drawings/_rels/drawing1.xml.rels',
      'xl/drawings/drawing1.xml',
    ])
    expect(chartPackageMetrics(exported)).toEqual(chartPackageMetrics(source))
    expect(readZipText(exported, 'xl/chartsheets/sheet1.xml')).toBe(chartSheetXml)
    expect(readZipText(exported, 'xl/charts/chart1.xml')).toBe(chartXml)
  })

  it('does not add a default legend to charts without one', () => {
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: {
        name: 'Chart Workbook',
        metadata: {
          charts: [
            {
              id: 'sales-chart',
              sheetName: 'Data',
              address: 'E1',
              source: { sheetName: 'Data', startAddress: 'A1', endAddress: 'B3' },
              chartType: 'line',
              seriesOrientation: 'columns',
              firstRowAsHeaders: true,
              firstColumnAsLabels: true,
              rows: 12,
              cols: 6,
            },
          ],
        },
      },
      sheets: [
        {
          id: 1,
          name: 'Data',
          order: 0,
          cells: [
            { address: 'A1', value: 'Month' },
            { address: 'B1', value: 'Revenue' },
            { address: 'A2', value: 'Jan' },
            { address: 'B2', value: 10 },
            { address: 'A3', value: 'Feb' },
            { address: 'B3', value: 12 },
          ],
        },
      ],
    }

    const imported = importXlsx(exportXlsx(snapshot), 'chart-workbook.xlsx')

    expect(imported.snapshot.workbook.metadata?.charts).toMatchObject([
      {
        id: 'sales-chart',
        sheetName: 'Data',
        address: 'E1',
        source: { sheetName: 'Data', startAddress: 'A1', endAddress: 'B3' },
        chartType: 'line',
      },
    ])
    expect(imported.snapshot.workbook.metadata?.charts?.[0]?.legendPosition).toBeUndefined()
  })
})

function buildWorkbookWithChartSheet(): Uint8Array {
  const workbook = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(
    workbook,
    XLSX.utils.aoa_to_sheet([
      ['Quarter', 'Revenue'],
      ['Q1', 10],
      ['Q2', 14],
    ]),
    'Data',
  )
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([['chart placeholder']]), 'Revenue Chart')

  const zip = unzipSync(XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' }))
  zip['xl/_rels/workbook.xml.rels'] = strToU8(
    readZipTextFromZip(zip, 'xl/_rels/workbook.xml.rels').replace(/<Relationship\b([^>]*)\/>/gu, (relationshipXml, attributes: string) =>
      readXmlAttribute(attributes, 'Target') === 'worksheets/sheet2.xml'
        ? `<Relationship Id="${readXmlAttribute(attributes, 'Id') ?? 'rId2'}" Type="${chartSheetRelationshipType}" Target="chartsheets/sheet1.xml"/>`
        : relationshipXml,
    ),
  )
  delete zip['xl/worksheets/sheet2.xml']
  delete zip['xl/worksheets/_rels/sheet2.xml.rels']

  zip['xl/chartsheets/sheet1.xml'] = strToU8(chartSheetXml)
  zip['xl/chartsheets/_rels/sheet1.xml.rels'] = strToU8(
    relationshipsXml([{ id: 'rId1', type: drawingRelationshipType, target: '../drawings/drawing1.xml' }]),
  )
  zip['xl/drawings/drawing1.xml'] = strToU8(drawingXml)
  zip['xl/drawings/_rels/drawing1.xml.rels'] = strToU8(
    relationshipsXml([{ id: 'rId1', type: chartRelationshipType, target: '../charts/chart1.xml' }]),
  )
  zip['xl/charts/chart1.xml'] = strToU8(chartXml)
  zip['[Content_Types].xml'] = strToU8(
    removeContentTypeOverride(
      upsertContentTypeOverride(
        upsertContentTypeOverride(
          upsertContentTypeOverride(readZipTextFromZip(zip, '[Content_Types].xml'), {
            partName: '/xl/chartsheets/sheet1.xml',
            contentType: chartSheetContentType,
          }),
          { partName: '/xl/drawings/drawing1.xml', contentType: drawingContentType },
        ),
        { partName: '/xl/charts/chart1.xml', contentType: chartContentType },
      ),
      '/xl/worksheets/sheet2.xml',
    ),
  )

  return zipSync(zip)
}

function chartPackageMetrics(bytes: Uint8Array): {
  chartRelationships: number
  chartSeries: number
  chartSheets: number
  chartSheetRelationships: number
  charts: number
  workbookChartSheetRelationships: number
} {
  const zip = unzipSync(bytes)
  const chartXmlParts = Object.entries(zip)
    .filter(([path]) => /^xl\/charts\/chart\d+\.xml$/u.test(path))
    .map(([, part]) => strFromU8(part))
  const drawingRelationshipsXml = Object.entries(zip)
    .filter(([path]) => /^xl\/drawings\/_rels\/drawing\d+\.xml\.rels$/u.test(path))
    .map(([, part]) => strFromU8(part))
    .join('')
  const chartSheetRelationshipsXml = Object.entries(zip)
    .filter(([path]) => /^xl\/chartsheets\/_rels\/sheet\d+\.xml\.rels$/u.test(path))
    .map(([, part]) => strFromU8(part))
    .join('')
  const workbookRelationshipsXml = readZipTextFromZip(zip, 'xl/_rels/workbook.xml.rels')
  return {
    chartRelationships: relationshipsWithType(drawingRelationshipsXml, chartRelationshipType).length,
    chartSeries: chartXmlParts.reduce((count, xml) => count + (xml.match(/<c:ser\b/gu)?.length ?? 0), 0),
    chartSheets: Object.keys(zip).filter((path) => /^xl\/chartsheets\/sheet\d+\.xml$/u.test(path)).length,
    chartSheetRelationships: relationshipsWithType(chartSheetRelationshipsXml, drawingRelationshipType).length,
    charts: chartXmlParts.length,
    workbookChartSheetRelationships: relationshipsWithType(workbookRelationshipsXml, chartSheetRelationshipType).length,
  }
}

function relationshipsXml(relationships: readonly { id: string; type: string; target: string }[]): string {
  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    `<Relationships xmlns="${relationshipNamespace}">`,
    ...relationships.map(
      (relationship) => `<Relationship Id="${relationship.id}" Type="${relationship.type}" Target="${relationship.target}"/>`,
    ),
    '</Relationships>',
  ].join('')
}

function relationshipsWithType(relationshipsXmlText: string, relationshipType: string): string[] {
  return [...relationshipsXmlText.matchAll(/<Relationship\b([^>]*)\/?>/gu)].flatMap((match) => {
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

function readXmlAttribute(attributes: string, name: string): string | null {
  return new RegExp(`\\b${name}=("|')([\\s\\S]*?)\\1`, 'u').exec(attributes)?.[2] ?? null
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

function removeContentTypeOverride(contentTypesXml: string, partName: string): string {
  return contentTypesXml.replace(new RegExp(`<Override\\b[^>]*\\bPartName="${partName}"[^>]*/>`, 'u'), '')
}

const chartSheetXml = [
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
  `<chartsheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${officeRelationshipNamespace}">`,
  '<sheetViews><sheetView workbookViewId="0"/></sheetViews>',
  '<drawing r:id="rId1"/>',
  '</chartsheet>',
].join('')

const drawingXml = [
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
  '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" ',
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">',
  '<xdr:absoluteAnchor>',
  '<xdr:pos x="0" y="0"/><xdr:ext cx="6000000" cy="4000000"/>',
  '<xdr:graphicFrame macro="">',
  '<xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Revenue Chart"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>',
  '<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>',
  '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">',
  `<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="${officeRelationshipNamespace}" r:id="rId1"/>`,
  '</a:graphicData></a:graphic>',
  '</xdr:graphicFrame><xdr:clientData/>',
  '</xdr:absoluteAnchor>',
  '</xdr:wsDr>',
].join('')

const chartXml = [
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
  `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${officeRelationshipNamespace}">`,
  '<c:chart><c:plotArea><c:layout/><c:lineChart>',
  '<c:grouping val="standard"/>',
  '<c:ser><c:idx val="0"/><c:order val="0"/>',
  '<c:tx><c:strRef><c:f>Data!$B$1</c:f></c:strRef></c:tx>',
  '<c:cat><c:strRef><c:f>Data!$A$2:$A$3</c:f></c:strRef></c:cat>',
  '<c:val><c:numRef><c:f>Data!$B$2:$B$3</c:f></c:numRef></c:val>',
  '</c:ser>',
  '<c:axId val="1"/><c:axId val="2"/>',
  '</c:lineChart>',
  '<c:catAx><c:axId val="1"/><c:crossAx val="2"/></c:catAx>',
  '<c:valAx><c:axId val="2"/><c:crossAx val="1"/></c:valAx>',
  '</c:plotArea></c:chart>',
  '</c:chartSpace>',
].join('')
