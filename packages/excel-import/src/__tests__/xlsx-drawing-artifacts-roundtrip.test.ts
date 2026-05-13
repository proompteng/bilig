import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { describe, expect, it } from 'vitest'
import * as XLSX from 'xlsx'

import type { WorkbookPreservedPackagePartSnapshot, WorkbookSnapshot } from '@bilig/protocol'
import { exportXlsx, importXlsx } from '../index.js'

const worksheetDrawingRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing'
const chartRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart'
const imageRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
const drawingContentType = 'application/vnd.openxmlformats-officedocument.drawing+xml'
const imageContentType = 'image/png'
const relationshipNamespace = 'http://schemas.openxmlformats.org/package/2006/relationships'
const officeRelationshipNamespace = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

describe('worksheet drawing artifacts roundtrip', () => {
  it('preserves embedded images and non-chart drawing parts across XLSX round trips', () => {
    const source = buildWorkbookWithEmbeddedImageAndShape()

    const imported = importXlsx(source, 'drawing-artifacts.xlsx')
    const exported = exportXlsx(imported.snapshot)

    expect(imported.warnings).toEqual([])
    expect(imported.snapshot.workbook.metadata?.drawingArtifacts?.parts.map((part) => part.path).toSorted()).toEqual(
      expect.arrayContaining([
        'xl/drawings/drawing1.xml',
        'xl/drawings/_rels/drawing1.xml.rels',
        'xl/drawings/drawing2.xml',
        'xl/media/image1.png',
      ]),
    )
    expect(imported.snapshot.sheets[0]?.metadata?.drawingArtifacts).toEqual({ relationshipTarget: '../drawings/drawing1.xml' })
    expect(imported.snapshot.sheets[1]?.metadata?.drawingArtifacts).toEqual({ relationshipTarget: '../drawings/drawing2.xml' })
    expect(drawingAndMediaPaths(exported)).toEqual(drawingAndMediaPaths(source))
    expect(strFromU8(unzipSync(exported)['xl/drawings/drawing1.xml'] ?? new Uint8Array())).toContain('<xdr:pic>')
    expect(strFromU8(unzipSync(exported)['xl/drawings/drawing2.xml'] ?? new Uint8Array())).toContain('<xdr:cxnSp')
    expect([...readZipBytes(exported, 'xl/media/image1.png')]).toEqual([137, 80, 78, 71, 13, 10, 26, 10])
  })

  it('keeps preserved drawings separate from generated chart drawings', () => {
    const exported = exportXlsx(buildSnapshotWithImageArtifactAndGeneratedChart())
    const exportedZip = unzipSync(exported)
    const imported = importXlsx(exported, 'drawing-chart-collision.xlsx')
    const coverRelationshipsXml = strFromU8(exportedZip['xl/worksheets/_rels/sheet1.xml.rels'] ?? new Uint8Array())
    const chartRelationshipsXml = strFromU8(exportedZip['xl/worksheets/_rels/sheet2.xml.rels'] ?? new Uint8Array())
    const coverDrawingTarget = relationshipTargetWithType(coverRelationshipsXml, worksheetDrawingRelationshipType)
    const chartDrawingTarget = relationshipTargetWithType(chartRelationshipsXml, worksheetDrawingRelationshipType)

    expect(coverDrawingTarget).not.toBe(chartDrawingTarget)
    expect(strFromU8(exportedZip['xl/drawings/drawing1.xml'] ?? new Uint8Array())).toContain('<c:chart ')
    expect(strFromU8(exportedZip['xl/drawings/drawing1.xml'] ?? new Uint8Array())).not.toContain('<xdr:pic>')
    expect(strFromU8(exportedZip['xl/drawings/drawing2.xml'] ?? new Uint8Array())).toContain('<xdr:pic>')
    expect(strFromU8(exportedZip['xl/drawings/drawing2.xml'] ?? new Uint8Array())).not.toContain('<c:chart ')
    expect(
      relationshipsWithType(strFromU8(exportedZip['xl/drawings/_rels/drawing1.xml.rels'] ?? new Uint8Array()), chartRelationshipType),
    ).toHaveLength(1)
    expect(
      relationshipsWithType(strFromU8(exportedZip['xl/drawings/_rels/drawing2.xml.rels'] ?? new Uint8Array()), imageRelationshipType),
    ).toHaveLength(1)
    expect(imported.snapshot.workbook.metadata?.charts).toEqual([
      expect.objectContaining({
        id: 'Chart 1',
        sheetName: 'Chart',
      }),
    ])
  })
})

function buildSnapshotWithImageArtifactAndGeneratedChart(): WorkbookSnapshot {
  const drawingXml = [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">',
    '<xdr:twoCellAnchor editAs="oneCell">',
    '<xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>',
    '<xdr:to><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>3</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>',
    '<xdr:pic>',
    '<xdr:nvPicPr><xdr:cNvPr id="3" name="Picture 1"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>',
    '<xdr:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>',
    '<xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr>',
    '</xdr:pic>',
    '<xdr:clientData/>',
    '</xdr:twoCellAnchor>',
    '</xdr:wsDr>',
  ].join('')
  const drawingRelationshipsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${relationshipNamespace}"><Relationship Id="rId1" Type="${imageRelationshipType}" Target="../media/image1.png"/></Relationships>`
  const imageBytes = Uint8Array.from([137, 80, 78, 71, 13, 10, 26, 10])

  return {
    version: 1,
    workbook: {
      name: 'drawing-chart-collision',
      metadata: {
        charts: [
          {
            id: 'Chart 1',
            sheetName: 'Chart',
            address: 'E2',
            source: { sheetName: 'Chart', startAddress: 'A1', endAddress: 'B3' },
            chartType: 'line',
            rows: 8,
            cols: 5,
            seriesOrientation: 'columns',
            firstRowAsHeaders: true,
            firstColumnAsLabels: true,
          },
        ],
        drawingArtifacts: {
          parts: [
            encodedPart('xl/drawings/drawing1.xml', drawingXml),
            encodedPart('xl/drawings/_rels/drawing1.xml.rels', drawingRelationshipsXml),
            encodedPart('xl/media/image1.png', imageBytes),
          ],
          contentTypeDefaults: [{ extension: 'png', contentType: imageContentType }],
          contentTypeOverrides: [{ partName: '/xl/drawings/drawing1.xml', contentType: drawingContentType }],
        },
      },
    },
    sheets: [
      {
        id: 1,
        name: 'Cover',
        order: 0,
        metadata: {
          drawingArtifacts: { relationshipTarget: '../drawings/drawing1.xml' },
        },
        cells: [{ address: 'A1', value: 'logo' }],
      },
      {
        id: 2,
        name: 'Chart',
        order: 1,
        cells: [
          { address: 'A1', value: 'Year' },
          { address: 'B1', value: 'Files' },
          { address: 'A2', value: '2024' },
          { address: 'B2', value: 12 },
          { address: 'A3', value: '2025' },
          { address: 'B3', value: 15 },
        ],
      },
    ],
  }
}

function encodedPart(path: string, value: string | Uint8Array): WorkbookPreservedPackagePartSnapshot {
  const bytes = typeof value === 'string' ? new TextEncoder().encode(value) : value
  return {
    path,
    storage: 'base64',
    dataBase64: Buffer.from(bytes).toString('base64'),
    byteLength: bytes.byteLength,
  }
}

function relationshipTargetWithType(xml: string, relationshipType: string): string | undefined {
  return [...xml.matchAll(/<Relationship\b([^>]*)\/?>/gu)]
    .map((match) => match[1] ?? '')
    .find((attributes) => attributes.includes(`Type="${relationshipType}"`))
    ?.match(/\bTarget="([^"]+)"/u)?.[1]
}

function relationshipsWithType(xml: string, relationshipType: string): string[] {
  return [...xml.matchAll(/<Relationship\b([^>]*)\/?>/gu)]
    .map((match) => match[1] ?? '')
    .filter((attributes) => attributes.includes(`Type="${relationshipType}"`))
}

function buildWorkbookWithEmbeddedImageAndShape(): Uint8Array {
  const workbook = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([['logo']]), 'Cover')
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([['connector']]), 'Flows')

  const zip = unzipSync(XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' }))
  addWorksheetDrawing(zip, {
    sheetIndex: 1,
    relationshipId: 'rIdDrawing1',
    target: '../drawings/drawing1.xml',
    drawingXml: [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">',
      '<xdr:twoCellAnchor editAs="oneCell">',
      '<xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>',
      '<xdr:to><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>3</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>',
      '<xdr:pic>',
      '<xdr:nvPicPr><xdr:cNvPr id="3" name="Picture 1"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>',
      '<xdr:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>',
      '<xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr>',
      '</xdr:pic>',
      '<xdr:clientData/>',
      '</xdr:twoCellAnchor>',
      '</xdr:wsDr>',
    ].join(''),
    relationshipsXml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${relationshipNamespace}"><Relationship Id="rId1" Type="${imageRelationshipType}" Target="../media/image1.png"/></Relationships>`,
  })
  addWorksheetDrawing(zip, {
    sheetIndex: 2,
    relationshipId: 'rIdDrawing2',
    target: '../drawings/drawing2.xml',
    drawingXml: [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">',
      '<xdr:twoCellAnchor>',
      '<xdr:from><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>',
      '<xdr:to><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>8</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>',
      '<xdr:cxnSp macro="">',
      '<xdr:nvCxnSpPr><xdr:cNvPr id="5" name="Straight Connector 1"/><xdr:cNvCxnSpPr/></xdr:nvCxnSpPr>',
      '<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm><a:prstGeom prst="line"><a:avLst/></a:prstGeom></xdr:spPr>',
      '</xdr:cxnSp>',
      '<xdr:clientData/>',
      '</xdr:twoCellAnchor>',
      '</xdr:wsDr>',
    ].join(''),
  })

  zip['xl/media/image1.png'] = Uint8Array.from([137, 80, 78, 71, 13, 10, 26, 10])
  const contentTypesXml = strFromU8(zip['[Content_Types].xml'] ?? new Uint8Array())
  zip['[Content_Types].xml'] = strToU8(
    contentTypesXml
      .replace('</Types>', `<Override PartName="/xl/drawings/drawing1.xml" ContentType="${drawingContentType}"/></Types>`)
      .replace('</Types>', `<Override PartName="/xl/drawings/drawing2.xml" ContentType="${drawingContentType}"/></Types>`)
      .replace('</Types>', `<Default Extension="png" ContentType="${imageContentType}"/></Types>`),
  )

  return zipSync(zip)
}

function addWorksheetDrawing(
  zip: Record<string, Uint8Array>,
  input: {
    readonly sheetIndex: number
    readonly relationshipId: string
    readonly target: string
    readonly drawingXml: string
    readonly relationshipsXml?: string
  },
): void {
  const sheetPath = `xl/worksheets/sheet${String(input.sheetIndex)}.xml`
  const relsPath = `xl/worksheets/_rels/sheet${String(input.sheetIndex)}.xml.rels`
  const drawingPath = `xl/drawings/drawing${String(input.sheetIndex)}.xml`
  const drawingRelsPath = `xl/drawings/_rels/drawing${String(input.sheetIndex)}.xml.rels`
  const sheetXml = ensureOfficeRelationshipNamespace(strFromU8(zip[sheetPath] ?? new Uint8Array()))
  zip[sheetPath] = strToU8(sheetXml.replace('</worksheet>', `<drawing r:id="${input.relationshipId}"/></worksheet>`))
  zip[relsPath] = strToU8(
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${relationshipNamespace}"><Relationship Id="${input.relationshipId}" Type="${worksheetDrawingRelationshipType}" Target="${input.target}"/></Relationships>`,
  )
  zip[drawingPath] = strToU8(input.drawingXml)
  if (input.relationshipsXml) {
    zip[drawingRelsPath] = strToU8(input.relationshipsXml)
  }
}

function ensureOfficeRelationshipNamespace(sheetXml: string): string {
  return /\sxmlns:r=("|')[\s\S]*?\1/u.test(sheetXml)
    ? sheetXml
    : sheetXml.replace(/<worksheet\b([^>]*)>/u, `<worksheet$1 xmlns:r="${officeRelationshipNamespace}">`)
}

function drawingAndMediaPaths(bytes: Uint8Array): string[] {
  return Object.keys(unzipSync(bytes))
    .filter((path) => path.startsWith('xl/drawings/') || path.startsWith('xl/media/'))
    .toSorted()
}

function readZipBytes(bytes: Uint8Array, path: string): Uint8Array {
  return unzipSync(bytes)[path] ?? new Uint8Array()
}
