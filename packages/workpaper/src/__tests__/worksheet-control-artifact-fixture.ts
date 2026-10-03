import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { exportXlsx } from '@bilig/excel-import'
import type { WorkbookSnapshot } from '@bilig/protocol'

const relationshipNamespace = 'http://schemas.openxmlformats.org/package/2006/relationships'
const officeRelationshipNamespace = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const drawingMlSpreadsheetNamespace = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'
const markupCompatibilityNamespace = 'http://schemas.openxmlformats.org/markup-compatibility/2006'
const spreadsheetControlNamespace = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main'
const vmlDrawingRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing'
const controlRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/ctrlProp'
const vmlDrawingContentType = 'application/vnd.openxmlformats-officedocument.vmlDrawing'
const controlPropertiesContentType = 'application/vnd.ms-excel.controlproperties+xml'

export function buildWorkbookWithWorksheetControl(): Uint8Array {
  const zip = unzipSync(exportXlsx(buildWorkbook()))
  zip['xl/worksheets/sheet1.xml'] = strToU8(addWorksheetControlXml(readZipTextFromZip(zip, 'xl/worksheets/sheet1.xml')))
  zip['xl/worksheets/_rels/sheet1.xml.rels'] = strToU8(
    [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      `<Relationships xmlns="${relationshipNamespace}">`,
      `<Relationship Id="rId3" Type="${vmlDrawingRelationshipType}" Target="../drawings/vmlDrawing1.vml"/>`,
      `<Relationship Id="rId4" Type="${controlRelationshipType}" Target="../ctrlProps/ctrlProp1.xml"/>`,
      '</Relationships>',
    ].join(''),
  )
  zip['xl/ctrlProps/ctrlProp1.xml'] = strToU8(controlPropertiesXml)
  zip['xl/drawings/vmlDrawing1.vml'] = strToU8(vmlDrawingXml)
  zip['[Content_Types].xml'] = strToU8(
    upsertContentTypeOverride(addContentTypeDefault(readZipTextFromZip(zip, '[Content_Types].xml'), 'vml', vmlDrawingContentType), {
      contentType: controlPropertiesContentType,
      partName: '/xl/ctrlProps/ctrlProp1.xml',
    }),
  )
  return zipSync(zip)
}

function buildWorkbook(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: { name: 'Worksheet control' },
    sheets: [
      {
        id: 1,
        name: 'Model',
        order: 0,
        cells: [{ address: 'A1', value: 'Control fixture' }],
      },
    ],
  }
}

function addWorksheetControlXml(sheetXml: string): string {
  return withWorksheetControlNamespaces(sheetXml).replace('</worksheet>', `${worksheetControlsXml}</worksheet>`)
}

function withWorksheetControlNamespaces(sheetXml: string): string {
  return sheetXml.replace(/<worksheet\b[^>]*>/u, (rootOpenTag) => {
    let output = rootOpenTag
    output = upsertXmlAttribute(output, 'xmlns:r', officeRelationshipNamespace)
    output = upsertXmlAttribute(output, 'xmlns:xdr', drawingMlSpreadsheetNamespace)
    output = upsertXmlAttribute(output, 'xmlns:x14', spreadsheetControlNamespace)
    output = upsertXmlAttribute(output, 'xmlns:mc', markupCompatibilityNamespace)
    const ignorable = readXmlAttribute(output, 'mc:Ignorable')
    if (!ignorable) {
      return upsertXmlAttribute(output, 'mc:Ignorable', 'x14')
    }
    if (ignorable.split(/\s+/u).includes('x14')) {
      return output
    }
    return output.replace(/\smc:Ignorable=(["'])([\s\S]*?)\1/u, ` mc:Ignorable="x14 ${ignorable}"`)
  })
}

function upsertXmlAttribute(rootOpenTag: string, name: string, value: string): string {
  if (new RegExp(`\\s${escapeRegExp(name)}=("|')`, 'u').test(rootOpenTag)) {
    return rootOpenTag
  }
  return rootOpenTag.replace(/>$/u, ` ${name}="${value}">`)
}

function readXmlAttribute(xml: string, name: string): string | null {
  return new RegExp(`\\s${escapeRegExp(name)}=("|')([\\s\\S]*?)\\1`, 'u').exec(xml)?.[2] ?? null
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/gu, '\\$&')
}

function readZipTextFromZip(zip: Record<string, Uint8Array>, path: string): string {
  const bytes = zip[path]
  if (!bytes) {
    throw new Error(`Missing XLSX part: ${path}`)
  }
  return strFromU8(bytes)
}

function addContentTypeDefault(contentTypesXml: string, extension: string, contentType: string): string {
  if (new RegExp(`<Default\\b[^>]*\\bExtension=(["'])${extension}\\1`, 'u').test(contentTypesXml)) {
    return contentTypesXml
  }
  return contentTypesXml.replace('</Types>', `<Default Extension="${extension}" ContentType="${contentType}"/></Types>`)
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

const worksheetControlsXml = [
  '<legacyDrawing r:id="rId3"/>',
  `<mc:AlternateContent xmlns:mc="${markupCompatibilityNamespace}">`,
  '<mc:Choice Requires="x14">',
  '<controls>',
  `<mc:AlternateContent xmlns:mc="${markupCompatibilityNamespace}">`,
  '<mc:Choice Requires="x14">',
  '<control shapeId="1025" r:id="rId4" name="Button 1">',
  '<controlPr defaultSize="0" print="0" autoFill="0" autoPict="0" macro="[0]!WriteHelloWorld">',
  '<anchor moveWithCells="1">',
  '<from><xdr:col>6</xdr:col><xdr:colOff>142875</xdr:colOff><xdr:row>3</xdr:row><xdr:rowOff>66675</xdr:rowOff></from>',
  '<to><xdr:col>11</xdr:col><xdr:colOff>85725</xdr:colOff><xdr:row>4</xdr:row><xdr:rowOff>152400</xdr:rowOff></to>',
  '</anchor>',
  '</controlPr>',
  '</control>',
  '</mc:Choice>',
  '</mc:AlternateContent>',
  '</controls>',
  '</mc:Choice>',
  '</mc:AlternateContent>',
].join('')

const controlPropertiesXml = [
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
  `<formControlPr xmlns="${spreadsheetControlNamespace}" objectType="Button" lockText="1"/>`,
].join('')

const vmlDrawingXml = [
  '<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" ',
  'xmlns:x="urn:schemas-microsoft-com:office:excel">',
  '<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>',
  '<v:shapetype id="_x0000_t201" coordsize="21600,21600" o:spt="201" path="m,l,21600r21600,l21600,xe">',
  '<v:stroke joinstyle="miter"/>',
  '<v:path shadowok="f" o:extrusionok="f" strokeok="f" fillok="f" o:connecttype="rect"/>',
  '<o:lock v:ext="edit" shapetype="t"/>',
  '</v:shapetype>',
  '<v:shape id="_x0000_s1025" type="#_x0000_t201" style="position:absolute;margin-left:299.25pt;margin-top:50.25pt;width:235.5pt;height:21.75pt;z-index:1;mso-wrap-style:tight" ',
  'o:button="t" fillcolor="buttonFace [67]" strokecolor="windowText [64]" o:insetmode="auto">',
  '<v:fill color2="buttonFace [67]" o:detectmouseclick="t"/>',
  '<o:lock v:ext="edit" rotation="t"/>',
  '<v:textbox style="mso-direction-alt:auto" o:singleclick="f"><div style="text-align:center">Press Me</div></v:textbox>',
  '<x:ClientData ObjectType="Button"><x:SizeWithCells/><x:Anchor>6, 11, 3, 5, 11, 7, 4, 12</x:Anchor>',
  '<x:PrintObject>False</x:PrintObject><x:AutoFill>False</x:AutoFill>',
  '<x:FmlaMacro>[0]!WriteHelloWorld</x:FmlaMacro><x:TextHAlign>Center</x:TextHAlign><x:TextVAlign>Center</x:TextVAlign>',
  '</x:ClientData>',
  '</v:shape>',
  '</xml>',
].join('')
