import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { describe, expect, it } from 'vitest'

import type { CellStyleRecord, WorkbookSnapshot } from '@bilig/protocol'
import { exportXlsx, importXlsx } from '../index.js'

describe('xlsx cell style roundtrip', () => {
  it('preserves visible cell styles when cell formats omit apply flags', () => {
    const source = removeApplyStyleFlags(exportXlsx(buildStyledWorkbook()))

    const imported = importXlsx(source, 'implicit-style-components.xlsx')

    expect(readFirstAppliedStyle(imported.snapshot)).toMatchObject(expectedVisibleStyle)

    const reimported = importXlsx(exportXlsx(imported.snapshot), 'implicit-style-components-roundtrip.xlsx')

    expect(readFirstAppliedStyle(reimported.snapshot)).toMatchObject(expectedVisibleStyle)
  })

  it('preserves raw theme and indexed style references for unchanged imported cells', () => {
    const source = buildRawStyleReferenceWorkbook()

    const exported = exportXlsx(importXlsx(source, 'raw-style-references.xlsx').snapshot)

    expect(readCellStyleParts(exported, 'xl/worksheets/sheet1.xml!A1')).toEqual(readCellStyleParts(source, 'xl/worksheets/sheet1.xml!A1'))
  })
})

const expectedVisibleStyle = {
  fill: { backgroundColor: '#1d3989' },
  font: {
    bold: true,
    italic: true,
    underline: true,
    color: '#ffffff',
  },
  borders: {
    top: { style: 'solid', weight: 'thin', color: '#808080' },
    right: { style: 'solid', weight: 'thin', color: '#808080' },
    bottom: { style: 'solid', weight: 'thin', color: '#808080' },
    left: { style: 'solid', weight: 'thin', color: '#808080' },
  },
} satisfies Partial<CellStyleRecord>

function buildStyledWorkbook(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: {
      name: 'Implicit style components',
      metadata: {
        styles: [
          {
            id: 'visible-review-style',
            fill: { backgroundColor: '#1d3989' },
            font: {
              bold: true,
              italic: true,
              underline: true,
              color: '#ffffff',
            },
            borders: {
              top: { style: 'solid', weight: 'thin', color: '#808080' },
              right: { style: 'solid', weight: 'thin', color: '#808080' },
              bottom: { style: 'solid', weight: 'thin', color: '#808080' },
              left: { style: 'solid', weight: 'thin', color: '#808080' },
            },
          },
        ],
      },
    },
    sheets: [
      {
        id: 1,
        name: 'Review',
        order: 0,
        cells: [{ address: 'A1', value: 'Input' }],
        metadata: {
          styleRanges: [
            {
              range: { sheetName: 'Review', startAddress: 'A1', endAddress: 'A1' },
              styleId: 'visible-review-style',
            },
          ],
        },
      },
    ],
  }
}

function removeApplyStyleFlags(bytes: Uint8Array): Uint8Array {
  const zip = unzipSync(bytes)
  const stylesXml = strFromU8(zip['xl/styles.xml'] ?? new Uint8Array())
  zip['xl/styles.xml'] = strToU8(stylesXml.replace(/\sapply(?:Font|Fill|Border)="1"/gu, ''))
  return zipSync(zip)
}

function readFirstAppliedStyle(snapshot: WorkbookSnapshot): CellStyleRecord | undefined {
  const styleRange = snapshot.sheets[0]?.metadata?.styleRanges?.[0]
  return snapshot.workbook.metadata?.styles?.find((style) => style.id === styleRange?.styleId)
}

function buildRawStyleReferenceWorkbook(): Uint8Array {
  const zip = unzipSync(exportXlsx(buildStyledWorkbook()))
  zip['xl/styles.xml'] = strToU8(rawStyleReferenceStylesXml)
  const sheetXml = strFromU8(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
  zip['xl/worksheets/sheet1.xml'] = strToU8(sheetXml.replace(/<c\b(?=[^>]*\br="A1")[^>]*>/u, (tag) => setXmlAttribute(tag, 's', '1')))
  return zipSync(zip)
}

function readCellStyleParts(bytes: Uint8Array, cellRef: string): { border: string; fill: string; font: string } {
  const [sheetPath, address] = cellRef.split('!')
  const zip = unzipSync(bytes)
  const stylesXml = strFromU8(zip['xl/styles.xml'] ?? new Uint8Array())
  const sheetXml = strFromU8(zip[sheetPath ?? ''] ?? new Uint8Array())
  const styleId = new RegExp(`<c\\b(?=[^>]*\\br="${address ?? ''}")[^>]*\\bs="([^"]+)"`, 'u').exec(sheetXml)?.[1]
  const xf = listElements(stylesXml, 'cellXfs', 'xf')[Number(styleId ?? '0')] ?? ''
  const fontId = Number(readXmlAttribute(xf, 'fontId') ?? 0)
  const fillId = Number(readXmlAttribute(xf, 'fillId') ?? 0)
  const borderId = Number(readXmlAttribute(xf, 'borderId') ?? 0)
  return {
    border: normalizeBorder(listElements(stylesXml, 'borders', 'border')[borderId] ?? ''),
    fill: normalizeFill(listElements(stylesXml, 'fills', 'fill')[fillId] ?? ''),
    font: normalizeFont(listElements(stylesXml, 'fonts', 'font')[fontId] ?? ''),
  }
}

function listElements(xml: string, parent: string, tag: string): string[] {
  const section = new RegExp(`<${parent}\\b[^>]*>([\\s\\S]*?)<\\/${parent}>`, 'u').exec(xml)?.[1] ?? ''
  return [...section.matchAll(new RegExp(`<${tag}\\b[^>]*(?:\\/>|>[\\s\\S]*?<\\/${tag}>)`, 'gu'))].map((match) =>
    match[0].replace(/\s+/gu, ' '),
  )
}

function normalizeFill(fill: string): string {
  return [
    `pattern=${readXmlAttribute(firstTag(fill, 'patternFill'), 'patternType') || 'solid'}`,
    firstTag(fill, 'fgColor'),
    firstTag(fill, 'bgColor'),
    firstTag(fill, 'gradientFill'),
  ].join('|')
}

function normalizeBorder(border: string): string {
  return ['left', 'right', 'top', 'bottom', 'diagonal']
    .map((edge) => {
      const edgeXml = firstTag(border, edge)
      const style = readXmlAttribute(edgeXml, 'style')
      return style ? `${edge}:${style}:${firstTag(edgeXml, 'color')}` : ''
    })
    .filter(Boolean)
    .join('|')
}

function normalizeFont(font: string): string {
  const parts = [
    /<b\b/u.test(font) ? 'bold' : '',
    /<i\b/u.test(font) ? 'italic' : '',
    /<u\b/u.test(font) ? 'underline' : '',
    /<strike\b/u.test(font) ? 'strike' : '',
  ].filter(Boolean)
  const color = firstTag(font, 'color')
  if (/\brgb="/u.test(color)) {
    parts.push(color)
  }
  return parts.join('|')
}

function firstTag(xml: string, tag: string): string {
  return new RegExp(`<${tag}\\b[^>]*(?:\\/>|>[\\s\\S]*?<\\/${tag}>)`, 'u').exec(xml)?.[0] ?? ''
}

function readXmlAttribute(xml: string, name: string): string | undefined {
  return new RegExp(`\\b${name}="([^"]*)"`, 'u').exec(xml)?.[1]
}

function setXmlAttribute(xml: string, name: string, value: string): string {
  if (new RegExp(`\\b${name}=`, 'u').test(xml)) {
    return xml.replace(new RegExp(`\\b${name}="[^"]*"`, 'u'), `${name}="${value}"`)
  }
  return xml.replace(/\/?>$/u, (suffix) => ` ${name}="${value}"${suffix}`)
}

const rawStyleReferenceStylesXml = [
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
  '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
  '<fonts count="2"><font><sz val="11"/><name val="Aptos"/></font><font><i/><u/><color rgb="FF0000FF"/></font></fonts>',
  '<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor theme="0"/><bgColor rgb="FF000000"/></patternFill></fill></fills>',
  '<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left style="thin"><color indexed="64"/></left><right style="thin"><color theme="0"/></right><top style="thin"><color theme="0"/></top><bottom style="thin"><color indexed="64"/></bottom><diagonal/></border></borders>',
  '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>',
  '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0"/></cellXfs>',
  '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>',
  '</styleSheet>',
].join('')
