import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { describe, expect, it } from 'vitest'

import { SpreadsheetEngine } from '@bilig/core'
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

  it('reads imported cell styles through workbook sheet relationships', () => {
    const imported = importXlsx(buildRelationshipMappedStyleWorkbook(), 'relationship-mapped-styles.xlsx')

    expect(readAppliedStyle(imported.snapshot, 'First', 'A1')).toMatchObject(expectedHeaderStyle)
    expect(readAppliedStyle(imported.snapshot, 'Second', 'A1')).toBeUndefined()
  })

  it('reads imported cell styles for relationship mapped sheets with trailing spaces', () => {
    const imported = importXlsx(buildTrailingSpaceRelationshipMappedStyleWorkbook(), 'trailing-space-relationship-mapped-styles.xlsx')

    expect(readAppliedStyle(imported.snapshot, 'First ', 'A1')).toMatchObject(expectedHeaderStyle)
    expect(readAppliedStyle(imported.snapshot, 'Second', 'A1')).toBeUndefined()
  })

  it('preserves row and column default style indexes for unchanged imported cells', () => {
    const source = buildAxisStyleReferenceWorkbook()

    const exported = exportXlsx(importXlsx(source, 'axis-style-references.xlsx').snapshot)

    expect(readColumnAttributes(exported, 1)).toMatchObject(readColumnAttributes(source, 1))
    expect(readRowAttributes(exported, 1)).toMatchObject(readRowAttributes(source, 1))
    expect(readRowAttributes(exported, 2)).toMatchObject(readRowAttributes(source, 2))
    expect(readColumnStyleNumberFormat(exported, 1)).toBe(readColumnStyleNumberFormat(source, 1))
    expect(readRowStyleNumberFormat(exported, 1)).toBe(readRowStyleNumberFormat(source, 1))
    expect(readCellXml(exported, 'xl/worksheets/sheet1.xml!A3')).toBeDefined()
    expect(readCellXml(exported, 'xl/worksheets/sheet1.xml!B2')).toBeDefined()
  })

  it('keeps repeated row style metadata compact while preserving export fidelity', () => {
    const source = buildRepeatedRowStyleMetadataWorkbook()

    const imported = importXlsx(source, 'repeated-row-style-metadata.xlsx')
    const metadata = imported.snapshot.sheets[0]?.metadata

    expect(metadata?.rows).toBeUndefined()
    expect(metadata?.rowMetadata).toEqual([
      {
        start: 0,
        count: repeatedRowStyleMetadataCount,
        styleIndex: 1,
        customFormat: true,
      },
    ])

    const engine = new SpreadsheetEngine({ workbookName: 'repeated-row-style-metadata' })
    expect(() => engine.importSnapshot(imported.snapshot)).not.toThrow()

    const exported = exportXlsx(imported.snapshot)
    expect(readRowAttributes(exported, 1)).toEqual({ style: '1', customFormat: '1' })
    expect(readRowAttributes(exported, repeatedRowStyleMetadataCount)).toEqual({ style: '1', customFormat: '1' })
  })

  it('preserves style-only blank cells on otherwise empty imported worksheets', () => {
    const source = buildStyleOnlyBlankCellWorkbook()

    const exported = exportXlsx(importXlsx(source, 'style-only-blank-cells.xlsx').snapshot)

    expect(readCellStyleParts(exported, 'xl/worksheets/sheet1.xml!A2')).toEqual(readCellStyleParts(source, 'xl/worksheets/sheet1.xml!A2'))
    expect(readCellStyleParts(exported, 'xl/worksheets/sheet1.xml!B2')).toEqual(readCellStyleParts(source, 'xl/worksheets/sheet1.xml!B2'))
  })

  it('exports many imported style-only blank cells with bounded memory', () => {
    const imported = importXlsx(buildManyStyleOnlyBlankCellsWorkbook(), 'many-style-only-blank-cells.xlsx')
    collectGarbage()
    const beforeRss = process.memoryUsage().rss
    const start = performance.now()

    const exported = exportXlsx(imported.snapshot)
    const durationMs = performance.now() - start
    collectGarbage()
    const rssDelta = process.memoryUsage().rss - beforeRss

    expect(imported.snapshot.sheets[0]?.cells).toHaveLength(manyStyleOnlyBlankCellRowCount * 2)
    expect(readCellXml(exported, `xl/worksheets/sheet1.xml!B${String(manyStyleOnlyBlankCellRowCount)}`)).toMatch(
      /^<c\b(?=[^>]*\br="B6000")(?=[^>]*\bs="\d+")[^>]*\/>$/u,
    )
    expect(durationMs).toBeLessThan(3_000 * readBenchmarkTolerance())
    expect(rssDelta).toBeLessThan(512 * 1024 * 1024)
  }, 15_000)

  it('exports inserted formatted blank cells as style-only blank cells', () => {
    const exported = exportXlsx(buildBlankFormattedCellWorkbook())
    const cellXml = readCellXml(exported, 'xl/worksheets/sheet1.xml!B2')

    expect(cellXml).toMatch(/^<c\b(?=[^>]*\br="B2")(?=[^>]*\bs="\d+")[^>]*\/>$/u)
    expect(cellXml).not.toContain('t="z"')
  })

  it('imports self-closing formatted blank cells as blank format cells', () => {
    const zip = unzipSync(exportXlsx(buildBlankFormattedCellWorkbook()))
    const sheetXml = strFromU8(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())

    const imported = importXlsx(zipSync(zip), 'self-closing-blank-format.xlsx')

    expect(imported.snapshot.sheets[0]?.cells.find((cell) => cell.address === 'B2')).toEqual({
      address: 'B2',
      format: 'dd/mm/yyyy;@',
    })
    expect(sheetXml).toMatch(/<c\b(?=[^>]*\br="B2")(?=[^>]*\bs="\d+")[^>]*\/>/u)
  })

  it('imports explicit empty formatted blank cells as blank format cells', () => {
    const zip = unzipSync(exportXlsx(buildBlankFormattedCellWorkbook()))
    const sheetXml = strFromU8(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
    zip['xl/worksheets/sheet1.xml'] = strToU8(sheetXml.replace(/<c r="B2" s="([^"]+)"\/>/u, '<c r="B2" s="$1" t="z"></c>'))

    const imported = importXlsx(zipSync(zip), 'empty-blank-format.xlsx')

    expect(imported.snapshot.sheets[0]?.cells.find((cell) => cell.address === 'B2')).toEqual({
      address: 'B2',
      format: 'dd/mm/yyyy;@',
    })
  })

  it('keeps literal numeric entity text across export round trips', () => {
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: { name: 'literal-entity-text' },
      sheets: [
        {
          id: 1,
          name: 'Data',
          order: 0,
          cells: [{ address: 'A1', value: 'RIS maintenance &#8211; year 5' }],
        },
      ],
    }

    const imported = importXlsx(exportXlsx(snapshot), 'literal-entity-text.xlsx')

    expect(imported.snapshot.sheets[0]?.cells.find((cell) => cell.address === 'A1')).toEqual({
      address: 'A1',
      value: 'RIS maintenance &#8211; year 5',
    })
  })

  it('keeps multiline text normalized across export round trips', () => {
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: { name: 'multiline-text-roundtrip' },
      sheets: [
        {
          id: 1,
          name: 'Data',
          order: 0,
          cells: [{ address: 'A1', value: 'Line 1\r\nLine 2\r\nLine 3' }],
        },
      ],
    }

    const imported = importXlsx(exportXlsx(snapshot), 'multiline-text-roundtrip.xlsx')

    expect(imported.snapshot.sheets[0]?.cells.find((cell) => cell.address === 'A1')).toEqual({
      address: 'A1',
      value: 'Line 1\nLine 2\nLine 3',
    })
  })

  it('keeps semantic style ranges valid when raw style artifacts are restored', () => {
    const exported = exportXlsx(buildPartiallyIndexedStyleArtifactWorkbook())
    const imported = importXlsx(exported, 'partial-style-artifacts.xlsx')

    expect(readAppliedStyle(imported.snapshot, 'Report', 'B1')).toMatchObject(expectedHeaderStyle)
    expect(readAppliedStyle(imported.snapshot, 'Report', 'C1')).toMatchObject(expectedHeaderStyle)
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

const expectedHeaderStyle = {
  font: { bold: true },
  alignment: { horizontal: 'center', vertical: 'middle' },
  borders: {
    top: { style: 'solid', weight: 'thin', color: '#000000' },
    right: { style: 'solid', weight: 'thin', color: '#000000' },
    bottom: { style: 'solid', weight: 'thin', color: '#000000' },
    left: { style: 'solid', weight: 'thin', color: '#000000' },
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

function buildBlankFormattedCellWorkbook(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: { name: 'blank-format' },
    sheets: [
      {
        id: 1,
        name: 'Data',
        order: 0,
        cells: [
          { address: 'A1', value: 'Total' },
          { address: 'B2', format: 'dd/mm/yyyy;@' },
        ],
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

function readAppliedStyle(snapshot: WorkbookSnapshot, sheetName: string, address: string): CellStyleRecord | undefined {
  const sheet = snapshot.sheets.find((entry) => entry.name === sheetName)
  const styleRange = sheet?.metadata?.styleRanges?.find(
    (entry) => entry.range.sheetName === sheetName && rangeContainsAddress(entry.range.startAddress, entry.range.endAddress, address),
  )
  return snapshot.workbook.metadata?.styles?.find((style) => style.id === styleRange?.styleId)
}

function rangeContainsAddress(startAddress: string, endAddress: string, address: string): boolean {
  const start = parseA1Address(startAddress)
  const end = parseA1Address(endAddress)
  const target = parseA1Address(address)
  return (
    target.row >= Math.min(start.row, end.row) &&
    target.row <= Math.max(start.row, end.row) &&
    target.col >= Math.min(start.col, end.col) &&
    target.col <= Math.max(start.col, end.col)
  )
}

function parseA1Address(address: string): { readonly row: number; readonly col: number } {
  const match = /^([A-Z]+)([1-9][0-9]*)$/u.exec(address)
  if (!match) {
    throw new Error(`Invalid A1 address: ${address}`)
  }
  const letters = match[1]
  let col = 0
  for (let index = 0; index < letters.length; index += 1) {
    col = col * 26 + letters.charCodeAt(index) - 64
  }
  col -= 1
  return { row: Number(match[2]) - 1, col }
}

function buildRawStyleReferenceWorkbook(): Uint8Array {
  const zip = unzipSync(exportXlsx(buildStyledWorkbook()))
  zip['xl/styles.xml'] = strToU8(rawStyleReferenceStylesXml)
  const sheetXml = strFromU8(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
  zip['xl/worksheets/sheet1.xml'] = strToU8(sheetXml.replace(/<c\b(?=[^>]*\br="A1")[^>]*>/u, (tag) => setXmlAttribute(tag, 's', '1')))
  return zipSync(zip)
}

function buildRelationshipMappedStyleWorkbook(): Uint8Array {
  const zip = unzipSync(exportXlsx(buildTwoSheetWorkbook()))
  zip['xl/styles.xml'] = strToU8(headerStyleReferenceStylesXml)
  const firstSheetXml = strFromU8(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array()).replace(/<c\b(?=[^>]*\br="A1")[^>]*>/u, (tag) =>
    setXmlAttribute(tag, 's', '1'),
  )
  const relationshipsXml = strFromU8(zip['xl/_rels/workbook.xml.rels'] ?? new Uint8Array()).replace(
    /<Relationship\b(?=[^>]*\bId="rId1")[^>]*\/>/u,
    (tag) => setXmlAttribute(tag, 'Target', 'worksheets/sheet7.xml'),
  )
  zip['xl/_rels/workbook.xml.rels'] = strToU8(relationshipsXml)
  delete zip['xl/worksheets/sheet1.xml']
  zip['xl/worksheets/sheet7.xml'] = strToU8(firstSheetXml)

  const reordered: Record<string, Uint8Array> = {}
  for (const path of ['xl/worksheets/sheet2.xml', 'xl/worksheets/sheet7.xml']) {
    const entry = zip[path]
    if (entry) {
      reordered[path] = entry
    }
  }
  for (const [path, entry] of Object.entries(zip)) {
    if (!(path in reordered)) {
      reordered[path] = entry
    }
  }
  return zipSync(reordered)
}

function buildTrailingSpaceRelationshipMappedStyleWorkbook(): Uint8Array {
  const zip = unzipSync(exportXlsx(buildTwoSheetWorkbookWithTrailingSpaceName()))
  zip['xl/styles.xml'] = strToU8(headerStyleReferenceStylesXml)
  const firstSheetXml = strFromU8(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array()).replace(/<c\b(?=[^>]*\br="A1")[^>]*>/u, (tag) =>
    setXmlAttribute(tag, 's', '1'),
  )
  const secondSheetXml = strFromU8(zip['xl/worksheets/sheet2.xml'] ?? new Uint8Array())
  const relationshipsXml = strFromU8(zip['xl/_rels/workbook.xml.rels'] ?? new Uint8Array())
    .replace(/<Relationship\b(?=[^>]*\bId="rId1")[^>]*\/>/u, (tag) => setXmlAttribute(tag, 'Target', 'worksheets/sheet7.xml'))
    .replace(/<Relationship\b(?=[^>]*\bId="rId2")[^>]*\/>/u, (tag) => setXmlAttribute(tag, 'Target', 'worksheets/sheet1.xml'))
  zip['xl/_rels/workbook.xml.rels'] = strToU8(relationshipsXml)
  delete zip['xl/worksheets/sheet2.xml']
  zip['xl/worksheets/sheet1.xml'] = strToU8(secondSheetXml)
  zip['xl/worksheets/sheet7.xml'] = strToU8(firstSheetXml)

  const reordered: Record<string, Uint8Array> = {}
  for (const path of ['xl/worksheets/sheet1.xml', 'xl/worksheets/sheet7.xml']) {
    const entry = zip[path]
    if (entry) {
      reordered[path] = entry
    }
  }
  for (const [path, entry] of Object.entries(zip)) {
    if (!(path in reordered)) {
      reordered[path] = entry
    }
  }
  return zipSync(reordered)
}

function buildTwoSheetWorkbook(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: { name: 'Relationship mapped styles' },
    sheets: [
      {
        id: 1,
        name: 'First',
        order: 0,
        cells: [{ address: 'A1', value: 'Styled' }],
      },
      {
        id: 2,
        name: 'Second',
        order: 1,
        cells: [{ address: 'A1', value: 'Plain' }],
      },
    ],
  }
}

function buildTwoSheetWorkbookWithTrailingSpaceName(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: { name: 'Relationship mapped trailing styles' },
    sheets: [
      {
        id: 1,
        name: 'First ',
        order: 0,
        cells: [{ address: 'A1', value: 'Styled' }],
      },
      {
        id: 2,
        name: 'Second',
        order: 1,
        cells: [{ address: 'A1', value: 'Plain' }],
      },
    ],
  }
}

function buildAxisStyleReferenceWorkbook(): Uint8Array {
  const zip = unzipSync(exportXlsx(buildStyledWorkbook()))
  zip['xl/styles.xml'] = strToU8(axisStyleReferenceStylesXml)
  const sheetXml = strFromU8(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
    .replace(/\s+s="[^"]*"/u, '')
    .replace(/<sheetData\b/u, '<cols><col min="1" max="1" width="20" customWidth="1" style="1" customFormat="1"/></cols><sheetData')
    .replace(/<row\b(?=[^>]*\br="1")[^>]*>/u, (tag) => `${tag.slice(0, -1)} s="1" customFormat="1">`)
    .replace(
      /<\/sheetData>/u,
      '<row r="2" s="1" customFormat="1"><c r="B2"/></row><row r="3"><c r="A3"/></row><row r="4" customFormat="1"/></sheetData>',
    )
  zip['xl/worksheets/sheet1.xml'] = strToU8(sheetXml)
  return zipSync(zip)
}

function buildStyleOnlyBlankCellWorkbook(): Uint8Array {
  const zip = unzipSync(exportXlsx(buildStyledWorkbook()))
  zip['xl/styles.xml'] = strToU8(axisStyleReferenceStylesXml)
  const sheetXml = strFromU8(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
    .replace(/<dimension\b[^>]*\/>/u, '<dimension ref="A2:B2"/>')
    .replace(/<sheetData\b[^>]*>[\s\S]*?<\/sheetData>/u, '<sheetData><row r="2"><c r="A2" s="1"/><c r="B2" s="1"/></row></sheetData>')
  zip['xl/worksheets/sheet1.xml'] = strToU8(sheetXml)
  return zipSync(zip)
}

const manyStyleOnlyBlankCellRowCount = 6_000
const repeatedRowStyleMetadataCount = 6_000

function buildManyStyleOnlyBlankCellsWorkbook(): Uint8Array {
  const zip = unzipSync(exportXlsx(buildStyledWorkbook()))
  zip['xl/styles.xml'] = strToU8(axisStyleReferenceStylesXml)
  const rows = Array.from({ length: manyStyleOnlyBlankCellRowCount }, (_entry, index) => {
    const rowNumber = index + 1
    return `<row r="${String(rowNumber)}"><c r="A${String(rowNumber)}" s="1"><v>${String(rowNumber)}</v></c><c r="B${String(
      rowNumber,
    )}" s="1"/></row>`
  }).join('')
  const sheetXml = strFromU8(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
    .replace(/<dimension\b[^>]*\/>/u, `<dimension ref="A1:B${String(manyStyleOnlyBlankCellRowCount)}"/>`)
    .replace(/<sheetData\b[^>]*>[\s\S]*?<\/sheetData>/u, `<sheetData>${rows}</sheetData>`)
  zip['xl/worksheets/sheet1.xml'] = strToU8(sheetXml)
  return zipSync(zip)
}

function buildRepeatedRowStyleMetadataWorkbook(): Uint8Array {
  const zip = unzipSync(exportXlsx(buildStyledWorkbook()))
  zip['xl/styles.xml'] = strToU8(axisStyleReferenceStylesXml)
  const rows = Array.from({ length: repeatedRowStyleMetadataCount }, (_entry, index) => {
    const rowNumber = index + 1
    const cellXml = rowNumber === 1 ? '<c r="A1"><v>1</v></c>' : ''
    return `<row r="${String(rowNumber)}" s="1" customFormat="1">${cellXml}</row>`
  }).join('')
  const sheetXml = strFromU8(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
    .replace(/<dimension\b[^>]*\/>/u, `<dimension ref="A1:A${String(repeatedRowStyleMetadataCount)}"/>`)
    .replace(/<sheetData\b[^>]*>[\s\S]*?<\/sheetData>/u, `<sheetData>${rows}</sheetData>`)
  zip['xl/worksheets/sheet1.xml'] = strToU8(sheetXml)
  return zipSync(zip)
}

function buildPartiallyIndexedStyleArtifactWorkbook(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: {
      name: 'Partial style artifacts',
      metadata: {
        styles: [
          {
            id: 'header-style',
            font: { bold: true },
            alignment: { horizontal: 'center', vertical: 'middle' },
            borders: {
              top: { style: 'solid', weight: 'thin', color: '#000000' },
              right: { style: 'solid', weight: 'thin', color: '#000000' },
              bottom: { style: 'solid', weight: 'thin', color: '#000000' },
              left: { style: 'solid', weight: 'thin', color: '#000000' },
            },
          },
        ],
        styleArtifacts: {
          stylesXml: headerStyleReferenceStylesXml,
        },
      },
    },
    sheets: [
      {
        id: 1,
        name: 'Report',
        order: 0,
        cells: [
          { address: 'B1', value: 'Current' },
          { address: 'C1', value: 'Prior' },
        ],
        metadata: {
          styleRanges: [
            {
              range: { sheetName: 'Report', startAddress: 'B1', endAddress: 'C1' },
              styleId: 'header-style',
            },
          ],
          styleArtifacts: {
            cellStyleIndexes: [{ address: 'C1', styleIndex: 1 }],
          },
        },
      },
    ],
  }
}

function readCellXml(bytes: Uint8Array, cellRef: string): string | undefined {
  const [sheetPath, address] = cellRef.split('!')
  const sheetXml = strFromU8(unzipSync(bytes)[sheetPath ?? ''] ?? new Uint8Array())
  return new RegExp(`<c\\b(?=[^>]*\\br="${address ?? ''}")[^>]*(?:\\/>|>[\\s\\S]*?<\\/c>)`, 'u').exec(sheetXml)?.[0]
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

function readColumnAttributes(bytes: Uint8Array, columnNumber: number): { style?: string; customFormat?: string } {
  const sheetXml = strFromU8(unzipSync(bytes)['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
  const colXml = new RegExp(`<col\\b(?=[^>]*\\bmin="${String(columnNumber)}")(?=[^>]*\\bmax="${String(columnNumber)}")[^>]*\\/>`, 'u').exec(
    sheetXml,
  )?.[0]
  return {
    ...(colXml ? { style: readXmlAttribute(colXml, 'style') } : {}),
    ...(colXml ? { customFormat: readXmlAttribute(colXml, 'customFormat') } : {}),
  }
}

function readRowAttributes(bytes: Uint8Array, rowNumber: number): { style?: string; customFormat?: string } {
  const sheetXml = strFromU8(unzipSync(bytes)['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
  const rowXml = new RegExp(`<row\\b(?=[^>]*\\br="${String(rowNumber)}")[^>]*>`, 'u').exec(sheetXml)?.[0]
  return {
    ...(rowXml ? { style: readXmlAttribute(rowXml, 's') } : {}),
    ...(rowXml ? { customFormat: readXmlAttribute(rowXml, 'customFormat') } : {}),
  }
}

function readColumnStyleNumberFormat(bytes: Uint8Array, columnNumber: number): string | undefined {
  const styleIndex = readColumnAttributes(bytes, columnNumber).style
  return styleIndex ? readStyleNumberFormat(bytes, Number(styleIndex)) : undefined
}

function readRowStyleNumberFormat(bytes: Uint8Array, rowNumber: number): string | undefined {
  const styleIndex = readRowAttributes(bytes, rowNumber).style
  return styleIndex ? readStyleNumberFormat(bytes, Number(styleIndex)) : undefined
}

function readStyleNumberFormat(bytes: Uint8Array, styleIndex: number): string | undefined {
  const stylesXml = strFromU8(unzipSync(bytes)['xl/styles.xml'] ?? new Uint8Array())
  const xf = listElements(stylesXml, 'cellXfs', 'xf')[styleIndex] ?? ''
  const numFmtId = readXmlAttribute(xf, 'numFmtId')
  if (!numFmtId) {
    return undefined
  }
  return [...stylesXml.matchAll(/<numFmt\b[^>]*numFmtId="([^"]+)"[^>]*formatCode="([^"]*)"\/?>(?:<\/numFmt>)?/gu)].find(
    (match) => match[1] === numFmtId,
  )?.[2]
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

function collectGarbage(): void {
  const bunValue = Reflect.get(globalThis, 'Bun')
  if (isRecord(bunValue) && isGarbageCollector(bunValue['gc'])) {
    bunValue['gc'](true)
    return
  }
  const nodeGc = Reflect.get(globalThis, 'gc')
  if (isGarbageCollector(nodeGc)) {
    nodeGc()
  }
}

function isRecord(value: unknown): value is Record<PropertyKey, unknown> {
  return typeof value === 'object' && value !== null
}

function isGarbageCollector(value: unknown): value is (force?: boolean) => void {
  return typeof value === 'function'
}

function readBenchmarkTolerance(): number {
  const raw = process.env.BILIG_BENCH_TOLERANCE
  if (!raw) {
    return 1
  }
  const tolerance = Number(raw)
  return Number.isFinite(tolerance) && tolerance > 0 ? tolerance : 1
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

const axisStyleReferenceStylesXml = [
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
  '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
  '<numFmts count="1"><numFmt numFmtId="164" formatCode="00000"/></numFmts>',
  '<fonts count="2"><font><sz val="11"/><name val="Aptos"/></font><font><i/><u/><color rgb="FF0000FF"/></font></fonts>',
  '<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor theme="0"/><bgColor rgb="FF000000"/></patternFill></fill></fills>',
  '<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left style="thin"><color indexed="64"/></left><right style="thin"><color theme="0"/></right><top style="thin"><color theme="0"/></top><bottom style="thin"><color indexed="64"/></bottom><diagonal/></border></borders>',
  '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>',
  '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="164" fontId="1" fillId="2" borderId="1" xfId="0" applyNumberFormat="1"/></cellXfs>',
  '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>',
  '</styleSheet>',
].join('')

const headerStyleReferenceStylesXml = [
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
  '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
  '<fonts count="2"><font><sz val="11"/><name val="Aptos"/></font><font><b/></font></fonts>',
  '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>',
  '<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left style="thin"><color rgb="FF000000"/></left><right style="thin"><color rgb="FF000000"/></right><top style="thin"><color rgb="FF000000"/></top><bottom style="thin"><color rgb="FF000000"/></bottom><diagonal/></border></borders>',
  '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>',
  '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf></cellXfs>',
  '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>',
  '</styleSheet>',
].join('')
