import { describe, expect, it } from 'vitest'
import { strToU8, unzipSync, zipSync } from 'fflate'
import * as XLSX from 'xlsx'

import { readBenchToleranceMultiplier } from '../../../../scripts/bench-tolerance.js'
import { importXlsx } from '../index.js'

describe('XLSX sparse ranges', () => {
  it('imports actual cells without scanning every coordinate in a broad sparse ref', () => {
    const imported = importXlsx(buildBroadSparseWorkbookBytes(), 'broad-sparse.xlsx')
    const sheet = imported.snapshot.sheets[0]

    expect(sheet?.cells).toEqual([{ address: 'XFD512', row: 511, col: 16_383, formula: '40+2', value: 42 }])
    expect(imported.preview.sheets[0]).toMatchObject({
      rowCount: 512,
      columnCount: 16_384,
      nonEmptyCellCount: 1,
    })
  }, 15_000)

  it('compacts styled blank XML cells when importing a broad styled template range', () => {
    const sparseBytes = buildStyledBlankWorkbookBytes({ includeBlankCells: false })
    const denseBytes = buildStyledBlankWorkbookBytes({ includeBlankCells: true })
    const sparseMs = measureImport(sparseBytes, 'styled-sparse-control.xlsx').durationMs
    const denseMeasurements = [
      measureImport(denseBytes, 'styled-blank-template.xlsx'),
      measureImport(denseBytes, 'styled-blank-template.xlsx'),
    ]
    const { imported, durationMs: denseMs } = denseMeasurements.reduce((best, current) =>
      current.durationMs < best.durationMs ? current : best,
    )
    const sheet = imported.snapshot.sheets[0]

    expect(sheet?.cells).toEqual([{ address: 'A1', value: 123 }])
    expect(imported.preview.sheets[0]).toMatchObject({
      rowCount: styledBlankRowCount,
      columnCount: styledBlankColumnCount,
      nonEmptyCellCount: 1,
    })
    expect(sheet?.metadata?.styleRanges).toHaveLength(1)
    expect(sheet?.metadata?.styleRanges?.[0]?.range).toEqual({
      sheetName: 'StyledBlanks',
      startAddress: 'A1',
      endAddress: 'Z20000',
    })
    expect(imported.snapshot.workbook.metadata?.styles?.[0]).toMatchObject({
      fill: { backgroundColor: '#ffcc00' },
    })
    const tolerance = readBenchmarkTolerance()
    expect(denseMs).toBeLessThan(Math.max(1_500 * tolerance, sparseMs * 12 * tolerance))
  }, 15_000)

  it('imports style-metadata-heavy workbooks without retaining inert style collections', () => {
    const bytes = buildMetadataHeavyStyleWorkbookBytes(200_000)
    collectGarbage()
    const beforeRss = process.memoryUsage().rss
    const start = performance.now()

    const imported = importXlsx(bytes, 'metadata-heavy-styles.xlsx')
    const durationMs = performance.now() - start
    const rssDelta = process.memoryUsage().rss - beforeRss

    expect(imported.snapshot.sheets[0]?.cells).toEqual([{ address: 'A1', row: 0, col: 0, value: 123 }])
    expect(imported.snapshot.workbook.metadata?.styles).toHaveLength(1)
    expect(imported.snapshot.sheets[0]?.metadata?.conditionalFormats).toHaveLength(1)
    expect(durationMs).toBeLessThan(1_500 * readBenchmarkTolerance())
    expect(rssDelta).toBeLessThan(256 * 1024 * 1024)
  }, 15_000)

  it('does not expand whole-worksheet column metadata into per-column snapshot entries', () => {
    const imported = importXlsx(buildWholeWorksheetColumnMetadataWorkbookBytes(), 'whole-column-metadata.xlsx')
    const sheet = imported.snapshot.sheets[0]

    expect(sheet?.cells).toEqual([{ address: 'A3040', row: 3039, col: 0, value: 1 }])
    expect(sheet?.metadata?.columns).toBeUndefined()
    expect(imported.preview.sheets[0]).toMatchObject({
      rowCount: 3_040,
      columnCount: 1,
      nonEmptyCellCount: 1,
    })
  })
})

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
  return readBenchToleranceMultiplier(process.env)
}

function measureImport(bytes: Uint8Array, fileName: string): { imported: ReturnType<typeof importXlsx>; durationMs: number } {
  const start = performance.now()
  const imported = importXlsx(bytes, fileName)
  return {
    imported,
    durationMs: performance.now() - start,
  }
}

function buildBroadSparseWorkbookBytes(): Uint8Array {
  const workbook = XLSX.utils.book_new()
  const sheet: XLSX.WorkSheet = {
    XFD512: { t: 'n', f: '40+2', v: 42 },
    '!ref': 'A1:XFD512',
  }
  XLSX.utils.book_append_sheet(workbook, sheet, 'Sparse')
  return XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' })
}

const styledBlankRowCount = 20_000
const styledBlankColumnCount = 26

function buildStyledBlankWorkbookBytes(options: { includeBlankCells: boolean }): Uint8Array {
  const workbook = XLSX.utils.book_new()
  const sheet = XLSX.utils.aoa_to_sheet([[123]])
  XLSX.utils.book_append_sheet(workbook, sheet, 'StyledBlanks')

  const zip = unzipSync(XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' }))
  zip['xl/worksheets/sheet1.xml'] = strToU8(buildStyledBlankWorksheetXml(options.includeBlankCells))
  zip['xl/styles.xml'] = strToU8(styledBlankWorkbookStylesXml)
  return zipSync(zip)
}

function buildMetadataHeavyStyleWorkbookBytes(styleCollectionCount: number): Uint8Array {
  const workbook = XLSX.utils.book_new()
  const sheet = XLSX.utils.aoa_to_sheet([[123]])
  XLSX.utils.book_append_sheet(workbook, sheet, 'HeavyStyles')

  const zip = unzipSync(XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' }))
  zip['xl/worksheets/sheet1.xml'] = strToU8(buildMetadataHeavyStyleWorksheetXml())
  zip['xl/styles.xml'] = strToU8(buildMetadataHeavyStylesXml(styleCollectionCount))
  return zipSync(zip)
}

function buildWholeWorksheetColumnMetadataWorkbookBytes(): Uint8Array {
  const workbook = XLSX.utils.book_new()
  const sheet = XLSX.utils.aoa_to_sheet([[1]])
  XLSX.utils.book_append_sheet(workbook, sheet, 'WideColumnMetadata')

  const zip = unzipSync(XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' }))
  zip['xl/worksheets/sheet1.xml'] = strToU8(
    [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
      '<dimension ref="A3040:A3040"/>',
      '<cols><col min="1" max="16384" width="10" customWidth="1"/></cols>',
      '<sheetData><row r="3040"><c r="A3040"><v>1</v></c></row></sheetData>',
      '</worksheet>',
    ].join(''),
  )
  return zipSync(zip)
}

function buildMetadataHeavyStyleWorksheetXml(): string {
  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
    '<dimension ref="A1:A1"/>',
    '<sheetData><row r="1"><c r="A1" s="1"><v>123</v></c></row></sheetData>',
    '<conditionalFormatting sqref="A1">',
    '<cfRule type="cellIs" priority="1" operator="greaterThan" dxfId="0"><formula>100</formula></cfRule>',
    '</conditionalFormatting>',
    '</worksheet>',
  ].join('')
}

function buildMetadataHeavyStylesXml(styleCollectionCount: number): string {
  const cellStyleXfs: string[] = []
  const cellStyles: string[] = []
  for (let index = 0; index < styleCollectionCount; index += 1) {
    cellStyleXfs.push('<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>')
    cellStyles.push(`<cellStyle name="Style${String(index)}" xfId="${String(index)}" builtinId="0"/>`)
  }

  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
    '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>',
    '<fills count="3">',
    '<fill><patternFill patternType="none"/></fill>',
    '<fill><patternFill patternType="gray125"/></fill>',
    '<fill><patternFill patternType="solid"><fgColor rgb="FFFFCC00"/><bgColor indexed="64"/></patternFill></fill>',
    '</fills>',
    '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>',
    `<cellStyleXfs count="${String(styleCollectionCount)}">${cellStyleXfs.join('')}</cellStyleXfs>`,
    '<cellXfs count="2">',
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>',
    '<xf numFmtId="0" fontId="0" fillId="2" borderId="0" applyFill="1"/>',
    '</cellXfs>',
    '<dxfs count="1">',
    '<dxf><fill><patternFill patternType="solid"><fgColor rgb="FFFF0000"/></patternFill></fill></dxf>',
    '</dxfs>',
    `<cellStyles count="${String(styleCollectionCount)}">${cellStyles.join('')}</cellStyles>`,
    '</styleSheet>',
  ].join('')
}

function buildStyledBlankWorksheetXml(includeBlankCells: boolean): string {
  const rows: string[] = []
  for (let row = 1; row <= styledBlankRowCount; row += 1) {
    const cells: string[] = []
    const columnCount = includeBlankCells ? styledBlankColumnCount : row === 1 ? 1 : 0
    for (let column = 0; column < columnCount; column += 1) {
      const address = `${encodeColumnName(column)}${String(row)}`
      cells.push(address === 'A1' ? '<c r="A1" s="1"><v>123</v></c>' : `<c r="${address}" s="1"/>`)
    }
    if (cells.length > 0) {
      rows.push(`<row r="${String(row)}">${cells.join('')}</row>`)
    }
  }
  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
    `<dimension ref="A1:${encodeColumnName(styledBlankColumnCount - 1)}${String(styledBlankRowCount)}"/>`,
    `<sheetData>${rows.join('')}</sheetData>`,
    '</worksheet>',
  ].join('')
}

function encodeColumnName(index: number): string {
  let value = index + 1
  let output = ''
  while (value > 0) {
    const remainder = (value - 1) % 26
    output = String.fromCharCode(65 + remainder) + output
    value = Math.floor((value - 1) / 26)
  }
  return output
}

const styledBlankWorkbookStylesXml = [
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
  '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
  '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>',
  '<fills count="3">',
  '<fill><patternFill patternType="none"/></fill>',
  '<fill><patternFill patternType="gray125"/></fill>',
  '<fill><patternFill patternType="solid"><fgColor rgb="FFFFCC00"/><bgColor indexed="64"/></patternFill></fill>',
  '</fills>',
  '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>',
  '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>',
  '<cellXfs count="2">',
  '<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>',
  '<xf numFmtId="0" fontId="0" fillId="2" borderId="0" applyFill="1"/>',
  '</cellXfs>',
  '</styleSheet>',
].join('')
