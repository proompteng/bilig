import { describe, expect, it } from 'vitest'

import { parseHeadlessLargeSimpleWorksheetFromChunks } from '../xlsx-large-simple-headless-worksheet-scanner.js'
import { ImportedWorkbookStringPool } from '../xlsx-large-simple-string-pool.js'
import { parseLargeSimpleWorksheetCellsFromChunks } from '../xlsx-large-simple-worksheet-stream-scanner.js'

const encoder = new TextEncoder()

describe('large simple worksheet stream scanners', () => {
  it('retains tag-open boundaries across chunks in headless scans', () => {
    const scan = parseHeadlessLargeSimpleWorksheetFromChunks(splitAfterTagOpen(worksheetXml()), 0, { hasSharedStrings: false })

    expect(scan?.cellCount).toBe(2)
    expect(scan?.valueCellCount).toBe(2)
    expect(scan?.usedRange).toEqual({ startRow: 0, startColumn: 0, endRow: 0, endColumn: 1 })
  })

  it('counts headless dimensions and conditional format ranges from attribute bytes', () => {
    const scan = parseHeadlessLargeSimpleWorksheetFromChunks(splitAfterTagOpen(headlessMetadataWorksheetXml()), 0, {
      hasSharedStrings: false,
    })

    expect(scan).toMatchObject({
      cellCount: 2,
      valueCellCount: 2,
      rowCount: 12,
      columnCount: 3,
      mergeCount: 2,
      tableCount: 2,
      conditionalFormatCount: 4,
      dataValidationCount: 3,
      usedRange: { startRow: 0, startColumn: 0, endRow: 11, endColumn: 2 },
    })
  })

  it('counts large split metadata in headless scans without retaining metadata bodies', () => {
    const retainedBufferLengths: number[] = []
    const scan = parseHeadlessLargeSimpleWorksheetFromChunks(splitLargeMergeCellsWorksheetXml(), 0, {
      hasSharedStrings: false,
      onRetainedBufferLength: (length) => retainedBufferLengths.push(length),
    })

    expect(scan).toMatchObject({
      cellCount: 1,
      valueCellCount: 1,
      mergeCount: 2,
      rowCount: 1,
      columnCount: 1,
    })
    expect(retainedBufferLengths.length).toBeGreaterThan(0)
    expect(Math.max(...retainedBufferLengths)).toBeLessThan(1024)
  })

  it('retains tag-open boundaries across chunks in materialized scans', () => {
    const scan = parseLargeSimpleWorksheetCellsFromChunks(splitAfterTagOpen(worksheetXml()), 0, { hasSharedStrings: false })

    expect(scan?.cellScan.cellCount).toBe(2)
    expect(scan?.cellScan.arena.materializeSheetCells(0)).toEqual([
      { address: 'A1', value: 1 },
      { address: 'B1', value: 2 },
    ])
  })

  it('can defer style coordinates while retaining required style indexes', () => {
    const scan = parseLargeSimpleWorksheetCellsFromChunks(splitAfterTagOpen(styledWorksheetXml()), 0, {
      hasSharedStrings: false,
      retainStyleCoordinates: false,
    })
    const requiredStyleIndexes = new Set<number>()
    const coordinates: Array<{ row: number; column: number; styleIndex: number }> = []

    scan?.cellScan.styleIndexes.collectRequiredStyleIndexes(requiredStyleIndexes)
    scan?.cellScan.styleIndexes.forEach((row, column, styleIndex) => coordinates.push({ row, column, styleIndex }))

    expect(scan?.cellScan.cellCount).toBe(1)
    expect(scan?.cellScan.blankStyleCellCount).toBe(1)
    expect(scan?.cellScan.styleIndexes.hasCoordinateStorage).toBe(false)
    expect([...requiredStyleIndexes].toSorted((left, right) => left - right)).toEqual([3, 4])
    expect(coordinates).toEqual([])
  })

  it('rescans style coordinates without materializing cell values', () => {
    const scan = parseLargeSimpleWorksheetCellsFromChunks(splitAfterTagOpen(styledInlineWorksheetXml()), 0, {
      hasSharedStrings: false,
      retainCells: false,
      retainStyleIndexes: true,
      retainStyleCoordinates: true,
    })
    const coordinates: Array<{ row: number; column: number; styleIndex: number }> = []

    scan?.cellScan.styleIndexes.forEach((row, column, styleIndex) => coordinates.push({ row, column, styleIndex }))

    expect(scan?.cellScan.cellCount).toBe(2)
    expect(scan?.cellScan.blankStyleCellCount).toBe(1)
    expect(scan?.cellScan.arena.materializeSheetCells(0)).toEqual([])
    expect(coordinates).toEqual([
      { row: 0, column: 0, styleIndex: 3 },
      { row: 0, column: 1, styleIndex: 4 },
      { row: 1, column: 0, styleIndex: 5 },
    ])
  })

  it('streams large split merge metadata in materialized scans without retaining metadata bodies', () => {
    const retainedBufferLengths: number[] = []
    const scan = parseLargeSimpleWorksheetCellsFromChunks(splitLargeMergeCellsWorksheetXml(), 0, {
      hasSharedStrings: false,
      onRetainedBufferLength: (length) => retainedBufferLengths.push(length),
    })

    expect(scan?.cellScan.cellCount).toBe(1)
    expect(scan?.cellScan.mergeCount).toBe(2)
    expect(scan?.metadata?.merges).toEqual([
      { startAddress: 'A1', endAddress: 'B1' },
      { startAddress: 'A2', endAddress: 'B2' },
    ])
    expect(retainedBufferLengths.length).toBeGreaterThan(0)
    expect(Math.max(...retainedBufferLengths)).toBeLessThan(1024)
  })

  it('streams large split table metadata in materialized scans without retaining metadata bodies', () => {
    const retainedBufferLengths: number[] = []
    const scan = parseLargeSimpleWorksheetCellsFromChunks(splitLargeTablePartsWorksheetXml(), 0, {
      hasSharedStrings: false,
      onRetainedBufferLength: (length) => retainedBufferLengths.push(length),
    })

    expect(scan?.cellScan.cellCount).toBe(1)
    expect(scan?.cellScan.tableCount).toBe(2)
    expect(scan?.metadata?.tableRelationshipIds).toEqual(['rIdTable1', 'rIdTable2'])
    expect(retainedBufferLengths.length).toBeGreaterThan(0)
    expect(Math.max(...retainedBufferLengths)).toBeLessThan(1024)
  })

  it('rejects unterminated streamed metadata in headless and materialized scans', () => {
    expect(
      parseHeadlessLargeSimpleWorksheetFromChunks(splitUnterminatedMergeCellsWorksheetXml(), 0, { hasSharedStrings: false }),
    ).toBeNull()
    expect(parseLargeSimpleWorksheetCellsFromChunks(splitUnterminatedMergeCellsWorksheetXml(), 0, { hasSharedStrings: false })).toBeNull()
  })

  it('preserves streamed scalar and shared-string values when parsing value byte ranges', () => {
    const scan = parseLargeSimpleWorksheetCellsFromChunks(splitAfterTagOpen(scalarWorksheetXml()), 0, {
      hasSharedStrings: true,
      sharedStrings: [{ text: 'Shared label', rich: false }],
    })

    expect(scan?.cellScan.arena.materializeSheetCells(0)).toEqual([
      { address: 'A1', value: 42 },
      { address: 'B1', value: -7 },
      { address: 'C1', value: Number('0.12345678901234568') },
      { address: 'D1', value: Number('1.25E-7') },
      { address: 'E1', value: true },
      { address: 'F1', value: false },
      { address: 'G1', value: '#DIV/0!' },
      { address: 'H1', value: 'Shared label' },
    ])
  })

  it('infers omitted row and cell refs from stream order', () => {
    const headless = parseHeadlessLargeSimpleWorksheetFromChunks(splitAfterTagOpen(implicitAddressWorksheetXml()), 0, {
      hasSharedStrings: true,
    })
    const scan = parseLargeSimpleWorksheetCellsFromChunks(splitAfterTagOpen(implicitAddressWorksheetXml()), 0, {
      hasSharedStrings: true,
      sharedStrings: [{ text: 'Shared label', rich: false }],
    })

    expect(headless).toMatchObject({
      cellCount: 4,
      valueCellCount: 4,
      formulaCellCount: 1,
      rowCount: 2,
      columnCount: 2,
      usedRange: { startRow: 0, startColumn: 0, endRow: 1, endColumn: 1 },
    })
    expect(scan?.cellScan.arena.materializeSheetCells(0)).toEqual([
      { address: 'A1', value: 1 },
      { address: 'B1', value: 2 },
      { address: 'A2', value: 'Shared label' },
      { address: 'B2', value: 3, formula: 'A1+B1' },
    ])
    expect(scan?.metadata?.rows?.entries).toEqual([{ id: 'row:0', index: 0, size: 16 }])
  })

  it('collects exact worksheet metadata records without retaining metadata XML', () => {
    const scan = parseLargeSimpleWorksheetCellsFromChunks(splitAfterTagOpen(metadataWorksheetXml()), 0, {
      hasSharedStrings: false,
      sheetName: 'Data',
    })

    expect(scan?.metadataXml).toBeUndefined()
    expect(scan?.cellScan.mergeCount).toBe(1)
    expect(scan?.metadata).toEqual({
      columns: {
        entries: [
          { id: 'col:0', index: 0, size: 75, hidden: true },
          { id: 'col:1', index: 1, size: 75, hidden: true },
        ],
        metadata: [
          {
            start: 0,
            count: 2,
            size: 75,
            xlsxWidth: 12.5,
            styleIndex: 3,
            hidden: true,
            customWidth: true,
            bestFit: false,
            outlineLevel: 1,
          },
        ],
      },
      conditionalFormats: [
        {
          id: 'xlsx-cf:Data:A1:B2:1',
          range: { sheetName: 'Data', startAddress: 'A1', endAddress: 'B2' },
          rule: { kind: 'cellIs', operator: 'greaterThan', values: [0] },
          style: {},
          priority: 1,
        },
      ],
      drawingRelationshipId: 'rIdDrawing1',
      filters: [
        {
          sheetName: 'Data',
          startAddress: 'A1',
          endAddress: 'B2',
          criteria: [
            {
              colId: 0,
              filters: { values: ['Open'] },
            },
          ],
        },
      ],
      hyperlinks: [{ ref: 'A1:B1', location: 'Summary!A1', tooltip: 'Jump', display: 'Summary' }],
      dataValidations: [
        {
          range: { sheetName: 'Data', startAddress: 'A1', endAddress: 'A1' },
          rule: { kind: 'list', values: ['Open', 'Closed'] },
          allowBlank: true,
        },
        {
          range: { sheetName: 'Data', startAddress: 'B1', endAddress: 'B1' },
          rule: { kind: 'whole', operator: 'between', values: [1, 10] },
        },
        {
          range: { sheetName: 'Data', startAddress: 'B2', endAddress: 'B2' },
          rule: { kind: 'whole', operator: 'between', values: [1, 10] },
        },
      ],
      rows: {
        entries: [{ id: 'row:1', index: 1, size: 24, hidden: true }],
        metadata: [
          {
            start: 1,
            count: 1,
            size: 24,
            xlsxHeight: 18,
            styleIndex: 4,
            hidden: true,
            customFormat: true,
            customHeight: true,
            collapsed: false,
            thickTop: true,
          },
        ],
      },
      merges: [{ startAddress: 'A1', endAddress: 'B1' }],
      printPageSetup: {
        printOptionsXml: '<printOptions horizontalCentered="1"/>',
        pageMarginsXml: '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75"/>',
        pageSetupXml: '<pageSetup orientation="landscape" r:id="rIdPrinterSettings1"/>',
        headerFooterXml: '<headerFooter><oddFooter>Page &amp;P</oddFooter></headerFooter>',
        rowBreaksXml: '<rowBreaks count="1"><brk id="10" max="16383" man="1"/></rowBreaks>',
        colBreaksXml: '<colBreaks count="1"><brk id="2" max="1048575" man="1"/></colBreaks>',
      },
      sheetFormatPr: { defaultRowHeight: 15, outlineLevelRow: 1 },
      tableRelationshipIds: ['rIdTable1'],
    })
  })

  it('keeps raw conditional-format XML only when style artifacts are required', () => {
    const scan = parseLargeSimpleWorksheetCellsFromChunks(splitAfterTagOpen(styledConditionalFormatWorksheetXml()), 0, {
      hasSharedStrings: false,
      sheetName: 'Data',
    })

    expect(scan?.metadataXml).toBeUndefined()
    expect(scan?.metadata?.conditionalFormats).toBeUndefined()
    expect(scan?.metadata?.conditionalFormattingXml).toEqual([
      '<conditionalFormatting sqref="A1"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>0</formula></cfRule></conditionalFormatting>',
    ])
  })

  it('shares repeated inline strings and formulas through the import string pool', () => {
    const pool = new ImportedWorkbookStringPool()
    const first = parseLargeSimpleWorksheetCellsFromChunks(splitAfterTagOpen(repeatedStringFormulaWorksheetXml()), 0, {
      hasSharedStrings: false,
      stringPool: pool,
    })
    const second = parseLargeSimpleWorksheetCellsFromChunks(splitAfterTagOpen(repeatedStringFormulaWorksheetXml()), 1, {
      hasSharedStrings: false,
      stringPool: pool,
    })

    expect(first?.cellScan.arena.materializeSheetCells(0)).toEqual([
      { address: 'A1', value: 'Repeated label' },
      { address: 'B1', value: 1, formula: 'A1&"!"' },
    ])
    expect(second?.cellScan.arena.materializeSheetCells(1)).toEqual([
      { address: 'A1', value: 'Repeated label' },
      { address: 'B1', value: 1, formula: 'A1&"!"' },
    ])
    expect(pool.count).toBe(2)
  })

  it('retains hyperlink XML when range expansion would lose fidelity', () => {
    const scan = parseLargeSimpleWorksheetCellsFromChunks(splitAfterTagOpen(oversizedHyperlinkWorksheetXml()), 0, {
      hasSharedStrings: false,
    })

    expect(scan?.metadata?.hyperlinks).toBeUndefined()
    expect(scan?.metadataXml).toContain('<hyperlinks>')
  })

  it('rejects unsupported data validation rules instead of dropping them from streamed metadata', () => {
    expect(
      parseHeadlessLargeSimpleWorksheetFromChunks(splitAfterTagOpen(unsupportedDataValidationWorksheetXml()), 0, {
        hasSharedStrings: false,
      }),
    ).toBeNull()
    expect(
      parseLargeSimpleWorksheetCellsFromChunks(splitAfterTagOpen(unsupportedDataValidationWorksheetXml()), 0, {
        hasSharedStrings: false,
        sheetName: 'Data',
      }),
    ).toBeNull()
  })
})

function splitAfterTagOpen(xml: string): (onChunk: (chunk: Uint8Array) => void) => boolean {
  const bytes = encoder.encode(xml)
  return (onChunk) => {
    let start = 0
    for (let index = 0; index < bytes.byteLength; index += 1) {
      if (bytes[index] !== 60) {
        continue
      }
      onChunk(bytes.subarray(start, index + 1))
      start = index + 1
    }
    if (start < bytes.byteLength) {
      onChunk(bytes.subarray(start))
    }
    return true
  }
}

function worksheetXml(): string {
  return [
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
    '<dimension ref="A1:B1"/>',
    '<sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row></sheetData>',
    '</worksheet>',
  ].join('')
}

function styledWorksheetXml(): string {
  return [
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
    '<dimension ref="A1:B1"/>',
    '<sheetData><row r="1"><c r="A1" s="3"><v>1</v></c><c r="B1" s="4"/></row></sheetData>',
    '</worksheet>',
  ].join('')
}

function styledInlineWorksheetXml(): string {
  return [
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
    '<dimension ref="A1:B2"/>',
    '<sheetData>',
    '<row r="1"><c r="A1" t="inlineStr" s="3"><is><t>Label</t></is></c><c r="B1" s="4"><v>2</v></c></row>',
    '<row r="2"><c r="A2" s="5"/></row>',
    '</sheetData>',
    '</worksheet>',
  ].join('')
}

function headlessMetadataWorksheetXml(): string {
  return [
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
    '<dimension ref=" $A$1:$C$12 "/>',
    '<sheetData><row r="1"><c r="A1"><v>1</v></c></row><row r="12"><c r="C12"><v>2</v></c></row></sheetData>',
    '<conditionalFormatting sqref="A1:A2 C1:C2">',
    '<cfRule type="cellIs" priority="1" operator="greaterThan"><formula>0</formula></cfRule>',
    '<cfRule type="cellIs" priority="2" operator="lessThan"><formula>9</formula></cfRule>',
    '</conditionalFormatting>',
    '<dataValidations count="2">',
    '<dataValidation type="list" sqref="A1"><formula1>"Open,Closed"</formula1></dataValidation>',
    '<dataValidation type="whole" operator="between" sqref="B1 B2"><formula1>1</formula1><formula2>10</formula2></dataValidation>',
    '</dataValidations>',
    '<mergeCells count="2"><mergeCell ref="A1:B1"/><mergeCell ref="A2:B2"/></mergeCells>',
    '<tableParts count="2"><tablePart r:id="rIdTable1"/><tablePart r:id="rIdTable2"/></tableParts>',
    '</worksheet>',
  ].join('')
}

function splitLargeMergeCellsWorksheetXml(): (onChunk: (chunk: Uint8Array) => void) => boolean {
  const chunks = [
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
    '<dimension ref="A1"/>',
    '<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>',
    '<mergeCells count="2"><mergeCell ref="A1:B1"/>',
    ' '.repeat(40_000),
    ' '.repeat(40_000),
    ' '.repeat(20_000),
    '<mergeCell ref="A2:B2"/></',
    'mergeCells>',
    '</worksheet>',
  ]
  return (onChunk) => {
    for (const chunk of chunks) {
      onChunk(encoder.encode(chunk))
    }
    return true
  }
}

function splitLargeTablePartsWorksheetXml(): (onChunk: (chunk: Uint8Array) => void) => boolean {
  const chunks = [
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
    '<dimension ref="A1"/>',
    '<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>',
    '<tableParts count="2"><tablePart r:id="rIdTable1"/>',
    ' '.repeat(40_000),
    ' '.repeat(40_000),
    ' '.repeat(20_000),
    '<tablePart r:id="rIdTable2"/></',
    'tableParts>',
    '</worksheet>',
  ]
  return (onChunk) => {
    for (const chunk of chunks) {
      onChunk(encoder.encode(chunk))
    }
    return true
  }
}

function splitUnterminatedMergeCellsWorksheetXml(): (onChunk: (chunk: Uint8Array) => void) => boolean {
  const chunks = [
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
    '<dimension ref="A1"/>',
    '<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>',
    '<mergeCells count="1"><mergeCell ref="A1:B1"/>',
  ]
  return (onChunk) => {
    for (const chunk of chunks) {
      onChunk(encoder.encode(chunk))
    }
    return true
  }
}

function scalarWorksheetXml(): string {
  return [
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
    '<dimension ref="A1:H1"/>',
    '<sheetData><row r="1">',
    '<c r="A1"><v>42</v></c>',
    '<c r="B1"><v> -7 </v></c>',
    '<c r="C1"><v>0.12345678901234568</v></c>',
    '<c r="D1"><v>1.25E-7</v></c>',
    '<c r="E1" t="b"><v>true</v></c>',
    '<c r="F1" t="b"><v>0</v></c>',
    '<c r="G1" t="e"><v>#DIV/0!</v></c>',
    '<c r="H1" t="s"><v>0</v></c>',
    '</row></sheetData>',
    '</worksheet>',
  ].join('')
}

function implicitAddressWorksheetXml(): string {
  return [
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
    '<dimension ref="A1:B2"/>',
    '<sheetData>',
    '<row ht="12" customHeight="1"><c><v>1</v></c><c><v>2</v></c></row>',
    '<row><c t="s"><v>0</v></c><c><f>A1+B1</f><v>3</v></c></row>',
    '</sheetData>',
    '</worksheet>',
  ].join('')
}

function metadataWorksheetXml(): string {
  return [
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
    '<dimension ref="A1:B2"/>',
    '<sheetFormatPr defaultRowHeight="15" outlineLevelRow="1"/>',
    '<cols><col min="1" max="2" width="12.5" style="3" hidden="1" customWidth="1" bestFit="0" outlineLevel="1"/></cols>',
    '<sheetData>',
    '<row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row>',
    '<row r="2" ht="18" s="4" hidden="true" customFormat="1" customHeight="1" collapsed="0" thickTop="1">',
    '<c r="A2"><v>3</v></c>',
    '</row>',
    '</sheetData>',
    '<autoFilter ref="A1:B2"><filterColumn colId="0"><filters><filter val="Open"/></filters></filterColumn></autoFilter>',
    '<conditionalFormatting sqref="A1:B2"><cfRule type="cellIs" priority="1" operator="greaterThan"><formula>0</formula></cfRule></conditionalFormatting>',
    '<mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells>',
    '<hyperlinks><hyperlink ref="A1:B1" location="Summary!A1" tooltip="Jump" display="Summary"/></hyperlinks>',
    '<dataValidations count="2">',
    '<dataValidation type="list" allowBlank="1" sqref="A1"><formula1>"Open,Closed"</formula1></dataValidation>',
    '<dataValidation type="whole" operator="between" sqref="B1 B2"><formula1>1</formula1><formula2>10</formula2></dataValidation>',
    '</dataValidations>',
    '<printOptions horizontalCentered="1"/>',
    '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75"/>',
    '<pageSetup orientation="landscape" r:id="rIdPrinterSettings1"/>',
    '<headerFooter><oddFooter>Page &amp;P</oddFooter></headerFooter>',
    '<rowBreaks count="1"><brk id="10" max="16383" man="1"/></rowBreaks>',
    '<colBreaks count="1"><brk id="2" max="1048575" man="1"/></colBreaks>',
    '<drawing r:id="rIdDrawing1"/>',
    '<tableParts count="1"><tablePart r:id="rIdTable1"/></tableParts>',
    '</worksheet>',
  ].join('')
}

function styledConditionalFormatWorksheetXml(): string {
  return [
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
    '<dimension ref="A1"/>',
    '<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>',
    '<conditionalFormatting sqref="A1"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>0</formula></cfRule></conditionalFormatting>',
    '</worksheet>',
  ].join('')
}

function repeatedStringFormulaWorksheetXml(): string {
  return [
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
    '<dimension ref="A1:B1"/>',
    '<sheetData><row r="1">',
    '<c r="A1" t="inlineStr"><is><t>Repeated label</t></is></c>',
    '<c r="B1"><f>A1&amp;&quot;!&quot;</f><v>1</v></c>',
    '</row></sheetData>',
    '</worksheet>',
  ].join('')
}

function oversizedHyperlinkWorksheetXml(): string {
  return [
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
    '<dimension ref="A1:A2000"/>',
    '<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>',
    '<hyperlinks><hyperlink ref="A1:A2000" location="Summary!A1"/></hyperlinks>',
    '</worksheet>',
  ].join('')
}

function unsupportedDataValidationWorksheetXml(): string {
  return [
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
    '<dimension ref="A1"/>',
    '<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>',
    '<dataValidations count="1"><dataValidation type="custom" sqref="A1"><formula1>A1&gt;0</formula1></dataValidation></dataValidations>',
    '</worksheet>',
  ].join('')
}
