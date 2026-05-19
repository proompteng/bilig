import { describe, expect, it } from 'vitest'

import { parseHeadlessLargeSimpleWorksheetFromChunks } from '../xlsx-large-simple-headless-worksheet-scanner.js'
import { parseLargeSimpleWorksheetCellsFromChunks } from '../xlsx-large-simple-worksheet-stream-scanner.js'

const encoder = new TextEncoder()

describe('large simple worksheet stream scanners', () => {
  it('retains tag-open boundaries across chunks in headless scans', () => {
    const scan = parseHeadlessLargeSimpleWorksheetFromChunks(splitAfterTagOpen(worksheetXml()), 0, { hasSharedStrings: false })

    expect(scan?.cellCount).toBe(2)
    expect(scan?.valueCellCount).toBe(2)
    expect(scan?.usedRange).toEqual({ startRow: 0, startColumn: 0, endRow: 0, endColumn: 1 })
  })

  it('retains tag-open boundaries across chunks in materialized scans', () => {
    const scan = parseLargeSimpleWorksheetCellsFromChunks(splitAfterTagOpen(worksheetXml()), 0, { hasSharedStrings: false })

    expect(scan?.cellScan.cellCount).toBe(2)
    expect(scan?.cellScan.arena.materializeSheetCells(0)).toEqual([
      { address: 'A1', value: 1 },
      { address: 'B1', value: 2 },
    ])
  })

  it('collects exact worksheet metadata records without retaining metadata XML', () => {
    const scan = parseLargeSimpleWorksheetCellsFromChunks(splitAfterTagOpen(metadataWorksheetXml()), 0, { hasSharedStrings: false })

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
      sheetFormatPr: { defaultRowHeight: 15, outlineLevelRow: 1 },
    })
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
    '<mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells>',
    '</worksheet>',
  ].join('')
}
