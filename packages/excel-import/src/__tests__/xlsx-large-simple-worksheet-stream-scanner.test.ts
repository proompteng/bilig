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
