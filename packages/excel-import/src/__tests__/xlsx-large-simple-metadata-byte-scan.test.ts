import { describe, expect, it } from 'vitest'

import {
  appendLargeSimpleColumnMetadataFromBytes,
  readLargeSimpleDrawingRelationshipIdTagFromBytes,
  readLargeSimpleSheetFormatPrTagFromBytes,
  readLargeSimpleTableRelationshipIdsFromBytes,
} from '../xlsx-large-simple-metadata-byte-scan.js'

const encoder = new TextEncoder()

describe('large simple metadata byte scan', () => {
  it('parses column metadata without decoding the cols XML span', () => {
    const bytes = encoder.encode(
      '<cols><col min="1" max="2" width="12.5" style="3" hidden="1" customWidth="1" bestFit="0" outlineLevel="1"/></cols>',
    )
    const entries: Parameters<typeof appendLargeSimpleColumnMetadataFromBytes>[0] = []
    const metadata: Parameters<typeof appendLargeSimpleColumnMetadataFromBytes>[1] = []

    appendLargeSimpleColumnMetadataFromBytes(entries, metadata, bytes, 0, bytes.byteLength)

    expect(entries).toEqual([
      { id: 'col:0', index: 0, size: 75, hidden: true },
      { id: 'col:1', index: 1, size: 75, hidden: true },
    ])
    expect(metadata).toEqual([
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
    ])
  })

  it('parses sheet format, drawing, and table relationship metadata from bytes', () => {
    const sheetFormatPr = encoder.encode('<sheetFormatPr defaultRowHeight="15" outlineLevelRow="1"/>')
    const drawing = encoder.encode('<drawing r:id="rIdDrawing1"/>')
    const tableParts = encoder.encode('<tableParts><tablePart r:id="rIdTable1"/></tableParts>')

    expect(readLargeSimpleSheetFormatPrTagFromBytes(sheetFormatPr, 0, sheetFormatPr.byteLength)).toEqual({
      defaultRowHeight: 15,
      outlineLevelRow: 1,
    })
    expect(readLargeSimpleDrawingRelationshipIdTagFromBytes(drawing, 0, drawing.byteLength)).toBe('rIdDrawing1')
    expect(readLargeSimpleTableRelationshipIdsFromBytes(tableParts, 0, tableParts.byteLength)).toEqual(['rIdTable1'])
  })
})
