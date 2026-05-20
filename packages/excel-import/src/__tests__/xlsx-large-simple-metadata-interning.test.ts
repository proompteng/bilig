import { describe, expect, it } from 'vitest'

import { internLargeSimpleWorksheetMetadata } from '../xlsx-large-simple-metadata-interning.js'
import { ImportedWorkbookStringPool } from '../xlsx-large-simple-string-pool.js'
import type { LargeSimpleWorksheetScannedMetadata } from '../xlsx-large-simple-worksheet-metadata.js'

describe('large simple XLSX metadata interning', () => {
  it('pools streamed metadata strings without dropping fidelity-only fields', () => {
    const pool = new ImportedWorkbookStringPool()
    const metadata: LargeSimpleWorksheetScannedMetadata = {
      cellMetadataRefs: [
        { address: 'A1', cm: '1', vm: '1' },
        { address: 'A1', cm: '1' },
      ],
      columns: {
        entries: [{ id: 'col:0', index: 0, size: 64 }],
        metadata: [{ start: 0, count: 1, size: 64 }],
      },
      conditionalFormats: [
        {
          id: 'xlsx-cf:Data:A1:B2:1',
          range: { sheetName: 'Data', startAddress: 'A1', endAddress: 'B2' },
          rule: { kind: 'formula', formula: '=A1>0' },
          style: {},
          priority: 1,
        },
      ],
      conditionalFormattingXml: ['<conditionalFormatting sqref="A1:B2"/>'],
      controlArtifacts: {
        controlsXml: '<oleObjects><oleObject r:id="rIdControl"/></oleObjects>',
        worksheetRootOpenTag: '<worksheet xmlns:r="relationships">',
        legacyDrawingRelationshipId: 'rIdLegacy',
      },
      drawingRelationshipId: 'rIdDrawing',
      filters: [{ sheetName: 'Data', startAddress: 'A1', endAddress: 'B2' }],
      hyperlinks: [
        {
          ref: 'A1',
          relationshipId: 'rIdHyperlink',
          location: '#Data!A1',
          tooltip: 'Open row',
          display: 'Open row',
        },
      ],
      legacyDrawingRelationshipId: 'rIdLegacy',
      merges: [{ startAddress: 'A1', endAddress: 'B2' }],
      printPageSetup: {
        printOptionsXml: '<printOptions horizontalCentered="1"/>',
        pageMarginsXml: '<pageMargins left="0.7" right="0.7"/>',
      },
      rows: {
        entries: [{ id: 'row:0', index: 0, size: 20 }],
        metadata: [{ start: 0, count: 1, size: 20 }],
      },
      sheetFormatPr: { defaultRowHeight: 15 },
      sheetSlicerListExtXml: '<ext><x14:slicerList/></ext>',
      tableRelationshipIds: ['rIdTable'],
    }

    const interned = internLargeSimpleWorksheetMetadata(metadata, pool)

    expect(interned).toEqual(metadata)
    expect(interned?.cellMetadataRefs).toEqual(metadata.cellMetadataRefs)
    expect(interned?.controlArtifacts).toEqual(metadata.controlArtifacts)
    expect(interned?.legacyDrawingRelationshipId).toBe('rIdLegacy')
    expect(interned?.sheetSlicerListExtXml).toBe('<ext><x14:slicerList/></ext>')
    expect(pool.count).toBeLessThan(30)
  })
})
