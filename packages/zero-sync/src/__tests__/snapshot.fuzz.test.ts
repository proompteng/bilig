import { describe, expect, it } from 'vitest'
import * as fc from 'fast-check'
import { runProperty } from '@bilig/test-fuzz'
import { createEmptyWorkbookSnapshot, projectWorkbookToSnapshot } from '../snapshot.js'

describe('zero-sync snapshot fuzz', () => {
  it('should preserve projected workbook rows, format metadata, and fallback workbook metadata', async () => {
    await runProperty({
      suite: 'zero-sync/snapshot/projected-workbook-parity',
      arbitrary: fc.record({
        documentId: fc.constantFrom('doc-a', 'doc-b'),
        workbookName: fc.constantFrom('Projected Book', 'Ops Workbook'),
        value: fc.integer({ min: -200, max: 200 }),
        formatCode: fc.constantFrom('0.00', '0%', '@'),
        formula: fc.constantFrom('A1+1', 'B2*2', '1+2'),
      }),
      predicate: async ({ documentId, workbookName, value, formatCode, formula }) => {
        const projected = projectWorkbookToSnapshot(
          {
            name: workbookName,
            compatibilityMode: 'excel-modern',
            recalcEpoch: 7,
            snapshot: {
              ...createEmptyWorkbookSnapshot(documentId),
              workbook: {
                name: documentId,
                metadata: {
                  tables: [
                    {
                      name: 'FallbackTable',
                      sheetName: 'Sheet1',
                      startAddress: 'A1',
                      endAddress: 'A2',
                      columnNames: ['A'],
                      headerRow: true,
                      totalsRow: false,
                    },
                  ],
                },
              },
            },
            numberFormats: [{ id: 'fmt-1', code: formatCode, kind: 'number' }],
            styles: [{ id: 'style-1', recordJSON: { fill: { backgroundColor: '#ffee00' } } }],
            sheets: [
              {
                id: 1,
                name: 'Sheet1',
                sortOrder: 0,
                freezeRows: 1,
                freezeCols: 1,
                rowMetadata: [{ startIndex: 0, count: 1, size: 28 }],
                columnMetadata: [{ startIndex: 0, count: 1, size: 120 }],
                styleRanges: [{ startRow: 0, endRow: 0, startCol: 0, endCol: 0, styleId: 'style-1' }],
                formatRanges: [{ startRow: 0, endRow: 0, startCol: 0, endCol: 0, formatId: 'fmt-1' }],
                cells: [
                  { rowNum: 0, colNum: 0, inputValue: value, explicitFormatId: 'fmt-1' },
                  { rowNum: 0, colNum: 1, formula },
                ],
              },
            ],
          },
          documentId,
        )

        expect(projected).not.toBeNull()
        expect(projected?.workbook.name).toBe(workbookName)
        expect(projected?.workbook.metadata?.tables?.[0]?.name).toBe('FallbackTable')
        expect(projected?.workbook.metadata?.formats).toEqual([{ id: 'fmt-1', code: formatCode, kind: 'number' }])
        expect(projected?.sheets[0]?.cells).toEqual([
          { address: 'A1', value, format: formatCode },
          { address: 'B1', formula },
        ])
        expect(projected?.sheets[0]?.metadata?.freezePane).toEqual({ rows: 1, cols: 1 })
        expect(projected?.sheets[0]?.metadata?.styleRanges).toEqual([
          {
            range: {
              sheetName: 'Sheet1',
              startAddress: 'A1',
              endAddress: 'A1',
            },
            styleId: 'style-1',
          },
        ])
        expect(projected?.sheets[0]?.metadata?.formatRanges).toEqual([
          {
            range: {
              sheetName: 'Sheet1',
              startAddress: 'A1',
              endAddress: 'A1',
            },
            formatId: 'fmt-1',
          },
        ])
      },
    })
  })
})
