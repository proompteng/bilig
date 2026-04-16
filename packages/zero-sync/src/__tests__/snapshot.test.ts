import { describe, expect, it } from 'vitest'
import { projectWorkbookToSnapshot } from '../snapshot'

describe('projectWorkbookToSnapshot', () => {
  it('reconstructs workbook metadata and sheet formatting from normalized Zero rows', () => {
    const projected = projectWorkbookToSnapshot(
      {
        id: 'doc-1',
        name: 'Projected Book',
        compatibilityMode: 'excel-modern',
        recalcEpoch: 7,
        snapshot: {
          version: 1,
          workbook: {
            name: 'Warm Snapshot',
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
          sheets: [
            {
              name: 'Sheet1',
              order: 0,
              cells: [],
            },
          ],
        },
        calculationSettings: {
          mode: 'automatic',
          recalcEpoch: 7,
        },
        workbookMetadataEntries: [{ key: 'locale', value: 'en-US' }],
        definedNames: [
          {
            name: 'Sales',
            value: {
              kind: 'range-ref',
              sheetName: 'Sheet1',
              startAddress: 'A1',
              endAddress: 'A2',
            },
          },
        ],
        styles: [
          {
            id: 'style-1',
            recordJSON: {
              fill: {
                backgroundColor: '#ffee00',
              },
            },
          },
        ],
        numberFormats: [
          {
            id: 'fmt-1',
            code: '0.00',
            kind: 'number',
          },
        ],
        sheets: [
          {
            name: 'Sheet1',
            sortOrder: 0,
            freezeRows: 1,
            freezeCols: 2,
            cells: [
              {
                address: 'A1',
                inputValue: 42,
                explicitFormatId: 'fmt-1',
              },
            ],
            rowMetadata: [{ startIndex: 0, count: 1, size: 28 }],
            columnMetadata: [{ startIndex: 0, count: 1, size: 144 }],
            styleRanges: [
              {
                startRow: 0,
                endRow: 0,
                startCol: 0,
                endCol: 0,
                styleId: 'style-1',
              },
            ],
            formatRanges: [
              {
                startRow: 0,
                endRow: 0,
                startCol: 0,
                endCol: 0,
                formatId: 'fmt-1',
              },
            ],
          },
        ],
      },
      'doc-1',
    )

    expect(projected?.workbook.name).toBe('Projected Book')
    expect(projected?.workbook.metadata?.properties).toEqual([{ key: 'locale', value: 'en-US' }])
    expect(projected?.workbook.metadata?.definedNames).toEqual([
      {
        name: 'Sales',
        value: {
          kind: 'range-ref',
          sheetName: 'Sheet1',
          startAddress: 'A1',
          endAddress: 'A2',
        },
      },
    ])
    expect(projected?.workbook.metadata?.styles).toEqual([
      {
        id: 'style-1',
        fill: {
          backgroundColor: '#ffee00',
        },
      },
    ])
    expect(projected?.workbook.metadata?.formats).toEqual([
      {
        id: 'fmt-1',
        code: '0.00',
        kind: 'number',
      },
    ])
    expect(projected?.workbook.metadata?.tables?.[0]?.name).toBe('FallbackTable')
    expect(projected?.sheets[0]?.cells[0]).toEqual({
      address: 'A1',
      value: 42,
      format: '0.00',
    })
    expect(projected?.sheets[0]?.metadata?.freezePane).toEqual({ rows: 1, cols: 2 })
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
  })
})
