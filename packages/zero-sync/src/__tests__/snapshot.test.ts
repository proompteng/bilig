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

  it('reconstructs workbook snapshots from Zero schema client row names', () => {
    const projected = projectWorkbookToSnapshot(
      {
        id: 'doc-1',
        name: 'Zero Query Book',
        calcMode: 'manual',
        compatibilityMode: 'odf-1.4',
        recalcEpoch: 11,
        styles: [
          {
            styleId: 'style-1',
            styleJson: {
              fill: { backgroundColor: '#ffee00' },
              font: { bold: true },
            },
          },
        ],
        numberFormats: [
          {
            formatId: 'fmt-1',
            code: '$#,##0.00',
            kind: 'currency',
          },
        ],
        sheets: [
          {
            sheetId: 12,
            name: 'Forecast',
            sortOrder: 0,
            freezeRows: 1,
            freezeCols: 0,
            cells: [
              {
                address: 'A1',
                rowNum: 0,
                colNum: 0,
                inputValue: 1250,
                styleId: 'style-1',
                explicitFormatId: 'fmt-1',
              },
              {
                rowNum: 1,
                colNum: 1,
                formula: 'A1*2',
                formatId: 'fmt-1',
              },
            ],
            rowMetadata: [{ startIndex: 0, count: 1, size: 28 }],
            columnMetadata: [{ startIndex: 0, count: 1, size: 144 }],
          },
        ],
      },
      'doc-1',
    )

    expect(projected).toEqual({
      version: 1,
      workbook: {
        name: 'Zero Query Book',
        metadata: {
          styles: [
            {
              id: 'style-1',
              fill: { backgroundColor: '#ffee00' },
              font: { bold: true },
            },
          ],
          formats: [{ id: 'fmt-1', code: '$#,##0.00', kind: 'currency' }],
          calculationSettings: { mode: 'manual', compatibilityMode: 'odf-1.4' },
          volatileContext: { recalcEpoch: 11 },
        },
      },
      sheets: [
        {
          id: 12,
          name: 'Forecast',
          order: 0,
          cells: [
            { address: 'A1', value: 1250, format: '$#,##0.00' },
            { address: 'B2', formula: 'A1*2', format: '$#,##0.00' },
          ],
          metadata: {
            rowMetadata: [{ start: 0, count: 1, size: 28 }],
            columnMetadata: [{ start: 0, count: 1, size: 144 }],
            styleRanges: [
              {
                range: {
                  sheetName: 'Forecast',
                  startAddress: 'A1',
                  endAddress: 'A1',
                },
                styleId: 'style-1',
              },
            ],
            formatRanges: [
              {
                range: {
                  sheetName: 'Forecast',
                  startAddress: 'A1',
                  endAddress: 'A1',
                },
                formatId: 'fmt-1',
              },
              {
                range: {
                  sheetName: 'Forecast',
                  startAddress: 'B2',
                  endAddress: 'B2',
                },
                formatId: 'fmt-1',
              },
            ],
            freezePane: { rows: 1, cols: 0 },
          },
        },
      ],
    })
  })

  it('preserves fallback axis style metadata when normalized rows omit it', () => {
    const projected = projectWorkbookToSnapshot(
      {
        id: 'doc-1',
        name: 'Projected Book',
        snapshot: {
          version: 1,
          workbook: { name: 'Warm Snapshot' },
          sheets: [
            {
              name: 'Sheet1',
              order: 0,
              cells: [],
              metadata: {
                rowMetadata: [{ start: 0, count: 1, size: 28, styleIndex: 7, customFormat: true }],
                columnMetadata: [{ start: 0, count: 1, size: 144, styleIndex: 9, customFormat: true }],
              },
            },
          ],
        },
        calculationSettings: {
          mode: 'automatic',
          recalcEpoch: 0,
        },
        sheets: [
          {
            name: 'Sheet1',
            sortOrder: 0,
            freezeRows: 0,
            freezeCols: 0,
            cells: [],
            rowMetadata: [{ startIndex: 0, count: 1, size: 28 }],
            columnMetadata: [{ startIndex: 0, count: 1, size: 144 }],
            styleRanges: [],
            formatRanges: [],
          },
        ],
      },
      'doc-1',
    )

    expect(projected?.sheets[0]?.metadata?.rowMetadata).toEqual([{ start: 0, count: 1, size: 28, styleIndex: 7, customFormat: true }])
    expect(projected?.sheets[0]?.metadata?.columnMetadata).toEqual([{ start: 0, count: 1, size: 144, styleIndex: 9, customFormat: true }])
  })

  it('treats present empty normalized collections as authoritative deletes', () => {
    const projected = projectWorkbookToSnapshot(
      {
        id: 'doc-1',
        name: 'Projected Book',
        recalcEpoch: undefined,
        snapshot: {
          version: 1,
          workbook: {
            name: 'Warm Snapshot',
            metadata: {
              properties: [{ key: 'locale', value: 'en-US' }],
              definedNames: [{ name: 'Rate', value: 0.12 }],
              styles: [{ id: 'style-old', fill: { backgroundColor: '#ff0000' } }],
              formats: [{ id: 'fmt-old', code: '0.00', kind: 'number' }],
              calculationSettings: { mode: 'manual', compatibilityMode: 'excel-modern' },
              volatileContext: { recalcEpoch: 12 },
              tables: [
                {
                  name: 'Table1',
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
              id: 1,
              name: 'Sheet1',
              order: 0,
              cells: [{ address: 'A1', value: 'stale' }],
              metadata: {
                rowMetadata: [{ start: 0, count: 1, size: 28, styleIndex: 7 }],
                columnMetadata: [{ start: 0, count: 1, size: 144, styleIndex: 9 }],
                styleRanges: [{ range: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' }, styleId: 'style-old' }],
                formatRanges: [{ range: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' }, formatId: 'fmt-old' }],
                freezePane: { rows: 1, cols: 1 },
                validations: [{ range: { sheetName: 'Sheet1', startAddress: 'B1', endAddress: 'B2' }, type: 'whole' }],
                notes: [{ ref: 'C1', text: 'snapshot-only note' }],
              },
            },
          ],
        },
        calculationSettings: {},
        workbookMetadataEntries: [],
        definedNames: [],
        styles: [],
        numberFormats: [],
        sheets: [
          {
            name: 'Sheet1',
            sortOrder: 0,
            freezeRows: 0,
            freezeCols: 0,
            cells: [],
            rowMetadata: [],
            columnMetadata: [],
            styleRanges: [],
            formatRanges: [],
          },
        ],
      },
      'doc-1',
    )

    expect(projected?.workbook.metadata?.properties).toBeUndefined()
    expect(projected?.workbook.metadata?.definedNames).toBeUndefined()
    expect(projected?.workbook.metadata?.styles).toBeUndefined()
    expect(projected?.workbook.metadata?.formats).toBeUndefined()
    expect(projected?.workbook.metadata?.calculationSettings).toBeUndefined()
    expect(projected?.workbook.metadata?.volatileContext).toBeUndefined()
    expect(projected?.workbook.metadata?.tables?.[0]?.name).toBe('Table1')
    expect(projected?.sheets[0]?.cells).toEqual([])
    expect(projected?.sheets[0]?.metadata?.rowMetadata).toBeUndefined()
    expect(projected?.sheets[0]?.metadata?.columnMetadata).toBeUndefined()
    expect(projected?.sheets[0]?.metadata?.styleRanges).toBeUndefined()
    expect(projected?.sheets[0]?.metadata?.formatRanges).toBeUndefined()
    expect(projected?.sheets[0]?.metadata?.freezePane).toBeUndefined()
    expect(projected?.sheets[0]?.metadata?.validations).toEqual([
      { range: { sheetName: 'Sheet1', startAddress: 'B1', endAddress: 'B2' }, type: 'whole' },
    ])
    expect(projected?.sheets[0]?.metadata?.notes).toEqual([{ ref: 'C1', text: 'snapshot-only note' }])
  })

  it('preserves warm snapshot data when normalized collections are omitted from a partial projection', () => {
    const projected = projectWorkbookToSnapshot(
      {
        id: 'doc-1',
        name: 'Projected Book',
        snapshot: {
          version: 1,
          workbook: {
            name: 'Warm Snapshot',
            metadata: {
              properties: [{ key: 'locale', value: 'en-US' }],
              definedNames: [{ name: 'Rate', value: 0.12 }],
              styles: [{ id: 'style-old', fill: { backgroundColor: '#ff0000' } }],
              formats: [{ id: 'fmt-old', code: '0.00', kind: 'number' }],
              calculationSettings: { mode: 'manual', compatibilityMode: 'excel-modern' },
              volatileContext: { recalcEpoch: 12 },
            },
          },
          sheets: [
            {
              id: 1,
              name: 'Sheet1',
              order: 0,
              cells: [{ address: 'A1', value: 'warm' }],
              metadata: {
                rowMetadata: [{ start: 0, count: 1, size: 28 }],
                styleRanges: [{ range: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' }, styleId: 'style-old' }],
                freezePane: { rows: 1, cols: 1 },
              },
            },
          ],
        },
        sheets: [
          {
            name: 'Sheet1',
            sortOrder: 0,
          },
        ],
      },
      'doc-1',
    )

    expect(projected?.workbook.metadata?.properties).toEqual([{ key: 'locale', value: 'en-US' }])
    expect(projected?.workbook.metadata?.definedNames).toEqual([{ name: 'Rate', value: 0.12 }])
    expect(projected?.workbook.metadata?.styles).toEqual([{ id: 'style-old', fill: { backgroundColor: '#ff0000' } }])
    expect(projected?.workbook.metadata?.formats).toEqual([{ id: 'fmt-old', code: '0.00', kind: 'number' }])
    expect(projected?.workbook.metadata?.calculationSettings).toEqual({ mode: 'manual', compatibilityMode: 'excel-modern' })
    expect(projected?.workbook.metadata?.volatileContext).toEqual({ recalcEpoch: 12 })
    expect(projected?.sheets[0]?.cells).toEqual([{ address: 'A1', value: 'warm' }])
    expect(projected?.sheets[0]?.metadata?.rowMetadata).toEqual([{ start: 0, count: 1, size: 28 }])
    expect(projected?.sheets[0]?.metadata?.styleRanges).toEqual([
      { range: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' }, styleId: 'style-old' },
    ])
    expect(projected?.sheets[0]?.metadata?.freezePane).toEqual({ rows: 1, cols: 1 })
  })

  it('sanitizes projected style records from persisted JSON', () => {
    const projected = projectWorkbookToSnapshot(
      {
        id: 'doc-1',
        name: 'Projected Book',
        styles: [
          {
            id: 'style-1',
            recordJSON: {
              fill: { backgroundColor: '#ffee00' },
              font: {
                family: 'Inter',
                size: 13,
                bold: true,
                underline: false,
                color: '#111111',
                shadow: true,
              },
              alignment: {
                horizontal: 'right',
                vertical: 'middle',
                indent: 2,
                textRotation: 45,
                wrap: true,
                fake: 'ignored',
              },
              borders: {
                top: { style: 'solid', weight: 'thin', color: '#333333' },
                bottom: { style: 'wave', weight: 'thin', color: '#333333' },
              },
              protection: { locked: true, hidden: false },
              arbitrary: { trusted: false },
            },
          },
          {
            id: 'style-unsafe',
            recordJSON: {
              fill: { backgroundColor: 42 },
              font: { size: Number.NaN, bold: 'yes' },
              alignment: { horizontal: 'diagonal', readingOrder: Number.POSITIVE_INFINITY },
              protection: { locked: 1 },
            },
          },
          {
            id: '',
            recordJSON: { fill: { backgroundColor: '#000000' } },
          },
          {
            id: 'style-non-object',
            recordJSON: null,
          },
        ],
        sheets: [],
      },
      'doc-1',
    )

    expect(projected?.workbook.metadata?.styles).toEqual([
      {
        id: 'style-1',
        fill: { backgroundColor: '#ffee00' },
        font: {
          family: 'Inter',
          size: 13,
          bold: true,
          underline: false,
          color: '#111111',
        },
        alignment: {
          horizontal: 'right',
          vertical: 'middle',
          indent: 2,
          textRotation: 45,
          wrap: true,
        },
        borders: {
          top: { style: 'solid', weight: 'thin', color: '#333333' },
        },
        protection: { locked: true, hidden: false },
      },
      { id: 'style-unsafe' },
    ])
  })

  it('drops unsafe projected sheet metadata bounds', () => {
    const projected = projectWorkbookToSnapshot(
      {
        id: 'doc-1',
        name: 'Projected Book',
        sheets: [
          {
            name: 'Sheet1',
            sortOrder: 0,
            freezeRows: 1.5,
            freezeCols: Number.MAX_SAFE_INTEGER + 1,
            cells: [],
            rowMetadata: [
              { startIndex: 0, count: 1, size: 28 },
              { startIndex: -1, count: 1, size: 30 },
              { startIndex: 2, count: 0, size: 30 },
            ],
            columnMetadata: [{ startIndex: 0, count: Number.MAX_SAFE_INTEGER + 1, size: 144 }],
            styleRanges: [
              { startRow: 2, endRow: 1, startCol: 0, endCol: 0, styleId: 'style-1' },
              { startRow: 0, endRow: 0, startCol: 1.5, endCol: 2, styleId: 'style-1' },
            ],
            formatRanges: [{ startRow: 0, endRow: 0, startCol: 2, endCol: 1, formatId: 'fmt-1' }],
          },
        ],
      },
      'doc-1',
    )

    expect(projected?.sheets[0]?.metadata?.freezePane).toBeUndefined()
    expect(projected?.sheets[0]?.metadata?.rowMetadata).toEqual([{ start: 0, count: 1, size: 28 }])
    expect(projected?.sheets[0]?.metadata?.columnMetadata).toBeUndefined()
    expect(projected?.sheets[0]?.metadata?.styleRanges).toBeUndefined()
    expect(projected?.sheets[0]?.metadata?.formatRanges).toBeUndefined()
  })

  it('drops malformed defined-name values', () => {
    const projected = projectWorkbookToSnapshot(
      {
        id: 'doc-1',
        name: 'Projected Book',
        definedNames: [
          { name: 'Rate', value: 0.12 },
          { name: 'Scalar', value: { kind: 'scalar', value: 'ok' } },
          { name: 'Cell', value: { kind: 'cell-ref', sheetName: 'Sheet1', address: 'A1' } },
          { name: 'BadNumber', value: Number.NaN },
          { name: 'BadScalar', value: { kind: 'scalar', value: Number.POSITIVE_INFINITY } },
          { name: 'BadObject', value: { kind: 'range-ref', sheetName: 'Sheet1', startAddress: 'A1' } },
          { name: 'ArbitraryObject', value: { formula: 'A1' } },
        ],
        sheets: [],
      },
      'doc-1',
    )

    expect(projected?.workbook.metadata?.definedNames).toEqual([
      { name: 'Rate', value: 0.12 },
      { name: 'Scalar', value: { kind: 'scalar', value: 'ok' } },
      { name: 'Cell', value: { kind: 'cell-ref', sheetName: 'Sheet1', address: 'A1' } },
    ])
  })

  it('drops unsafe projected sheet and cell coordinates', () => {
    const projected = projectWorkbookToSnapshot(
      {
        id: 'doc-1',
        name: 'Projected Book',
        recalcEpoch: 2.5,
        sheets: [
          {
            id: Number.MAX_SAFE_INTEGER + 1,
            name: 'BadOrder',
            sortOrder: 1.5,
            cells: [{ rowNum: 0, colNum: 0, inputValue: 'ignored' }],
          },
          {
            id: 2,
            name: 'Sheet1',
            sortOrder: 0,
            cells: [
              { rowNum: 1, colNum: 2, inputValue: 'kept' },
              { rowNum: -1, colNum: 0, inputValue: 'dropped' },
              { rowNum: 0.5, colNum: 0, inputValue: 'dropped' },
            ],
            rowMetadata: [{ startIndex: 0, count: 1, size: -10 }],
          },
        ],
      },
      'doc-1',
    )

    expect(projected?.workbook.metadata?.volatileContext).toBeUndefined()
    expect(projected?.sheets).toEqual([
      {
        id: 2,
        name: 'Sheet1',
        order: 0,
        cells: [{ address: 'C2', value: 'kept' }],
        metadata: { rowMetadata: [{ start: 0, count: 1 }] },
      },
    ])
  })
})
