import { describe, expect, it } from 'vitest'
import * as fc from 'fast-check'
import * as XLSX from 'xlsx'
import { formatAddress } from '@bilig/formula'
import type { WorkbookSnapshot } from '@bilig/protocol'
import { runProperty } from '@bilig/test-fuzz'
import { exportXlsx, importCsv, importXlsx } from '../index.js'
import { decodeExcelEscapedText, encodeExcelEscapedText } from '../xlsx-escaped-text.js'

type CsvCellSpec =
  | { kind: 'empty'; raw: '' }
  | { kind: 'number'; raw: string; value: number }
  | { kind: 'text'; raw: string; value: string }
  | { kind: 'formula'; raw: string; formula: string }

type XlsxCellSpec =
  | { kind: 'empty' }
  | { kind: 'number'; value: number }
  | { kind: 'boolean'; value: boolean }
  | { kind: 'text'; value: string }
  | { kind: 'formula'; formula: string }

type XlsxEscapedTextCellSpec = {
  address: string
  value: string
}

type XlsxEscapedTextWorkbookSpec = {
  fileStem: string
  cells: XlsxEscapedTextCellSpec[]
  rows: AxisEntrySnapshot[]
  columns: AxisEntrySnapshot[]
}

type AxisEntrySnapshot = NonNullable<NonNullable<WorkbookSnapshot['sheets'][number]['metadata']>['rows']>[number]

describe('excel import fuzz', () => {
  it('should encode generated Excel escaped text literals reversibly', async () => {
    await runProperty({
      suite: 'excel-import/xlsx/escaped-text-codec',
      arbitrary: escapedTextArbitrary,
      predicate: async (value) => {
        expect(decodeExcelEscapedText(encodeExcelEscapedText(value))).toBe(value)
      },
      parameters: { numRuns: 120 },
    })
  })

  it('should preserve csv preview and snapshot semantics for generated small tables', async () => {
    await runProperty({
      suite: 'excel-import/csv/generated-tables',
      arbitrary: fc
        .record({
          fileStem: fc.constantFrom('metrics', 'report', 'summary', 'imported-book'),
          rows: fc.array(fc.array(csvCellSpecArbitrary, { minLength: 1, maxLength: 4 }), { minLength: 1, maxLength: 5 }),
        })
        .filter((record) => record.rows.some((row) => row.some((cell) => cell.kind !== 'empty'))),
      predicate: async ({ fileStem, rows }) => {
        const normalizedRows = rows.map((row) => (row.length > 4 ? row.slice(0, 4) : row))
        const width = Math.max(...normalizedRows.map((row) => row.length))
        const paddedRows = normalizedRows.map((row) =>
          row.concat(Array.from({ length: width - row.length }, () => ({ kind: 'empty', raw: '' }) as const)),
        )
        const csv = paddedRows.map((row) => row.map((cell) => cell.raw).join(',')).join('\n')
        const imported = importCsv(csv, `${fileStem}.csv`)

        expect(imported.workbookName).toBe(fileStem)
        expect(imported.sheetNames).toEqual([fileStem])
        expect(imported.preview.sheetCount).toBe(1)
        expect(imported.preview.sheets[0]?.rowCount).toBe(paddedRows.length)
        expect(imported.preview.sheets[0]?.columnCount).toBe(width)
        expect(imported.preview.sheets[0]?.previewRows).toEqual(paddedRows.map((row) => row.map((cell) => cell.raw)))

        const expectedCells = paddedRows.flatMap((row, rowIndex) =>
          row.flatMap((cell, colIndex) => {
            const address = formatAddress(rowIndex, colIndex)
            switch (cell.kind) {
              case 'empty':
                return []
              case 'number':
                return [{ address, value: cell.value }]
              case 'text':
                return [{ address, value: cell.value }]
              case 'formula':
                return [{ address, formula: cell.formula }]
            }
          }),
        )

        expect(imported.preview.sheets[0]?.nonEmptyCellCount).toBe(expectedCells.length)
        expect(imported.snapshot.sheets[0]?.cells).toEqual(expectedCells)
      },
    })
  })

  it('should preserve generated escaped text and axis entries across XLSX export/import cycles', async () => {
    await runProperty({
      suite: 'excel-import/xlsx/escaped-text-axis-roundtrip',
      arbitrary: xlsxEscapedTextWorkbookSpecArbitrary,
      predicate: async (spec) => {
        const snapshot = buildEscapedTextWorkbookSnapshot(spec)
        const imported = importXlsx(exportXlsx(snapshot), `${spec.fileStem}.xlsx`).snapshot
        const reimported = importXlsx(exportXlsx(imported), `${spec.fileStem}-again.xlsx`).snapshot

        expect(escapedTextWorkbookDigest(imported)).toEqual(escapedTextWorkbookDigest(snapshot))
        expect(escapedTextWorkbookDigest(reimported)).toEqual(escapedTextWorkbookDigest(imported))
      },
      parameters: { numRuns: 80 },
    })
  })

  it('should preserve xlsx workbook semantics for generated multi-sheet inputs', async () => {
    await runProperty({
      suite: 'excel-import/xlsx/generated-workbooks',
      arbitrary: fc.record({
        fileStem: fc.constantFrom('Quarterly Report', 'Budget Workbook', 'Ops Review'),
        sheets: fc.uniqueArray(
          fc.record({
            name: fc.constantFrom('Alpha', 'Beta', 'Gamma'),
            cells: fc.array(fc.array(xlsxCellSpecArbitrary, { minLength: 1, maxLength: 3 }), { minLength: 1, maxLength: 3 }),
          }),
          {
            minLength: 1,
            maxLength: 2,
            selector: (sheet) => sheet.name,
          },
        ),
      }),
      predicate: async ({ fileStem, sheets }) => {
        const workbook = XLSX.utils.book_new()
        const expectedSheets = sheets.map((sheet, order) => {
          const rowCount = sheet.cells.length
          const colCount = Math.max(...sheet.cells.map((row) => row.length))
          const matrix = sheet.cells.map((row) => [
            ...row,
            ...Array.from({ length: colCount - row.length }, () => ({ kind: 'empty' }) as const),
          ])
          const aoa = matrix.map((row) =>
            row.map((cell) => {
              switch (cell.kind) {
                case 'empty':
                case 'formula':
                  return undefined
                case 'number':
                case 'boolean':
                case 'text':
                  return cell.value
              }
            }),
          )
          const worksheet = XLSX.utils.aoa_to_sheet(aoa)
          worksheet['!ref'] = `${formatAddress(0, 0)}:${formatAddress(rowCount - 1, colCount - 1)}`
          const expectedCells = matrix.flatMap((row, rowIndex) =>
            row.flatMap((cell, colIndex) => {
              const address = formatAddress(rowIndex, colIndex)
              if (cell.kind === 'formula') {
                worksheet[address] = { t: 'n', f: cell.formula }
                return [{ address, formula: cell.formula }]
              }
              if (cell.kind === 'empty') {
                return []
              }
              return [{ address, value: cell.value }]
            }),
          )
          XLSX.utils.book_append_sheet(workbook, worksheet, sheet.name)
          return {
            name: sheet.name,
            order,
            expectedCells,
          }
        })

        const imported = importXlsx(XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' }), `${fileStem}.xlsx`)

        expect(imported.workbookName).toBe(fileStem)
        expect(imported.sheetNames).toEqual(expectedSheets.map((sheet) => sheet.name))
        expect(imported.snapshot.sheets.map((sheet) => ({ name: sheet.name, order: sheet.order }))).toEqual(
          expectedSheets.map((sheet) => ({ name: sheet.name, order: sheet.order })),
        )
        expectedSheets.forEach((expectedSheet, index) => {
          expect(normalizeImportedCells(imported.snapshot.sheets[index]?.cells ?? [])).toEqual(expectedSheet.expectedCells)
        })
      },
    })
  })
})

// Helpers

const csvCellSpecArbitrary = fc.oneof<CsvCellSpec>(
  fc.constant({ kind: 'empty', raw: '' }),
  fc.integer({ min: -999, max: 999 }).map((value) => ({
    kind: 'number' as const,
    raw: String(value),
    value,
  })),
  fc.constantFrom('north', 'south', 'ready', 'done').map((value) => ({
    kind: 'text' as const,
    raw: value,
    value,
  })),
  fc.constantFrom('A1', 'B2', 'C3', 'D4').map((formula) => ({
    kind: 'formula' as const,
    raw: `=${formula}`,
    formula,
  })),
)

const escapedTextArbitrary = fc
  .oneof(
    fc.constantFrom(
      'simple',
      ' leading and trailing ',
      '<tag attr="value">&text</tag>',
      'quote " apostrophe \' ampersand &',
      'line\nbreak',
      'tab\tseparated',
      'literal _x000D_ token',
      '_x000D_',
      '_x005F_x000D_',
      '_x0041_',
      'control \u0001 value',
    ),
    fc
      .array(fc.constantFrom('alpha', '&', '<', '>', '"', "'", '_', '_x0009_', '_x005F_', '\n', '\t', '\u0001'), {
        minLength: 1,
        maxLength: 6,
      })
      .map((parts) => parts.join('')),
  )
  .filter((value) => value.length > 0)

const xlsxEscapedTextWorkbookSpecArbitrary = fc.record({
  fileStem: fc.constantFrom('Escaped Text', 'Import Fidelity', 'Axis Entries'),
  cells: fc.uniqueArray(
    fc.record({
      address: fc.constantFrom('A1', 'B1', 'C2', 'D3', 'E4', 'B5'),
      value: escapedTextArbitrary,
    }),
    {
      minLength: 1,
      maxLength: 6,
      selector: (cell) => cell.address,
    },
  ),
  rows: fc.uniqueArray(axisEntryArbitrary('row'), {
    minLength: 0,
    maxLength: 4,
    selector: (entry) => entry.index,
  }),
  columns: fc.uniqueArray(axisEntryArbitrary('col'), {
    minLength: 0,
    maxLength: 4,
    selector: (entry) => entry.index,
  }),
})

const xlsxCellSpecArbitrary = fc.oneof<XlsxCellSpec>(
  fc.constant({ kind: 'empty' }),
  fc.integer({ min: -999, max: 999 }).map((value) => ({ kind: 'number' as const, value })),
  fc.boolean().map((value) => ({ kind: 'boolean' as const, value })),
  fc.constantFrom('alpha', 'beta', 'gamma', 'delta').map((value) => ({ kind: 'text' as const, value })),
  fc.constantFrom('1+2', 'A1+1', 'B2*2').map((formula) => ({ kind: 'formula' as const, formula })),
)

function axisEntryArbitrary(idPrefix: 'row' | 'col'): fc.Arbitrary<AxisEntrySnapshot> {
  return fc
    .record({
      index: fc.integer({ min: 0, max: 4 }),
      metadata: fc.oneof(
        fc.record({ size: fc.integer({ min: 12, max: 96 }) }),
        fc.record({ hidden: fc.constant(true) }),
        fc.record({ size: fc.integer({ min: 12, max: 96 }), hidden: fc.constant(true) }),
      ),
    })
    .map(({ index, metadata }) => generatedAxisEntry(idPrefix, index, metadata))
}

function buildEscapedTextWorkbookSnapshot(spec: XlsxEscapedTextWorkbookSpec): WorkbookSnapshot {
  return {
    version: 1,
    workbook: { name: spec.fileStem },
    sheets: [
      {
        id: 1,
        name: 'Data',
        order: 0,
        cells: spec.cells.toSorted(compareCellsByAddress),
        metadata: {
          rows: spec.rows.toSorted((left, right) => left.index - right.index),
          columns: spec.columns.toSorted((left, right) => left.index - right.index),
        },
      },
    ],
  }
}

function escapedTextWorkbookDigest(snapshot: WorkbookSnapshot): unknown {
  return snapshot.sheets
    .toSorted((left, right) => left.order - right.order)
    .map((sheet) => ({
      name: sheet.name,
      order: sheet.order,
      cells: sheet.cells.map((cell) => ({ address: cell.address, value: cell.value })).toSorted(compareCellsByAddress),
      rows: normalizeAxisEntries(sheet.metadata?.rows),
      columns: normalizeAxisEntries(sheet.metadata?.columns),
    }))
}

function compareCellsByAddress(left: { address: string }, right: { address: string }): number {
  const leftCell = XLSX.utils.decode_cell(left.address)
  const rightCell = XLSX.utils.decode_cell(right.address)
  return leftCell.r - rightCell.r || leftCell.c - rightCell.c
}

function normalizeAxisEntries(entries: readonly AxisEntrySnapshot[] | undefined): AxisEntrySnapshot[] {
  return (entries ?? []).map((entry) => normalizedAxisEntry(entry)).toSorted((left, right) => left.index - right.index)
}

function generatedAxisEntry(
  idPrefix: 'row' | 'col',
  index: number,
  metadata: { readonly size?: number | undefined; readonly hidden?: boolean | undefined },
): AxisEntrySnapshot {
  const entry: AxisEntrySnapshot = {
    id: `${idPrefix}:${String(index)}`,
    index,
  }
  if (metadata.size !== undefined) {
    entry.size = metadata.size
  }
  if (metadata.hidden !== undefined) {
    entry.hidden = metadata.hidden
  }
  return entry
}

function normalizedAxisEntry(entry: AxisEntrySnapshot): AxisEntrySnapshot {
  const normalized: AxisEntrySnapshot = {
    id: entry.id,
    index: entry.index,
  }
  if (entry.size !== undefined && entry.size !== null) {
    normalized.size = entry.size
  }
  if (entry.hidden !== undefined && entry.hidden !== null) {
    normalized.hidden = entry.hidden
  }
  return normalized
}

function normalizeImportedCells(
  cells: ReadonlyArray<{ address: string; value?: unknown; formula?: string; format?: string }>,
): Array<{ address: string; value?: unknown; formula?: string }> {
  return cells.map((cell) => {
    if (cell.formula !== undefined) {
      return { address: cell.address, formula: cell.formula }
    }
    return { address: cell.address, ...(cell.value !== undefined ? { value: cell.value } : {}) }
  })
}
