import { describe, expect, it } from 'vitest'
import fc from 'fast-check'
import * as XLSX from 'xlsx'
import { formatAddress } from '@bilig/formula'
import { runProperty } from '@bilig/test-fuzz'
import { importCsv, importXlsx } from '../index.js'

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

describe('excel import fuzz', () => {
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

const xlsxCellSpecArbitrary = fc.oneof<XlsxCellSpec>(
  fc.constant({ kind: 'empty' }),
  fc.integer({ min: -999, max: 999 }).map((value) => ({ kind: 'number' as const, value })),
  fc.boolean().map((value) => ({ kind: 'boolean' as const, value })),
  fc.constantFrom('alpha', 'beta', 'gamma', 'delta').map((value) => ({ kind: 'text' as const, value })),
  fc.constantFrom('1+2', 'A1+1', 'B2*2').map((formula) => ({ kind: 'formula' as const, formula })),
)

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
