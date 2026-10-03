import { readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'

import { SpreadsheetEngine } from '@bilig/core'
import { exportXlsx, importXlsx } from '@bilig/excel-import'
import { isMacosExcelInstalled, runMacosExcelInspectionOracle, runMacosExcelStructuralOperationOracle } from '@bilig/excel-fixtures'
import { ValueTag, type CellValue, type WorkbookSnapshot } from '@bilig/protocol'
import { describe, expect, it } from 'vitest'

import { createExcelAccessibleTempDir, removeMacosExcelTestDir } from './macos-excel-oracle-test-utils.js'

const inspectedCells = [
  'A1',
  'B1',
  'C1',
  'D1',
  'A2',
  'B2',
  'C2',
  'D2',
  'A3',
  'B3',
  'C3',
  'D3',
  'A4',
  'B4',
  'C4',
  'D4',
  'A5',
  'B5',
  'C5',
  'D5',
  'A6',
  'B6',
  'C6',
  'D6',
] as const

const expectedSortedRows = [
  ['East', 50, 'invoice-005', 100],
  ['West', 40, 'invoice-002', 80],
  ['East', 30, 'invoice-003', 60],
  ['West', 20, 'invoice-004', 40],
  ['East', 10, 'invoice-001', 20],
] as const

const expectedSortedFormulas = [
  { address: 'D2', formula: 'B2*2' },
  { address: 'D3', formula: 'B3*2' },
  { address: 'D4', formula: 'B4*2' },
  { address: 'D5', formula: 'B5*2' },
  { address: 'D6', formula: 'B6*2' },
] as const

describe('macOS Desktop Excel sort oracle', () => {
  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'matches Desktop Excel table-body sort row-bundle semantics',
    async () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = createExcelAccessibleTempDir('bilig-headless-excel-sort-oracle-')
      try {
        const excelWorkbookPath = join(tempDir, 'excel-table-sort-oracle.xlsx')
        writeFileSync(excelWorkbookPath, exportXlsx(tableLedgerSnapshot()))

        const excelResult = runMacosExcelStructuralOperationOracle({
          workbookPath: excelWorkbookPath,
          worksheetName: 'Ledger',
          operations: [
            {
              kind: 'applyTableSort',
              tableName: 'Sales',
              keys: [{ key: 'B2:B6', order: 'descending' }],
              header: 'yes',
              orientation: 'rows',
            },
          ],
          inspectCells: inspectedCells,
          saveWorkbook: true,
          timeoutMs: 90_000,
        })

        expect(excelRows(excelResult.cells)).toEqual(expectedSortedRows)
        expect(oracleFormulas(excelResult.cells)).toEqual(expectedSortedFormulas)

        const excelTruth = importXlsx(new Uint8Array(readFileSync(excelWorkbookPath)), 'excel-table-sort-oracle.xlsx')
        expect(importedRows(excelTruth.snapshot)).toEqual(expectedSortedRows)
        expect(snapshotFormulas(excelTruth.snapshot)).toEqual(expectedSortedFormulas)
        expect(excelTruth.snapshot.workbook.metadata?.tables?.[0]).toMatchObject({
          name: 'Sales',
          startAddress: 'A1',
          endAddress: 'D6',
          totalsRow: false,
        })
        const excelSortState = excelTruth.snapshot.workbook.metadata?.tables?.[0]?.sortState
        expect(excelSortState).toContain('ref="A2:D6"')
        expect(excelSortState).toContain('descending="1"')
        expect(excelSortState).toContain('ref="B2:B6"')
        expect(excelTruth.snapshot.sheets[0]?.metadata?.commentThreads?.[0]).toMatchObject({
          address: 'C2',
          comments: [{ body: 'Approved largest invoice' }],
        })
        expect(excelTruth.snapshot.sheets[0]?.metadata?.validations?.[0]?.range).toEqual({
          sheetName: 'Ledger',
          startAddress: 'B6',
          endAddress: 'B6',
        })

        const headless = new SpreadsheetEngine({ workbookName: 'headless-table-sort-oracle' })
        await headless.ready()
        headless.importSnapshot(tableLedgerSnapshot())
        expect(headless.sortTable('Ledger', 'Sales', [{ keyAddress: 'B1', direction: 'desc' }])).toBe(true)

        expect(engineRows(headless)).toEqual(expectedSortedRows)
        expect(engineFormulas(headless)).toEqual(expectedSortedFormulas)
        expect(headless.getCommentThreads('Ledger')).toEqual([
          {
            threadId: 'thread-largest-invoice',
            sheetName: 'Ledger',
            address: 'C2',
            comments: [{ id: 'comment-largest-invoice', body: 'Approved largest invoice', authorDisplayName: 'Audit' }],
          },
        ])
        expect(headless.getNote('Ledger', 'A2')).toEqual({ sheetName: 'Ledger', address: 'A2', text: 'largest invoice note' })
        expect(headless.getDataValidation('Ledger', { sheetName: 'Ledger', startAddress: 'B6', endAddress: 'B6' })).toMatchObject({
          rule: { kind: 'whole', operator: 'greaterThan', values: [0] },
        })
        expect(headless.getTable('Sales')?.sortState).toBe('<sortState ref="A2:D6"><sortCondition descending="1" ref="B2:B6"/></sortState>')

        const headlessWorkbookPath = join(tempDir, 'headless-table-sort-oracle.xlsx')
        writeFileSync(headlessWorkbookPath, exportXlsx(headless.exportSnapshot()))
        const headlessExcel = runMacosExcelInspectionOracle({
          workbookPath: headlessWorkbookPath,
          worksheetName: 'Ledger',
          formulaCells: [],
          inspectCells: inspectedCells,
          saveWorkbook: true,
          timeoutMs: 90_000,
        })

        expect(headlessExcel.cells.map(({ address, value }) => ({ address, value }))).toEqual(
          excelResult.cells.map(({ address, value }) => ({ address, value })),
        )
        expect(oracleFormulas(headlessExcel.cells)).toEqual(expectedSortedFormulas)
        const headlessExcelTruth = importXlsx(new Uint8Array(readFileSync(headlessWorkbookPath)), 'headless-table-sort-oracle.xlsx')
        expect(snapshotFormulas(headlessExcelTruth.snapshot)).toEqual(expectedSortedFormulas)
        expect(headlessExcelTruth.snapshot.workbook.metadata?.tables?.[0]?.sortState).toContain('<sortState')
        expect(headlessExcelTruth.snapshot.sheets[0]?.metadata?.validations).toEqual([
          {
            range: { sheetName: 'Ledger', startAddress: 'B6', endAddress: 'B6' },
            rule: { kind: 'whole', operator: 'greaterThan', values: [0] },
          },
        ])
        expect(headlessExcelTruth.snapshot.sheets[0]?.metadata?.commentThreads?.[0]).toMatchObject({
          address: 'C2',
          comments: [{ body: 'Approved largest invoice' }],
        })
      } finally {
        removeMacosExcelTestDir(tempDir)
      }
    },
    120_000,
  )
})

function tableLedgerSnapshot(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: {
      name: 'Table sort oracle',
      metadata: {
        tables: [
          {
            name: 'Sales',
            sheetName: 'Ledger',
            startAddress: 'A1',
            endAddress: 'D6',
            columnNames: ['Region', 'Amount', 'Invoice', 'Double'],
            columns: [{ name: 'Region' }, { name: 'Amount' }, { name: 'Invoice' }, { name: 'Double' }],
            headerRow: true,
            totalsRow: false,
          },
        ],
      },
    },
    sheets: [
      {
        id: 1,
        name: 'Ledger',
        order: 0,
        metadata: {
          commentThreads: [
            {
              threadId: 'thread-largest-invoice',
              sheetName: 'Ledger',
              address: 'C6',
              comments: [{ id: 'comment-largest-invoice', body: 'Approved largest invoice', authorDisplayName: 'Audit' }],
            },
          ],
          notes: [{ sheetName: 'Ledger', address: 'A6', text: 'largest invoice note' }],
          validations: [
            {
              range: { sheetName: 'Ledger', startAddress: 'B6', endAddress: 'B6' },
              rule: { kind: 'whole', operator: 'greaterThan', values: [0] },
            },
          ],
        },
        cells: [
          { address: 'A1', value: 'Region' },
          { address: 'B1', value: 'Amount' },
          { address: 'C1', value: 'Invoice' },
          { address: 'D1', value: 'Double' },
          { address: 'A2', value: 'East' },
          { address: 'B2', value: 10 },
          { address: 'C2', value: 'invoice-001' },
          { address: 'D2', formula: 'B2*2', value: 20 },
          { address: 'A3', value: 'West' },
          { address: 'B3', value: 40 },
          { address: 'C3', value: 'invoice-002' },
          { address: 'D3', formula: 'B3*2', value: 80 },
          { address: 'A4', value: 'East' },
          { address: 'B4', value: 30 },
          { address: 'C4', value: 'invoice-003' },
          { address: 'D4', formula: 'B4*2', value: 60 },
          { address: 'A5', value: 'West' },
          { address: 'B5', value: 20 },
          { address: 'C5', value: 'invoice-004' },
          { address: 'D5', formula: 'B5*2', value: 40 },
          { address: 'A6', value: 'East' },
          { address: 'B6', value: 50 },
          { address: 'C6', value: 'invoice-005' },
          { address: 'D6', formula: 'B6*2', value: 100 },
        ],
      },
    ],
  }
}

function excelRows(cells: readonly { readonly value: { readonly value?: unknown } }[]): unknown[][] {
  return [4, 8, 12, 16, 20].map((start) => [
    cells[start]?.value.value,
    cells[start + 1]?.value.value,
    cells[start + 2]?.value.value,
    cells[start + 3]?.value.value,
  ])
}

function importedRows(snapshot: WorkbookSnapshot): unknown[][] {
  const sheet = snapshot.sheets[0]
  const values = new Map(sheet?.cells.map((cell) => [cell.address, cell.value]) ?? [])
  return [
    [values.get('A2'), values.get('B2'), values.get('C2'), values.get('D2')],
    [values.get('A3'), values.get('B3'), values.get('C3'), values.get('D3')],
    [values.get('A4'), values.get('B4'), values.get('C4'), values.get('D4')],
    [values.get('A5'), values.get('B5'), values.get('C5'), values.get('D5')],
    [values.get('A6'), values.get('B6'), values.get('C6'), values.get('D6')],
  ]
}

function engineRows(engine: SpreadsheetEngine): unknown[][] {
  return engine.getRangeValues({ sheetName: 'Ledger', startAddress: 'A2', endAddress: 'D6' }).map((row) => row.map(cellValue))
}

function oracleFormulas(
  cells: readonly { readonly address: string; readonly formula?: string }[],
): readonly { readonly address: string; readonly formula: string }[] {
  return expectedSortedFormulas.map(({ address }) => ({
    address,
    formula: normalizeFormula(cells.find((cell) => cell.address === address)?.formula),
  }))
}

function snapshotFormulas(snapshot: WorkbookSnapshot): readonly { readonly address: string; readonly formula: string }[] {
  const formulas = new Map(snapshot.sheets[0]?.cells.map((cell) => [cell.address, cell.formula]) ?? [])
  return expectedSortedFormulas.map(({ address }) => ({
    address,
    formula: normalizeFormula(formulas.get(address)),
  }))
}

function engineFormulas(engine: SpreadsheetEngine): readonly { readonly address: string; readonly formula: string }[] {
  return expectedSortedFormulas.map(({ address }) => ({
    address,
    formula: normalizeFormula(engine.getCell('Ledger', address).formula),
  }))
}

function normalizeFormula(formula: string | undefined): string {
  if (formula === undefined) {
    throw new Error('Missing expected formula')
  }
  return formula.startsWith('=') ? formula.slice(1) : formula
}

function cellValue(value: CellValue): unknown {
  switch (value.tag) {
    case ValueTag.Number:
    case ValueTag.String:
    case ValueTag.Boolean:
      return value.value
    case ValueTag.Empty:
    case ValueTag.Error:
      return null
    default:
      return null
  }
}
