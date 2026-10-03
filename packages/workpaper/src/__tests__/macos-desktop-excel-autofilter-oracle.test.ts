import { readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'

import { SpreadsheetEngine } from '@bilig/core'
import { exportXlsx, importXlsx } from '@bilig/excel-import'
import { isMacosExcelInstalled, runMacosExcelInspectionOracle, runMacosExcelStructuralOperationOracle } from '@bilig/excel-fixtures'
import { ValueTag, type WorkbookSnapshot } from '@bilig/protocol'
import { describe, expect, it } from 'vitest'

import { createExcelAccessibleTempDir, removeMacosExcelTestDir } from './macos-excel-oracle-test-utils.js'

const inspectedCells = ['F2', 'F3'] as const
const expectedExcelSubtotal = { kind: 'number', value: 90 }
const expectedEngineSubtotal = { tag: ValueTag.Number, value: 90 }

describe('macOS Desktop Excel AutoFilter oracle', () => {
  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'matches Desktop Excel table AutoFilter execution and SUBTOTAL row visibility',
    async () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = createExcelAccessibleTempDir('bilig-headless-excel-autofilter-oracle-')
      try {
        const excelWorkbookPath = join(tempDir, 'excel-table-autofilter-oracle.xlsx')
        writeFileSync(excelWorkbookPath, exportXlsx(tableFilterSnapshot()))

        const excelResult = runMacosExcelStructuralOperationOracle({
          workbookPath: excelWorkbookPath,
          worksheetName: 'Ledger',
          operations: [{ kind: 'applyTableAutoFilter', tableName: 'Sales', field: 1, criteria1: 'East' }],
          inspectCells: inspectedCells,
          saveWorkbook: true,
          timeoutMs: 90_000,
        })

        expect(excelResult.cells.map((cell) => cell.value)).toEqual([expectedExcelSubtotal, expectedExcelSubtotal])

        const excelTruth = importXlsx(new Uint8Array(readFileSync(excelWorkbookPath)), 'excel-table-autofilter-oracle.xlsx')
        expect(excelTruth.snapshot.workbook.metadata?.tables?.[0]?.autoFilter).toEqual({
          sheetName: 'Ledger',
          startAddress: 'A1',
          endAddress: 'D6',
          criteria: [{ colId: 0, filters: { values: ['East'] } }],
        })
        expect(rowFilterHiddenStarts(excelTruth.snapshot)).toEqual([2, 4])

        const headless = new SpreadsheetEngine({ workbookName: 'headless-table-autofilter-oracle' })
        await headless.ready()
        headless.importSnapshot(tableFilterSnapshot())
        expect(headless.applyTableAutoFilter('Ledger', 'Sales', [{ colId: 0, filters: { values: ['East'] } }])).toBe(true)

        expect(headless.getCellValue('Ledger', 'F2')).toEqual(expectedEngineSubtotal)
        expect(headless.getCellValue('Ledger', 'F3')).toEqual(expectedEngineSubtotal)
        expect(rowFilterHiddenStarts(headless.exportSnapshot())).toEqual([2, 4])

        const headlessWorkbookPath = join(tempDir, 'headless-table-autofilter-oracle.xlsx')
        writeFileSync(headlessWorkbookPath, exportXlsx(headless.exportSnapshot()))
        const headlessExcel = runMacosExcelInspectionOracle({
          workbookPath: headlessWorkbookPath,
          worksheetName: 'Ledger',
          formulaCells: [],
          inspectCells: inspectedCells,
          saveWorkbook: true,
          timeoutMs: 90_000,
        })

        expect(headlessExcel.cells.map((cell) => cell.value)).toEqual(excelResult.cells.map((cell) => cell.value))
        const headlessExcelTruth = importXlsx(new Uint8Array(readFileSync(headlessWorkbookPath)), 'headless-table-autofilter-oracle.xlsx')
        expect(headlessExcelTruth.snapshot.workbook.metadata?.tables?.[0]?.autoFilter).toEqual(
          excelTruth.snapshot.workbook.metadata?.tables?.[0]?.autoFilter,
        )
        expect(rowFilterHiddenStarts(headlessExcelTruth.snapshot)).toEqual([2, 4])
      } finally {
        removeMacosExcelTestDir(tempDir)
      }
    },
    120_000,
  )
})

function tableFilterSnapshot(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: {
      name: 'Table AutoFilter oracle',
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
        cells: [
          { address: 'A1', value: 'Region' },
          { address: 'B1', value: 'Amount' },
          { address: 'C1', value: 'Invoice' },
          { address: 'D1', value: 'Double' },
          { address: 'F1', value: 'Visible total' },
          { address: 'A2', value: 'East' },
          { address: 'B2', value: 10 },
          { address: 'C2', value: 'invoice-001' },
          { address: 'D2', formula: 'B2*2', value: 20 },
          { address: 'F2', formula: 'SUBTOTAL(9,B2:B6)', value: 150 },
          { address: 'A3', value: 'West' },
          { address: 'B3', value: 40 },
          { address: 'C3', value: 'invoice-002' },
          { address: 'D3', formula: 'B3*2', value: 80 },
          { address: 'F3', formula: 'SUBTOTAL(109,B2:B6)', value: 150 },
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

function rowFilterHiddenStarts(snapshot: WorkbookSnapshot): number[] {
  return (
    snapshot.sheets[0]?.metadata?.rowMetadata
      ?.flatMap((record) => (record.filterHidden === true ? [record.start] : []))
      .toSorted((left, right) => left - right) ?? []
  )
}
