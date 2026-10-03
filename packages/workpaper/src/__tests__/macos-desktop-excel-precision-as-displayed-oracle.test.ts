import { readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'

import { SpreadsheetEngine } from '@bilig/core'
import { exportXlsx } from '@bilig/excel-import'
import { isMacosExcelInstalled, runMacosExcelInspectionOracle, type NormalizedFormulaValue } from '@bilig/excel-fixtures'
import { ValueTag, type CellValue, type WorkbookSnapshot } from '@bilig/protocol'
import { describe, expect, it } from 'vitest'

import { createExcelAccessibleTempDir, removeMacosExcelTestDir } from './macos-excel-oracle-test-utils.js'

const inspectedCells = ['A2', 'B2', 'C2', 'A3', 'B3', 'C3'] as const
const expectedExcelValues: readonly NormalizedFormulaValue[] = [
  { kind: 'number', value: 1.234 },
  { kind: 'number', value: 2.5 },
  { kind: 'number', value: 25 },
  { kind: 'number', value: 0.12345 },
  { kind: 'number', value: 0.123 },
  { kind: 'number', value: 12.3 },
] as const

describe('macOS Desktop Excel precision-as-displayed oracle', () => {
  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'matches Desktop Excel precision-as-displayed recalculation for decimal and percent inputs',
    async () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = createExcelAccessibleTempDir('bilig-headless-precision-displayed-oracle-')
      try {
        const sourcePath = join(tempDir, 'precision-as-displayed.xlsx')
        writeFileSync(sourcePath, exportXlsx(precisionAsDisplayedSnapshot()))

        const excelResult = runMacosExcelInspectionOracle({
          workbookPath: sourcePath,
          worksheetName: 'Cases',
          formulaCells: [],
          inspectCells: inspectedCells,
          saveWorkbook: true,
          timeoutMs: 90_000,
        })

        expect(excelResult.cells.map((cell) => cell.value)).toEqual(expectedExcelValues)

        const headless = new SpreadsheetEngine({ workbookName: 'headless-precision-as-displayed' })
        await headless.ready()
        headless.importSnapshot(precisionAsDisplayedSnapshot())
        headless.recalculateNow()

        expect(inspectedCells.map((address) => normalizedCellValue(headless.getCellValue('Cases', address)))).toEqual(expectedExcelValues)

        const headlessPath = join(tempDir, 'headless-precision-as-displayed.xlsx')
        writeFileSync(headlessPath, exportXlsx(headless.exportSnapshot()))
        const headlessExcel = runMacosExcelInspectionOracle({
          workbookPath: headlessPath,
          worksheetName: 'Cases',
          formulaCells: [],
          inspectCells: inspectedCells,
          saveWorkbook: true,
          timeoutMs: 90_000,
        })

        expect(headlessExcel.cells.map((cell) => cell.value)).toEqual(excelResult.cells.map((cell) => cell.value))
        expect(readFileSync(headlessPath).byteLength).toBeGreaterThan(0)
      } finally {
        removeMacosExcelTestDir(tempDir)
      }
    },
    120_000,
  )
})

function precisionAsDisplayedSnapshot(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: {
      name: 'Precision As Displayed Oracle',
      metadata: {
        calculationSettings: {
          mode: 'automatic',
          compatibilityMode: 'excel-modern',
          fullPrecision: false,
        },
      },
    },
    sheets: [
      {
        id: 1,
        name: 'Cases',
        order: 0,
        cells: [
          { address: 'A1', value: 'Decimal input' },
          { address: 'B1', value: 'Rounded double' },
          { address: 'C1', value: 'Dependent on rounded formula' },
          { address: 'A2', value: 1.234, format: '0.0' },
          { address: 'B2', formula: 'A2*2', value: 0, format: '0.0' },
          { address: 'C2', formula: 'B2*10', value: 0 },
          { address: 'A3', value: 0.12345 },
          { address: 'B3', formula: 'A3', value: 0, format: '0.0%' },
          { address: 'C3', formula: 'B3*100', value: 0 },
        ],
      },
    ],
  }
}

function normalizedCellValue(value: CellValue): NormalizedFormulaValue {
  switch (value.tag) {
    case ValueTag.Number:
      return { kind: 'number', value: value.value }
    case ValueTag.Boolean:
      return { kind: 'boolean', value: value.value }
    case ValueTag.String:
      return { kind: 'string', value: value.value }
    case ValueTag.Error:
      return { kind: 'error', value: String(value.code) }
    case ValueTag.Empty:
      return { kind: 'blank' }
  }
}
