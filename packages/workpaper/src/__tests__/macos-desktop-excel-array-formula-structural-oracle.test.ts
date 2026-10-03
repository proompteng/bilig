import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { exportXlsx, importXlsx } from '@bilig/excel-import'
import { isMacosExcelInstalled, runMacosExcelStructuralOperationOracle } from '@bilig/excel-fixtures'
import type { WorkbookSnapshot } from '@bilig/protocol'
import { describe, expect, it } from 'vitest'

import { WorkPaper } from '../index.js'

const workbookConfig = { maxRows: 16, maxColumns: 8, useColumnIndex: true }
const expectedInsertedArrayFormulaMetadata = {
  formulas: [{ address: 'D3', formulaXml: '<f t="array" ref="D3:D5">TRANSPOSE(A3:C3)</f>' }],
}
const expectedInsertedArrayFormulaCells = [
  { address: 'D3', formula: '=TRANSPOSE(A3:C3)', rawValue: 'number\t10.0', value: { kind: 'number', value: 10 } },
  { address: 'D4', formula: '=TRANSPOSE(A3:C3)', rawValue: 'number\t20.0', value: { kind: 'number', value: 20 } },
  { address: 'D5', formula: '=TRANSPOSE(A3:C3)', rawValue: 'number\t30.0', value: { kind: 'number', value: 30 } },
] as const

describe('macOS Desktop Excel array formula structural oracle', () => {
  it('preserves native array formula metadata through a headless runtime export', () => {
    const imported = importXlsx(exportXlsx(arrayFormulaSnapshot()), 'headless-array-formula-source.xlsx')
    const workbook = WorkPaper.buildFromSnapshot(imported.snapshot, workbookConfig)
    try {
      const roundTrip = importXlsx(exportXlsx(workbook.exportSnapshot()), 'headless-array-formula-runtime-roundtrip.xlsx')
      expect(roundTrip.snapshot.sheets[0]?.metadata?.arrayFormulas).toEqual(arrayFormulaMetadata())
    } finally {
      workbook.dispose()
    }
  })

  it('retargets native array formula metadata through headless structural row inserts', () => {
    const imported = importXlsx(exportXlsx(arrayFormulaSnapshot()), 'headless-array-formula-source.xlsx')
    const workbook = WorkPaper.buildFromSnapshot(imported.snapshot, workbookConfig)
    try {
      const sheetId = workbook.getSheetId('Forecast')
      if (sheetId === undefined) {
        throw new Error('Forecast sheet is missing')
      }
      workbook.addRows(sheetId, 0, 1)

      const roundTrip = importXlsx(exportXlsx(workbook.exportSnapshot()), 'headless-array-formula-insert-roundtrip.xlsx')
      expect(roundTrip.snapshot.sheets[0]?.metadata?.arrayFormulas).toEqual(expectedInsertedArrayFormulaMetadata)
    } finally {
      workbook.dispose()
    }
  })

  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'matches Desktop Excel native array formula metadata after structural row inserts',
    () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = mkdtempSync(join(tmpdir(), 'bilig-headless-excel-array-formula-structural-oracle-'))
      try {
        const excelWorkbookPath = join(tempDir, 'excel-array-formula-structural-oracle.xlsx')
        writeFileSync(excelWorkbookPath, exportXlsx(arrayFormulaSnapshot()))

        const excelResult = runMacosExcelStructuralOperationOracle({
          workbookPath: excelWorkbookPath,
          worksheetName: 'Forecast',
          operations: [{ kind: 'insertRows', range: '1:1' }],
          inspectCells: ['D3', 'D4', 'D5'],
          saveWorkbook: true,
        })
        expect(excelResult.cells).toEqual(expectedInsertedArrayFormulaCells)

        const excelTruth = importXlsx(new Uint8Array(readFileSync(excelWorkbookPath)), 'excel-array-formula-structural-oracle.xlsx')
        expect(excelTruth.snapshot.sheets[0]?.metadata?.arrayFormulas).toEqual(expectedInsertedArrayFormulaMetadata)

        const workbook = WorkPaper.buildFromSnapshot(arrayFormulaSnapshot(), workbookConfig)
        try {
          const sheetId = workbook.getSheetId('Forecast')
          if (sheetId === undefined) {
            throw new Error('Forecast sheet is missing')
          }
          workbook.addRows(sheetId, 0, 1)

          const headless = importXlsx(exportXlsx(workbook.exportSnapshot()), 'headless-array-formula-structural-oracle.xlsx')
          expect(headless.snapshot.sheets[0]?.metadata?.arrayFormulas).toEqual(excelTruth.snapshot.sheets[0]?.metadata?.arrayFormulas)
        } finally {
          workbook.dispose()
        }
      } finally {
        rmSync(tempDir, { recursive: true, force: true })
      }
    },
    60_000,
  )
})

function arrayFormulaSnapshot(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: { name: 'Array formula metadata' },
    sheets: [
      {
        id: 1,
        name: 'Forecast',
        order: 0,
        cells: [
          { address: 'A2', value: 10 },
          { address: 'B2', value: 20 },
          { address: 'C2', value: 30 },
          { address: 'D2', formula: '=TRANSPOSE(A2:C2)', value: 10 },
          { address: 'D3', value: 20 },
          { address: 'D4', value: 30 },
        ],
        metadata: { arrayFormulas: arrayFormulaMetadata() },
      },
    ],
  }
}

function arrayFormulaMetadata() {
  return {
    formulas: [{ address: 'D2', formulaXml: '<f t="array" ref="D2:D4">TRANSPOSE(A2:C2)</f>' }],
  }
}
