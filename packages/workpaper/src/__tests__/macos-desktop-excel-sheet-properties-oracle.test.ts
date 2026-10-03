import { readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'

import { exportXlsx, importXlsx } from '@bilig/excel-import'
import { isMacosExcelInstalled, runMacosExcelInspectionOracle, runMacosExcelStructuralOperationOracle } from '@bilig/excel-fixtures'
import type { WorkbookSnapshot } from '@bilig/protocol'
import { describe, expect, it } from 'vitest'

import { WorkPaper } from '../index.js'
import { createExcelAccessibleTempDir, removeMacosExcelTestDir } from './macos-excel-oracle-test-utils.js'

const sheetPr = {
  xml: '<sheetPr codeName="Sheet8"><outlinePr summaryBelow="0" summaryRight="0"/><pageSetUpPr fitToPage="1"/></sheetPr>',
}

describe('macOS Desktop Excel worksheet sheetPr oracle', () => {
  it('preserves imported worksheet sheetPr properties through WorkPaper export', () => {
    const workpaper = WorkPaper.buildFromSnapshot(importXlsx(buildSheetPropertiesWorkbookBytes(), 'sheet-properties-source.xlsx').snapshot)
    try {
      const sheet = workpaper.getSheetId('Report')
      if (sheet === undefined) {
        throw new Error('Expected Report sheet to be available')
      }
      workpaper.addRows(sheet, 0, 1)

      const reimported = importXlsx(exportXlsx(workpaper.exportSnapshot()), 'sheet-properties-headless-roundtrip.xlsx')

      expect(reimported.snapshot.sheets[0]?.metadata?.sheetPr).toEqual(sheetPr)
      expect(reimported.snapshot.sheets[0]?.metadata?.tabColor).toEqual({ rgb: 'FFFF0000' })
    } finally {
      workpaper.dispose()
    }
  })

  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'matches Desktop Excel worksheet sheetPr properties after structural row inserts',
    () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = createExcelAccessibleTempDir('bilig-headless-excel-sheet-properties-oracle-')
      try {
        const excelWorkbookPath = join(tempDir, 'excel-sheet-properties-structural-source.xlsx')
        writeFileSync(excelWorkbookPath, buildSheetPropertiesWorkbookBytes())

        const excelResult = runMacosExcelStructuralOperationOracle({
          workbookPath: excelWorkbookPath,
          worksheetName: 'Report',
          operations: [{ kind: 'insertRows', range: '1:1' }],
          inspectCells: ['A1', 'A2', 'A3'],
          saveWorkbook: true,
          timeoutMs: 90_000,
        })

        const excelTruth = importXlsx(new Uint8Array(readFileSync(excelWorkbookPath)), 'excel-sheet-properties-structural-truth.xlsx')
        expect(excelTruth.snapshot.sheets[0]?.metadata?.sheetPr).toEqual(sheetPr)

        const workpaper = WorkPaper.buildFromSnapshot(
          importXlsx(buildSheetPropertiesWorkbookBytes(), 'headless-sheet-properties-source.xlsx').snapshot,
        )
        try {
          const sheet = workpaper.getSheetId('Report')
          if (sheet === undefined) {
            throw new Error('Expected Report sheet to be available')
          }
          workpaper.addRows(sheet, 0, 1)

          const headlessPath = join(tempDir, 'headless-sheet-properties-structural.xlsx')
          writeFileSync(headlessPath, exportXlsx(workpaper.exportSnapshot()))
          const headlessExcel = runMacosExcelInspectionOracle({
            workbookPath: headlessPath,
            worksheetName: 'Report',
            formulaCells: [],
            inspectCells: ['A1', 'A2', 'A3'],
            saveWorkbook: true,
            timeoutMs: 90_000,
          })
          expect(headlessExcel.cells).toEqual(excelResult.cells)

          const headlessTruth = importXlsx(new Uint8Array(readFileSync(headlessPath)), 'headless-sheet-properties-structural-truth.xlsx')
          expect(headlessTruth.snapshot.sheets[0]?.metadata?.sheetPr).toEqual(excelTruth.snapshot.sheets[0]?.metadata?.sheetPr)
          expect(headlessTruth.snapshot.sheets[0]?.metadata?.tabColor).toEqual(excelTruth.snapshot.sheets[0]?.metadata?.tabColor)
        } finally {
          workpaper.dispose()
        }
      } finally {
        removeMacosExcelTestDir(tempDir)
      }
    },
    120_000,
  )
})

function buildSheetPropertiesWorkbookBytes(): Uint8Array {
  return exportXlsx(sheetPropertiesSnapshot())
}

function sheetPropertiesSnapshot(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: { name: 'Worksheet properties oracle' },
    sheets: [
      {
        id: 1,
        name: 'Report',
        order: 0,
        metadata: {
          sheetPr,
          tabColor: { rgb: 'FFFF0000' },
        },
        cells: [
          { address: 'A1', value: 'Metric' },
          { address: 'A2', value: 'Revenue' },
        ],
      },
    ],
  }
}
