import { readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'

import { SpreadsheetEngine } from '@bilig/core'
import { exportXlsx, importXlsx } from '@bilig/excel-import'
import { isMacosExcelInstalled, runMacosExcelInspectionOracle, runMacosExcelStructuralOperationOracle } from '@bilig/excel-fixtures'
import type { WorkbookSnapshot } from '@bilig/protocol'
import { describe, expect, it } from 'vitest'

import { createExcelAccessibleTempDir, removeMacosExcelTestDir } from './macos-excel-oracle-test-utils.js'

const ignoredErrorsXml =
  '<ignoredErrors><ignoredError sqref="B2:B4 D5" numberStoredAsText="1"/><ignoredError sqref="C3" formula="1"/></ignoredErrors>'

describe('macOS Desktop Excel ignored errors oracle', () => {
  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'matches Desktop Excel ignoredErrors sqref movement after structural row inserts',
    async () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = createExcelAccessibleTempDir('bilig-headless-excel-ignored-errors-oracle-')
      try {
        const excelWorkbookPath = join(tempDir, 'excel-ignored-errors-oracle.xlsx')
        writeFileSync(excelWorkbookPath, exportXlsx(ignoredErrorsSnapshot()))

        runMacosExcelStructuralOperationOracle({
          workbookPath: excelWorkbookPath,
          worksheetName: 'Review',
          operations: [{ kind: 'insertRows', range: '1:1' }],
          inspectCells: ['B2', 'B3', 'B5', 'C4', 'D6'],
          saveWorkbook: true,
          timeoutMs: 90_000,
        })

        const excelTruth = importXlsx(new Uint8Array(readFileSync(excelWorkbookPath)), 'excel-ignored-errors-oracle.xlsx')
        expect(ignoredErrorSqrefs(excelTruth.snapshot.sheets[0]?.metadata?.ignoredErrors?.xml)).toEqual(['B3:B5 D6', 'C4'])

        const headless = new SpreadsheetEngine({ workbookName: 'headless-ignored-errors-oracle' })
        await headless.ready()
        headless.importSnapshot(ignoredErrorsSnapshot())
        headless.insertRows('Review', 0, 1)
        expect(ignoredErrorSqrefs(headless.exportSnapshot().sheets[0]?.metadata?.ignoredErrors?.xml)).toEqual(['B3:B5 D6', 'C4'])

        const headlessWorkbookPath = join(tempDir, 'headless-ignored-errors-oracle.xlsx')
        writeFileSync(headlessWorkbookPath, exportXlsx(headless.exportSnapshot()))
        runMacosExcelInspectionOracle({
          workbookPath: headlessWorkbookPath,
          worksheetName: 'Review',
          formulaCells: [],
          inspectCells: ['B2', 'B3', 'B5', 'C4', 'D6'],
          saveWorkbook: true,
          timeoutMs: 90_000,
        })

        const headlessExcelTruth = importXlsx(new Uint8Array(readFileSync(headlessWorkbookPath)), 'headless-ignored-errors-oracle.xlsx')
        expect(ignoredErrorSqrefs(headlessExcelTruth.snapshot.sheets[0]?.metadata?.ignoredErrors?.xml)).toEqual(['B3:B5 D6', 'C4'])
      } finally {
        removeMacosExcelTestDir(tempDir)
      }
    },
    120_000,
  )
})

function ignoredErrorSqrefs(xml: string | undefined): string[] {
  if (!xml) {
    throw new Error('Expected ignoredErrors XML')
  }
  return [...xml.matchAll(/\bignoredError\b[^>]*\bsqref="([^"]*)"/gu)].map((match) => match[1] ?? '')
}

function ignoredErrorsSnapshot(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: { name: 'Ignored errors oracle' },
    sheets: [
      {
        id: 1,
        name: 'Review',
        order: 0,
        metadata: {
          ignoredErrors: { xml: ignoredErrorsXml },
        },
        cells: [
          { address: 'A1', value: 'Header' },
          { address: 'B2', value: '001' },
          { address: 'B3', value: '002' },
          { address: 'B4', value: '003' },
          { address: 'C3', formula: 'A1', value: 'Header' },
          { address: 'D5', value: '004' },
        ],
      },
    ],
  }
}
