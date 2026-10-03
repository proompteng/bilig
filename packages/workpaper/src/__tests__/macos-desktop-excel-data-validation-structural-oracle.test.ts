import { readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'

import { exportXlsx, importXlsx } from '@bilig/excel-import'
import { isMacosExcelInstalled, runMacosExcelInspectionOracle, runMacosExcelStructuralOperationOracle } from '@bilig/excel-fixtures'
import type { WorkbookDataValidationSnapshot, WorkbookSnapshot } from '@bilig/protocol'
import { describe, expect, it } from 'vitest'

import { WorkPaper } from '../index.js'
import { createExcelAccessibleTempDir, removeMacosExcelTestDir } from './macos-excel-oracle-test-utils.js'

const workbookConfig = { maxRows: 24, maxColumns: 8, useColumnIndex: true }

describe('macOS Desktop Excel data validation structural oracle', () => {
  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'matches Desktop Excel formula list validation source refs after row inserts',
    async () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = createExcelAccessibleTempDir('bilig-headless-excel-data-validation-structural-oracle-')
      try {
        const source = dataValidationSourceWorkbook()
        const excelPath = join(tempDir, 'excel-data-validation-source-insert.xlsx')
        writeFileSync(excelPath, exportXlsx(source))

        runMacosExcelStructuralOperationOracle({
          workbookPath: excelPath,
          worksheetName: 'Entry',
          operations: [{ kind: 'insertRows', range: '1:1' }],
          inspectCells: ['A2', 'B3'],
          saveWorkbook: true,
          timeoutMs: 90_000,
        })

        const excelTruth = importXlsx(new Uint8Array(readFileSync(excelPath)), 'excel-data-validation-source-insert.xlsx')
        expect(normalizedEntryValidations(excelTruth.snapshot)).toEqual([
          {
            range: { sheetName: 'Entry', startAddress: 'B3', endAddress: 'B5' },
            rule: {
              kind: 'list',
              source: { kind: 'formula', formula: '=OFFSET(A2,0,0,3,1)' },
            },
            allowBlank: true,
          },
        ])

        const workpaper = WorkPaper.buildFromSnapshot(source, workbookConfig)
        try {
          const entrySheet = workpaper.getSheetId('Entry')
          if (entrySheet === undefined) {
            throw new Error('Expected Entry sheet in data-validation oracle workbook')
          }
          workpaper.addRows(entrySheet, 0, 1)

          const headlessPath = join(tempDir, 'headless-data-validation-source-insert.xlsx')
          writeFileSync(headlessPath, exportXlsx(workpaper.exportSnapshot()))
          runMacosExcelInspectionOracle({
            workbookPath: headlessPath,
            worksheetName: 'Entry',
            formulaCells: [],
            inspectCells: ['A2', 'B3'],
            saveWorkbook: true,
            timeoutMs: 90_000,
          })

          const headlessTruth = importXlsx(new Uint8Array(readFileSync(headlessPath)), 'headless-data-validation-source-insert.xlsx')
          expect(normalizedEntryValidations(headlessTruth.snapshot)).toEqual(normalizedEntryValidations(excelTruth.snapshot))
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

function normalizedEntryValidations(snapshot: WorkbookSnapshot): readonly WorkbookDataValidationSnapshot[] {
  const normalized: WorkbookDataValidationSnapshot[] = []
  for (const validation of snapshot.sheets.find((sheet) => sheet.name === 'Entry')?.metadata?.validations ?? []) {
    if (validation.rule.kind !== 'list' || validation.rule.source?.kind !== 'formula') {
      normalized.push(validation)
      continue
    }
    normalized.push({
      ...validation,
      rule: {
        ...validation.rule,
        source: {
          kind: 'formula',
          formula: validation.rule.source.formula.replaceAll('$', ''),
        },
      },
    })
  }
  return normalized
}

function dataValidationSourceWorkbook(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: { name: 'data-validation-source-structural-oracle' },
    sheets: [
      {
        id: 1,
        name: 'Entry',
        order: 0,
        cells: [
          { address: 'A1', value: 'Draft' },
          { address: 'A2', value: 'Review' },
          { address: 'A3', value: 'Final' },
          { address: 'B2', value: 'Draft' },
        ],
        metadata: {
          validations: [
            {
              range: { sheetName: 'Entry', startAddress: 'B2', endAddress: 'B4' },
              rule: {
                kind: 'list',
                source: { kind: 'formula', formula: '=OFFSET($A$1,0,0,3,1)' },
              },
              allowBlank: true,
            },
          ],
        },
      },
    ],
  }
}
