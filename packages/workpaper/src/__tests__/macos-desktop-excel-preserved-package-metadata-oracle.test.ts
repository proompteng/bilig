import { Buffer } from 'node:buffer'
import { readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'

import { exportXlsx, importXlsx } from '@bilig/excel-import'
import { isMacosExcelInstalled, runMacosExcelInspectionOracle, runMacosExcelStructuralOperationOracle } from '@bilig/excel-fixtures'
import type { WorkbookMetadataSnapshot, WorkbookSnapshot } from '@bilig/protocol'
import { writeSimpleXlsxWorkbook } from '@bilig/xlsx'
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { describe, expect, it } from 'vitest'

import { WorkPaper } from '../index.js'
import { createExcelAccessibleTempDir, removeMacosExcelTestDir } from './macos-excel-oracle-test-utils.js'

describe('macOS Desktop Excel preserved package metadata oracle', () => {
  it('renames raw worksheet chart package formulas on the WorkPaper fast path', () => {
    const importedSource = importXlsx(
      buildWorksheetChartPackageStructuralSourceXlsx(),
      'chart-package-rename-fast-path-source.xlsx',
    ).snapshot
    const sourceFormulaRefs = normalizedChartFormulaRefs(importedSource)
    expect(sourceFormulaRefs).toEqual(['Data!$B$1', 'Data!$A$2:$A$3', 'Data!$B$2:$B$3'])

    const workpaper = WorkPaper.buildFromSnapshot(importedSource)
    try {
      const dataSheet = workpaper.getSheetId('Data')
      if (dataSheet === undefined) {
        throw new Error('Expected Data sheet')
      }
      workpaper.renameSheet(dataSheet, 'Revenue Data')

      expect(normalizedChartFormulaRefs(workpaper.exportSnapshot())).toEqual([
        "'Revenue Data'!$B$1",
        "'Revenue Data'!$A$2:$A$3",
        "'Revenue Data'!$B$2:$B$3",
      ])
    } finally {
      workpaper.dispose()
    }
  })

  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'preserves Desktop Excel workbook and sheet view state after a headless edit',
    () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = createExcelAccessibleTempDir('bilig-headless-excel-package-metadata-oracle-')
      try {
        const sourcePath = join(tempDir, 'view-state-source.xlsx')
        writeFileSync(sourcePath, exportXlsx(viewStateSnapshot()))

        const excelInitial = runMacosExcelInspectionOracle({
          workbookPath: sourcePath,
          worksheetName: 'Report',
          formulaCells: [],
          inspectCells: ['A1', 'B1'],
          saveWorkbook: true,
          timeoutMs: 90_000,
        })
        expect(excelInitial.cells.map((cell) => cell.value)).toEqual([
          { kind: 'string', value: 'view-state' },
          { kind: 'string', value: '' },
        ])

        const excelTruth = importXlsx(new Uint8Array(readFileSync(sourcePath)), 'excel-saved-view-state-source.xlsx')
        expect(viewStateSummary(excelTruth.snapshot)).toEqual({
          workbookView: true,
          sheetView: true,
        })

        const workpaper = WorkPaper.buildFromSnapshot(excelTruth.snapshot)
        try {
          const sheet = workpaper.getSheetId('Report')
          if (sheet === undefined) {
            throw new Error('Expected Report sheet')
          }
          workpaper.setCellContents({ sheet, row: 0, col: 1 }, 'headless edit')

          const headlessSnapshot = workpaper.exportSnapshot()
          expect(viewStateXml(headlessSnapshot)).toEqual(viewStateXml(excelTruth.snapshot))

          const headlessPath = join(tempDir, 'view-state-headless-edit.xlsx')
          writeFileSync(headlessPath, exportXlsx(headlessSnapshot))
          const excelHeadless = runMacosExcelInspectionOracle({
            workbookPath: headlessPath,
            worksheetName: 'Report',
            formulaCells: [],
            inspectCells: ['A1', 'B1'],
            saveWorkbook: true,
            timeoutMs: 90_000,
          })
          expect(excelHeadless.cells.map((cell) => cell.value)).toEqual([
            { kind: 'string', value: 'view-state' },
            { kind: 'string', value: 'headless edit' },
          ])

          const headlessTruth = importXlsx(new Uint8Array(readFileSync(headlessPath)), 'excel-saved-view-state-headless.xlsx')
          expect(viewStateSummary(headlessTruth.snapshot)).toEqual(viewStateSummary(excelTruth.snapshot))
        } finally {
          workpaper.dispose()
        }
      } finally {
        removeMacosExcelTestDir(tempDir)
      }
    },
    120_000,
  )

  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'matches Desktop Excel workbook view tab indexes after deleting a prior sheet',
    () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = createExcelAccessibleTempDir('bilig-headless-excel-workbook-view-delete-oracle-')
      try {
        const sourceBytes = exportXlsx(workbookViewSheetDeletionSnapshot())
        const importedSource = importXlsx(sourceBytes, 'workbook-view-delete-source.xlsx').snapshot
        expect(workbookViewTabIndexes(importedSource)).toEqual({ activeTab: '2', firstSheet: '1' })

        const excelWorkbookPath = join(tempDir, 'excel-workbook-view-delete-source.xlsx')
        writeFileSync(excelWorkbookPath, sourceBytes)
        const excelResult = runMacosExcelStructuralOperationOracle({
          workbookPath: excelWorkbookPath,
          worksheetName: 'Report',
          operations: [{ kind: 'deleteSheet', name: 'Data' }],
          inspectCells: ['A1'],
          saveWorkbook: true,
          timeoutMs: 120_000,
        })
        expect(excelResult.cells[0]?.value).toEqual({ kind: 'string', value: 'report' })

        const excelTruth = importXlsx(new Uint8Array(readFileSync(excelWorkbookPath)), 'excel-workbook-view-delete-truth.xlsx')
        expect(excelTruth.snapshot.sheets.map((sheet) => sheet.name)).toEqual(['Inputs', 'Report'])
        const excelTabIndexes = workbookViewTabIndexes(excelTruth.snapshot)
        expect(excelTabIndexes).toEqual({ activeTab: '1', firstSheet: undefined })

        const workpaper = WorkPaper.buildFromSnapshot(importedSource)
        try {
          const dataSheet = workpaper.getSheetId('Data')
          if (dataSheet === undefined) {
            throw new Error('Expected Data sheet')
          }
          workpaper.removeSheet(dataSheet)

          const headlessSnapshot = workpaper.exportSnapshot()
          expect(headlessSnapshot.sheets.map((sheet) => sheet.name)).toEqual(['Inputs', 'Report'])
          expect(workbookViewTabIndexes(headlessSnapshot)).toEqual(excelTabIndexes)

          const headlessPath = join(tempDir, 'headless-workbook-view-delete.xlsx')
          writeFileSync(headlessPath, exportXlsx(headlessSnapshot))
          const headlessExcel = runMacosExcelInspectionOracle({
            workbookPath: headlessPath,
            worksheetName: 'Report',
            formulaCells: [],
            inspectCells: ['A1'],
            saveWorkbook: true,
            timeoutMs: 120_000,
          })
          expect(headlessExcel.cells).toEqual(excelResult.cells)

          const headlessTruth = importXlsx(new Uint8Array(readFileSync(headlessPath)), 'headless-workbook-view-delete-truth.xlsx')
          expect(workbookViewTabIndexes(headlessTruth.snapshot)).toEqual(excelTabIndexes)
        } finally {
          workpaper.dispose()
        }
      } finally {
        removeMacosExcelTestDir(tempDir)
      }
    },
    180_000,
  )

  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'matches Desktop Excel workbook view tab indexes after moving a sheet tab',
    () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = createExcelAccessibleTempDir('bilig-headless-excel-workbook-view-move-oracle-')
      try {
        const sourceBytes = exportXlsx(workbookViewSheetReorderSnapshot())
        const importedSource = importXlsx(sourceBytes, 'workbook-view-move-source.xlsx').snapshot
        expect(importedSource.sheets.map((sheet) => sheet.name)).toEqual(['Data', 'Inputs', 'Report'])
        expect(workbookViewTabIndexes(importedSource)).toEqual({ activeTab: '2', firstSheet: '1' })

        const excelWorkbookPath = join(tempDir, 'excel-workbook-view-move-source.xlsx')
        writeFileSync(excelWorkbookPath, sourceBytes)
        const excelResult = runMacosExcelStructuralOperationOracle({
          workbookPath: excelWorkbookPath,
          worksheetName: 'Report',
          operations: [{ kind: 'moveSheet', name: 'Report', before: 'Inputs' }],
          inspectCells: ['A1'],
          saveWorkbook: true,
          timeoutMs: 120_000,
        })
        expect(excelResult.cells[0]?.value).toEqual({ kind: 'string', value: 'report' })

        const excelTruth = importXlsx(new Uint8Array(readFileSync(excelWorkbookPath)), 'excel-workbook-view-move-truth.xlsx')
        expect(excelTruth.snapshot.sheets.map((sheet) => sheet.name)).toEqual(['Data', 'Report', 'Inputs'])
        const excelTabIndexes = workbookViewTabIndexes(excelTruth.snapshot)
        expect(excelTabIndexes).toEqual({ activeTab: '1', firstSheet: '1' })

        const workpaper = WorkPaper.buildFromSnapshot(importedSource)
        try {
          const reportSheet = workpaper.getSheetId('Report')
          if (reportSheet === undefined) {
            throw new Error('Expected Report sheet')
          }
          workpaper.moveSheet(reportSheet, 1)

          const headlessSnapshot = workpaper.exportSnapshot()
          expect(headlessSnapshot.sheets.map((sheet) => sheet.name)).toEqual(['Data', 'Report', 'Inputs'])
          expect(workbookViewTabIndexes(headlessSnapshot)).toEqual(excelTabIndexes)

          const headlessPath = join(tempDir, 'headless-workbook-view-move.xlsx')
          writeFileSync(headlessPath, exportXlsx(headlessSnapshot))
          const headlessExcel = runMacosExcelInspectionOracle({
            workbookPath: headlessPath,
            worksheetName: 'Report',
            formulaCells: [],
            inspectCells: ['A1'],
            saveWorkbook: true,
            timeoutMs: 120_000,
          })
          expect(headlessExcel.cells).toEqual(excelResult.cells)

          const headlessTruth = importXlsx(new Uint8Array(readFileSync(headlessPath)), 'headless-workbook-view-move-truth.xlsx')
          expect(workbookViewTabIndexes(headlessTruth.snapshot)).toEqual(excelTabIndexes)
        } finally {
          workpaper.dispose()
        }
      } finally {
        removeMacosExcelTestDir(tempDir)
      }
    },
    180_000,
  )

  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'matches Desktop Excel workbook view first visible tab after moving a prior sheet past it',
    () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = createExcelAccessibleTempDir('bilig-headless-excel-workbook-view-first-sheet-move-oracle-')
      try {
        const sourceBytes = exportXlsx(workbookViewSheetReorderSnapshot())
        const importedSource = importXlsx(sourceBytes, 'workbook-view-first-sheet-move-source.xlsx').snapshot
        expect(importedSource.sheets.map((sheet) => sheet.name)).toEqual(['Data', 'Inputs', 'Report'])
        expect(workbookViewTabIndexes(importedSource)).toEqual({ activeTab: '2', firstSheet: '1' })

        const excelWorkbookPath = join(tempDir, 'excel-workbook-view-first-sheet-move-source.xlsx')
        writeFileSync(excelWorkbookPath, sourceBytes)
        const excelResult = runMacosExcelStructuralOperationOracle({
          workbookPath: excelWorkbookPath,
          worksheetName: 'Report',
          operations: [{ kind: 'moveSheet', name: 'Data', after: 'Report' }],
          inspectCells: ['A1'],
          saveWorkbook: true,
          timeoutMs: 120_000,
        })
        expect(excelResult.cells[0]?.value).toEqual({ kind: 'string', value: 'report' })

        const excelTruth = importXlsx(new Uint8Array(readFileSync(excelWorkbookPath)), 'excel-workbook-view-first-sheet-move-truth.xlsx')
        expect(excelTruth.snapshot.sheets.map((sheet) => sheet.name)).toEqual(['Inputs', 'Report', 'Data'])
        const excelTabIndexes = workbookViewTabIndexes(excelTruth.snapshot)

        const workpaper = WorkPaper.buildFromSnapshot(importedSource)
        try {
          const dataSheet = workpaper.getSheetId('Data')
          if (dataSheet === undefined) {
            throw new Error('Expected Data sheet')
          }
          workpaper.moveSheet(dataSheet, 2)

          const headlessSnapshot = workpaper.exportSnapshot()
          expect(headlessSnapshot.sheets.map((sheet) => sheet.name)).toEqual(['Inputs', 'Report', 'Data'])
          expect(workbookViewTabIndexes(headlessSnapshot)).toEqual(excelTabIndexes)

          const headlessPath = join(tempDir, 'headless-workbook-view-first-sheet-move.xlsx')
          writeFileSync(headlessPath, exportXlsx(headlessSnapshot))
          const headlessExcel = runMacosExcelInspectionOracle({
            workbookPath: headlessPath,
            worksheetName: 'Report',
            formulaCells: [],
            inspectCells: ['A1'],
            saveWorkbook: true,
            timeoutMs: 120_000,
          })
          expect(headlessExcel.cells).toEqual(excelResult.cells)

          const headlessTruth = importXlsx(new Uint8Array(readFileSync(headlessPath)), 'headless-workbook-view-first-sheet-move-truth.xlsx')
          expect(workbookViewTabIndexes(headlessTruth.snapshot)).toEqual(excelTabIndexes)
        } finally {
          workpaper.dispose()
        }
      } finally {
        removeMacosExcelTestDir(tempDir)
      }
    },
    180_000,
  )

  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'matches Desktop Excel calc-chain sheet ids after deleting a prior sheet',
    () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = createExcelAccessibleTempDir('bilig-headless-excel-calc-chain-delete-oracle-')
      try {
        const sourceBytes = buildCalcChainSheetDeletionXlsx()
        const importedSource = importXlsx(sourceBytes, 'calc-chain-delete-source.xlsx').snapshot
        expect(calcChainSummary(importedSource)).toEqual(['1:Data:A1', '2:Inputs:A1', '3:Report:A1'])

        const excelWorkbookPath = join(tempDir, 'excel-calc-chain-delete-source.xlsx')
        writeFileSync(excelWorkbookPath, sourceBytes)
        const excelResult = runMacosExcelStructuralOperationOracle({
          workbookPath: excelWorkbookPath,
          worksheetName: 'Report',
          operations: [{ kind: 'deleteSheet', name: 'Data' }],
          inspectCells: ['A1'],
          saveWorkbook: true,
          timeoutMs: 120_000,
        })
        expect(excelResult.cells[0]?.value).toEqual({ kind: 'number', value: 12 })

        const excelTruth = importXlsx(new Uint8Array(readFileSync(excelWorkbookPath)), 'excel-calc-chain-delete-truth.xlsx')
        expect(excelTruth.snapshot.sheets.map((sheet) => sheet.name)).toEqual(['Inputs', 'Report'])
        const excelCalcChain = calcChainSummary(excelTruth.snapshot)
        expect(excelCalcChain).toEqual(['2:Inputs:A1', '3:Report:A1'])

        const workpaper = WorkPaper.buildFromSnapshot(importedSource)
        try {
          const dataSheet = workpaper.getSheetId('Data')
          if (dataSheet === undefined) {
            throw new Error('Expected Data sheet')
          }
          workpaper.removeSheet(dataSheet)

          const headlessSnapshot = workpaper.exportSnapshot()
          expect(headlessSnapshot.sheets.map((sheet) => sheet.name)).toEqual(['Inputs', 'Report'])
          expect(calcChainSummary(headlessSnapshot)).toEqual(excelCalcChain)
        } finally {
          workpaper.dispose()
        }
      } finally {
        removeMacosExcelTestDir(tempDir)
      }
    },
    180_000,
  )

  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'matches Desktop Excel preserved style artifacts and view refs after structural row inserts',
    () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = createExcelAccessibleTempDir('bilig-headless-excel-package-metadata-structure-oracle-')
      try {
        const excelWorkbookPath = join(tempDir, 'excel-preserved-package-structure-source.xlsx')
        const sourceBytes = exportXlsx(structuralPreservedMetadataSnapshot())
        writeFileSync(excelWorkbookPath, sourceBytes)

        const excelResult = runMacosExcelStructuralOperationOracle({
          workbookPath: excelWorkbookPath,
          worksheetName: 'Report',
          operations: [{ kind: 'insertRows', range: '1:1' }],
          inspectCells: ['A1', 'B3', 'C4'],
          saveWorkbook: true,
          timeoutMs: 90_000,
        })

        const excelTruth = importXlsx(new Uint8Array(readFileSync(excelWorkbookPath)), 'excel-preserved-package-structure-truth.xlsx')
        expect(styleArtifactRefs(excelTruth.snapshot)).toEqual({
          styled: ['B3', 'C4'],
          blank: [],
        })
        expect(sheetViewRefs(excelTruth.snapshot)).toEqual({
          topLeftCell: 'B2',
          activeCell: 'C3',
          sqref: 'C3 D4:E5',
        })

        const workpaper = WorkPaper.buildFromSnapshot(importXlsx(sourceBytes, 'headless-preserved-package-structure-source.xlsx').snapshot)
        try {
          const sheet = workpaper.getSheetId('Report')
          if (sheet === undefined) {
            throw new Error('Expected Report sheet')
          }
          workpaper.addRows(sheet, 0, 1)

          const headlessSnapshot = workpaper.exportSnapshot()
          expect(styleArtifactRefs(headlessSnapshot)).toEqual(styleArtifactRefs(excelTruth.snapshot))
          expect(sheetViewRefs(headlessSnapshot)).toEqual(sheetViewRefs(excelTruth.snapshot))

          const headlessPath = join(tempDir, 'headless-preserved-package-structure.xlsx')
          writeFileSync(headlessPath, exportXlsx(headlessSnapshot))
          const headlessExcel = runMacosExcelInspectionOracle({
            workbookPath: headlessPath,
            worksheetName: 'Report',
            formulaCells: [],
            inspectCells: ['A1', 'B3', 'C4'],
            saveWorkbook: true,
            timeoutMs: 90_000,
          })
          expect(headlessExcel.cells).toEqual(excelResult.cells)

          const headlessTruth = importXlsx(new Uint8Array(readFileSync(headlessPath)), 'headless-preserved-package-structure-truth.xlsx')
          expect(styleArtifactRefs(headlessTruth.snapshot)).toEqual(styleArtifactRefs(excelTruth.snapshot))
          expect(sheetViewRefs(headlessTruth.snapshot)).toEqual(sheetViewRefs(excelTruth.snapshot))
        } finally {
          workpaper.dispose()
        }
      } finally {
        removeMacosExcelTestDir(tempDir)
      }
    },
    120_000,
  )

  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'matches Desktop Excel preserved pivot package refs after structural row inserts',
    () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = createExcelAccessibleTempDir('bilig-headless-excel-pivot-package-structure-oracle-')
      try {
        const sourceBytes = exportXlsx(pivotPackageStructuralSourceSnapshot())
        const importedSource = importXlsx(sourceBytes, 'pivot-package-structure-source.xlsx').snapshot
        expect(pivotLocationRefs(importedSource)).toEqual(['B2:C5'])

        const excelWorkbookPath = join(tempDir, 'excel-pivot-package-structure-source.xlsx')
        writeFileSync(excelWorkbookPath, sourceBytes)
        runMacosExcelStructuralOperationOracle({
          workbookPath: excelWorkbookPath,
          worksheetName: 'Pivot',
          operations: [{ kind: 'insertRows', range: '1:1' }],
          inspectCells: ['A1'],
          saveWorkbook: true,
          timeoutMs: 120_000,
        })

        const excelTruth = importXlsx(new Uint8Array(readFileSync(excelWorkbookPath)), 'excel-pivot-package-structure-truth.xlsx')
        expect(pivotLocationRefs(excelTruth.snapshot)).toEqual(['B3:C6'])

        const workpaper = WorkPaper.buildFromSnapshot(importedSource)
        try {
          const pivotSheet = workpaper.getSheetId('Pivot')
          if (pivotSheet === undefined) {
            throw new Error('Expected Pivot sheet')
          }
          workpaper.addRows(pivotSheet, 0, 1)

          const headlessSnapshot = workpaper.exportSnapshot()
          expect(pivotLocationRefs(headlessSnapshot)).toEqual(pivotLocationRefs(excelTruth.snapshot))

          const headlessPath = join(tempDir, 'headless-pivot-package-structure.xlsx')
          writeFileSync(headlessPath, exportXlsx(headlessSnapshot))
          runMacosExcelInspectionOracle({
            workbookPath: headlessPath,
            worksheetName: 'Pivot',
            formulaCells: [],
            inspectCells: ['A1'],
            saveWorkbook: true,
            timeoutMs: 120_000,
          })
          const headlessTruth = importXlsx(new Uint8Array(readFileSync(headlessPath)), 'headless-pivot-package-structure-truth.xlsx')
          expect(pivotLocationRefs(headlessTruth.snapshot)).toEqual(pivotLocationRefs(excelTruth.snapshot))
        } finally {
          workpaper.dispose()
        }
      } finally {
        removeMacosExcelTestDir(tempDir)
      }
    },
    180_000,
  )

  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'matches Desktop Excel preserved pivot cache source refs after structural source row inserts',
    () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = createExcelAccessibleTempDir('bilig-headless-excel-pivot-cache-source-structure-oracle-')
      try {
        const sourceBytes = exportXlsx(pivotPackageStructuralSourceSnapshot())
        const importedSource = importXlsx(sourceBytes, 'pivot-cache-source-structure-source.xlsx').snapshot
        expect(pivotCacheSourceRefs(importedSource)).toEqual(['A1:D4'])

        const excelWorkbookPath = join(tempDir, 'excel-pivot-cache-source-structure-source.xlsx')
        writeFileSync(excelWorkbookPath, sourceBytes)
        runMacosExcelStructuralOperationOracle({
          workbookPath: excelWorkbookPath,
          worksheetName: 'Data',
          operations: [{ kind: 'insertRows', range: '1:1' }],
          inspectCells: ['A1'],
          saveWorkbook: true,
          timeoutMs: 120_000,
        })

        const excelTruth = importXlsx(new Uint8Array(readFileSync(excelWorkbookPath)), 'excel-pivot-cache-source-structure-truth.xlsx')
        const excelSourceRefs = pivotCacheSourceRefs(excelTruth.snapshot)
        expect(excelSourceRefs).not.toEqual(['A1:D4'])

        const workpaper = WorkPaper.buildFromSnapshot(importedSource)
        try {
          const dataSheet = workpaper.getSheetId('Data')
          if (dataSheet === undefined) {
            throw new Error('Expected Data sheet')
          }
          workpaper.addRows(dataSheet, 0, 1)

          const headlessSnapshot = workpaper.exportSnapshot()
          expect(pivotCacheSourceRefs(headlessSnapshot)).toEqual(excelSourceRefs)

          const headlessPath = join(tempDir, 'headless-pivot-cache-source-structure.xlsx')
          writeFileSync(headlessPath, exportXlsx(headlessSnapshot))
          runMacosExcelInspectionOracle({
            workbookPath: headlessPath,
            worksheetName: 'Data',
            formulaCells: [],
            inspectCells: ['A1'],
            saveWorkbook: true,
            timeoutMs: 120_000,
          })
          const headlessTruth = importXlsx(new Uint8Array(readFileSync(headlessPath)), 'headless-pivot-cache-source-structure-truth.xlsx')
          expect(pivotCacheSourceRefs(headlessTruth.snapshot)).toEqual(excelSourceRefs)
        } finally {
          workpaper.dispose()
        }
      } finally {
        removeMacosExcelTestDir(tempDir)
      }
    },
    180_000,
  )

  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'matches Desktop Excel preserved pivot cache source sheet after source sheet rename',
    () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = createExcelAccessibleTempDir('bilig-headless-excel-pivot-cache-source-rename-oracle-')
      try {
        const sourceBytes = exportXlsx(pivotPackageStructuralSourceSnapshot())
        const importedSource = importXlsx(sourceBytes, 'pivot-cache-source-rename-source.xlsx').snapshot
        expect(pivotCacheSourceSheets(importedSource)).toEqual(['Data'])

        const excelWorkbookPath = join(tempDir, 'excel-pivot-cache-source-rename-source.xlsx')
        writeFileSync(excelWorkbookPath, sourceBytes)
        runMacosExcelStructuralOperationOracle({
          workbookPath: excelWorkbookPath,
          worksheetName: 'Data',
          operations: [{ kind: 'renameSheet', newName: 'Revenue Data' }],
          inspectCells: ['A1'],
          saveWorkbook: true,
          timeoutMs: 120_000,
        })

        const excelTruth = importXlsx(new Uint8Array(readFileSync(excelWorkbookPath)), 'excel-pivot-cache-source-rename-truth.xlsx')
        const excelSourceSheets = pivotCacheSourceSheets(excelTruth.snapshot)
        expect(excelSourceSheets).toEqual(['Revenue Data'])

        const workpaper = WorkPaper.buildFromSnapshot(importedSource)
        try {
          const dataSheet = workpaper.getSheetId('Data')
          if (dataSheet === undefined) {
            throw new Error('Expected Data sheet')
          }
          workpaper.renameSheet(dataSheet, 'Revenue Data')

          const headlessSnapshot = workpaper.exportSnapshot()
          expect(pivotCacheSourceSheets(headlessSnapshot)).toEqual(excelSourceSheets)

          const headlessPath = join(tempDir, 'headless-pivot-cache-source-rename.xlsx')
          writeFileSync(headlessPath, exportXlsx(headlessSnapshot))
          runMacosExcelInspectionOracle({
            workbookPath: headlessPath,
            worksheetName: 'Revenue Data',
            formulaCells: [],
            inspectCells: ['A1'],
            saveWorkbook: true,
            timeoutMs: 120_000,
          })
          const headlessTruth = importXlsx(new Uint8Array(readFileSync(headlessPath)), 'headless-pivot-cache-source-rename-truth.xlsx')
          expect(pivotCacheSourceSheets(headlessTruth.snapshot)).toEqual(excelSourceSheets)
        } finally {
          workpaper.dispose()
        }
      } finally {
        removeMacosExcelTestDir(tempDir)
      }
    },
    180_000,
  )

  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'matches Desktop Excel raw worksheet chart package formulas after structural source row inserts',
    () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = createExcelAccessibleTempDir('bilig-headless-excel-chart-package-structure-oracle-')
      try {
        const sourceBytes = buildWorksheetChartPackageStructuralSourceXlsx()
        const importedSource = importXlsx(sourceBytes, 'chart-package-structure-source.xlsx').snapshot
        const sourceFormulaRefs = normalizedChartFormulaRefs(importedSource)
        expect(sourceFormulaRefs).toEqual(['Data!$B$1', 'Data!$A$2:$A$3', 'Data!$B$2:$B$3'])

        const excelWorkbookPath = join(tempDir, 'excel-chart-package-structure-source.xlsx')
        writeFileSync(excelWorkbookPath, sourceBytes)
        runMacosExcelStructuralOperationOracle({
          workbookPath: excelWorkbookPath,
          worksheetName: 'Data',
          operations: [{ kind: 'insertRows', range: '2:2' }],
          inspectCells: ['A1'],
          saveWorkbook: true,
          timeoutMs: 120_000,
        })

        const excelTruth = importXlsx(new Uint8Array(readFileSync(excelWorkbookPath)), 'excel-chart-package-structure-truth.xlsx')
        const excelFormulaRefs = normalizedChartFormulaRefs(excelTruth.snapshot)
        expect(excelFormulaRefs).not.toEqual(sourceFormulaRefs)

        const workpaper = WorkPaper.buildFromSnapshot(importedSource)
        try {
          const dataSheet = workpaper.getSheetId('Data')
          if (dataSheet === undefined) {
            throw new Error('Expected Data sheet')
          }
          workpaper.addRows(dataSheet, 1, 1)

          const headlessSnapshot = workpaper.exportSnapshot()
          expect(normalizedChartFormulaRefs(headlessSnapshot)).toEqual(excelFormulaRefs)

          const headlessPath = join(tempDir, 'headless-chart-package-structure.xlsx')
          writeFileSync(headlessPath, exportXlsx(headlessSnapshot))
          runMacosExcelInspectionOracle({
            workbookPath: headlessPath,
            worksheetName: 'Data',
            formulaCells: [],
            inspectCells: ['A1'],
            saveWorkbook: true,
            timeoutMs: 120_000,
          })
          const headlessTruth = importXlsx(new Uint8Array(readFileSync(headlessPath)), 'headless-chart-package-structure-truth.xlsx')
          expect(normalizedChartFormulaRefs(headlessTruth.snapshot)).toEqual(excelFormulaRefs)
        } finally {
          workpaper.dispose()
        }
      } finally {
        removeMacosExcelTestDir(tempDir)
      }
    },
    180_000,
  )

  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'matches Desktop Excel raw worksheet chart package formulas after source sheet rename',
    () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = createExcelAccessibleTempDir('bilig-headless-excel-chart-package-rename-oracle-')
      try {
        const sourceBytes = buildWorksheetChartPackageStructuralSourceXlsx()
        const importedSource = importXlsx(sourceBytes, 'chart-package-rename-source.xlsx').snapshot
        const sourceFormulaRefs = normalizedChartFormulaRefs(importedSource)
        expect(sourceFormulaRefs).toEqual(['Data!$B$1', 'Data!$A$2:$A$3', 'Data!$B$2:$B$3'])

        const excelWorkbookPath = join(tempDir, 'excel-chart-package-rename-source.xlsx')
        writeFileSync(excelWorkbookPath, sourceBytes)
        runMacosExcelStructuralOperationOracle({
          workbookPath: excelWorkbookPath,
          worksheetName: 'Data',
          operations: [{ kind: 'renameSheet', newName: 'Revenue Data' }],
          inspectCells: ['A1'],
          saveWorkbook: true,
          timeoutMs: 120_000,
        })

        const excelTruth = importXlsx(new Uint8Array(readFileSync(excelWorkbookPath)), 'excel-chart-package-rename-truth.xlsx')
        const excelFormulaRefs = normalizedChartFormulaRefs(excelTruth.snapshot)
        expect(excelFormulaRefs).toEqual(["'Revenue Data'!$B$1", "'Revenue Data'!$A$2:$A$3", "'Revenue Data'!$B$2:$B$3"])

        const workpaper = WorkPaper.buildFromSnapshot(importedSource)
        try {
          const dataSheet = workpaper.getSheetId('Data')
          if (dataSheet === undefined) {
            throw new Error('Expected Data sheet')
          }
          workpaper.renameSheet(dataSheet, 'Revenue Data')

          const headlessSnapshot = workpaper.exportSnapshot()
          expect(normalizedChartFormulaRefs(headlessSnapshot)).toEqual(excelFormulaRefs)

          const headlessPath = join(tempDir, 'headless-chart-package-rename.xlsx')
          writeFileSync(headlessPath, exportXlsx(headlessSnapshot))
          runMacosExcelInspectionOracle({
            workbookPath: headlessPath,
            worksheetName: 'Revenue Data',
            formulaCells: [],
            inspectCells: ['A1'],
            saveWorkbook: true,
            timeoutMs: 120_000,
          })
          const headlessTruth = importXlsx(new Uint8Array(readFileSync(headlessPath)), 'headless-chart-package-rename-truth.xlsx')
          expect(normalizedChartFormulaRefs(headlessTruth.snapshot)).toEqual(excelFormulaRefs)
        } finally {
          workpaper.dispose()
        }
      } finally {
        removeMacosExcelTestDir(tempDir)
      }
    },
    180_000,
  )
})

function viewStateSnapshot(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: {
      name: 'Desktop Excel package metadata oracle',
      metadata: {
        viewState: {
          bookViewsXml: '<bookViews><workbookView activeTab="0" firstSheet="0"/></bookViews>',
        },
      },
    },
    sheets: [
      {
        id: 1,
        name: 'Report',
        order: 0,
        metadata: {
          viewState: {
            sheetViewsXml: '<sheetViews><sheetView workbookViewId="0" tabSelected="1"/></sheetViews>',
          },
        },
        cells: [{ address: 'A1', value: 'view-state' }],
      },
    ],
  }
}

function workbookViewSheetDeletionSnapshot(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: {
      name: 'Desktop Excel workbook view delete oracle',
      metadata: {
        viewState: {
          bookViewsXml: '<bookViews><workbookView activeTab="2" firstSheet="1"/></bookViews>',
        },
      },
    },
    sheets: [
      {
        id: 1,
        name: 'Data',
        order: 0,
        cells: [{ address: 'A1', value: 'data' }],
      },
      {
        id: 2,
        name: 'Inputs',
        order: 1,
        cells: [{ address: 'A1', value: 'inputs' }],
      },
      {
        id: 3,
        name: 'Report',
        order: 2,
        metadata: {
          viewState: {
            sheetViewsXml: '<sheetViews><sheetView workbookViewId="0" tabSelected="1"/></sheetViews>',
          },
        },
        cells: [{ address: 'A1', value: 'report' }],
      },
    ],
  }
}

function workbookViewSheetReorderSnapshot(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: {
      name: 'Workbook view sheet reorder',
      metadata: {
        viewState: {
          bookViewsXml: '<bookViews><workbookView activeTab="2" firstSheet="1"/></bookViews>',
        },
      },
    },
    sheets: [
      {
        id: 1,
        name: 'Data',
        order: 0,
        cells: [{ address: 'A1', value: 'data' }],
      },
      {
        id: 2,
        name: 'Inputs',
        order: 1,
        cells: [{ address: 'A1', value: 'inputs' }],
      },
      {
        id: 3,
        name: 'Report',
        order: 2,
        cells: [{ address: 'A1', value: 'report' }],
      },
    ],
  }
}

function calcChainSheetDeletionSnapshot(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: {
      name: 'Desktop Excel calc chain delete oracle',
    },
    sheets: [
      {
        id: 1,
        name: 'Data',
        order: 0,
        cells: [{ address: 'A1', formula: '1+1', value: 2 }],
      },
      {
        id: 2,
        name: 'Inputs',
        order: 1,
        cells: [{ address: 'A1', formula: '10+1', value: 11 }],
      },
      {
        id: 3,
        name: 'Report',
        order: 2,
        cells: [{ address: 'A1', formula: 'Inputs!A1+1', value: 12 }],
      },
    ],
  }
}

function buildCalcChainSheetDeletionXlsx(): Uint8Array {
  const zip = unzipSync(exportXlsx(calcChainSheetDeletionSnapshot()))
  zip['xl/calcChain.xml'] = strToU8(
    [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
      '<c r="A1" i="1"/>',
      '<c r="A1" i="2"/>',
      '<c r="A1" i="3"/>',
      '</calcChain>',
    ].join(''),
  )
  addCalcChainWorkbookRelationship(zip)
  addCalcChainContentType(zip)
  return zipSync(zip)
}

function addCalcChainWorkbookRelationship(zip: Record<string, Uint8Array>): void {
  const path = 'xl/_rels/workbook.xml.rels'
  const xml = strFromU8(zip[path] ?? new Uint8Array())
  if (xml.includes('/calcChain')) {
    return
  }
  zip[path] = strToU8(
    xml.replace(
      '</Relationships>',
      '<Relationship Id="rIdCalcChain1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/></Relationships>',
    ),
  )
}

function addCalcChainContentType(zip: Record<string, Uint8Array>): void {
  const path = '[Content_Types].xml'
  const xml = strFromU8(zip[path] ?? new Uint8Array())
  if (xml.includes('calcChain+xml')) {
    return
  }
  zip[path] = strToU8(
    xml.replace(
      '</Types>',
      '<Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/></Types>',
    ),
  )
}

function structuralPreservedMetadataSnapshot(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: {
      name: 'Desktop Excel structural package metadata oracle',
      metadata: {
        styleArtifacts: {
          stylesXml: headerStyleReferenceStylesXml,
        },
        viewState: {
          bookViewsXml: '<bookViews><workbookView activeTab="0" firstSheet="0"/></bookViews>',
        },
      },
    },
    sheets: [
      {
        id: 1,
        name: 'Report',
        order: 0,
        metadata: {
          styleArtifacts: {
            cellStyleIndexes: [
              { address: 'B2', styleIndex: 1 },
              { address: 'C3', styleIndex: 1 },
            ],
            blankCellAddresses: ['C3'],
          },
          viewState: {
            sheetViewsXml:
              '<sheetViews><sheetView workbookViewId="0" topLeftCell="B2" tabSelected="1"><selection activeCell="C3" sqref="C3 D4:E5"/></sheetView></sheetViews>',
          },
        },
        cells: [
          { address: 'A1', value: 'preserved-structure' },
          { address: 'B2', value: 'styled' },
        ],
      },
    ],
  }
}

function pivotPackageStructuralSourceSnapshot(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: {
      name: 'Desktop Excel pivot package structural oracle',
      metadata: {
        pivots: [
          {
            name: 'SalesByRegion',
            sheetName: 'Pivot',
            address: 'B2',
            source: { sheetName: 'Data', startAddress: 'A1', endAddress: 'D4' },
            sourceKind: 'worksheet',
            groupBy: ['Region'],
            values: [{ sourceColumn: 'Sales', summarizeBy: 'sum' }],
            cacheFields: ['Region', 'Notes', 'Product', 'Sales'],
            cachedRecords: [
              ['East', 'priority', 'Widget', 20],
              ['West', 'priority', 'Widget', 7],
              ['East', 'priority', 'Gizmo', 5],
            ],
            rows: 4,
            cols: 2,
          },
        ],
      },
    },
    sheets: [
      {
        id: 1,
        name: 'Data',
        order: 0,
        cells: [
          { address: 'A1', value: 'Region' },
          { address: 'B1', value: 'Notes' },
          { address: 'C1', value: 'Product' },
          { address: 'D1', value: 'Sales' },
          { address: 'A2', value: 'East' },
          { address: 'B2', value: 'priority' },
          { address: 'C2', value: 'Widget' },
          { address: 'D2', value: 20 },
          { address: 'A3', value: 'West' },
          { address: 'B3', value: 'priority' },
          { address: 'C3', value: 'Widget' },
          { address: 'D3', value: 7 },
          { address: 'A4', value: 'East' },
          { address: 'B4', value: 'priority' },
          { address: 'C4', value: 'Gizmo' },
          { address: 'D4', value: 5 },
        ],
      },
      {
        id: 2,
        name: 'Pivot',
        order: 1,
        cells: [],
      },
    ],
  }
}

function viewStateXml(snapshot: WorkbookSnapshot): { readonly workbook: string | undefined; readonly sheet: string | undefined } {
  return {
    workbook: snapshot.workbook.metadata?.viewState?.bookViewsXml,
    sheet: snapshot.sheets[0]?.metadata?.viewState?.sheetViewsXml,
  }
}

function viewStateSummary(snapshot: WorkbookSnapshot): { readonly workbookView: boolean; readonly sheetView: boolean } {
  const xml = viewStateXml(snapshot)
  return {
    workbookView: xml.workbook?.includes('<workbookView') ?? false,
    sheetView: xml.sheet?.includes('<sheetView') ?? false,
  }
}

function workbookViewTabIndexes(snapshot: WorkbookSnapshot): {
  readonly activeTab: string | undefined
  readonly firstSheet: string | undefined
} {
  const xml = snapshot.workbook.metadata?.viewState?.bookViewsXml ?? ''
  return {
    activeTab: readXmlAttribute(xml, 'activeTab'),
    firstSheet: readXmlAttribute(xml, 'firstSheet'),
  }
}

function calcChainSummary(snapshot: WorkbookSnapshot): string[] {
  return (snapshot.workbook.metadata?.formulaAudit?.calcChain?.cells ?? []).map(
    (cell) => `${String(cell.sheetIndex)}:${cell.sheetName ?? ''}:${cell.address}`,
  )
}

function styleArtifactRefs(snapshot: WorkbookSnapshot): { readonly styled: string[]; readonly blank: string[] } {
  const artifacts = snapshot.sheets[0]?.metadata?.styleArtifacts
  return {
    styled: (artifacts?.cellStyleIndexes ?? []).map((entry) => entry.address).toSorted(),
    blank: (artifacts?.blankCellAddresses ?? []).toSorted(),
  }
}

function sheetViewRefs(snapshot: WorkbookSnapshot): {
  readonly topLeftCell: string | undefined
  readonly activeCell: string | undefined
  readonly sqref: string | undefined
} {
  const xml = snapshot.sheets[0]?.metadata?.viewState?.sheetViewsXml ?? ''
  return {
    topLeftCell: readXmlAttribute(xml, 'topLeftCell'),
    activeCell: readXmlAttribute(xml, 'activeCell'),
    sqref: readXmlAttribute(xml, 'sqref'),
  }
}

function pivotLocationRefs(snapshot: WorkbookSnapshot): string[] {
  return (snapshot.workbook.metadata?.pivotArtifacts?.parts ?? [])
    .filter((part) => part.path.startsWith('xl/pivotTables/'))
    .map((part) => readXmlAttribute(part.xml, 'ref'))
    .filter((ref): ref is string => ref !== undefined)
    .toSorted()
}

function pivotCacheSourceRefs(snapshot: WorkbookSnapshot): string[] {
  return (snapshot.workbook.metadata?.pivotArtifacts?.parts ?? [])
    .filter((part) => part.path.startsWith('xl/pivotCache/pivotCacheDefinition'))
    .map((part) => readXmlAttribute(part.xml, 'ref'))
    .filter((ref): ref is string => ref !== undefined)
    .toSorted()
}

function pivotCacheSourceSheets(snapshot: WorkbookSnapshot): string[] {
  return (snapshot.workbook.metadata?.pivotArtifacts?.parts ?? [])
    .filter((part) => part.path.startsWith('xl/pivotCache/pivotCacheDefinition'))
    .map((part) => readXmlAttribute(part.xml, 'sheet'))
    .filter((sheet): sheet is string => sheet !== undefined)
    .toSorted()
}

function normalizedChartFormulaRefs(snapshot: WorkbookSnapshot): string[] {
  return chartFormulaRefs(snapshot.workbook.metadata, 'xl/charts/chart1.xml').map(normalizeChartFormulaRef)
}

function chartFormulaRefs(metadata: WorkbookMetadataSnapshot | undefined, path: string): string[] {
  const part = [...(metadata?.chartArtifacts?.parts ?? []), ...(metadata?.drawingArtifacts?.parts ?? [])].find(
    (candidate) => candidate.path === path,
  )
  if (!part) {
    return []
  }
  const xml = Buffer.from(part.dataBase64, 'base64').toString('utf8')
  return [...xml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?f\b[^>]*>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?f>/gu)].map((match) =>
    decodeXmlText(match[1] ?? ''),
  )
}

function normalizeChartFormulaRef(formula: string): string {
  return formula.replace(/'Data'!/gu, 'Data!')
}

function readXmlAttribute(xml: string, name: string): string | undefined {
  return new RegExp(`\\b${name}="([^"]*)"`, 'u').exec(xml)?.[1]
}

function relationshipsXml(relationships: readonly { id: string; type: string; target: string }[]): string {
  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    `<Relationships xmlns="${relationshipNamespace}">`,
    ...relationships.map(
      (relationship) => `<Relationship Id="${relationship.id}" Type="${relationship.type}" Target="${relationship.target}"/>`,
    ),
    '</Relationships>',
  ].join('')
}

function decodeXmlText(value: string): string {
  return value
    .replace(/&quot;/gu, '"')
    .replace(/&apos;/gu, "'")
    .replace(/&lt;/gu, '<')
    .replace(/&gt;/gu, '>')
    .replace(/&amp;/gu, '&')
}

function buildWorksheetChartPackageStructuralSourceXlsx(): Uint8Array {
  const zip = unzipSync(
    writeSimpleXlsxWorkbook({
      sheets: [
        {
          name: 'Data',
          cells: [
            { address: 'A1', row: 0, col: 0, value: 'Quarter' },
            { address: 'B1', row: 0, col: 1, value: 'Revenue' },
            { address: 'A2', row: 1, col: 0, value: 'Q1' },
            { address: 'B2', row: 1, col: 1, value: 10 },
            { address: 'A3', row: 2, col: 0, value: 'Q2' },
            { address: 'B3', row: 2, col: 1, value: 14 },
          ],
        },
      ],
    }),
  )
  zip['xl/worksheets/sheet1.xml'] = strToU8(addWorksheetDrawing(readZipTextFromZip(zip, 'xl/worksheets/sheet1.xml'), 'rId1'))
  zip['xl/worksheets/_rels/sheet1.xml.rels'] = strToU8(
    relationshipsXml([{ id: 'rId1', type: drawingRelationshipType, target: '../drawings/drawing1.xml' }]),
  )
  zip['xl/drawings/drawing1.xml'] = strToU8(worksheetChartDrawingXml)
  zip['xl/drawings/_rels/drawing1.xml.rels'] = strToU8(
    relationshipsXml([{ id: 'rId1', type: chartRelationshipType, target: '../charts/chart1.xml' }]),
  )
  zip['xl/charts/chart1.xml'] = strToU8(unsupportedWorksheetChartXml)
  zip['[Content_Types].xml'] = strToU8(
    upsertContentTypeOverride(
      upsertContentTypeOverride(readZipTextFromZip(zip, '[Content_Types].xml'), {
        partName: '/xl/drawings/drawing1.xml',
        contentType: drawingContentType,
      }),
      { partName: '/xl/charts/chart1.xml', contentType: chartContentType },
    ),
  )

  return zipSync(zip)
}

function addWorksheetDrawing(sheetXml: string, relationshipId: string): string {
  const withRelationshipNamespace = /xmlns:r=/u.test(sheetXml)
    ? sheetXml
    : sheetXml.replace(
        /<worksheet\b([^>]*)>/u,
        `<worksheet$1 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">`,
      )
  return withRelationshipNamespace.replace('</worksheet>', `<drawing r:id="${relationshipId}"/></worksheet>`)
}

function readZipTextFromZip(zip: Record<string, Uint8Array>, path: string): string {
  const bytes = zip[path]
  if (!bytes) {
    throw new Error(`Missing XLSX part: ${path}`)
  }
  return strFromU8(bytes)
}

function upsertContentTypeOverride(contentTypesXml: string, input: { readonly partName: string; readonly contentType: string }): string {
  if (contentTypesXml.includes(`PartName="${input.partName}"`)) {
    return contentTypesXml
  }
  return contentTypesXml.replace('</Types>', `<Override PartName="${input.partName}" ContentType="${input.contentType}"/></Types>`)
}

const headerStyleReferenceStylesXml = [
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
  '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
  '<fonts count="2"><font><sz val="11"/><name val="Aptos"/></font><font><b/></font></fonts>',
  '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>',
  '<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left style="thin"><color rgb="FF000000"/></left><right style="thin"><color rgb="FF000000"/></right><top style="thin"><color rgb="FF000000"/></top><bottom style="thin"><color rgb="FF000000"/></bottom><diagonal/></border></borders>',
  '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>',
  '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf></cellXfs>',
  '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>',
  '</styleSheet>',
].join('')

const relationshipNamespace = 'http://schemas.openxmlformats.org/package/2006/relationships'
const officeRelationshipNamespace = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const drawingRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing'
const chartRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart'
const drawingContentType = 'application/vnd.openxmlformats-officedocument.drawing+xml'
const chartContentType = 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml'

const worksheetChartDrawingXml = [
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
  '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" ',
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">',
  '<xdr:twoCellAnchor>',
  '<xdr:from><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>',
  '<xdr:to><xdr:col>8</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>12</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>',
  '<xdr:graphicFrame macro="">',
  '<xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Unsupported Chart"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>',
  '<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>',
  '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">',
  `<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="${officeRelationshipNamespace}" r:id="rId1"/>`,
  '</a:graphicData></a:graphic>',
  '</xdr:graphicFrame><xdr:clientData/>',
  '</xdr:twoCellAnchor>',
  '</xdr:wsDr>',
].join('')

const unsupportedWorksheetChartXml = [
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
  `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${officeRelationshipNamespace}">`,
  '<c:lang val="en-US"/>',
  '<c:chart><c:plotArea><c:layout/><c:doughnutChart>',
  '<c:varyColors val="1"/>',
  '<c:ser><c:idx val="0"/><c:order val="0"/>',
  '<c:tx><c:strRef><c:f>Data!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Revenue</c:v></c:pt></c:strCache></c:strRef></c:tx>',
  '<c:cat><c:strRef><c:f>Data!$A$2:$A$3</c:f><c:strCache><c:ptCount val="2"/><c:pt idx="0"><c:v>Q1</c:v></c:pt><c:pt idx="1"><c:v>Q2</c:v></c:pt></c:strCache></c:strRef></c:cat>',
  '<c:val><c:numRef><c:f>Data!$B$2:$B$3</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="2"/><c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>14</c:v></c:pt></c:numCache></c:numRef></c:val>',
  '</c:ser>',
  '<c:firstSliceAng val="0"/><c:holeSize val="50"/>',
  '</c:doughnutChart></c:plotArea>',
  '<c:legend><c:legendPos val="r"/><c:layout/></c:legend><c:plotVisOnly val="1"/></c:chart>',
  '<c:printSettings><c:headerFooter/><c:pageMargins b="0.75" l="0.7" r="0.7" t="0.75" header="0.3" footer="0.3"/><c:pageSetup/></c:printSettings>',
  '</c:chartSpace>',
].join('')
