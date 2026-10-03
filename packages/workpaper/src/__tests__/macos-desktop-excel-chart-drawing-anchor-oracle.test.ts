import { readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'

import { exportXlsx, importXlsx } from '@bilig/excel-import'
import { isMacosExcelInstalled, runMacosExcelInspectionOracle, runMacosExcelStructuralOperationOracle } from '@bilig/excel-fixtures'
import type { WorkbookChartAnchorSnapshot, WorkbookSnapshot } from '@bilig/protocol'
import { strFromU8, unzipSync } from 'fflate'
import { describe, expect, it } from 'vitest'

import { WorkPaper } from '../index.js'
import { createExcelAccessibleTempDir, removeMacosExcelTestDir } from './macos-excel-oracle-test-utils.js'

const expectedShiftedAnchor: WorkbookChartAnchorSnapshot = {
  kind: 'twoCell',
  editAs: 'oneCell',
  from: { row: 2, col: 4, rowOffset: 12345, colOffset: 67890 },
  to: { row: 10, col: 9, rowOffset: 22222, colOffset: 33333 },
}

describe('macOS Desktop Excel chart drawing anchor oracle', () => {
  it('preserves supported chart drawing anchors and rewrites them after headless row inserts', () => {
    const imported = importXlsx(exportXlsx(workbookWithOffsetChartAnchor()), 'chart-anchor-source.xlsx')
    expect(importedChartAnchor(imported.snapshot)).toEqual({
      kind: 'twoCell',
      editAs: 'oneCell',
      from: { row: 1, col: 4, rowOffset: 12345, colOffset: 67890 },
      to: { row: 9, col: 9, rowOffset: 22222, colOffset: 33333 },
    })

    const workpaper = WorkPaper.buildFromSnapshot(imported.snapshot)
    try {
      const dashboard = workpaper.getSheetId('Dashboard')
      if (dashboard === undefined) {
        throw new Error('Expected Dashboard sheet to be available')
      }
      workpaper.addRows(dashboard, 0, 1)

      const exported = exportXlsx(workpaper.exportSnapshot())
      const reimported = importXlsx(exported, 'headless-chart-anchor-roundtrip.xlsx')

      expect(importedChartAnchor(reimported.snapshot)).toEqual(expectedShiftedAnchor)
      expect(readDrawingXml(exported)).toContain('editAs="oneCell"')
      expect(readDrawingXml(exported)).toContain('<xdr:rowOff>12345</xdr:rowOff>')
      expect(readDrawingXml(exported)).toContain('<xdr:colOff>67890</xdr:colOff>')
    } finally {
      workpaper.dispose()
    }
  })

  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'matches Desktop Excel chart drawing anchors after structural row inserts',
    () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = createExcelAccessibleTempDir('bilig-headless-excel-chart-anchor-oracle-')
      try {
        const sourcePath = join(tempDir, 'chart-anchor-source.xlsx')
        writeFileSync(sourcePath, exportXlsx(workbookWithOffsetChartAnchor()))

        const excelResult = runMacosExcelStructuralOperationOracle({
          workbookPath: sourcePath,
          worksheetName: 'Dashboard',
          operations: [{ kind: 'insertRows', range: '1:1' }],
          inspectCells: ['A1', 'A2'],
          saveWorkbook: true,
          timeoutMs: 90_000,
        })
        expect(excelResult.cells.map((cell) => cell.value)).toEqual([
          { kind: 'string', value: '' },
          { kind: 'string', value: 'dashboard' },
        ])

        const excelTruth = importXlsx(new Uint8Array(readFileSync(sourcePath)), 'excel-chart-anchor-truth.xlsx')
        expect(importedChartAnchor(excelTruth.snapshot)).toEqual(expectedShiftedAnchor)

        const workpaper = WorkPaper.buildFromSnapshot(
          importXlsx(exportXlsx(workbookWithOffsetChartAnchor()), 'headless-source.xlsx').snapshot,
        )
        try {
          const dashboard = workpaper.getSheetId('Dashboard')
          if (dashboard === undefined) {
            throw new Error('Expected Dashboard sheet to be available')
          }
          workpaper.addRows(dashboard, 0, 1)

          const headlessPath = join(tempDir, 'headless-chart-anchor-oracle.xlsx')
          writeFileSync(headlessPath, exportXlsx(workpaper.exportSnapshot()))
          const headlessExcel = runMacosExcelInspectionOracle({
            workbookPath: headlessPath,
            worksheetName: 'Dashboard',
            formulaCells: [],
            inspectCells: ['A1', 'A2'],
            saveWorkbook: true,
            timeoutMs: 90_000,
          })
          expect(headlessExcel.cells).toEqual(excelResult.cells)

          const headlessTruth = importXlsx(new Uint8Array(readFileSync(headlessPath)), 'excel-saved-headless-chart-anchor.xlsx')
          expect(importedChartAnchor(headlessTruth.snapshot)).toEqual(importedChartAnchor(excelTruth.snapshot))
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

function importedChartAnchor(snapshot: WorkbookSnapshot): WorkbookChartAnchorSnapshot {
  const anchor = snapshot.workbook.metadata?.charts?.find((chart) => chart.id === 'Revenue Chart')?.anchor
  if (!anchor) {
    throw new Error('Expected imported chart anchor metadata')
  }
  return anchor
}

function readDrawingXml(bytes: Uint8Array): string {
  return strFromU8(unzipSync(bytes)['xl/drawings/drawing1.xml'] ?? new Uint8Array())
}

function workbookWithOffsetChartAnchor(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: {
      name: 'Desktop Excel chart anchor oracle',
      metadata: {
        charts: [
          {
            id: 'Revenue Chart',
            sheetName: 'Dashboard',
            address: 'E2',
            source: { sheetName: 'Data', startAddress: 'A1', endAddress: 'B4' },
            chartType: 'line',
            anchor: {
              kind: 'twoCell',
              editAs: 'oneCell',
              from: { row: 1, col: 4, rowOffset: 12345, colOffset: 67890 },
              to: { row: 9, col: 9, rowOffset: 22222, colOffset: 33333 },
            },
            rows: 8,
            cols: 5,
            title: 'Revenue',
            seriesOrientation: 'columns',
            firstRowAsHeaders: true,
            firstColumnAsLabels: true,
            legendPosition: 'right',
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
          { address: 'A1', value: 'Month' },
          { address: 'B1', value: 'Revenue' },
          { address: 'A2', value: 'Jan' },
          { address: 'B2', value: 10 },
          { address: 'A3', value: 'Feb' },
          { address: 'B3', value: 20 },
          { address: 'A4', value: 'Mar' },
          { address: 'B4', value: 30 },
        ],
      },
      {
        id: 2,
        name: 'Dashboard',
        order: 1,
        cells: [{ address: 'A1', value: 'dashboard' }],
      },
    ],
  }
}
