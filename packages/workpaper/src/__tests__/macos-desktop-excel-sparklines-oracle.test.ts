import { readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'

import { exportXlsx, importXlsx } from '@bilig/excel-import'
import {
  isMacosExcelInstalled,
  runMacosExcelInspectionOracle,
  runMacosExcelStructuralOperationOracle,
  type NormalizedFormulaValue,
} from '@bilig/excel-fixtures'
import type { WorkbookSnapshot } from '@bilig/protocol'
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { describe, expect, it } from 'vitest'

import { WorkPaper } from '../index.js'
import { createExcelAccessibleTempDir, removeMacosExcelTestDir } from './macos-excel-oracle-test-utils.js'

const sparklineExtensionUri = '{05C60535-1F16-4fd2-B633-F4F36F0B64E0}'
const x14Namespace = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main'
const xmNamespace = 'http://schemas.microsoft.com/office/excel/2006/main'
const sparklineExtensionXml = `<ext uri="${sparklineExtensionUri}" xmlns:x14="${x14Namespace}"><x14:sparklineGroups xmlns:xm="${xmNamespace}"><x14:sparklineGroup type="line" displayEmptyCellsAs="gap" markers="1"><x14:colorSeries rgb="FF376092"/><x14:colorNegative rgb="FFD00000"/><x14:colorAxis rgb="FF000000"/><x14:colorMarkers rgb="FF376092"/><x14:colorFirst rgb="FF376092"/><x14:colorLast rgb="FF376092"/><x14:colorHigh rgb="FF376092"/><x14:colorLow rgb="FF376092"/><x14:sparklines><x14:sparkline><xm:f>Data!A2:D2</xm:f><xm:sqref>E2</xm:sqref></x14:sparkline><x14:sparkline><xm:f>Data!A3:D3</xm:f><xm:sqref>E3</xm:sqref></x14:sparkline></x14:sparklines></x14:sparklineGroup></x14:sparklineGroups></ext>`
const crossSheetSparklineExtensionXml = `<ext uri="${sparklineExtensionUri}" xmlns:x14="${x14Namespace}"><x14:sparklineGroups xmlns:xm="${xmNamespace}"><x14:sparklineGroup type="line" displayEmptyCellsAs="gap"><x14:sparklines><x14:sparkline><xm:f>Data!A1:D1</xm:f><xm:sqref>E2</xm:sqref></x14:sparkline></x14:sparklines></x14:sparklineGroup></x14:sparklineGroups></ext>`
const expectedInsertedRowValues: readonly NormalizedFormulaValue[] = [
  { kind: 'string', value: '' },
  { kind: 'number', value: 10 },
  { kind: 'string', value: '' },
  { kind: 'number', value: 18 },
  { kind: 'string', value: '' },
] as const

describe('macOS Desktop Excel sparkline oracle', () => {
  it('preserves and structurally rewrites imported sparkline extension XML through WorkPaper export', () => {
    const workpaper = WorkPaper.buildFromSnapshot(importXlsx(buildWorkbookWithSparklineExtension(), 'sparkline-source.xlsx').snapshot)
    try {
      const sheet = workpaper.getSheetId('Data')
      if (sheet === undefined) {
        throw new Error('Expected Data sheet to be available')
      }
      workpaper.addRows(sheet, 1, 1)

      const exported = exportXlsx(workpaper.exportSnapshot())
      const reimported = importXlsx(exported, 'sparkline-headless-roundtrip.xlsx')

      expect(sparklineRefs(reimported.snapshot.sheets[0]?.metadata?.sparklines?.xml)).toEqual([
        { formula: 'Data!A3:D3', sqref: 'E3' },
        { formula: 'Data!A4:D4', sqref: 'E4' },
      ])
      expect(sparklineCount(exported)).toBe(2)
    } finally {
      workpaper.dispose()
    }
  })

  it('rewrites cross-sheet sparkline source refs when the source sheet changes', () => {
    const workpaper = WorkPaper.buildFromSnapshot(
      importXlsx(buildWorkbookWithCrossSheetSparklineExtension(), 'cross-sheet-sparkline-source.xlsx').snapshot,
    )
    try {
      const sheet = workpaper.getSheetId('Data')
      if (sheet === undefined) {
        throw new Error('Expected Data sheet to be available')
      }
      workpaper.addRows(sheet, 0, 1)

      const exported = exportXlsx(workpaper.exportSnapshot())
      const reimported = importXlsx(exported, 'cross-sheet-sparkline-headless-roundtrip.xlsx')

      expect(sparklineRefs(reimported.snapshot.sheets[1]?.metadata?.sparklines?.xml)).toEqual([{ formula: 'Data!A2:D2', sqref: 'E2' }])
      expect(sparklineCount(exported, 'xl/worksheets/sheet2.xml')).toBe(1)
    } finally {
      workpaper.dispose()
    }
  })

  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'matches Desktop Excel sparkline source and output refs after structural row inserts',
    () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = createExcelAccessibleTempDir('bilig-headless-excel-sparklines-oracle-')
      try {
        const sourcePath = join(tempDir, 'excel-sparklines-structural-source.xlsx')
        writeFileSync(sourcePath, buildWorkbookWithSparklineExtension())

        const excelResult = runMacosExcelStructuralOperationOracle({
          workbookPath: sourcePath,
          worksheetName: 'Data',
          operations: [{ kind: 'insertRows', range: '2:2' }],
          inspectCells: ['A2', 'A3', 'E3', 'A4', 'E4'],
          saveWorkbook: true,
          timeoutMs: 90_000,
        })
        expect(excelResult.cells.map((cell) => cell.value)).toEqual(expectedInsertedRowValues)

        const excelTruth = importXlsx(new Uint8Array(readFileSync(sourcePath)), 'excel-sparklines-structural-truth.xlsx')
        const excelSparklineRefs = sparklineRefs(excelTruth.snapshot.sheets[0]?.metadata?.sparklines?.xml)
        expect(excelSparklineRefs).toHaveLength(2)

        const workpaper = WorkPaper.buildFromSnapshot(
          importXlsx(buildWorkbookWithSparklineExtension(), 'headless-sparklines-source.xlsx').snapshot,
        )
        try {
          const sheet = workpaper.getSheetId('Data')
          if (sheet === undefined) {
            throw new Error('Expected Data sheet to be available')
          }
          workpaper.addRows(sheet, 1, 1)

          const headlessPath = join(tempDir, 'headless-sparklines-structural.xlsx')
          writeFileSync(headlessPath, exportXlsx(workpaper.exportSnapshot()))
          const headlessExcel = runMacosExcelInspectionOracle({
            workbookPath: headlessPath,
            worksheetName: 'Data',
            formulaCells: [],
            inspectCells: ['A2', 'A3', 'E3', 'A4', 'E4'],
            saveWorkbook: true,
            timeoutMs: 90_000,
          })
          expect(headlessExcel.cells).toEqual(excelResult.cells)

          const headlessTruth = importXlsx(new Uint8Array(readFileSync(headlessPath)), 'headless-sparklines-structural-truth.xlsx')
          expect(sparklineRefs(headlessTruth.snapshot.sheets[0]?.metadata?.sparklines?.xml)).toEqual(excelSparklineRefs)
          expect(sparklineCount(new Uint8Array(readFileSync(headlessPath)))).toBe(2)
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
    'matches Desktop Excel cross-sheet sparkline source refs after source sheet row inserts',
    () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = createExcelAccessibleTempDir('bilig-headless-excel-cross-sheet-sparklines-oracle-')
      try {
        const sourcePath = join(tempDir, 'excel-cross-sheet-sparklines-source.xlsx')
        writeFileSync(sourcePath, buildWorkbookWithCrossSheetSparklineExtension())

        const excelResult = runMacosExcelStructuralOperationOracle({
          workbookPath: sourcePath,
          worksheetName: 'Data',
          operations: [{ kind: 'insertRows', range: '1:1' }],
          inspectCells: ['A1', 'A2', 'D2'],
          saveWorkbook: true,
          timeoutMs: 90_000,
        })
        expect(excelResult.cells.map((cell) => cell.value)).toEqual([
          { kind: 'string', value: '' },
          { kind: 'number', value: 10 },
          { kind: 'number', value: 40 },
        ])

        const excelTruth = importXlsx(new Uint8Array(readFileSync(sourcePath)), 'excel-cross-sheet-sparklines-truth.xlsx')
        const excelSparklineRefs = sparklineRefs(excelTruth.snapshot.sheets[1]?.metadata?.sparklines?.xml)
        expect(excelSparklineRefs).toEqual([{ formula: 'Data!A2:D2', sqref: 'E2' }])

        const workpaper = WorkPaper.buildFromSnapshot(
          importXlsx(buildWorkbookWithCrossSheetSparklineExtension(), 'headless-cross-sheet-sparklines-source.xlsx').snapshot,
        )
        try {
          const sheet = workpaper.getSheetId('Data')
          if (sheet === undefined) {
            throw new Error('Expected Data sheet to be available')
          }
          workpaper.addRows(sheet, 0, 1)

          const headlessPath = join(tempDir, 'headless-cross-sheet-sparklines.xlsx')
          writeFileSync(headlessPath, exportXlsx(workpaper.exportSnapshot()))
          const headlessExcel = runMacosExcelInspectionOracle({
            workbookPath: headlessPath,
            worksheetName: 'Data',
            formulaCells: [],
            inspectCells: ['A1', 'A2', 'D2'],
            saveWorkbook: true,
            timeoutMs: 90_000,
          })
          expect(headlessExcel.cells).toEqual(excelResult.cells)

          const headlessTruth = importXlsx(new Uint8Array(readFileSync(headlessPath)), 'headless-cross-sheet-sparklines-truth.xlsx')
          expect(sparklineRefs(headlessTruth.snapshot.sheets[1]?.metadata?.sparklines?.xml)).toEqual(excelSparklineRefs)
          expect(sparklineCount(new Uint8Array(readFileSync(headlessPath)), 'xl/worksheets/sheet2.xml')).toBe(1)
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

function buildWorkbookWithSparklineExtension(): Uint8Array {
  const zip = unzipSync(exportXlsx(buildSparklineWorkbook()))
  const sheetPath = 'xl/worksheets/sheet1.xml'
  const sheetXml = strFromU8(zip[sheetPath] ?? new Uint8Array())
  zip[sheetPath] = strToU8(sheetXml.replace('</worksheet>', `<extLst>${sparklineExtensionXml}</extLst></worksheet>`))
  return zipSync(zip)
}

function buildWorkbookWithCrossSheetSparklineExtension(): Uint8Array {
  const zip = unzipSync(exportXlsx(buildCrossSheetSparklineWorkbook()))
  const sheetPath = 'xl/worksheets/sheet2.xml'
  const sheetXml = strFromU8(zip[sheetPath] ?? new Uint8Array())
  zip[sheetPath] = strToU8(sheetXml.replace('</worksheet>', `<extLst>${crossSheetSparklineExtensionXml}</extLst></worksheet>`))
  return zipSync(zip)
}

function buildSparklineWorkbook(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: {
      name: 'Desktop Excel sparkline oracle',
    },
    sheets: [
      {
        id: 1,
        name: 'Data',
        order: 0,
        cells: [
          { address: 'A1', value: 'Q1' },
          { address: 'B1', value: 'Q2' },
          { address: 'C1', value: 'Q3' },
          { address: 'D1', value: 'Q4' },
          { address: 'E1', value: 'Trend' },
          { address: 'A2', value: 10 },
          { address: 'B2', value: 20 },
          { address: 'C2', value: 15 },
          { address: 'D2', value: 30 },
          { address: 'E2', value: '' },
          { address: 'A3', value: 18 },
          { address: 'B3', value: 12 },
          { address: 'C3', value: 24 },
          { address: 'D3', value: 28 },
          { address: 'E3', value: '' },
        ],
      },
    ],
  }
}

function buildCrossSheetSparklineWorkbook(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: {
      name: 'Desktop Excel cross-sheet sparkline oracle',
    },
    sheets: [
      {
        id: 1,
        name: 'Data',
        order: 0,
        cells: [
          { address: 'A1', value: 10 },
          { address: 'B1', value: 20 },
          { address: 'C1', value: 30 },
          { address: 'D1', value: 40 },
        ],
      },
      {
        id: 2,
        name: 'Dashboard',
        order: 1,
        cells: [{ address: 'E2', value: '' }],
      },
    ],
  }
}

function sparklineRefs(xml: string | undefined): Array<{ readonly formula: string; readonly sqref: string }> {
  if (!xml) {
    throw new Error('Expected sparkline XML')
  }
  const matches = [
    ...xml.matchAll(
      /<(?:[A-Za-z_][\w.-]*:)?sparkline><(?:[A-Za-z_][\w.-]*:)?f>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?f><(?:[A-Za-z_][\w.-]*:)?sqref>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?sqref><\/(?:[A-Za-z_][\w.-]*:)?sparkline>/gu,
    ),
  ]
  return matches.map((match) => ({
    formula: match[1] ?? '',
    sqref: match[2] ?? '',
  }))
}

function sparklineCount(bytes: Uint8Array, sheetPath = 'xl/worksheets/sheet1.xml'): number {
  return worksheetXml(bytes, sheetPath).match(/<(?:[A-Za-z_][\w.-]*:)?sparkline\b/gu)?.length ?? 0
}

function worksheetXml(bytes: Uint8Array, sheetPath = 'xl/worksheets/sheet1.xml'): string {
  return strFromU8(unzipSync(bytes)[sheetPath] ?? new Uint8Array())
}
