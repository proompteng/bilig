import { readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'

import { exportXlsx, importXlsx } from '@bilig/excel-import'
import {
  isMacosExcelInstalled,
  runMacosExcelInspectionOracle,
  runMacosExcelStructuralOperationOracle,
  type NormalizedFormulaValue,
} from '@bilig/excel-fixtures'
import { writeSimpleXlsxWorkbook } from '@bilig/xlsx'
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { describe, expect, it } from 'vitest'

import { WorkPaper } from '../index.js'
import { createExcelAccessibleTempDir, removeMacosExcelTestDir } from './macos-excel-oracle-test-utils.js'

const inspectedCells = ['A1', 'B1', 'C1', 'D1'] as const
const expectedValues: readonly NormalizedFormulaValue[] = [
  { kind: 'number', value: 10 },
  { kind: 'number', value: 20 },
  { kind: 'number', value: 30 },
  { kind: 'string', value: '' },
] as const
const expectedInsertedRowSqrefs = ['A2:A4', 'B2:B4', 'C2:C4'] as const

describe('macOS Desktop Excel conditional format artifact oracle', () => {
  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'preserves Desktop Excel advanced visual conditional-format rules after a headless edit',
    () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = createExcelAccessibleTempDir('bilig-headless-excel-cf-artifacts-oracle-')
      try {
        const sourcePath = join(tempDir, 'excel-advanced-conditional-format-oracle.xlsx')
        writeFileSync(sourcePath, buildAdvancedConditionalFormattingWorkbook())

        const excelResult = runMacosExcelInspectionOracle({
          workbookPath: sourcePath,
          worksheetName: 'Dashboard',
          formulaCells: [],
          inspectCells: inspectedCells,
          saveWorkbook: true,
          timeoutMs: 90_000,
        })
        expect(excelResult.cells.map((cell) => cell.value)).toEqual(expectedValues)

        const excelTruth = importXlsx(new Uint8Array(readFileSync(sourcePath)), 'excel-advanced-conditional-format-oracle.xlsx')
        expectConditionalFormatArtifacts(excelTruth.snapshot.sheets[0]?.metadata?.conditionalFormatArtifacts?.xml)

        const workpaper = WorkPaper.buildFromSnapshot(excelTruth.snapshot)
        const sheet = workpaper.getSheetId('Dashboard')
        if (sheet === undefined) {
          throw new Error('Expected Dashboard sheet to be available')
        }
        workpaper.setCellContents({ sheet, row: 0, col: 3 }, 'headless edit')
        const headlessSnapshot = workpaper.exportSnapshot()
        expectConditionalFormatArtifacts(headlessSnapshot.sheets[0]?.metadata?.conditionalFormatArtifacts?.xml)

        const headlessPath = join(tempDir, 'headless-advanced-conditional-format-oracle.xlsx')
        writeFileSync(headlessPath, exportXlsx(headlessSnapshot))
        const headlessExcel = runMacosExcelInspectionOracle({
          workbookPath: headlessPath,
          worksheetName: 'Dashboard',
          formulaCells: [],
          inspectCells: inspectedCells,
          saveWorkbook: true,
          timeoutMs: 90_000,
        })
        expect(headlessExcel.cells.map((cell) => cell.value)).toEqual([
          { kind: 'number', value: 10 },
          { kind: 'number', value: 20 },
          { kind: 'number', value: 30 },
          { kind: 'string', value: 'headless edit' },
        ])

        const headlessExcelTruth = importXlsx(
          new Uint8Array(readFileSync(headlessPath)),
          'headless-advanced-conditional-format-oracle.xlsx',
        )
        expectConditionalFormatArtifacts(headlessExcelTruth.snapshot.sheets[0]?.metadata?.conditionalFormatArtifacts?.xml)
        expect(readWorksheetXml(headlessPath).match(/<conditionalFormatting\b/gu)).toHaveLength(3)
      } finally {
        removeMacosExcelTestDir(tempDir)
      }
    },
    120_000,
  )

  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'matches Desktop Excel advanced visual conditional-format artifact ranges after structural row inserts',
    () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = createExcelAccessibleTempDir('bilig-headless-excel-cf-artifacts-structural-oracle-')
      try {
        const sourcePath = join(tempDir, 'excel-advanced-conditional-format-structural-oracle.xlsx')
        writeFileSync(sourcePath, buildAdvancedConditionalFormattingWorkbook())

        const excelResult = runMacosExcelStructuralOperationOracle({
          workbookPath: sourcePath,
          worksheetName: 'Dashboard',
          operations: [{ kind: 'insertRows', range: '1:1' }],
          inspectCells: ['A1', 'A2', 'B2', 'C2'],
          saveWorkbook: true,
          timeoutMs: 90_000,
        })
        expect(excelResult.cells.map((cell) => cell.value)).toEqual([
          { kind: 'string', value: '' },
          { kind: 'number', value: 10 },
          { kind: 'number', value: 20 },
          { kind: 'number', value: 30 },
        ])

        const excelTruth = importXlsx(new Uint8Array(readFileSync(sourcePath)), 'excel-advanced-cf-structural-oracle.xlsx')
        const excelTruthArtifacts = excelTruth.snapshot.sheets[0]?.metadata?.conditionalFormatArtifacts?.xml
        expect(extractConditionalFormatSqrefs(excelTruthArtifacts)).toEqual(expectedInsertedRowSqrefs)

        const workpaper = WorkPaper.buildFromSnapshot(
          importXlsx(buildAdvancedConditionalFormattingWorkbook(), 'headless-cf-source.xlsx').snapshot,
        )
        try {
          const sheet = workpaper.getSheetId('Dashboard')
          if (sheet === undefined) {
            throw new Error('Expected Dashboard sheet to be available')
          }
          workpaper.addRows(sheet, 0, 1)
          expect(extractConditionalFormatSqrefs(workpaper.exportSnapshot().sheets[0]?.metadata?.conditionalFormatArtifacts?.xml)).toEqual(
            expectedInsertedRowSqrefs,
          )

          const headlessPath = join(tempDir, 'headless-advanced-conditional-format-structural-oracle.xlsx')
          writeFileSync(headlessPath, exportXlsx(workpaper.exportSnapshot()))
          const headlessExcel = runMacosExcelInspectionOracle({
            workbookPath: headlessPath,
            worksheetName: 'Dashboard',
            formulaCells: [],
            inspectCells: ['A1', 'A2', 'B2', 'C2'],
            saveWorkbook: true,
            timeoutMs: 90_000,
          })
          expect(headlessExcel.cells).toEqual(excelResult.cells)

          const headlessExcelTruth = importXlsx(new Uint8Array(readFileSync(headlessPath)), 'headless-advanced-cf-structural-oracle.xlsx')
          expect(extractConditionalFormatSqrefs(headlessExcelTruth.snapshot.sheets[0]?.metadata?.conditionalFormatArtifacts?.xml)).toEqual(
            extractConditionalFormatSqrefs(excelTruthArtifacts),
          )
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
    'matches Desktop Excel namespace-qualified conditional-format artifact ranges after structural row inserts',
    () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = createExcelAccessibleTempDir('bilig-headless-excel-cf-prefixed-structural-oracle-')
      try {
        const sourcePath = join(tempDir, 'excel-prefixed-conditional-format-structural-oracle.xlsx')
        writeFileSync(sourcePath, buildPrefixedConditionalFormattingWorkbook())

        const excelResult = runMacosExcelStructuralOperationOracle({
          workbookPath: sourcePath,
          worksheetName: 'Dashboard',
          operations: [{ kind: 'insertRows', range: '1:1' }],
          inspectCells: ['A1', 'A2', 'B2', 'C2'],
          saveWorkbook: true,
          timeoutMs: 90_000,
        })
        expect(excelResult.cells.map((cell) => cell.value)).toEqual([
          { kind: 'string', value: '' },
          { kind: 'number', value: 10 },
          { kind: 'number', value: 20 },
          { kind: 'number', value: 30 },
        ])

        const excelTruth = importXlsx(new Uint8Array(readFileSync(sourcePath)), 'excel-prefixed-cf-structural-oracle.xlsx')
        const excelTruthArtifacts = excelTruth.snapshot.sheets[0]?.metadata?.conditionalFormatArtifacts?.xml
        expect(extractConditionalFormatSqrefs(excelTruthArtifacts)).toEqual(expectedInsertedRowSqrefs)

        const workpaper = WorkPaper.buildFromSnapshot(
          importXlsx(buildPrefixedConditionalFormattingWorkbook(), 'headless-prefixed-cf-source.xlsx').snapshot,
        )
        try {
          const sheet = workpaper.getSheetId('Dashboard')
          if (sheet === undefined) {
            throw new Error('Expected Dashboard sheet to be available')
          }
          workpaper.addRows(sheet, 0, 1)
          expect(extractConditionalFormatSqrefs(workpaper.exportSnapshot().sheets[0]?.metadata?.conditionalFormatArtifacts?.xml)).toEqual(
            expectedInsertedRowSqrefs,
          )

          const headlessPath = join(tempDir, 'headless-prefixed-conditional-format-structural-oracle.xlsx')
          writeFileSync(headlessPath, exportXlsx(workpaper.exportSnapshot()))
          const headlessExcel = runMacosExcelInspectionOracle({
            workbookPath: headlessPath,
            worksheetName: 'Dashboard',
            formulaCells: [],
            inspectCells: ['A1', 'A2', 'B2', 'C2'],
            saveWorkbook: true,
            timeoutMs: 90_000,
          })
          expect(headlessExcel.cells).toEqual(excelResult.cells)

          const headlessExcelTruth = importXlsx(new Uint8Array(readFileSync(headlessPath)), 'headless-prefixed-cf-structural-oracle.xlsx')
          expect(extractConditionalFormatSqrefs(headlessExcelTruth.snapshot.sheets[0]?.metadata?.conditionalFormatArtifacts?.xml)).toEqual(
            extractConditionalFormatSqrefs(excelTruthArtifacts),
          )
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
    'matches Desktop Excel conditional-format artifact formulas after structural row inserts',
    () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = createExcelAccessibleTempDir('bilig-headless-excel-cf-formula-structural-oracle-')
      try {
        const sourcePath = join(tempDir, 'excel-formula-conditional-format-structural-oracle.xlsx')
        writeFileSync(sourcePath, buildFormulaConditionalFormattingWorkbook())

        const excelResult = runMacosExcelStructuralOperationOracle({
          workbookPath: sourcePath,
          worksheetName: 'Dashboard',
          operations: [{ kind: 'insertRows', range: '1:1' }],
          inspectCells: ['A1', 'A2', 'A3', 'A4'],
          saveWorkbook: true,
          timeoutMs: 90_000,
        })
        expect(excelResult.cells.map((cell) => cell.value)).toEqual([
          { kind: 'string', value: '' },
          { kind: 'number', value: 10 },
          { kind: 'number', value: 20 },
          { kind: 'number', value: 30 },
        ])

        const excelTruth = importXlsx(new Uint8Array(readFileSync(sourcePath)), 'excel-formula-cf-structural-oracle.xlsx')
        const excelTruthArtifacts = excelTruth.snapshot.sheets[0]?.metadata?.conditionalFormatArtifacts?.xml
        expect(extractConditionalFormatSqrefs(excelTruthArtifacts)).toEqual(['A2:A4'])
        expect(extractConditionalFormatFormulas(excelTruthArtifacts)).toEqual(['A2>15'])

        const workpaper = WorkPaper.buildFromSnapshot(
          importXlsx(buildFormulaConditionalFormattingWorkbook(), 'headless-formula-cf-source.xlsx').snapshot,
        )
        try {
          const sheet = workpaper.getSheetId('Dashboard')
          if (sheet === undefined) {
            throw new Error('Expected Dashboard sheet to be available')
          }
          workpaper.addRows(sheet, 0, 1)
          const headlessArtifacts = workpaper.exportSnapshot().sheets[0]?.metadata?.conditionalFormatArtifacts?.xml
          expect(extractConditionalFormatSqrefs(headlessArtifacts)).toEqual(['A2:A4'])
          expect(extractConditionalFormatFormulas(headlessArtifacts)).toEqual(['A2>15'])

          const headlessPath = join(tempDir, 'headless-formula-conditional-format-structural-oracle.xlsx')
          writeFileSync(headlessPath, exportXlsx(workpaper.exportSnapshot()))
          const headlessExcel = runMacosExcelInspectionOracle({
            workbookPath: headlessPath,
            worksheetName: 'Dashboard',
            formulaCells: [],
            inspectCells: ['A1', 'A2', 'A3', 'A4'],
            saveWorkbook: true,
            timeoutMs: 90_000,
          })
          expect(headlessExcel.cells).toEqual(excelResult.cells)

          const headlessExcelTruth = importXlsx(new Uint8Array(readFileSync(headlessPath)), 'headless-formula-cf-structural-oracle.xlsx')
          const headlessExcelArtifacts = headlessExcelTruth.snapshot.sheets[0]?.metadata?.conditionalFormatArtifacts?.xml
          expect(extractConditionalFormatSqrefs(headlessExcelArtifacts)).toEqual(extractConditionalFormatSqrefs(excelTruthArtifacts))
          expect(extractConditionalFormatFormulas(headlessExcelArtifacts)).toEqual(extractConditionalFormatFormulas(excelTruthArtifacts))
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
    'matches Desktop Excel cross-sheet conditional-format artifact formulas after target sheet row inserts',
    () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = createExcelAccessibleTempDir('bilig-headless-excel-cf-cross-sheet-formula-oracle-')
      try {
        const sourcePath = join(tempDir, 'excel-cross-sheet-formula-conditional-format-oracle.xlsx')
        writeFileSync(sourcePath, buildCrossSheetFormulaConditionalFormattingWorkbook())

        const excelResult = runMacosExcelStructuralOperationOracle({
          workbookPath: sourcePath,
          worksheetName: 'Inputs',
          operations: [{ kind: 'insertRows', range: '1:1' }],
          inspectCells: ['A1', 'A2', 'A3', 'A4'],
          saveWorkbook: true,
          timeoutMs: 90_000,
        })
        expect(excelResult.cells.map((cell) => cell.value)).toEqual([
          { kind: 'string', value: '' },
          { kind: 'number', value: 10 },
          { kind: 'number', value: 20 },
          { kind: 'number', value: 30 },
        ])

        const excelTruth = importXlsx(new Uint8Array(readFileSync(sourcePath)), 'excel-cross-sheet-cf-structural-oracle.xlsx')
        const excelTruthArtifacts = excelTruth.snapshot.sheets[0]?.metadata?.conditionalFormatArtifacts?.xml
        expect(extractConditionalFormatSqrefs(excelTruthArtifacts)).toEqual(['A1:A3'])
        expect(extractConditionalFormatFormulas(excelTruthArtifacts)).toEqual(['Inputs!A2>15'])

        const workpaper = WorkPaper.buildFromSnapshot(
          importXlsx(buildCrossSheetFormulaConditionalFormattingWorkbook(), 'headless-cross-sheet-formula-cf-source.xlsx').snapshot,
        )
        try {
          const sheet = workpaper.getSheetId('Inputs')
          if (sheet === undefined) {
            throw new Error('Expected Inputs sheet to be available')
          }
          workpaper.addRows(sheet, 0, 1)
          const headlessArtifacts = workpaper.exportSnapshot().sheets[0]?.metadata?.conditionalFormatArtifacts?.xml
          expect(extractConditionalFormatSqrefs(headlessArtifacts)).toEqual(['A1:A3'])
          expect(extractConditionalFormatFormulas(headlessArtifacts)).toEqual(['Inputs!A2>15'])

          const headlessPath = join(tempDir, 'headless-cross-sheet-formula-conditional-format-oracle.xlsx')
          writeFileSync(headlessPath, exportXlsx(workpaper.exportSnapshot()))
          const headlessExcel = runMacosExcelInspectionOracle({
            workbookPath: headlessPath,
            worksheetName: 'Inputs',
            formulaCells: [],
            inspectCells: ['A1', 'A2', 'A3', 'A4'],
            saveWorkbook: true,
            timeoutMs: 90_000,
          })
          expect(headlessExcel.cells).toEqual(excelResult.cells)

          const headlessExcelTruth = importXlsx(
            new Uint8Array(readFileSync(headlessPath)),
            'headless-cross-sheet-cf-structural-oracle.xlsx',
          )
          const headlessExcelArtifacts = headlessExcelTruth.snapshot.sheets[0]?.metadata?.conditionalFormatArtifacts?.xml
          expect(extractConditionalFormatSqrefs(headlessExcelArtifacts)).toEqual(extractConditionalFormatSqrefs(excelTruthArtifacts))
          expect(extractConditionalFormatFormulas(headlessExcelArtifacts)).toEqual(extractConditionalFormatFormulas(excelTruthArtifacts))
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
    'matches Desktop Excel x14 conditional-format artifact ranges after owner sheet row inserts',
    () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = createExcelAccessibleTempDir('bilig-headless-excel-cf-x14-sqref-oracle-')
      try {
        const sourcePath = join(tempDir, 'excel-x14-conditional-format-source.xlsx')
        writeFileSync(sourcePath, buildCrossSheetFormulaConditionalFormattingWorkbook())

        runMacosExcelInspectionOracle({
          workbookPath: sourcePath,
          worksheetName: 'Dashboard',
          formulaCells: [],
          inspectCells: ['A1', 'A2', 'A3'],
          saveWorkbook: true,
          timeoutMs: 90_000,
        })
        const x14Source = importXlsx(new Uint8Array(readFileSync(sourcePath)), 'excel-x14-conditional-format-source.xlsx')
        const x14SourceArtifacts = x14Source.snapshot.sheets[0]?.metadata?.conditionalFormatArtifacts?.xml
        expect(extractConditionalFormatSqrefs(x14SourceArtifacts)).toEqual(['A1:A3'])
        expect(extractConditionalFormatFormulas(x14SourceArtifacts)).toEqual(['Inputs!A1>15'])

        const excelResult = runMacosExcelStructuralOperationOracle({
          workbookPath: sourcePath,
          worksheetName: 'Dashboard',
          operations: [{ kind: 'insertRows', range: '1:1' }],
          inspectCells: ['A1', 'A2', 'A3', 'A4'],
          saveWorkbook: true,
          timeoutMs: 90_000,
        })
        expect(excelResult.cells.map((cell) => cell.value)).toEqual([
          { kind: 'string', value: '' },
          { kind: 'number', value: 10 },
          { kind: 'number', value: 20 },
          { kind: 'number', value: 30 },
        ])

        const excelTruth = importXlsx(new Uint8Array(readFileSync(sourcePath)), 'excel-x14-cf-owner-structural-oracle.xlsx')
        const excelTruthArtifacts = excelTruth.snapshot.sheets[0]?.metadata?.conditionalFormatArtifacts?.xml
        expect(extractConditionalFormatSqrefs(excelTruthArtifacts)).toEqual(['A2:A4'])
        expect(extractConditionalFormatFormulas(excelTruthArtifacts)).toEqual(['Inputs!A1>15'])

        const workpaper = WorkPaper.buildFromSnapshot(x14Source.snapshot)
        try {
          const sheet = workpaper.getSheetId('Dashboard')
          if (sheet === undefined) {
            throw new Error('Expected Dashboard sheet to be available')
          }
          workpaper.addRows(sheet, 0, 1)
          const headlessArtifacts = workpaper.exportSnapshot().sheets[0]?.metadata?.conditionalFormatArtifacts?.xml
          expect(extractConditionalFormatSqrefs(headlessArtifacts)).toEqual(['A2:A4'])
          expect(extractConditionalFormatFormulas(headlessArtifacts)).toEqual(['Inputs!A1>15'])

          const headlessPath = join(tempDir, 'headless-x14-conditional-format-oracle.xlsx')
          writeFileSync(headlessPath, exportXlsx(workpaper.exportSnapshot()))
          const headlessExcel = runMacosExcelInspectionOracle({
            workbookPath: headlessPath,
            worksheetName: 'Dashboard',
            formulaCells: [],
            inspectCells: ['A1', 'A2', 'A3', 'A4'],
            saveWorkbook: true,
            timeoutMs: 90_000,
          })
          expect(headlessExcel.cells).toEqual(excelResult.cells)

          const headlessExcelTruth = importXlsx(new Uint8Array(readFileSync(headlessPath)), 'headless-x14-cf-owner-oracle.xlsx')
          const headlessExcelArtifacts = headlessExcelTruth.snapshot.sheets[0]?.metadata?.conditionalFormatArtifacts?.xml
          expect(extractConditionalFormatSqrefs(headlessExcelArtifacts)).toEqual(extractConditionalFormatSqrefs(excelTruthArtifacts))
          expect(extractConditionalFormatFormulas(headlessExcelArtifacts)).toEqual(extractConditionalFormatFormulas(excelTruthArtifacts))
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

function expectConditionalFormatArtifacts(xml: string | undefined): void {
  expect(xml).toContain('type="dataBar"')
  expect(xml).toContain('type="colorScale"')
  expect(xml).toContain('type="iconSet"')
}

function extractConditionalFormatSqrefs(xml: string | undefined): string[] {
  if (!xml) {
    throw new Error('Expected conditional format artifact XML')
  }
  return [
    ...[...xml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?conditionalFormatting\b[^>]*\bsqref=("|')([\s\S]*?)\1/gu)].map((match) => match[2] ?? ''),
    ...[...xml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?sqref\b[^>]*>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?sqref>/gu)].map((match) =>
      decodeXmlText(match[1] ?? ''),
    ),
  ]
}

function extractConditionalFormatFormulas(xml: string | undefined): string[] {
  if (!xml) {
    throw new Error('Expected conditional format artifact XML')
  }
  const standardFormulas = [...xml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?formula\b[^>]*>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?formula>/gu)]
  const x14Formulas = [...xml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?f\b[^>]*>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?f>/gu)]
  return standardFormulas.concat(x14Formulas).map((match) => decodeXmlText(match[1] ?? ''))
}

function decodeXmlText(value: string): string {
  return value
    .replace(/&quot;/gu, '"')
    .replace(/&apos;/gu, "'")
    .replace(/&lt;/gu, '<')
    .replace(/&gt;/gu, '>')
    .replace(/&amp;/gu, '&')
}

function buildAdvancedConditionalFormattingWorkbook(): Uint8Array {
  const zip = baseConditionalFormattingZip()
  zip['xl/worksheets/sheet1.xml'] = strToU8(advancedConditionalFormattingWorksheetXml)
  return zipSync(zip)
}

function buildPrefixedConditionalFormattingWorkbook(): Uint8Array {
  const zip = baseConditionalFormattingZip()
  zip['xl/worksheets/sheet1.xml'] = strToU8(prefixedAdvancedConditionalFormattingWorksheetXml)
  return zipSync(zip)
}

function baseConditionalFormattingZip(): Record<string, Uint8Array> {
  return unzipSync(
    writeSimpleXlsxWorkbook({
      sheets: [
        {
          name: 'Dashboard',
          cells: [
            { address: 'A1', row: 0, col: 0, value: 10 },
            { address: 'B1', row: 0, col: 1, value: 20 },
            { address: 'C1', row: 0, col: 2, value: 30 },
            { address: 'A2', row: 1, col: 0, value: 20 },
            { address: 'B2', row: 1, col: 1, value: 40 },
            { address: 'C2', row: 1, col: 2, value: 60 },
            { address: 'A3', row: 2, col: 0, value: 30 },
            { address: 'B3', row: 2, col: 1, value: 60 },
            { address: 'C3', row: 2, col: 2, value: 90 },
          ],
        },
      ],
    }),
  )
}

function buildFormulaConditionalFormattingWorkbook(): Uint8Array {
  return exportXlsx({
    version: 1,
    workbook: { name: 'Desktop Excel conditional format formula oracle' },
    sheets: [
      {
        id: 1,
        name: 'Dashboard',
        order: 0,
        cells: [
          { address: 'A1', value: 10 },
          { address: 'A2', value: 20 },
          { address: 'A3', value: 30 },
        ],
        metadata: {
          conditionalFormats: [
            {
              id: 'formula-highlight',
              range: { sheetName: 'Dashboard', startAddress: 'A1', endAddress: 'A3' },
              rule: { kind: 'formula', formula: '=A1>15' },
              style: { fill: { backgroundColor: '#ffeb84' } },
              priority: 1,
            },
          ],
        },
      },
    ],
  })
}

function buildCrossSheetFormulaConditionalFormattingWorkbook(): Uint8Array {
  return exportXlsx({
    version: 1,
    workbook: { name: 'Desktop Excel cross-sheet conditional format formula oracle' },
    sheets: [
      {
        id: 1,
        name: 'Dashboard',
        order: 0,
        cells: [
          { address: 'A1', value: 10 },
          { address: 'A2', value: 20 },
          { address: 'A3', value: 30 },
        ],
        metadata: {
          conditionalFormats: [
            {
              id: 'cross-sheet-formula-highlight',
              range: { sheetName: 'Dashboard', startAddress: 'A1', endAddress: 'A3' },
              rule: { kind: 'formula', formula: '=Inputs!A1>15' },
              style: { fill: { backgroundColor: '#ffeb84' } },
              priority: 1,
            },
          ],
        },
      },
      {
        id: 2,
        name: 'Inputs',
        order: 1,
        cells: [
          { address: 'A1', value: 10 },
          { address: 'A2', value: 20 },
          { address: 'A3', value: 30 },
        ],
      },
    ],
  })
}

function readWorksheetXml(path: string): string {
  return strFromU8(unzipSync(new Uint8Array(readFileSync(path)))['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
}

const advancedConditionalFormattingWorksheetXml = [
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
  '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
  '<dimension ref="A1:D3"/>',
  '<sheetData>',
  '<row r="1"><c r="A1"><v>10</v></c><c r="B1"><v>20</v></c><c r="C1"><v>30</v></c></row>',
  '<row r="2"><c r="A2"><v>20</v></c><c r="B2"><v>40</v></c><c r="C2"><v>60</v></c></row>',
  '<row r="3"><c r="A3"><v>30</v></c><c r="B3"><v>60</v></c><c r="C3"><v>90</v></c></row>',
  '</sheetData>',
  '<conditionalFormatting sqref="A1:A3">',
  '<cfRule type="dataBar" priority="1">',
  '<dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF63C384"/></dataBar>',
  '</cfRule>',
  '</conditionalFormatting>',
  '<conditionalFormatting sqref="B1:B3">',
  '<cfRule type="colorScale" priority="2">',
  '<colorScale><cfvo type="min"/><cfvo type="percentile" val="50"/><cfvo type="max"/>',
  '<color rgb="FFF8696B"/><color rgb="FFFFEB84"/><color rgb="FF63BE7B"/></colorScale>',
  '</cfRule>',
  '</conditionalFormatting>',
  '<conditionalFormatting sqref="C1:C3">',
  '<cfRule type="iconSet" priority="3">',
  '<iconSet iconSet="3TrafficLights1"><cfvo type="percent" val="0"/><cfvo type="percent" val="33"/>',
  '<cfvo type="percent" val="67"/></iconSet>',
  '</cfRule>',
  '</conditionalFormatting>',
  '</worksheet>',
].join('')

const prefixedAdvancedConditionalFormattingWorksheetXml = advancedConditionalFormattingWorksheetXml
  .replace(
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
  )
  .replaceAll('<conditionalFormatting ', '<x:conditionalFormatting ')
  .replaceAll('</conditionalFormatting>', '</x:conditionalFormatting>')
