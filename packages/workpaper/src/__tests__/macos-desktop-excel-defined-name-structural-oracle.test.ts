import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { SpreadsheetEngine } from '@bilig/core'
import { exportXlsx, importXlsx } from '@bilig/excel-import'
import { isMacosExcelInstalled, runMacosExcelInspectionOracle, runMacosExcelStructuralOperationOracle } from '@bilig/excel-fixtures'
import { ErrorCode, ValueTag } from '@bilig/protocol'
import { describe, expect, it } from 'vitest'

import { WorkPaper, type WorkPaperCellAddress } from '../index.js'

const workbookConfig = { maxRows: 16, maxColumns: 8, useColumnIndex: true }

describe('macOS Desktop Excel defined-name structural oracle', () => {
  it('preserves structurally deleted defined references as #REF! through export/import', async () => {
    const cases = await buildStructuralDefinedNameReferenceCases()

    for (const { engine, testCase } of cases) {
      testCase.applyHeadlessDelete(engine)

      expect(engine.getDefinedName(testCase.name)).toEqual({
        name: testCase.name,
        value: { kind: 'formula', formula: '=Data!#REF!' },
      })
      expect(engine.getCell(testCase.sheetName, testCase.expectedAddress).formula).toBe(testCase.expectedFormula)
      expect(engine.getCellValue(testCase.sheetName, testCase.expectedAddress)).toEqual({
        tag: ValueTag.Error,
        code: ErrorCode.Ref,
      })

      const imported = importXlsx(exportXlsx(engine.exportSnapshot()), `headless-${testCase.fileName}`)
      expect(imported.snapshot.workbook.metadata?.definedNames).toEqual([
        { name: testCase.name, value: { kind: 'formula', formula: '=Data!#REF!' } },
      ])

      const restored = WorkPaper.buildFromSnapshot(imported.snapshot, workbookConfig)
      try {
        expect(restored.getCellValue(addressToCell(testCase.expectedAddress))).toEqual({ tag: ValueTag.Error, code: ErrorCode.Ref })
      } finally {
        restored.dispose()
      }
    }
  })

  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'matches Desktop Excel when structural deletes invalidate defined references',
    async () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = mkdtempSync(join(tmpdir(), 'bilig-headless-excel-defined-name-ref-delete-oracle-'))
      try {
        const cases = await buildStructuralDefinedNameReferenceCases()
        for (const { engine, testCase } of cases) {
          const workbookPath = join(tempDir, testCase.fileName)
          writeFileSync(workbookPath, exportXlsx(engine.exportSnapshot()))

          const excelResult = runMacosExcelStructuralOperationOracle({
            workbookPath,
            worksheetName: testCase.sheetName,
            operations: [testCase.excelOperation],
            inspectCells: [testCase.expectedAddress],
            saveWorkbook: true,
          })

          expect(excelResult.cells).toEqual([
            {
              address: testCase.expectedAddress,
              formula: testCase.expectedExcelFormula,
              rawValue: 'error\t#REF!',
              value: { kind: 'error', value: String(ErrorCode.Ref) },
            },
          ])

          const imported = importXlsx(new Uint8Array(readFileSync(workbookPath)), `recalculated-${testCase.fileName}`)
          expect(imported.snapshot.workbook.metadata?.definedNames).toEqual([
            { name: testCase.name, value: { kind: 'formula', formula: '=Data!#REF!' } },
          ])

          const restored = WorkPaper.buildFromSnapshot(imported.snapshot, workbookConfig)
          try {
            expect(restored.getCellValue(addressToCell(testCase.expectedAddress))).toEqual({ tag: ValueTag.Error, code: ErrorCode.Ref })
          } finally {
            restored.dispose()
          }
        }
      } finally {
        rmSync(tempDir, { recursive: true, force: true })
      }
    },
    60_000,
  )

  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'matches Desktop Excel when deleting a sheet invalidates workbook-level defined names',
    async () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = mkdtempSync(join(tmpdir(), 'bilig-headless-excel-defined-name-sheet-delete-oracle-'))
      try {
        const engine = await buildDeletedSheetDefinedNameEngine()
        const sourceSnapshot = engine.exportSnapshot()
        const excelPath = join(tempDir, 'excel-defined-name-sheet-delete-oracle.xlsx')
        writeFileSync(excelPath, exportXlsx(sourceSnapshot))

        const excelResult = runMacosExcelStructuralOperationOracle({
          workbookPath: excelPath,
          worksheetName: 'Report',
          operations: [{ kind: 'deleteSheet', name: 'Data' }],
          inspectCells: ['A1', 'A2', 'A3'],
          saveWorkbook: true,
          timeoutMs: 90_000,
        })
        expect(excelResult.cells).toEqual([
          {
            address: 'A1',
            formula: '=RateCell*2',
            rawValue: 'error\t#REF!',
            value: { kind: 'error', value: String(ErrorCode.Ref) },
          },
          {
            address: 'A2',
            formula: '=SUM(SalesRange)',
            rawValue: 'error\t#REF!',
            value: { kind: 'error', value: String(ErrorCode.Ref) },
          },
          {
            address: 'A3',
            formula: '=FormulaRate+FormulaSum',
            rawValue: 'error\t#REF!',
            value: { kind: 'error', value: String(ErrorCode.Ref) },
          },
        ])

        const excelTruth = importXlsx(new Uint8Array(readFileSync(excelPath)), 'excel-defined-name-sheet-delete-truth.xlsx').snapshot
        expect(excelTruth.sheets.map((sheet) => sheet.name)).toEqual(['Report'])
        expect(excelTruth.workbook.metadata?.definedNames).toEqual(expectedSheetDeleteDefinedNames())

        const workpaper = WorkPaper.buildFromSnapshot(sourceSnapshot, workbookConfig)
        try {
          const dataSheet = workpaper.getSheetId('Data')
          if (dataSheet === undefined) {
            throw new Error('Expected Data sheet')
          }
          workpaper.removeSheet(dataSheet)

          const headlessPath = join(tempDir, 'headless-defined-name-sheet-delete-oracle.xlsx')
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

          const headlessTruth = importXlsx(
            new Uint8Array(readFileSync(headlessPath)),
            'headless-defined-name-sheet-delete-truth.xlsx',
          ).snapshot
          expect(headlessTruth.sheets.map((sheet) => sheet.name)).toEqual(['Report'])
          expect(headlessTruth.workbook.metadata?.definedNames).toEqual(excelTruth.workbook.metadata?.definedNames)
        } finally {
          workpaper.dispose()
        }
      } finally {
        rmSync(tempDir, { recursive: true, force: true })
      }
    },
    120_000,
  )
})

async function buildStructuralDefinedNameReferenceCases(): Promise<
  Array<{ readonly testCase: StructuralDefinedNameReferenceCase; readonly engine: SpreadsheetEngine }>
> {
  return Promise.all(
    structuralDefinedNameReferenceCases().map(async (testCase) => ({
      testCase,
      engine: await testCase.buildEngine(),
    })),
  )
}

interface StructuralDefinedNameReferenceCase {
  readonly fileName: string
  readonly sheetName: string
  readonly name: string
  readonly expectedAddress: string
  readonly expectedFormula: string
  readonly expectedExcelFormula: string
  readonly excelOperation: { readonly kind: 'deleteRows'; readonly range: string }
  readonly buildEngine: () => Promise<SpreadsheetEngine>
  readonly applyHeadlessDelete: (engine: SpreadsheetEngine) => void
}

function structuralDefinedNameReferenceCases(): StructuralDefinedNameReferenceCase[] {
  return [
    {
      fileName: 'defined-name-range-ref-delete-oracle.xlsx',
      sheetName: 'Data',
      name: 'SalesRange',
      expectedAddress: 'D2',
      expectedFormula: 'SUM(SalesRange)',
      expectedExcelFormula: '=SUM(SalesRange)',
      excelOperation: { kind: 'deleteRows', range: '1:3' },
      buildEngine: buildDeletedRangeDefinedNameEngine,
      applyHeadlessDelete: (engine) => engine.deleteRows('Data', 0, 3),
    },
    {
      fileName: 'defined-name-cell-ref-delete-oracle.xlsx',
      sheetName: 'Data',
      name: 'RateCell',
      expectedAddress: 'D2',
      expectedFormula: 'RateCell*2',
      expectedExcelFormula: '=RateCell*2',
      excelOperation: { kind: 'deleteRows', range: '1:1' },
      buildEngine: buildDeletedCellDefinedNameEngine,
      applyHeadlessDelete: (engine) => engine.deleteRows('Data', 0, 1),
    },
  ]
}

async function buildDeletedRangeDefinedNameEngine(): Promise<SpreadsheetEngine> {
  const engine = new SpreadsheetEngine({ workbookName: 'defined-name-range-ref-delete-oracle' })
  await engine.ready()
  engine.createSheet('Data')
  engine.setRangeValues({ sheetName: 'Data', startAddress: 'A1', endAddress: 'A3' }, [[10], [20], [30]])
  engine.setDefinedName('SalesRange', { kind: 'range-ref', sheetName: 'Data', startAddress: 'A1', endAddress: 'A3' })
  engine.setCellFormula('Data', 'D5', 'SUM(SalesRange)')
  return engine
}

function cell(row: number, col: number): WorkPaperCellAddress {
  return { sheet: 1, row, col }
}

function addressToCell(address: string): WorkPaperCellAddress {
  const match = /^([A-Z]+)([1-9][0-9]*)$/u.exec(address)
  if (!match) {
    throw new Error(`Unexpected oracle address: ${address}`)
  }
  let col = 0
  for (const char of match[1]) {
    col = col * 26 + char.charCodeAt(0) - 64
  }
  return cell(Number(match[2]) - 1, col - 1)
}

async function buildDeletedCellDefinedNameEngine(): Promise<SpreadsheetEngine> {
  const engine = new SpreadsheetEngine({ workbookName: 'defined-name-cell-ref-delete-oracle' })
  await engine.ready()
  engine.createSheet('Data')
  engine.setCellValue('Data', 'A1', 5)
  engine.setDefinedName('RateCell', { kind: 'cell-ref', sheetName: 'Data', address: 'A1' })
  engine.setCellFormula('Data', 'D3', 'RateCell*2')
  return engine
}

async function buildDeletedSheetDefinedNameEngine(): Promise<SpreadsheetEngine> {
  const engine = new SpreadsheetEngine({ workbookName: 'defined-name-sheet-delete-oracle' })
  await engine.ready()
  engine.createSheet('Data')
  engine.createSheet('Report')
  engine.setRangeValues({ sheetName: 'Data', startAddress: 'A1', endAddress: 'A3' }, [[5], [10], [20]])
  engine.setDefinedName('RateCell', { kind: 'cell-ref', sheetName: 'Data', address: 'A1' })
  engine.setDefinedName('SalesRange', { kind: 'range-ref', sheetName: 'Data', startAddress: 'A1', endAddress: 'A3' })
  engine.setDefinedName('FormulaRate', { kind: 'formula', formula: '=Data!$A$1' })
  engine.setDefinedName('FormulaSum', { kind: 'formula', formula: '=SUM(Data!$A$1:$A$3)' })
  engine.workbook.setDefinedName('_xlnm.Print_Area', { kind: 'range-ref', sheetName: 'Data', startAddress: 'A1', endAddress: 'A3' }, 'Data')
  engine.workbook.setDefinedName(
    '_xlnm.Print_Area',
    { kind: 'range-ref', sheetName: 'Report', startAddress: 'A1', endAddress: 'A3' },
    'Report',
  )
  engine.setCellFormula('Report', 'A1', 'RateCell*2')
  engine.setCellFormula('Report', 'A2', 'SUM(SalesRange)')
  engine.setCellFormula('Report', 'A3', 'FormulaRate+FormulaSum')
  return engine
}

function expectedSheetDeleteDefinedNames() {
  return [
    {
      name: '_xlnm.Print_Area',
      scopeSheetName: 'Report',
      value: { kind: 'formula' as const, formula: '=Report!$A$1:$A$3' },
    },
    { name: 'FormulaRate', value: { kind: 'formula' as const, formula: '=#REF!' } },
    { name: 'FormulaSum', value: { kind: 'formula' as const, formula: '=SUM(#REF!)' } },
    { name: 'RateCell', value: { kind: 'formula' as const, formula: '=#REF!' } },
    { name: 'SalesRange', value: { kind: 'formula' as const, formula: '=#REF!' } },
  ]
}
