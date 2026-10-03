import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { SpreadsheetEngine } from '@bilig/core'
import { exportXlsx, importXlsx } from '@bilig/excel-import'
import {
  isMacosExcelInstalled,
  runMacosExcelInspectionOracle,
  runMacosExcelStructuralOperationOracle,
  type NormalizedFormulaValue,
} from '@bilig/excel-fixtures'
import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import { describe, expect, it } from 'vitest'

import { WorkPaper, type WorkPaperCellAddress } from '../index.js'

const workbookConfig = { maxRows: 16, maxColumns: 8, useColumnIndex: true }
const spacedHeaderOracleCells = [
  { address: 'D1', formula: '=SUM(Sales[Q1 Sales])', rawValue: 'number\t30.0', value: { kind: 'number', value: 30 } },
  { address: 'E1', formula: '=SUM(Sales[Units Sold])', rawValue: 'number\t5.0', value: { kind: 'number', value: 5 } },
  { address: 'F1', formula: '=SUM(SalesQ1)', rawValue: 'number\t30.0', value: { kind: 'number', value: 30 } },
  { address: 'G1', formula: '=SUM(SalesUnitsFormula)', rawValue: 'number\t5.0', value: { kind: 'number', value: 5 } },
] as const
const spacedHeaderRenameOracleCells = [
  { address: 'A1', formula: 'Q1 Revenue', rawValue: 'string\tQ1 Revenue', value: { kind: 'string', value: 'Q1 Revenue' } },
  { address: 'D1', formula: '=SUM(Sales[Q1 Revenue])', rawValue: 'number\t30.0', value: { kind: 'number', value: 30 } },
  { address: 'E1', formula: '=SUM(Sales[Units Sold])', rawValue: 'number\t5.0', value: { kind: 'number', value: 5 } },
  { address: 'F1', formula: '=SUM(SalesQ1)', rawValue: 'number\t30.0', value: { kind: 'number', value: 30 } },
  { address: 'G1', formula: '=SUM(SalesUnitsFormula)', rawValue: 'number\t5.0', value: { kind: 'number', value: 5 } },
] as const
const spacedHeaderDeleteOracleCells = [
  { address: 'A1', formula: 'Units Sold', rawValue: 'string\tUnits Sold', value: { kind: 'string', value: 'Units Sold' } },
  { address: 'C1', formula: '=SUM(#REF!)', rawValue: 'error\t#REF!', value: { kind: 'error', value: String(ErrorCode.Ref) } },
  { address: 'D1', formula: '=SUM(Sales[Units Sold])', rawValue: 'number\t5.0', value: { kind: 'number', value: 5 } },
  { address: 'E1', formula: '=SUM(SalesQ1)', rawValue: 'error\t#REF!', value: { kind: 'error', value: String(ErrorCode.Ref) } },
  { address: 'F1', formula: '=SUM(SalesUnitsFormula)', rawValue: 'number\t5.0', value: { kind: 'number', value: 5 } },
] as const
const spacedHeaderDeleteImportedCells = [
  { address: 'A1', value: { kind: 'string', value: 'Units Sold' } },
  { address: 'C1', value: { kind: 'error', value: String(ErrorCode.Ref) } },
  { address: 'D1', value: { kind: 'number', value: 5 } },
  { address: 'E1', value: { kind: 'error', value: String(ErrorCode.Ref) } },
  { address: 'F1', value: { kind: 'number', value: 5 } },
] as const
const specialHeaderOracleCells = [
  { address: 'G1', formula: '=SUM(Sales[Revenue, Net])', rawValue: 'number\t30.0', value: { kind: 'number', value: 30 } },
  { address: 'G2', formula: '=SUM(Sales[A:B])', rawValue: 'number\t5.0', value: { kind: 'number', value: 5 } },
  { address: 'G3', formula: "=SUM(Sales['# Units])", rawValue: 'number\t10.0', value: { kind: 'number', value: 10 } },
  { address: 'G4', formula: "=SUM(Sales[Owner''s Share])", rawValue: 'number\t1.0', value: { kind: 'number', value: 1 } },
  { address: 'G5', formula: "=SUM(Sales[A'[B']])", rawValue: 'number\t300.0', value: { kind: 'number', value: 300 } },
] as const

describe('macOS Desktop Excel structured-reference syntax oracle', () => {
  it('authors structured references for table headers with spaces before XLSX export', async () => {
    const engine = await buildSpacedHeaderStructuredReferenceEngine()

    expect(engine.getCell('Data', 'D1').formula).toBe('SUM(Sales[Q1 Sales])')
    expect(engine.getCellValue('Data', 'D1')).toEqual({ tag: ValueTag.Number, value: 30 })
    expect(engine.getCell('Data', 'E1').formula).toBe('SUM(Sales[Units Sold])')
    expect(engine.getCellValue('Data', 'E1')).toEqual({ tag: ValueTag.Number, value: 5 })
    expect(engine.getCell('Data', 'F1').formula).toBe('SUM(SalesQ1)')
    expect(engine.getCellValue('Data', 'F1')).toEqual({ tag: ValueTag.Number, value: 30 })
    expect(engine.getCell('Data', 'G1').formula).toBe('SUM(SalesUnitsFormula)')
    expect(engine.getCellValue('Data', 'G1')).toEqual({ tag: ValueTag.Number, value: 5 })

    const imported = importXlsx(exportXlsx(engine.exportSnapshot()), 'headless-structured-reference-spaced-headers.xlsx')
    expect(imported.snapshot.workbook.metadata?.definedNames).toEqual([
      { name: 'SalesQ1', value: { kind: 'formula', formula: '=Sales[Q1 Sales]' } },
      { name: 'SalesUnitsFormula', value: { kind: 'formula', formula: '=Sales[Units Sold]' } },
    ])
    expectImportedValues(imported.snapshot, spacedHeaderOracleCells)
  })

  it('authors escaped structured references for special table headers before XLSX export', async () => {
    const engine = await buildSpecialHeaderStructuredReferenceEngine()

    expect(engine.getCellValue('Data', 'G1')).toEqual({ tag: ValueTag.Number, value: 30 })
    expect(engine.getCellValue('Data', 'G2')).toEqual({ tag: ValueTag.Number, value: 5 })
    expect(engine.getCellValue('Data', 'G3')).toEqual({ tag: ValueTag.Number, value: 10 })
    expect(engine.getCellValue('Data', 'G4')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Data', 'G5')).toEqual({ tag: ValueTag.Number, value: 300 })

    const imported = importXlsx(exportXlsx(engine.exportSnapshot()), 'headless-structured-reference-special-headers.xlsx')
    expect(imported.snapshot.workbook.metadata?.definedNames).toEqual([
      { name: 'SalesBracketFormula', value: { kind: 'formula', formula: "=Sales[A'[B']]" } },
      { name: 'SalesUnits', value: { kind: 'formula', formula: "=Sales['# Units]" } },
    ])
    expectImportedValues(imported.snapshot, specialHeaderOracleCells)
  })

  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'matches Desktop Excel structured references for table headers with spaces',
    async () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = mkdtempSync(join(tmpdir(), 'bilig-headless-excel-structured-reference-spaced-headers-'))
      try {
        const workbookPath = join(tempDir, 'headless-structured-reference-spaced-headers.xlsx')
        const engine = await buildSpacedHeaderStructuredReferenceEngine()
        writeFileSync(workbookPath, exportXlsx(engine.exportSnapshot()))

        const excelResult = runMacosExcelInspectionOracle({
          workbookPath,
          worksheetName: 'Data',
          formulaCells: [],
          inspectCells: spacedHeaderOracleCells.map((cell) => cell.address),
          saveWorkbook: true,
        })
        expect(excelResult.cells).toEqual(spacedHeaderOracleCells)

        const imported = importXlsx(new Uint8Array(readFileSync(workbookPath)), 'headless-structured-reference-spaced-headers-saved.xlsx')
        expectImportedValues(imported.snapshot, spacedHeaderOracleCells)
      } finally {
        rmSync(tempDir, { recursive: true, force: true })
      }
    },
    60_000,
  )

  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'matches Desktop Excel structured references for special table headers',
    async () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = mkdtempSync(join(tmpdir(), 'bilig-headless-excel-structured-reference-special-headers-'))
      try {
        const workbookPath = join(tempDir, 'headless-structured-reference-special-headers.xlsx')
        const engine = await buildSpecialHeaderStructuredReferenceEngine()
        writeFileSync(workbookPath, exportXlsx(engine.exportSnapshot()))

        const excelResult = runMacosExcelInspectionOracle({
          workbookPath,
          worksheetName: 'Data',
          formulaCells: [],
          inspectCells: specialHeaderOracleCells.map((cell) => cell.address),
          saveWorkbook: true,
        })
        expect(excelResult.cells).toEqual(specialHeaderOracleCells)

        const imported = importXlsx(new Uint8Array(readFileSync(workbookPath)), 'headless-structured-reference-special-headers-saved.xlsx')
        expectImportedValues(imported.snapshot, specialHeaderOracleCells)
      } finally {
        rmSync(tempDir, { recursive: true, force: true })
      }
    },
    60_000,
  )

  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'matches Desktop Excel structural rewrite semantics for spaced table headers',
    async () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      await expectExcelStructuralScenario({
        workbookName: 'rename',
        operation: { kind: 'setCellValue', address: 'A1', value: 'Q1 Revenue' },
        expectedCells: spacedHeaderRenameOracleCells,
        expectedColumns: ['Q1 Revenue', 'Units Sold'],
        expectedDefinedNames: [
          { name: 'SalesQ1', value: { kind: 'formula' as const, formula: '=Sales[Q1 Revenue]' } },
          { name: 'SalesUnitsFormula', value: { kind: 'formula' as const, formula: '=Sales[Units Sold]' } },
        ],
      })
      await expectExcelStructuralScenario({
        workbookName: 'delete',
        operation: { kind: 'deleteColumns', range: 'A:A' },
        expectedCells: spacedHeaderDeleteOracleCells,
        expectedImportedCells: spacedHeaderDeleteImportedCells,
        expectedColumns: ['Units Sold'],
        expectedDefinedNames: [
          { name: 'SalesQ1', value: { kind: 'formula' as const, formula: '=#REF!' } },
          { name: 'SalesUnitsFormula', value: { kind: 'formula' as const, formula: '=Sales[Units Sold]' } },
        ],
      })
    },
    90_000,
  )
})

async function buildSpacedHeaderStructuredReferenceEngine(): Promise<SpreadsheetEngine> {
  const engine = new SpreadsheetEngine({ workbookName: 'structured-reference-spaced-headers' })
  await engine.ready()
  engine.createSheet('Data')
  engine.setRangeValues({ sheetName: 'Data', startAddress: 'A1', endAddress: 'B3' }, [
    ['Q1 Sales', 'Units Sold'],
    [10, 2],
    [20, 3],
  ])
  engine.setTable({
    name: 'Sales',
    sheetName: 'Data',
    startAddress: 'A1',
    endAddress: 'B3',
    columnNames: ['Q1 Sales', 'Units Sold'],
    headerRow: true,
    totalsRow: false,
  })
  engine.setDefinedName('SalesQ1', { kind: 'structured-ref', tableName: 'Sales', columnName: 'Q1 Sales' })
  engine.setDefinedName('SalesUnitsFormula', { kind: 'formula', formula: '=Sales[Units Sold]' })
  engine.setCellFormula('Data', 'D1', 'SUM(Sales[Q1 Sales])')
  engine.setCellFormula('Data', 'E1', 'SUM(Sales[Units Sold])')
  engine.setCellFormula('Data', 'F1', 'SUM(SalesQ1)')
  engine.setCellFormula('Data', 'G1', 'SUM(SalesUnitsFormula)')
  return engine
}

async function buildSpecialHeaderStructuredReferenceEngine(): Promise<SpreadsheetEngine> {
  const engine = new SpreadsheetEngine({ workbookName: 'structured-reference-special-headers' })
  await engine.ready()
  engine.createSheet('Data')
  engine.setRangeValues({ sheetName: 'Data', startAddress: 'A1', endAddress: 'E3' }, [
    ['Revenue, Net', 'A:B', '# Units', "Owner's Share", 'A[B]'],
    [10, 2, 4, 0.25, 100],
    [20, 3, 6, 0.75, 200],
  ])
  engine.setTable({
    name: 'Sales',
    sheetName: 'Data',
    startAddress: 'A1',
    endAddress: 'E3',
    columnNames: ['Revenue, Net', 'A:B', '# Units', "Owner's Share", 'A[B]'],
    headerRow: true,
    totalsRow: false,
  })
  engine.setDefinedName('SalesUnits', { kind: 'structured-ref', tableName: 'Sales', columnName: '# Units' })
  engine.setDefinedName('SalesBracketFormula', { kind: 'formula', formula: "=Sales[A'[B']]" })
  engine.setCellFormula('Data', 'G1', 'SUM(Sales[Revenue, Net])')
  engine.setCellFormula('Data', 'G2', 'SUM(Sales[A:B])')
  engine.setCellFormula('Data', 'G3', "SUM(Sales['# Units])")
  engine.setCellFormula('Data', 'G4', "SUM(Sales[Owner''s Share])")
  engine.setCellFormula('Data', 'G5', "SUM(Sales[A'[B']])")
  return engine
}

async function expectExcelStructuralScenario(args: {
  readonly workbookName: string
  readonly operation:
    | { readonly kind: 'setCellValue'; readonly address: string; readonly value: string }
    | { readonly kind: 'deleteColumns'; readonly range: string }
  readonly expectedCells: readonly {
    readonly address: string
    readonly formula: string
    readonly rawValue: string
    readonly value: NormalizedFormulaValue
  }[]
  readonly expectedImportedCells?: readonly { readonly address: string; readonly value: NormalizedFormulaValue }[]
  readonly expectedColumns: readonly string[]
  readonly expectedDefinedNames: readonly {
    readonly name: string
    readonly value: { readonly kind: 'formula'; readonly formula: string }
  }[]
}): Promise<void> {
  const tempDir = mkdtempSync(join(tmpdir(), `bilig-headless-excel-structured-reference-spaced-headers-${args.workbookName}-`))
  try {
    const workbookPath = join(tempDir, `headless-structured-reference-spaced-headers-${args.workbookName}.xlsx`)
    const engine = await buildSpacedHeaderStructuredReferenceEngine()
    writeFileSync(workbookPath, exportXlsx(engine.exportSnapshot()))

    const excelResult = runMacosExcelStructuralOperationOracle({
      workbookPath,
      worksheetName: 'Data',
      operations: [args.operation],
      inspectCells: args.expectedCells.map((cell) => cell.address),
      saveWorkbook: true,
    })
    expect(excelResult.cells).toEqual(args.expectedCells)

    const imported = importXlsx(
      new Uint8Array(readFileSync(workbookPath)),
      `headless-structured-reference-spaced-headers-${args.workbookName}-saved.xlsx`,
    )
    expect(imported.snapshot.workbook.metadata?.tables?.[0]).toMatchObject({
      name: 'Sales',
      sheetName: 'Data',
      startAddress: 'A1',
      endAddress: `${String.fromCharCode(64 + args.expectedColumns.length)}3`,
      columnNames: args.expectedColumns,
    })
    expect(imported.snapshot.workbook.metadata?.definedNames).toEqual(args.expectedDefinedNames)
    expectImportedValues(imported.snapshot, args.expectedImportedCells ?? args.expectedCells)
  } finally {
    rmSync(tempDir, { recursive: true, force: true })
  }
}

function expectImportedValues(
  snapshot: ReturnType<typeof importXlsx>['snapshot'],
  expectedCells: readonly { readonly address: string; readonly value: NormalizedFormulaValue }[],
): void {
  const workbook = WorkPaper.buildFromSnapshot(snapshot, workbookConfig)
  try {
    expect(expectedCells.map((cell) => normalizedCellValue(workbook.getCellValue(addressToCell(cell.address))))).toEqual(
      expectedCells.map((cell) => cell.value),
    )
  } finally {
    workbook.dispose()
  }
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
  return { sheet: 1, row: Number(match[2]) - 1, col: col - 1 }
}

function normalizedCellValue(value: CellValue): NormalizedFormulaValue {
  switch (value.tag) {
    case ValueTag.Empty:
      return { kind: 'blank' }
    case ValueTag.Boolean:
      return { kind: 'boolean', value: value.value }
    case ValueTag.Error:
      return { kind: 'error', value: String(value.code) }
    case ValueTag.Number:
      return { kind: 'number', value: value.value }
    case ValueTag.String:
      return { kind: 'string', value: value.value }
  }
}
