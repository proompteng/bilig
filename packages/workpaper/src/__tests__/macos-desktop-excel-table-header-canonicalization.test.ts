import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { SpreadsheetEngine } from '@bilig/core'
import { exportXlsx, importXlsx } from '@bilig/excel-import'
import { isMacosExcelInstalled, runMacosExcelStructuralOperationOracle, type NormalizedFormulaValue } from '@bilig/excel-fixtures'
import { ValueTag, type CellValue } from '@bilig/protocol'
import { describe, expect, it } from 'vitest'

import { WorkPaper, type WorkPaperCellAddress } from '../index.js'

const workbookConfig = { maxRows: 16, maxColumns: 8, useColumnIndex: true }

const blankHeaderOracleCells = [
  { address: 'B1', formula: 'Column1', rawValue: 'string\tColumn1', value: { kind: 'string', value: 'Column1' } },
  { address: 'E1', formula: '=SUM(Sales[Column1])', rawValue: 'number\t30.0', value: { kind: 'number', value: 30 } },
  { address: 'F1', formula: '=SUM(Sales[Margin])', rawValue: 'number\t5.0', value: { kind: 'number', value: 5 } },
  { address: 'G1', formula: '=SUM(SalesAmount)', rawValue: 'number\t30.0', value: { kind: 'number', value: 30 } },
  { address: 'H1', formula: '=SUM(SalesMarginFormula)', rawValue: 'number\t5.0', value: { kind: 'number', value: 5 } },
] as const
const duplicateHeaderOracleCells = [
  { address: 'C1', formula: 'Amount2', rawValue: 'string\tAmount2', value: { kind: 'string', value: 'Amount2' } },
  { address: 'E1', formula: '=SUM(Sales[Amount])', rawValue: 'number\t30.0', value: { kind: 'number', value: 30 } },
  { address: 'F1', formula: '=SUM(Sales[Amount2])', rawValue: 'number\t5.0', value: { kind: 'number', value: 5 } },
  { address: 'G1', formula: '=SUM(SalesAmount)', rawValue: 'number\t30.0', value: { kind: 'number', value: 30 } },
  { address: 'H1', formula: '=SUM(SalesMarginFormula)', rawValue: 'number\t5.0', value: { kind: 'number', value: 5 } },
] as const

describe('macOS Desktop Excel table-header canonicalization oracle', () => {
  it('canonicalizes blank table headers before XLSX export', async () => {
    const engine = await buildTableHeaderCanonicalizationEngine('table-header-blank-canonicalization-oracle')

    engine.setCellValue('Data', 'B1', '')

    expect(engine.getTable('Sales')).toMatchObject({
      startAddress: 'A1',
      endAddress: 'C3',
      columnNames: ['Region', 'Column1', 'Margin'],
    })
    expect(engine.getCellValue('Data', 'B1')).toEqual({ tag: ValueTag.String, value: 'Column1', stringId: expect.any(Number) })
    expect(engine.getDefinedName('SalesAmount')).toEqual({
      name: 'SalesAmount',
      value: { kind: 'structured-ref', tableName: 'Sales', columnName: 'Column1' },
    })
    expect(engine.getCellValue('Data', 'E1')).toEqual({ tag: ValueTag.Number, value: 30 })
    expect(engine.getCellValue('Data', 'G1')).toEqual({ tag: ValueTag.Number, value: 30 })

    const imported = importXlsx(exportXlsx(engine.exportSnapshot()), 'headless-table-header-blank-canonicalization-oracle.xlsx')
    expect(imported.snapshot.workbook.metadata?.tables?.[0]).toMatchObject({
      columnNames: ['Region', 'Column1', 'Margin'],
    })
    expect(imported.snapshot.workbook.metadata?.definedNames).toEqual([
      { name: 'SalesAmount', value: { kind: 'formula', formula: '=Sales[Column1]' } },
      { name: 'SalesMarginFormula', value: { kind: 'formula', formula: '=Sales[Margin]' } },
    ])
    expectImportedValues(imported.snapshot, blankHeaderOracleCells)
  })

  it('deduplicates renamed table headers before XLSX export', async () => {
    const engine = await buildTableHeaderCanonicalizationEngine('table-header-duplicate-canonicalization-oracle')

    engine.setCellValue('Data', 'C1', 'Amount')

    expect(engine.getTable('Sales')).toMatchObject({
      startAddress: 'A1',
      endAddress: 'C3',
      columnNames: ['Region', 'Amount', 'Amount2'],
    })
    expect(engine.getCellValue('Data', 'C1')).toEqual({ tag: ValueTag.String, value: 'Amount2', stringId: expect.any(Number) })
    expect(engine.getDefinedName('SalesMarginFormula')).toEqual({
      name: 'SalesMarginFormula',
      value: { kind: 'formula', formula: '=Sales[Amount2]' },
    })
    expect(engine.getCellValue('Data', 'F1')).toEqual({ tag: ValueTag.Number, value: 5 })
    expect(engine.getCellValue('Data', 'H1')).toEqual({ tag: ValueTag.Number, value: 5 })

    const imported = importXlsx(exportXlsx(engine.exportSnapshot()), 'headless-table-header-duplicate-canonicalization-oracle.xlsx')
    expect(imported.snapshot.workbook.metadata?.tables?.[0]).toMatchObject({
      columnNames: ['Region', 'Amount', 'Amount2'],
    })
    expect(imported.snapshot.workbook.metadata?.definedNames).toEqual([
      { name: 'SalesAmount', value: { kind: 'formula', formula: '=Sales[Amount]' } },
      { name: 'SalesMarginFormula', value: { kind: 'formula', formula: '=Sales[Amount2]' } },
    ])
    expectImportedValues(imported.snapshot, duplicateHeaderOracleCells)
  })

  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'matches Desktop Excel blank and duplicate table-header canonicalization',
    async () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      await expectExcelScenario({
        workbookName: 'blank',
        operation: { kind: 'setCellValue', address: 'B1', value: '' },
        expectedCells: blankHeaderOracleCells,
        expectedColumns: ['Region', 'Column1', 'Margin'],
        expectedDefinedNames: [
          { name: 'SalesAmount', value: { kind: 'formula' as const, formula: '=Sales[Column1]' } },
          { name: 'SalesMarginFormula', value: { kind: 'formula' as const, formula: '=Sales[Margin]' } },
        ],
      })
      await expectExcelScenario({
        workbookName: 'duplicate',
        operation: { kind: 'setCellValue', address: 'C1', value: 'Amount' },
        expectedCells: duplicateHeaderOracleCells,
        expectedColumns: ['Region', 'Amount', 'Amount2'],
        expectedDefinedNames: [
          { name: 'SalesAmount', value: { kind: 'formula' as const, formula: '=Sales[Amount]' } },
          { name: 'SalesMarginFormula', value: { kind: 'formula' as const, formula: '=Sales[Amount2]' } },
        ],
      })
    },
    90_000,
  )
})

async function expectExcelScenario(args: {
  readonly workbookName: string
  readonly operation: { readonly kind: 'setCellValue'; readonly address: string; readonly value: string }
  readonly expectedCells: readonly {
    readonly address: string
    readonly formula: string
    readonly rawValue: string
    readonly value: NormalizedFormulaValue
  }[]
  readonly expectedColumns: readonly string[]
  readonly expectedDefinedNames: readonly {
    readonly name: string
    readonly value: { readonly kind: 'formula'; readonly formula: string }
  }[]
}): Promise<void> {
  const tempDir = mkdtempSync(join(tmpdir(), `bilig-headless-excel-table-header-${args.workbookName}-canonicalization-`))
  try {
    const workbookPath = join(tempDir, `headless-table-header-${args.workbookName}-canonicalization.xlsx`)
    const engine = await buildTableHeaderCanonicalizationEngine(`table-header-${args.workbookName}-canonicalization-oracle`)
    writeFileSync(workbookPath, exportXlsx(engine.exportSnapshot()))

    const excelResult = runMacosExcelStructuralOperationOracle({
      workbookPath,
      worksheetName: 'Data',
      operations: [args.operation],
      inspectCells: args.expectedCells.map((cell) => cell.address),
      saveWorkbook: true,
    })
    expect(excelResult.cells).toEqual(args.expectedCells)

    const imported = importXlsx(new Uint8Array(readFileSync(workbookPath)), `headless-table-header-${args.workbookName}-canonicalized.xlsx`)
    expect(imported.snapshot.workbook.metadata?.tables?.[0]).toMatchObject({
      columnNames: args.expectedColumns,
    })
    expect(imported.snapshot.workbook.metadata?.definedNames).toEqual(args.expectedDefinedNames)
    expectImportedValues(imported.snapshot, args.expectedCells)
  } finally {
    rmSync(tempDir, { recursive: true, force: true })
  }
}

async function buildTableHeaderCanonicalizationEngine(workbookName: string): Promise<SpreadsheetEngine> {
  const engine = new SpreadsheetEngine({ workbookName })
  await engine.ready()
  engine.createSheet('Data')
  engine.setRangeValues({ sheetName: 'Data', startAddress: 'A1', endAddress: 'C3' }, [
    ['Region', 'Amount', 'Margin'],
    ['East', 10, 2],
    ['West', 20, 3],
  ])
  engine.setTable({
    name: 'Sales',
    sheetName: 'Data',
    startAddress: 'A1',
    endAddress: 'C3',
    columnNames: ['Region', 'Amount', 'Margin'],
    headerRow: true,
    totalsRow: false,
  })
  engine.setDefinedName('SalesAmount', { kind: 'structured-ref', tableName: 'Sales', columnName: 'Amount' })
  engine.setDefinedName('SalesMarginFormula', { kind: 'formula', formula: '=Sales[Margin]' })
  engine.setCellFormula('Data', 'E1', 'SUM(Sales[Amount])')
  engine.setCellFormula('Data', 'F1', 'SUM(Sales[Margin])')
  engine.setCellFormula('Data', 'G1', 'SUM(SalesAmount)')
  engine.setCellFormula('Data', 'H1', 'SUM(SalesMarginFormula)')
  return engine
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
