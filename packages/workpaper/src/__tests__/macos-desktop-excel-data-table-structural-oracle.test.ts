import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'

import { exportXlsx, importXlsx } from '@bilig/excel-import'
import {
  isMacosExcelInstalled,
  runMacosExcelInspectionOracle,
  runMacosExcelStructuralOperationOracle,
  type MacosExcelStructuralOperation,
} from '@bilig/excel-fixtures'
import { ValueTag, type CellValue, type WorkbookSheetDataTableFormulasSnapshot } from '@bilig/protocol'
import { describe, expect, it } from 'vitest'

import { WorkPaper, type WorkPaperCellAddress } from '../index.js'

const workbookConfig = { maxRows: 16, maxColumns: 8, useColumnIndex: true }

describe('macOS Desktop Excel data-table structural oracle', () => {
  it('retargets native data-table metadata through headless structural inserts', () => {
    for (const testCase of structuralDataTableCases()) {
      const workbook = WorkPaper.buildFromSnapshot(importXlsx(testCase.buildBiligXlsx(), testCase.fileName).snapshot, workbookConfig)
      try {
        testCase.applyHeadlessEdit(workbook)

        expect(workbook.exportSnapshot().sheets[0]?.metadata?.dataTableFormulas).toEqual(testCase.expectedMetadata)
        for (const expected of testCase.expectedCells) {
          expect(workbook.getCellFormula(addressToCell(expected.address))).toBe(expected.formula)
          expect(normalizedCellValue(workbook.getCellValue(addressToCell(expected.address)))).toEqual(expected.value)
        }

        const exported = exportXlsx(workbook.exportSnapshot())
        assertExportedDataTableXml(exported, testCase.expectedMetadata, testCase.staleDataTableAnchors)

        const roundTrip = importXlsx(exported, `roundtrip-${testCase.fileName}`)
        expect(roundTrip.snapshot.sheets[0]?.metadata?.dataTableFormulas).toEqual(testCase.expectedMetadata)
      } finally {
        workbook.dispose()
      }
    }
  })

  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'matches Desktop Excel native data-table metadata after structural inserts',
    () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = mkdtempSync(join(tmpdir(), 'bilig-headless-excel-data-table-structural-oracle-'))
      try {
        for (const testCase of structuralDataTableCases()) {
          const excelTruthPath = join(tempDir, `excel-${testCase.fileName}`)
          const baseWorkbook = testCase.buildBaseWorkbook(false)
          try {
            writeFileSync(excelTruthPath, exportXlsx(baseWorkbook.exportSnapshot()))
          } finally {
            baseWorkbook.dispose()
          }

          const excelTruth = runMacosExcelStructuralOperationOracle({
            workbookPath: excelTruthPath,
            worksheetName: 'DataTable',
            operations: [...testCase.createDataTableOperations, testCase.excelEdit],
            inspectCells: testCase.expectedCells.map(({ address }) => address),
            saveWorkbook: true,
          })
          const excelTruthImport = importXlsx(new Uint8Array(readFileSync(excelTruthPath)), `excel-${testCase.fileName}`)
          expect(excelTruthImport.snapshot.sheets[0]?.metadata?.dataTableFormulas).toEqual(testCase.expectedMetadata)

          const headlessWorkbook = WorkPaper.buildFromSnapshot(
            importXlsx(testCase.buildBiligXlsx(), testCase.fileName).snapshot,
            workbookConfig,
          )
          try {
            testCase.applyHeadlessEdit(headlessWorkbook)
            const headlessPath = join(tempDir, `headless-${testCase.fileName}`)
            writeFileSync(headlessPath, exportXlsx(headlessWorkbook.exportSnapshot()))

            const headlessExcel = runMacosExcelInspectionOracle({
              workbookPath: headlessPath,
              worksheetName: 'DataTable',
              formulaCells: [],
              inspectCells: testCase.expectedCells.map(({ address }) => address),
              saveWorkbook: true,
            })
            expect(headlessExcel.cells).toEqual(excelTruth.cells)

            const headlessImport = importXlsx(new Uint8Array(readFileSync(headlessPath)), `headless-${testCase.fileName}`)
            expect(headlessImport.snapshot.sheets[0]?.metadata?.dataTableFormulas).toEqual(testCase.expectedMetadata)
          } finally {
            headlessWorkbook.dispose()
          }
        }
      } finally {
        rmSync(tempDir, { recursive: true, force: true })
      }
    },
    60_000,
  )
})

interface StructuralDataTableCase {
  readonly fileName: string
  readonly buildBaseWorkbook: (includeOutputs: boolean) => WorkPaper
  readonly buildBiligXlsx: () => Uint8Array
  readonly createDataTableOperations: readonly MacosExcelStructuralOperation[]
  readonly excelEdit: MacosExcelStructuralOperation
  readonly applyHeadlessEdit: (workbook: WorkPaper) => void
  readonly expectedMetadata: WorkbookSheetDataTableFormulasSnapshot
  readonly expectedCells: readonly {
    readonly address: string
    readonly formula: string
    readonly value: NormalizedCellValue
  }[]
  readonly staleDataTableAnchors: readonly string[]
}

type NormalizedCellValue =
  | { readonly kind: 'blank' }
  | { readonly kind: 'boolean'; readonly value: boolean }
  | { readonly kind: 'error'; readonly value: string }
  | { readonly kind: 'number'; readonly value: number }
  | { readonly kind: 'string'; readonly value: string }

function structuralDataTableCases(): StructuralDataTableCase[] {
  return [
    {
      fileName: 'bilig-xlsx-two-variable-data-table-insert-row-oracle.xlsx',
      buildBaseWorkbook: buildTwoVariableDataTableWorkbook,
      buildBiligXlsx: buildBiligTwoVariableDataTableXlsx,
      createDataTableOperations: [{ kind: 'createDataTable', range: 'B2:D4', rowInput: 'A1', columnInput: 'A2' }],
      excelEdit: { kind: 'insertRows', range: '1:1' },
      applyHeadlessEdit: (workbook) => workbook.addRows(1, 0, 1),
      expectedMetadata: {
        formulas: [{ address: 'C4', formulaXml: '<f t="dataTable" ref="C4:D5" dt2D="1" dtr="1" r1="A2" r2="A3"/>' }],
      },
      expectedCells: [
        { address: 'C4', formula: '=MULTIPLE.OPERATIONS(B3,A2,C3,A3,B4)', value: { kind: 'number', value: 40 } },
        { address: 'D4', formula: '=MULTIPLE.OPERATIONS(B3,A2,D3,A3,B4)', value: { kind: 'number', value: 60 } },
        { address: 'C5', formula: '=MULTIPLE.OPERATIONS(B3,A2,C3,A3,B5)', value: { kind: 'number', value: 60 } },
        { address: 'D5', formula: '=MULTIPLE.OPERATIONS(B3,A2,D3,A3,B5)', value: { kind: 'number', value: 90 } },
      ],
      staleDataTableAnchors: ['C3'],
    },
    {
      fileName: 'bilig-xlsx-two-variable-data-table-insert-column-oracle.xlsx',
      buildBaseWorkbook: buildTwoVariableDataTableWorkbook,
      buildBiligXlsx: buildBiligTwoVariableDataTableXlsx,
      createDataTableOperations: [{ kind: 'createDataTable', range: 'B2:D4', rowInput: 'A1', columnInput: 'A2' }],
      excelEdit: { kind: 'insertColumns', range: 'A:A' },
      applyHeadlessEdit: (workbook) => workbook.addColumns(1, 0, 1),
      expectedMetadata: {
        formulas: [{ address: 'D3', formulaXml: '<f t="dataTable" ref="D3:E4" dt2D="1" dtr="1" r1="B1" r2="B2"/>' }],
      },
      expectedCells: [
        { address: 'D3', formula: '=MULTIPLE.OPERATIONS(C2,B1,D2,B2,C3)', value: { kind: 'number', value: 40 } },
        { address: 'E3', formula: '=MULTIPLE.OPERATIONS(C2,B1,E2,B2,C3)', value: { kind: 'number', value: 60 } },
        { address: 'D4', formula: '=MULTIPLE.OPERATIONS(C2,B1,D2,B2,C4)', value: { kind: 'number', value: 60 } },
        { address: 'E4', formula: '=MULTIPLE.OPERATIONS(C2,B1,E2,B2,C4)', value: { kind: 'number', value: 90 } },
      ],
      staleDataTableAnchors: ['C3'],
    },
    {
      fileName: 'bilig-xlsx-one-variable-data-table-insert-row-oracle.xlsx',
      buildBaseWorkbook: buildOneVariableDataTableWorkbook,
      buildBiligXlsx: buildBiligOneVariableDataTableXlsx,
      createDataTableOperations: [
        { kind: 'createDataTable', range: 'B1:D2', rowInput: 'A1' },
        { kind: 'createDataTable', range: 'A5:B8', columnInput: 'A1' },
      ],
      excelEdit: { kind: 'insertRows', range: '1:1' },
      applyHeadlessEdit: (workbook) => workbook.addRows(1, 0, 1),
      expectedMetadata: {
        formulas: [
          { address: 'C3', formulaXml: '<f t="dataTable" ref="C3:D3" dt2D="0" dtr="1" r1="A2" ca="1"/>' },
          { address: 'B7', formulaXml: '<f t="dataTable" ref="B7:B9" dt2D="0" dtr="0" r1="A2"/>' },
        ],
      },
      expectedCells: [
        { address: 'C3', formula: '=MULTIPLE.OPERATIONS(B3,A2,C2)', value: { kind: 'number', value: 30 } },
        { address: 'D3', formula: '=MULTIPLE.OPERATIONS(B3,A2,D2)', value: { kind: 'number', value: 40 } },
        { address: 'B7', formula: '=MULTIPLE.OPERATIONS(B6,A2,A7)', value: { kind: 'number', value: 20 } },
        { address: 'B8', formula: '=MULTIPLE.OPERATIONS(B6,A2,A8)', value: { kind: 'number', value: 30 } },
        { address: 'B9', formula: '=MULTIPLE.OPERATIONS(B6,A2,A9)', value: { kind: 'number', value: 40 } },
      ],
      staleDataTableAnchors: ['C2', 'B6'],
    },
    {
      fileName: 'bilig-xlsx-one-variable-data-table-insert-column-oracle.xlsx',
      buildBaseWorkbook: buildOneVariableDataTableWorkbook,
      buildBiligXlsx: buildBiligOneVariableDataTableXlsx,
      createDataTableOperations: [
        { kind: 'createDataTable', range: 'B1:D2', rowInput: 'A1' },
        { kind: 'createDataTable', range: 'A5:B8', columnInput: 'A1' },
      ],
      excelEdit: { kind: 'insertColumns', range: 'A:A' },
      applyHeadlessEdit: (workbook) => workbook.addColumns(1, 0, 1),
      expectedMetadata: {
        formulas: [
          { address: 'D2', formulaXml: '<f t="dataTable" ref="D2:E2" dt2D="0" dtr="1" r1="B1" ca="1"/>' },
          { address: 'C6', formulaXml: '<f t="dataTable" ref="C6:C8" dt2D="0" dtr="0" r1="B1"/>' },
        ],
      },
      expectedCells: [
        { address: 'D2', formula: '=MULTIPLE.OPERATIONS(C2,B1,D1)', value: { kind: 'number', value: 30 } },
        { address: 'E2', formula: '=MULTIPLE.OPERATIONS(C2,B1,E1)', value: { kind: 'number', value: 40 } },
        { address: 'C6', formula: '=MULTIPLE.OPERATIONS(C5,B1,B6)', value: { kind: 'number', value: 20 } },
        { address: 'C7', formula: '=MULTIPLE.OPERATIONS(C5,B1,B7)', value: { kind: 'number', value: 30 } },
        { address: 'C8', formula: '=MULTIPLE.OPERATIONS(C5,B1,B8)', value: { kind: 'number', value: 40 } },
      ],
      staleDataTableAnchors: ['C2', 'B6'],
    },
  ]
}

function buildTwoVariableDataTableWorkbook(includeOutputs: boolean): WorkPaper {
  return WorkPaper.buildFromSheets(
    {
      DataTable: [
        [1],
        [10, '=A3', 2, 3],
        ['=A1*A2', 20, includeOutputs ? 40 : null, includeOutputs ? 60 : null],
        [null, 30, includeOutputs ? 60 : null, includeOutputs ? 90 : null],
      ],
    },
    workbookConfig,
  )
}

function buildOneVariableDataTableWorkbook(includeOutputs: boolean): WorkPaper {
  return WorkPaper.buildFromSheets(
    {
      DataTable: [
        [1, 2, 3, 4],
        ['=A1*10', '=A2', includeOutputs ? 30 : null, includeOutputs ? 40 : null],
        [],
        [],
        [1, '=A1*10'],
        [2, includeOutputs ? 20 : null],
        [3, includeOutputs ? 30 : null],
        [4, includeOutputs ? 40 : null],
      ],
    },
    workbookConfig,
  )
}

function buildBiligTwoVariableDataTableXlsx(): Uint8Array {
  const workbook = buildTwoVariableDataTableWorkbook(true)
  try {
    const zip = unzipSync(exportXlsx(workbook.exportSnapshot()))
    const sheetXml = readZipTextFromZip(zip, 'xl/worksheets/sheet1.xml')
    zip['xl/worksheets/sheet1.xml'] = strToU8(
      sheetXml.replace(
        /<c\b[^>]*\br=(["'])C3\1[^>]*>[\s\S]*?<\/c>/u,
        '<c r="C3"><f t="dataTable" ref="C3:D4" dt2D="1" dtr="1" r1="A1" r2="A2"/><v>40</v></c>',
      ),
    )
    return zipSync(zip)
  } finally {
    workbook.dispose()
  }
}

function buildBiligOneVariableDataTableXlsx(): Uint8Array {
  const workbook = buildOneVariableDataTableWorkbook(true)
  try {
    const zip = unzipSync(exportXlsx(workbook.exportSnapshot()))
    const sheetXml = readZipTextFromZip(zip, 'xl/worksheets/sheet1.xml')
    zip['xl/worksheets/sheet1.xml'] = strToU8(
      sheetXml
        .replace(
          /<c\b[^>]*\br=(["'])C2\1[^>]*>[\s\S]*?<\/c>/u,
          '<c r="C2"><f t="dataTable" ref="C2:D2" dt2D="0" dtr="1" r1="A1" ca="1"/><v>30</v></c>',
        )
        .replace(
          /<c\b[^>]*\br=(["'])B6\1[^>]*>[\s\S]*?<\/c>/u,
          '<c r="B6"><f t="dataTable" ref="B6:B8" dt2D="0" dtr="0" r1="A1"/><v>20</v></c>',
        ),
    )
    return zipSync(zip)
  } finally {
    workbook.dispose()
  }
}

function assertExportedDataTableXml(
  bytes: Uint8Array,
  expectedMetadata: WorkbookSheetDataTableFormulasSnapshot,
  staleDataTableAnchors: readonly string[],
): void {
  expect(dataTableFormulaXml(bytes)).toEqual(expectedMetadata.formulas.map((formula) => formula.formulaXml))
  for (const formula of expectedMetadata.formulas) {
    expect(cellXml(bytes, formula.address)).toContain(formula.formulaXml)
  }
  for (const staleAnchor of staleDataTableAnchors) {
    expect(cellXml(bytes, staleAnchor)).not.toContain('dataTable')
  }
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

function normalizedCellValue(value: CellValue): NormalizedCellValue {
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

function dataTableFormulaXml(bytes: Uint8Array): string[] {
  return [...readZipText(bytes, 'xl/worksheets/sheet1.xml').matchAll(/<f\b[^>]*\/>|<f\b[^>]*>[\s\S]*?<\/f>/gu)]
    .map((match) => match[0])
    .filter((formulaXml) => /\bt=(["'])dataTable\1/u.test(formulaXml))
}

function cellXml(bytes: Uint8Array, address: string): string {
  const addressPattern = address.replace(/[.*+?^${}()|[\]\\]/gu, '\\$&')
  return (
    readZipText(bytes, 'xl/worksheets/sheet1.xml').match(
      new RegExp(`<c\\b(?=[^>]*\\br=(["'])${addressPattern}\\1)[^>]*>[\\s\\S]*?<\\/c>`, 'u'),
    )?.[0] ?? ''
  )
}

function readZipText(bytes: Uint8Array, path: string): string {
  return readZipTextFromZip(unzipSync(bytes), path)
}

function readZipTextFromZip(zip: Record<string, Uint8Array>, path: string): string {
  const bytes = zip[path]
  if (!bytes) {
    throw new Error(`Missing XLSX part: ${path}`)
  }
  return strFromU8(bytes)
}
