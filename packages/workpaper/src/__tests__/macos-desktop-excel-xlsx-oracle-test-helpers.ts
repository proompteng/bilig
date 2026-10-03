export { readRuntimeImage, SpreadsheetEngine } from '@bilig/core'
export {
  buildFormulaCellComparison,
  buildReportSummary,
  isMacosExcelInstalled,
  runMacosExcelInspectionOracle,
  runMacosExcelStructuralOperationOracle,
} from '@bilig/excel-fixtures'
export type { FormulaCellComparison, NormalizedFormulaValue } from '@bilig/excel-fixtures'
export { dataTableFormulasWarning, exportXlsx, importXlsx } from '@bilig/excel-import'
export { ErrorCode, ValueTag, type CellValue, type WorkbookSnapshot } from '@bilig/protocol'
export { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
export { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
export { tmpdir } from 'node:os'
export { join } from 'node:path'
export { describe, expect, it } from 'vitest'
export { WorkPaper, type WorkPaperCellAddress } from '../index.js'

import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'

import { SpreadsheetEngine } from '@bilig/core'
import { buildFormulaCellComparison, type FormulaCellComparison, type NormalizedFormulaValue } from '@bilig/excel-fixtures'
import { exportXlsx } from '@bilig/excel-import'
import { ErrorCode, ValueTag, type CellValue, type WorkbookSnapshot } from '@bilig/protocol'

import { WorkPaper, type WorkPaperCellAddress } from '../index.js'

export const workbookConfig = { maxRows: 16, maxColumns: 8, useColumnIndex: true }
export const indexImplicitIntersectionConfig = { maxRows: 10, maxColumns: 10, useColumnIndex: true }
export const oracleFormulaAddresses = ['C1', 'D1', 'E1', 'F1', 'G1', 'H1'] as const
export const expectedOracleCells = [
  { address: 'C1', formula: '=COUNTBLANK(A1:A5)', rawValue: 'number\t2.0', value: { kind: 'number', value: 2 } },
  { address: 'D1', formula: '=COUNTIF(A1:A5,"")', rawValue: 'number\t2.0', value: { kind: 'number', value: 2 } },
  { address: 'E1', formula: '=COUNTIF(A1:A5,"<>")', rawValue: 'number\t4.0', value: { kind: 'number', value: 4 } },
  { address: 'F1', formula: '=SUMIF(A1:A5,"",B1:B5)', rawValue: 'number\t50.0', value: { kind: 'number', value: 50 } },
  { address: 'G1', formula: '=SUMIF(A1:A5,"<>",B1:B5)', rawValue: 'number\t130.0', value: { kind: 'number', value: 130 } },
  { address: 'H1', formula: '=SUMIFS(B1:B5,A1:A5,"<>")', rawValue: 'number\t130.0', value: { kind: 'number', value: 130 } },
] as const
export const futureFunctionOracleAddresses = ['D2', 'E2', 'F2'] as const
export const expectedFutureFunctionOracleCells = [
  { address: 'D2', formula: '=TEXTJOIN("-",TRUE,A2:A4)', rawValue: 'string\ta-c', value: { kind: 'string', value: 'a-c' } },
  { address: 'E2', formula: '=XLOOKUP("b",B2:B4,C2:C4)', rawValue: 'number\t20.0', value: { kind: 'number', value: 20 } },
  { address: 'F2', formula: '=XMATCH("b",B2:B4,0)', rawValue: 'number\t2.0', value: { kind: 'number', value: 2 } },
] as const
export const singleImplicitIntersectionOracleAddresses = ['C1', 'C2', 'C3', 'D1'] as const
export const expectedSingleImplicitIntersectionOracleValues = [
  { address: 'C1', value: { kind: 'number', value: 1 } },
  { address: 'C2', value: { kind: 'number', value: 2 } },
  { address: 'C3', value: { kind: 'number', value: 3 } },
  { address: 'D1', value: { kind: 'number', value: 1 } },
] as const
export const expectedDesktopExcelSingleImplicitIntersectionOracleCells = [
  { address: 'C1', formula: '=A1:A3', rawValue: 'number\t1.0', value: { kind: 'number', value: 1 } },
  { address: 'C2', formula: '=A1:A3', rawValue: 'number\t2.0', value: { kind: 'number', value: 2 } },
  { address: 'C3', formula: '=A1:A3', rawValue: 'number\t3.0', value: { kind: 'number', value: 3 } },
  { address: 'D1', formula: '=SUM(@A1:A3)', rawValue: 'number\t1.0', value: { kind: 'number', value: 1 } },
] as const
export const dynamicSpillOracleAddresses = ['B1', 'B2', 'B3'] as const
export const expectedDynamicSpillOracleValues = [
  { address: 'B1', value: { kind: 'number', value: 2 } },
  { address: 'B2', value: { kind: 'number', value: 4 } },
  { address: 'B3', value: { kind: 'number', value: 6 } },
] as const
export const spillReferenceOracleAddresses = ['B1', 'B2', 'B3', 'D1', 'E1', 'F1'] as const
export const expectedSpillReferenceOracleCells = [
  { address: 'B1', formula: '=SEQUENCE(3,1,1,1)', rawValue: 'number\t1.0', value: { kind: 'number', value: 1 } },
  { address: 'B2', rawValue: 'number\t2.0', value: { kind: 'number', value: 2 } },
  { address: 'B3', rawValue: 'number\t3.0', value: { kind: 'number', value: 3 } },
  { address: 'D1', formula: '=SUM(B1#)', rawValue: 'number\t6.0', value: { kind: 'number', value: 6 } },
  { address: 'E1', formula: '=ROWS(B1#)', rawValue: 'number\t3.0', value: { kind: 'number', value: 3 } },
  { address: 'F1', formula: '=INDEX(B1#,2)', rawValue: 'number\t2.0', value: { kind: 'number', value: 2 } },
] as const
export const shrinkingSpillReferenceOracleAddresses = ['B1', 'B2', 'B3', 'D1', 'E1', 'F1'] as const
export const expectedShrinkingSpillReferenceOracleCells = [
  { address: 'B1', formula: '=SEQUENCE(A1,1,1,1)', rawValue: 'number\t1.0', value: { kind: 'number', value: 1 } },
  { address: 'B2', rawValue: 'blank\t', value: { kind: 'blank' } },
  { address: 'B3', rawValue: 'blank\t', value: { kind: 'blank' } },
  { address: 'D1', formula: '=SUM(B1#)', rawValue: 'number\t1.0', value: { kind: 'number', value: 1 } },
  { address: 'E1', formula: '=ROWS(B1#)', rawValue: 'number\t1.0', value: { kind: 'number', value: 1 } },
  {
    address: 'F1',
    formula: '=IFERROR(INDEX(B1#,2),"missing")',
    rawValue: 'string\tmissing',
    value: { kind: 'string', value: 'missing' },
  },
] as const
export const expectedDesktopExcelShrinkingSpillReferenceOracleCells = [
  expectedShrinkingSpillReferenceOracleCells[0],
  { address: 'B2', rawValue: 'string\t', value: { kind: 'string', value: '' } },
  { address: 'B3', rawValue: 'string\t', value: { kind: 'string', value: '' } },
  ...expectedShrinkingSpillReferenceOracleCells.slice(3),
] as const
export const blockedSpillReferenceOracleAddresses = ['B1', 'B2', 'B3', 'D1', 'E1', 'F1'] as const
export const expectedBlockedSpillReferenceOracleValues = [
  { address: 'B1', value: { kind: 'error', value: String(ErrorCode.Spill) } },
  { address: 'B2', value: { kind: 'number', value: 99 } },
  { address: 'B3', value: { kind: 'blank' } },
  { address: 'D1', value: { kind: 'error', value: String(ErrorCode.Spill) } },
  { address: 'E1', value: { kind: 'number', value: 1 } },
  { address: 'F1', value: { kind: 'string', value: 'missing' } },
] as const
export const expectedDesktopExcelBlockedSpillReferenceValues = [
  expectedBlockedSpillReferenceOracleValues[0],
  expectedBlockedSpillReferenceOracleValues[1],
  { address: 'B3', value: { kind: 'string', value: '' } },
  ...expectedBlockedSpillReferenceOracleValues.slice(3),
] as const
export const expectedUnblockedSpillReferenceOracleValues = [
  { address: 'B1', value: { kind: 'number', value: 1 } },
  { address: 'B2', value: { kind: 'number', value: 2 } },
  { address: 'B3', value: { kind: 'number', value: 3 } },
  { address: 'D1', value: { kind: 'number', value: 6 } },
  { address: 'E1', value: { kind: 'number', value: 3 } },
  { address: 'F1', value: { kind: 'number', value: 2 } },
] as const
export const horizontalStructuralSpillOracleAddresses = ['B1', 'C1', 'D1', 'E1', 'A3', 'A4', 'A5'] as const
export const expectedHorizontalStructuralSpillOracleValues = [
  { address: 'B1', value: { kind: 'number', value: 1 } },
  { address: 'C1', value: { kind: 'number', value: 2 } },
  { address: 'D1', value: { kind: 'number', value: 3 } },
  { address: 'E1', value: { kind: 'blank' } },
  { address: 'A3', value: { kind: 'number', value: 6 } },
  { address: 'A4', value: { kind: 'number', value: 3 } },
  { address: 'A5', value: { kind: 'number', value: 2 } },
] as const
export const expectedDesktopExcelHorizontalStructuralSpillValues = [
  ...expectedHorizontalStructuralSpillOracleValues.slice(0, 3),
  { address: 'E1', value: { kind: 'string', value: '' } },
  ...expectedHorizontalStructuralSpillOracleValues.slice(4),
] as const
export const textsplitErrorOracleAddresses = ['C1', 'D1', 'C2', 'D2'] as const
export const expectedTextsplitErrorOracleCells = [
  { address: 'C1', formula: '=TEXTSPLIT(A1,",","|")', rawValue: 'string\tred', value: { kind: 'string', value: 'red' } },
  { address: 'D1', rawValue: 'string\tblue', value: { kind: 'string', value: 'blue' } },
  { address: 'C2', rawValue: 'string\tgreen', value: { kind: 'string', value: 'green' } },
  { address: 'D2', rawValue: 'error\t#N/A', value: { kind: 'error', value: String(ErrorCode.NA) } },
] as const
export const chooseArrayIndexOracleAddresses = ['E1', 'F1', 'E2', 'F2', 'E3', 'F3', 'H1', 'H2', 'H3', 'H4', 'H6', 'H7'] as const
export const expectedChooseArrayIndexOracleValues = [
  { address: 'E1', value: { kind: 'string', value: 'a' } },
  { address: 'F1', value: { kind: 'number', value: 10 } },
  { address: 'E2', value: { kind: 'string', value: 'b' } },
  { address: 'F2', value: { kind: 'number', value: 20 } },
  { address: 'E3', value: { kind: 'string', value: 'c' } },
  { address: 'F3', value: { kind: 'number', value: 30 } },
  { address: 'H1', value: { kind: 'number', value: 660 } },
  { address: 'H2', value: { kind: 'number', value: 100 } },
  { address: 'H3', value: { kind: 'number', value: 200 } },
  { address: 'H4', value: { kind: 'number', value: 300 } },
  { address: 'H6', value: { kind: 'number', value: 600 } },
  { address: 'H7', value: { kind: 'number', value: 20 } },
] as const
export const structuralMoveColumnFormulaOracleCell = {
  address: 'F1',
  formula: '=SUM(B1:B1)',
  rawValue: 'number\t3.0',
  value: { kind: 'number', value: 3 },
} as const
export const tableColumnInsertOracleCells = [
  { address: 'B1', formula: 'Column1', rawValue: 'string\tColumn1', value: { kind: 'string', value: 'Column1' } },
  { address: 'F1', formula: '=SUM(Sales[Margin])', rawValue: 'number\t5.0', value: { kind: 'number', value: 5 } },
] as const
export const tableColumnDeleteOracleFormulaCells = [
  { address: 'D1', formula: '=SUM(#REF!)' },
  { address: 'E1', formula: '=SUM(Sales[Margin])' },
] as const
export const tableColumnDeleteDefinedNameOracleFormulaCells = [
  { address: 'D1', formula: '=SUM(SalesAmount)' },
  { address: 'E1', formula: '=SUM(SalesAmountFormula)' },
] as const
export const tableHeaderRenameOracleCells = [
  { address: 'B1', formula: 'Revenue', rawValue: 'string\tRevenue', value: { kind: 'string', value: 'Revenue' } },
  { address: 'E1', formula: '=SUM(Sales[Revenue])', rawValue: 'number\t30.0', value: { kind: 'number', value: 30 } },
  { address: 'F1', formula: '=SUM(Sales[Margin])', rawValue: 'number\t5.0', value: { kind: 'number', value: 5 } },
] as const
export const tableHeaderRenameDefinedNameOracleCells = [
  { address: 'B1', formula: 'Revenue', rawValue: 'string\tRevenue', value: { kind: 'string', value: 'Revenue' } },
  { address: 'E1', formula: '=SUM(SalesAmount)', rawValue: 'number\t30.0', value: { kind: 'number', value: 30 } },
  { address: 'F1', formula: '=SUM(SalesAmountFormula)', rawValue: 'number\t30.0', value: { kind: 'number', value: 30 } },
] as const
export const tableEmptyBodyOracleCell = {
  address: 'D1',
  formula: '=SUM(Sales[Amount])',
  rawValue: 'number\t0.0',
  value: { kind: 'number', value: 0 },
} as const
export const indexImplicitIntersectionOracleAddresses = ['E1', 'E2', 'E3', 'E4', 'A5', 'B5', 'C5', 'D5', 'E5', 'G1', 'H1', 'I1'] as const
export const expectedIndexImplicitIntersectionOracleCells = [
  { address: 'E1', formula: '=INDEX(A1:C3,0,2)', rawValue: 'number\t2.0', value: { kind: 'number', value: 2 } },
  { address: 'E2', formula: '=INDEX(A1:C3,0,2)', rawValue: 'number\t5.0', value: { kind: 'number', value: 5 } },
  { address: 'E3', formula: '=INDEX(A1:C3,0,2)', rawValue: 'number\t8.0', value: { kind: 'number', value: 8 } },
  { address: 'E4', formula: '=INDEX(A1:C3,0,2)', rawValue: 'error\t#VALUE!', value: { kind: 'error', value: String(ErrorCode.Value) } },
  { address: 'A5', formula: '=INDEX(A1:C3,2,0)', rawValue: 'number\t4.0', value: { kind: 'number', value: 4 } },
  { address: 'B5', formula: '=INDEX(A1:C3,2,0)', rawValue: 'number\t5.0', value: { kind: 'number', value: 5 } },
  { address: 'C5', formula: '=INDEX(A1:C3,2,0)', rawValue: 'number\t6.0', value: { kind: 'number', value: 6 } },
  { address: 'D5', formula: '=INDEX(A1:C3,2,0)', rawValue: 'error\t#VALUE!', value: { kind: 'error', value: String(ErrorCode.Value) } },
  { address: 'E5', formula: '=INDEX(A1:C3,0,0)', rawValue: 'error\t#VALUE!', value: { kind: 'error', value: String(ErrorCode.Value) } },
  { address: 'G1', formula: '=SUM(INDEX(A1:C3,0,2))', rawValue: 'number\t15.0', value: { kind: 'number', value: 15 } },
  { address: 'H1', formula: '=SUM(INDEX(A1:C3,2,0))', rawValue: 'number\t15.0', value: { kind: 'number', value: 15 } },
  { address: 'I1', formula: '=SUM(INDEX(A1:C3,0,0))', rawValue: 'number\t45.0', value: { kind: 'number', value: 45 } },
] as const
export const offsetImplicitIntersectionOracleAddresses = ['E1', 'E2', 'E3', 'E4', 'A5', 'B5', 'C5', 'D5', 'E5', 'G1', 'H1', 'I1'] as const
export const expectedOffsetImplicitIntersectionOracleCells = [
  { address: 'E1', formula: '=OFFSET(A1,0,1,3,1)', rawValue: 'number\t2.0', value: { kind: 'number', value: 2 } },
  { address: 'E2', formula: '=OFFSET(A1,0,1,3,1)', rawValue: 'number\t5.0', value: { kind: 'number', value: 5 } },
  { address: 'E3', formula: '=OFFSET(A1,0,1,3,1)', rawValue: 'number\t8.0', value: { kind: 'number', value: 8 } },
  { address: 'E4', formula: '=OFFSET(A1,0,1,3,1)', rawValue: 'error\t#VALUE!', value: { kind: 'error', value: String(ErrorCode.Value) } },
  { address: 'A5', formula: '=OFFSET(A2,0,0,1,3)', rawValue: 'number\t4.0', value: { kind: 'number', value: 4 } },
  { address: 'B5', formula: '=OFFSET(A2,0,0,1,3)', rawValue: 'number\t5.0', value: { kind: 'number', value: 5 } },
  { address: 'C5', formula: '=OFFSET(A2,0,0,1,3)', rawValue: 'number\t6.0', value: { kind: 'number', value: 6 } },
  { address: 'D5', formula: '=OFFSET(A2,0,0,1,3)', rawValue: 'error\t#VALUE!', value: { kind: 'error', value: String(ErrorCode.Value) } },
  { address: 'E5', formula: '=OFFSET(A1,0,0,3,3)', rawValue: 'error\t#VALUE!', value: { kind: 'error', value: String(ErrorCode.Value) } },
  { address: 'G1', formula: '=SUM(OFFSET(A1,0,1,3,1))', rawValue: 'number\t15.0', value: { kind: 'number', value: 15 } },
  { address: 'H1', formula: '=SUM(OFFSET(A2,0,0,1,3))', rawValue: 'number\t15.0', value: { kind: 'number', value: 15 } },
  { address: 'I1', formula: '=SUM(OFFSET(A1,0,0,3,3))', rawValue: 'number\t45.0', value: { kind: 'number', value: 45 } },
] as const
export const dataTableOracleAddresses = ['C3', 'D3', 'C4', 'D4'] as const
export const expectedDataTableOracleValues = [
  { address: 'C3', value: { kind: 'number', value: 40 } },
  { address: 'D3', value: { kind: 'number', value: 60 } },
  { address: 'C4', value: { kind: 'number', value: 60 } },
  { address: 'D4', value: { kind: 'number', value: 90 } },
] as const
export const expectedDataTableImportedFormulaByAddress = new Map([
  ['C3', '=MULTIPLE.OPERATIONS(B2,A1,C2,A2,B3)'],
  ['D3', '=MULTIPLE.OPERATIONS(B2,A1,D2,A2,B3)'],
  ['C4', '=MULTIPLE.OPERATIONS(B2,A1,C2,A2,B4)'],
  ['D4', '=MULTIPLE.OPERATIONS(B2,A1,D2,A2,B4)'],
] as const)
export const oneVariableDataTableOracleAddresses = ['C2', 'D2', 'B6', 'B7', 'B8'] as const
export const expectedOneVariableDataTableOracleValues = [
  { address: 'C2', value: { kind: 'number', value: 30 } },
  { address: 'D2', value: { kind: 'number', value: 40 } },
  { address: 'B6', value: { kind: 'number', value: 20 } },
  { address: 'B7', value: { kind: 'number', value: 30 } },
  { address: 'B8', value: { kind: 'number', value: 40 } },
] as const
export const expectedOneVariableDataTableImportedFormulaByAddress = new Map([
  ['C2', '=MULTIPLE.OPERATIONS(B2,A1,C1)'],
  ['D2', '=MULTIPLE.OPERATIONS(B2,A1,D1)'],
  ['B6', '=MULTIPLE.OPERATIONS(B5,A1,A6)'],
  ['B7', '=MULTIPLE.OPERATIONS(B5,A1,A7)'],
  ['B8', '=MULTIPLE.OPERATIONS(B5,A1,A8)'],
] as const)
export const aggregateOptionsOracleAddresses = ['B1', 'B2', 'B3', 'B4', 'C1'] as const
export const expectedAggregateOptionsOracleCells = [
  { address: 'B1', formula: '=AGGREGATE(9,3,A1:A5)', rawValue: 'number\t40.0', value: { kind: 'number', value: 40 } },
  { address: 'B2', formula: '=AGGREGATE(9,6,A1:A5)', rawValue: 'number\t120.0', value: { kind: 'number', value: 120 } },
  { address: 'B3', formula: '=AGGREGATE(9,4,A1:A5)', rawValue: 'error\t#DIV/0!', value: { kind: 'error', value: String(ErrorCode.Div0) } },
  { address: 'B4', formula: '=AGGREGATE(9,7,A1:A5)', rawValue: 'number\t100.0', value: { kind: 'number', value: 100 } },
  { address: 'C1', formula: '=SUBTOTAL(109,A1:A4)', rawValue: 'number\t40.0', value: { kind: 'number', value: 40 } },
] as const

export function buildOracleWorkbook(): WorkPaper {
  return WorkPaper.buildFromSheets(
    {
      Cases: [
        [
          'North',
          10,
          '=COUNTBLANK(A1:A5)',
          '=COUNTIF(A1:A5,"")',
          '=COUNTIF(A1:A5,"<>")',
          '=SUMIF(A1:A5,"",B1:B5)',
          '=SUMIF(A1:A5,"<>",B1:B5)',
          '=SUMIFS(B1:B5,A1:A5,"<>")',
        ],
        [null, 20],
        ['=IF(TRUE,"","x")', 30],
        [' ', 40],
        ['South', 50],
      ],
    },
    workbookConfig,
  )
}

export function buildAggregateOptionsOracleWorkbook(): WorkPaper {
  const snapshot: WorkbookSnapshot = {
    version: 1,
    workbook: { name: 'headless-aggregate-options-oracle' },
    sheets: [
      {
        id: 1,
        name: 'Cases',
        order: 0,
        metadata: {
          rows: [{ id: 'row:1', index: 1, hidden: true }],
        },
        cells: [
          { address: 'A1', value: 10 },
          { address: 'A2', value: 20 },
          { address: 'A3', value: 30 },
          { address: 'A4', formula: 'SUBTOTAL(9,A1:A3)' },
          { address: 'A5', formula: '1/0' },
          { address: 'B1', formula: 'AGGREGATE(9,3,A1:A5)' },
          { address: 'B2', formula: 'AGGREGATE(9,6,A1:A5)' },
          { address: 'B3', formula: 'AGGREGATE(9,4,A1:A5)' },
          { address: 'B4', formula: 'AGGREGATE(9,7,A1:A5)' },
          { address: 'C1', formula: 'SUBTOTAL(109,A1:A4)' },
        ],
      },
    ],
  }
  return WorkPaper.buildFromSnapshot(snapshot, workbookConfig)
}

export function buildIndexImplicitIntersectionOracleWorkbook(): WorkPaper {
  return WorkPaper.buildFromSheets(
    {
      Sheet1: [
        [1, 2, 3, null, '=INDEX(A1:C3,0,2)', null, '=SUM(INDEX(A1:C3,0,2))', '=SUM(INDEX(A1:C3,2,0))', '=SUM(INDEX(A1:C3,0,0))'],
        [4, 5, 6, null, '=INDEX(A1:C3,0,2)'],
        [7, 8, 9, null, '=INDEX(A1:C3,0,2)'],
        [null, null, null, null, '=INDEX(A1:C3,0,2)'],
        ['=INDEX(A1:C3,2,0)', '=INDEX(A1:C3,2,0)', '=INDEX(A1:C3,2,0)', '=INDEX(A1:C3,2,0)', '=INDEX(A1:C3,0,0)'],
      ],
    },
    indexImplicitIntersectionConfig,
  )
}

export function buildOffsetImplicitIntersectionOracleWorkbook(): WorkPaper {
  return WorkPaper.buildFromSheets(
    {
      Sheet1: [
        [1, 2, 3, null, '=OFFSET(A1,0,1,3,1)', null, '=SUM(OFFSET(A1,0,1,3,1))', '=SUM(OFFSET(A2,0,0,1,3))', '=SUM(OFFSET(A1,0,0,3,3))'],
        [4, 5, 6, null, '=OFFSET(A1,0,1,3,1)'],
        [7, 8, 9, null, '=OFFSET(A1,0,1,3,1)'],
        [null, null, null, null, '=OFFSET(A1,0,1,3,1)'],
        ['=OFFSET(A2,0,0,1,3)', '=OFFSET(A2,0,0,1,3)', '=OFFSET(A2,0,0,1,3)', '=OFFSET(A2,0,0,1,3)', '=OFFSET(A1,0,0,3,3)'],
      ],
    },
    indexImplicitIntersectionConfig,
  )
}

export function buildFutureFunctionOracleWorkbook(): WorkPaper {
  return WorkPaper.buildFromSheets(
    {
      Cases: [
        ['Label', 'Key', 'Value', 'Joined', 'Lookup', 'Match'],
        ['a', 'a', 10, '=TEXTJOIN("-",TRUE,A2:A4)', '=XLOOKUP("b",B2:B4,C2:C4)', '=XMATCH("b",B2:B4,0)'],
        [null, 'b', 20],
        ['c', 'c', 30],
      ],
    },
    workbookConfig,
  )
}

export function buildSingleImplicitIntersectionOracleWorkbook(): WorkPaper {
  return WorkPaper.buildFromSheets(
    {
      Cases: [
        [1, null, '=SINGLE(A1:A3)', '=SUM(SINGLE(A1:A3))'],
        [2, null, '=SINGLE(A1:A3)'],
        [3, null, '=SINGLE(A1:A3)'],
      ],
    },
    workbookConfig,
  )
}

export function buildDynamicSpillOracleWorkbook(): WorkPaper {
  return WorkPaper.buildFromSheets(
    {
      Cases: [[1, '=MAP(A1:A3,LAMBDA(x,x*2))'], [2], [3]],
    },
    workbookConfig,
  )
}

export function buildSpillReferenceOracleWorkbook(): WorkPaper {
  return WorkPaper.buildFromSheets(
    {
      Cases: [[1, '=SEQUENCE(3,1,1,1)', null, '=SUM(B1#)', '=ROWS(B1#)', '=INDEX(B1#,2)'], [2], [3]],
    },
    workbookConfig,
  )
}

export function buildShrinkingSpillReferenceOracleWorkbook(): WorkPaper {
  return WorkPaper.buildFromSheets(
    {
      Cases: [[3, '=SEQUENCE(A1,1,1,1)', null, '=SUM(B1#)', '=ROWS(B1#)', '=IFERROR(INDEX(B1#,2),"missing")'], [], []],
    },
    workbookConfig,
  )
}

export function buildHorizontalStructuralSpillOracleWorkbook(): WorkPaper {
  return WorkPaper.buildFromSheets(
    {
      Cases: [[null, '=SEQUENCE(1,3,1,1)'], [], ['=SUM(B1#)'], ['=COLUMNS(B1#)'], ['=IFERROR(INDEX(B1#,1,2),"missing")']],
    },
    workbookConfig,
  )
}

export function buildTwoDimensionalStructuralSpillOracleWorkbook(): WorkPaper {
  return WorkPaper.buildFromSheets(
    {
      Cases: [
        [],
        [null, '=SEQUENCE(2,3,1,1)', null, null, null, null, '=SUM(B2#)'],
        [null, null, null, null, null, null, '=ROWS(B2#)'],
        [null, null, null, null, null, null, '=COLUMNS(B2#)'],
        [null, null, null, null, null, null, '=IFERROR(INDEX(B2#,2,2),"missing")'],
      ],
    },
    workbookConfig,
  )
}

export function buildTextsplitErrorOracleWorkbook(): WorkPaper {
  return WorkPaper.buildFromSheets(
    {
      Cases: [['red,blue|green', null, '=TEXTSPLIT(A1,",","|")']],
    },
    workbookConfig,
  )
}

export function buildChooseArrayIndexOracleWorkbook(): WorkPaper {
  return WorkPaper.buildFromSheets(
    {
      ChooseRef: [
        ['a', 10, 100, null, '=CHOOSE({1,2},A1:A3,B1:B3)', null, null, '=SUM(CHOOSE({1,2},B1:B3,C1:C3))'],
        ['b', 20, 200, null, null, null, null, '=CHOOSE(2,A1:A3,C1:C3)'],
        ['c', 30, 300],
        [],
        [],
        [null, null, null, null, null, null, null, '=SUM(CHOOSE(2,B1:B3,C1:C3))'],
        [null, null, null, null, null, null, null, '=XLOOKUP("b",CHOOSE(1,A1:A3,C1:C3),CHOOSE(1,B1:B3,C1:C3),"missing",0)'],
      ],
    },
    workbookConfig,
  )
}

export function buildStructuralMoveColumnOracleWorkbook(): WorkPaper {
  return WorkPaper.buildFromSheets(
    {
      Cases: [[1, 2, 3, 4, 5, '=SUM(B1:C1)']],
    },
    workbookConfig,
  )
}

export function buildDataTableOracleWorkbook(includeOutputs: boolean): WorkPaper {
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

export function buildOneVariableDataTableOracleWorkbook(includeOutputs: boolean): WorkPaper {
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

export function buildBiligDataTableXlsx(): Uint8Array {
  const workbook = buildDataTableOracleWorkbook(true)
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

export function buildBiligOneVariableDataTableXlsx(): Uint8Array {
  const workbook = buildOneVariableDataTableOracleWorkbook(true)
  try {
    const zip = unzipSync(exportXlsx(workbook.exportSnapshot()))
    const sheetXml = readZipTextFromZip(zip, 'xl/worksheets/sheet1.xml')
    zip['xl/worksheets/sheet1.xml'] = strToU8(
      sheetXml
        .replace(
          /<c\b[^>]*\br=(["'])C2\1[^>]*>[\s\S]*?<\/c>/u,
          '<c r="C2"><f t="dataTable" ref="C2:D2" dt2D="0" dtr="1" r1="A1"/><v>30</v></c>',
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

export function readZipTextFromZip(zip: Record<string, Uint8Array>, path: string): string {
  const bytes = zip[path]
  if (!bytes) {
    throw new Error(`Missing XLSX part: ${path}`)
  }
  return strFromU8(bytes)
}

export async function buildTableColumnInsertOracleEngine(): Promise<SpreadsheetEngine> {
  const engine = new SpreadsheetEngine({ workbookName: 'table-column-insert-oracle' })
  await engine.ready()
  engine.createSheet('Data')
  engine.setRangeValues({ sheetName: 'Data', startAddress: 'A1', endAddress: 'C3' }, [
    ['Region', 'Revenue', 'Margin'],
    ['East', 10, 2],
    ['West', 20, 3],
  ])
  engine.setTable({
    name: 'Sales',
    sheetName: 'Data',
    startAddress: 'A1',
    endAddress: 'C3',
    columnNames: ['Region', 'Revenue', 'Margin'],
    headerRow: true,
    totalsRow: false,
  })
  engine.setCellFormula('Data', 'E1', 'SUM(Sales[Margin])')
  return engine
}

export async function buildTableColumnDeleteOracleEngine(): Promise<SpreadsheetEngine> {
  const engine = new SpreadsheetEngine({ workbookName: 'table-column-delete-oracle' })
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
  engine.setCellFormula('Data', 'E1', 'SUM(Sales[Amount])')
  engine.setCellFormula('Data', 'F1', 'SUM(Sales[Margin])')
  return engine
}

export async function buildTableColumnDeleteDefinedNameOracleEngine(): Promise<SpreadsheetEngine> {
  const engine = new SpreadsheetEngine({ workbookName: 'table-column-delete-defined-name-oracle' })
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
  engine.setDefinedName('SalesAmountFormula', { kind: 'formula', formula: '=Sales[Amount]' })
  engine.setCellFormula('Data', 'E1', 'SUM(SalesAmount)')
  engine.setCellFormula('Data', 'F1', 'SUM(SalesAmountFormula)')
  return engine
}

export async function buildTableHeaderRenameDefinedNameOracleEngine(): Promise<SpreadsheetEngine> {
  const engine = new SpreadsheetEngine({ workbookName: 'table-header-rename-defined-name-oracle' })
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
  engine.setDefinedName('SalesAmountFormula', { kind: 'formula', formula: '=Sales[Amount]' })
  engine.setCellFormula('Data', 'E1', 'SUM(SalesAmount)')
  engine.setCellFormula('Data', 'F1', 'SUM(SalesAmountFormula)')
  return engine
}

export async function buildTableHeaderRenameOracleEngine(): Promise<SpreadsheetEngine> {
  const engine = new SpreadsheetEngine({ workbookName: 'table-header-rename-oracle' })
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
  engine.setCellFormula('Data', 'E1', 'SUM(Sales[Amount])')
  engine.setCellFormula('Data', 'F1', 'SUM(Sales[Margin])')
  return engine
}

export async function buildTableEmptyBodyOracleEngine(): Promise<SpreadsheetEngine> {
  const engine = new SpreadsheetEngine({ workbookName: 'table-empty-body-oracle' })
  await engine.ready()
  engine.createSheet('Data')
  engine.setRangeValues({ sheetName: 'Data', startAddress: 'A1', endAddress: 'B2' }, [
    ['Region', 'Amount'],
    ['East', 10],
  ])
  engine.setTable({
    name: 'Sales',
    sheetName: 'Data',
    startAddress: 'A1',
    endAddress: 'B2',
    columnNames: ['Region', 'Amount'],
    headerRow: true,
    totalsRow: false,
  })
  engine.setCellFormula('Data', 'D1', 'SUM(Sales[Amount])')
  return engine
}

export function cell(row: number, col: number): WorkPaperCellAddress {
  return { sheet: 1, row, col }
}

export function buildHeadlessExcelComparisons(
  workbook: WorkPaper,
  excelCells: readonly { readonly address: string; readonly formula?: string; readonly value: NormalizedFormulaValue }[],
  workbookId = 'headless-oracle',
): FormulaCellComparison[] {
  return excelCells.map((excelCell) => {
    const address = addressToCell(excelCell.address)
    const formula = workbook.getCellFormula(address)
    if (!formula) {
      throw new Error(`Missing imported formula at ${excelCell.address}`)
    }
    return buildFormulaCellComparison({
      workbookId,
      sheet: 'Cases',
      address: excelCell.address,
      formula,
      ...(excelCell.formula !== undefined ? { excelOracleFormula: excelCell.formula } : {}),
      excelOracleValue: excelCell.value,
      actualBiligValue: normalizedCellValue(workbook.getCellValue(address)),
    })
  })
}

export function addressToCell(address: string): WorkPaperCellAddress {
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

export function normalizedCellValue(value: CellValue): NormalizedFormulaValue {
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
