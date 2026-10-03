import { readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'

import { exportXlsx, importXlsx } from '@bilig/excel-import'
import { isMacosExcelInstalled, runMacosExcelInspectionOracle } from '@bilig/excel-fixtures'
import type { WorkbookSnapshot } from '@bilig/protocol'
import { writeSimpleXlsxWorkbook } from '@bilig/xlsx'
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { describe, expect, it } from 'vitest'

import { WorkPaper } from '../index.js'
import { createExcelAccessibleTempDir, removeMacosExcelTestDir } from './macos-excel-oracle-test-utils.js'

const protectedRangeSecurityAttributes = [
  { name: 'password', value: 'AF2B' },
  { name: 'algorithmName', value: 'SHA-512' },
  { name: 'hashValue', value: 'QUJDREVGRw==' },
  { name: 'saltValue', value: 'SElKS0w=' },
  { name: 'spinCount', value: '100000' },
  { name: 'securityDescriptor', value: 'user&group' },
]

describe('macOS Desktop Excel protected ranges oracle', () => {
  it('preserves imported protected editable range security attributes through WorkPaper export', () => {
    const workpaper = WorkPaper.buildFromSnapshot(
      withUnlockedReviewCell(importXlsx(buildProtectedEditableRangeSecurityBytes(), 'protected-editable-range-source.xlsx').snapshot),
    )
    try {
      const sheet = workpaper.getSheetId('Protected')
      if (sheet === undefined) {
        throw new Error('Expected Protected sheet to be available')
      }
      workpaper.setCellContents({ sheet, row: 0, col: 0 }, 'Reviewed')

      const reimported = importXlsx(exportXlsx(workpaper.exportSnapshot()), 'protected-editable-range-headless.xlsx')

      expect(protectedRangeSecuritySummary(reimported.snapshot)).toEqual([
        {
          id: 'EditableInputs',
          range: 'B2:C3',
          attributes: Object.fromEntries(protectedRangeSecurityAttributes.map((attribute) => [attribute.name, attribute.value])),
        },
      ])
    } finally {
      workpaper.dispose()
    }
  })

  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'preserves Desktop Excel protected editable range security attributes after a headless edit',
    () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = createExcelAccessibleTempDir('bilig-headless-excel-protected-ranges-oracle-')
      try {
        const excelWorkbookPath = join(tempDir, 'excel-protected-editable-range-source.xlsx')
        writeFileSync(excelWorkbookPath, buildProtectedEditableRangeSecurityBytes())

        const excelResult = runMacosExcelInspectionOracle({
          workbookPath: excelWorkbookPath,
          worksheetName: 'Protected',
          formulaCells: [],
          inspectCells: ['A1', 'B2'],
          saveWorkbook: true,
          timeoutMs: 90_000,
        })
        expect(excelResult.cells[0]?.value).toEqual({ kind: 'string', value: 'Label' })
        expect(excelResult.cells[1]?.value).toEqual({ kind: 'number', value: 10 })

        const excelTruth = importXlsx(new Uint8Array(readFileSync(excelWorkbookPath)), 'excel-protected-range-truth.xlsx')
        const excelSummary = protectedRangeSecuritySummary(excelTruth.snapshot)
        expect(excelSummary).toEqual([
          {
            id: 'EditableInputs',
            range: 'B2:C3',
            attributes: Object.fromEntries(protectedRangeSecurityAttributes.map((attribute) => [attribute.name, attribute.value])),
          },
        ])

        const workpaper = WorkPaper.buildFromSnapshot(withUnlockedReviewCell(excelTruth.snapshot))
        try {
          const sheet = workpaper.getSheetId('Protected')
          if (sheet === undefined) {
            throw new Error('Expected Protected sheet to be available')
          }
          workpaper.setCellContents({ sheet, row: 0, col: 0 }, 'Reviewed')

          const headlessPath = join(tempDir, 'headless-protected-editable-range.xlsx')
          writeFileSync(headlessPath, exportXlsx(workpaper.exportSnapshot()))
          const headlessExcel = runMacosExcelInspectionOracle({
            workbookPath: headlessPath,
            worksheetName: 'Protected',
            formulaCells: [],
            inspectCells: ['A1', 'B2'],
            saveWorkbook: true,
            timeoutMs: 90_000,
          })
          expect(headlessExcel.cells[0]?.value).toEqual({ kind: 'string', value: 'Reviewed' })
          expect(headlessExcel.cells[1]?.value).toEqual({ kind: 'number', value: 10 })

          const headlessTruth = importXlsx(new Uint8Array(readFileSync(headlessPath)), 'headless-protected-range-truth.xlsx')
          expect(protectedRangeSecuritySummary(headlessTruth.snapshot)).toEqual(excelSummary)
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

function buildProtectedEditableRangeSecurityBytes(): Uint8Array {
  const zip = unzipSync(
    writeSimpleXlsxWorkbook({
      sheets: [
        {
          name: 'Protected',
          cells: [
            { address: 'A1', row: 0, col: 0, value: 'Label' },
            { address: 'B1', row: 0, col: 1, value: 'Input A' },
            { address: 'C1', row: 0, col: 2, value: 'Input B' },
            { address: 'A2', row: 1, col: 0, value: 'North' },
            { address: 'B2', row: 1, col: 1, value: 10 },
            { address: 'C2', row: 1, col: 2, value: 20 },
            { address: 'A3', row: 2, col: 0, value: 'South' },
            { address: 'B3', row: 2, col: 1, value: 30 },
            { address: 'C3', row: 2, col: 2, value: 40 },
          ],
        },
      ],
    }),
  )
  const sourceSheetXml = strFromU8(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
  zip['xl/worksheets/sheet1.xml'] = strToU8(
    sourceSheetXml.replace(
      '</sheetData>',
      [
        '</sheetData>',
        '<sheetProtection sheet="1"/>',
        '<protectedRanges>',
        '<protectedRange name="EditableInputs" sqref="B2:C3" password="AF2B" algorithmName="SHA-512" hashValue="QUJDREVGRw==" saltValue="SElKS0w=" spinCount="100000" securityDescriptor="user&amp;group"/>',
        '</protectedRanges>',
      ].join(''),
    ),
  )
  return zipSync(zip)
}

function withUnlockedReviewCell(snapshot: WorkbookSnapshot): WorkbookSnapshot {
  const clone = structuredClone(snapshot)
  clone.workbook.metadata = clone.workbook.metadata ?? {}
  clone.workbook.metadata.styles = [
    ...(clone.workbook.metadata.styles ?? []),
    { id: 'review-cell-unlocked', protection: { locked: false } },
  ]
  const protectedSheet = clone.sheets.find((sheet) => sheet.name === 'Protected')
  if (!protectedSheet) {
    throw new Error('Expected Protected sheet to be available')
  }
  protectedSheet.metadata = protectedSheet.metadata ?? {}
  protectedSheet.metadata.styleRanges = [
    ...(protectedSheet.metadata.styleRanges ?? []),
    {
      range: { sheetName: 'Protected', startAddress: 'A1', endAddress: 'A1' },
      styleId: 'review-cell-unlocked',
    },
  ]
  return clone
}

function protectedRangeSecuritySummary(snapshot: WorkbookSnapshot): Array<{
  readonly id: string
  readonly range: string
  readonly attributes: Record<string, string>
}> {
  return (snapshot.sheets[0]?.metadata?.protectedRanges ?? []).map((protection) => ({
    id: protection.id,
    range: `${protection.range.startAddress}:${protection.range.endAddress}`,
    attributes: Object.fromEntries((protection.xmlAttributes ?? []).map((attribute) => [attribute.name, attribute.value])),
  }))
}
