import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'

import { SpreadsheetEngine } from '@bilig/core'
import { exportXlsx, importXlsx } from '@bilig/excel-import'
import { writeSimpleXlsxWorkbook } from '@bilig/xlsx'
import {
  isMacosExcelInstalled,
  runMacosExcelInspectionOracle,
  runMacosExcelRejectedStructuralOperationOracle,
  type MacosExcelStructuralOperation,
} from '@bilig/excel-fixtures'
import { describe, expect, it } from 'vitest'

import { WorkPaper } from '../index.js'

const workbookConfig = { maxRows: 16, maxColumns: 8, useColumnIndex: true }

describe('macOS Desktop Excel workbook structure protection oracle', () => {
  it('blocks Bilig sheet topology mutations after importing workbook structure protection', async () => {
    const imported = importXlsx(buildWorkbookStructureProtectionBytes(), 'protected-workbook-structure.xlsx')
    const engine = new SpreadsheetEngine({ workbookName: 'protected-workbook-structure' })
    await engine.ready()
    engine.importSnapshot(imported.snapshot)

    expect(() => engine.createSheet('Added')).toThrow(/workbook structure is protected/)
    expect(() => engine.moveSheet('Report', 0)).toThrow(/workbook structure is protected/)
    expect(() => engine.renameSheet('Data', 'Source')).toThrow(/workbook structure is protected/)
    expect(() => engine.deleteSheet('Data')).toThrow(/workbook structure is protected/)
    expect(engine.exportSnapshot().sheets.map((sheet) => sheet.name)).toEqual(['Data', 'Report'])

    const workbook = WorkPaper.buildFromSnapshot(imported.snapshot, workbookConfig)
    try {
      const dataSheet = workbook.getSheetId('Data')!
      const reportSheet = workbook.getSheetId('Report')!

      expect(workbook.isItPossibleToAddSheet('Added')).toBe(false)
      expect(workbook.isItPossibleToRemoveSheet(dataSheet)).toBe(false)
      expect(workbook.isItPossibleToRenameSheet(dataSheet, 'Source')).toBe(false)

      expect(() => workbook.addSheet('Added')).toThrow(/Workbook structure is protected/)
      expect(() => workbook.moveSheet(reportSheet, 0)).toThrow(/Workbook structure is protected/)
      expect(() => workbook.renameSheet(dataSheet, 'Source')).toThrow(/Workbook structure is protected/)
      expect(() => workbook.removeSheet(dataSheet)).toThrow(/Workbook structure is protected/)
      expect(workbook.getSheetNames()).toEqual(['Data', 'Report'])
    } finally {
      workbook.dispose()
    }
  })

  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'matches Desktop Excel protected workbook rejection for protected sheet topology commands',
    () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = mkdtempSync(join(tmpdir(), 'bilig-headless-excel-workbook-protection-oracle-'))
      try {
        const operations: Array<{ readonly name: string; readonly operation: MacosExcelStructuralOperation }> = [
          { name: 'create-sheet', operation: { kind: 'createSheet', name: 'Added' } },
          { name: 'rename-sheet', operation: { kind: 'renameSheet', newName: 'Source' } },
          { name: 'move-sheet', operation: { kind: 'moveSheet', name: 'Report', before: 'Data' } },
        ]

        for (const { name, operation } of operations) {
          const workbookPath = join(tempDir, `${name}.xlsx`)
          writeFileSync(workbookPath, buildWorkbookStructureProtectionBytes())

          const excelResult = runMacosExcelRejectedStructuralOperationOracle({
            workbookPath,
            worksheetName: 'Data',
            operation,
            timeoutMs: 90_000,
          })

          expect(excelResult.excelVersion).toMatch(/^\d+\./u)
          expect(excelResult.errorMessage.length).toBeGreaterThan(0)
          expect(excelResult.sheetNames).toEqual(['Data', 'Report'])
        }

        const imported = importXlsx(buildWorkbookStructureProtectionBytes(), 'protected-workbook-structure.xlsx')
        const workbook = WorkPaper.buildFromSnapshot(imported.snapshot, workbookConfig)
        try {
          const dataSheet = workbook.getSheetId('Data')!
          expect(() => workbook.renameSheet(dataSheet, 'Source')).toThrow(/Workbook structure is protected/)

          const headlessPath = join(tempDir, 'headless-protected-workbook-structure.xlsx')
          writeFileSync(headlessPath, exportXlsx(workbook.exportSnapshot()))
          const headlessExcel = runMacosExcelInspectionOracle({
            workbookPath: headlessPath,
            worksheetName: 'Data',
            formulaCells: [],
            inspectCells: ['A1'],
            saveWorkbook: true,
            timeoutMs: 90_000,
          })
          expect(headlessExcel.cells[0]?.value).toEqual({ kind: 'string', value: 'Revenue' })
          expect(
            importXlsx(new Uint8Array(readFileSync(headlessPath)), 'headless-protected-workbook-structure.xlsx').snapshot.sheets,
          ).toHaveLength(2)
        } finally {
          workbook.dispose()
        }
      } finally {
        rmSync(tempDir, { recursive: true, force: true })
      }
    },
    180_000,
  )
})

function buildWorkbookStructureProtectionBytes(): Uint8Array {
  const zip = unzipSync(
    writeSimpleXlsxWorkbook({
      sheets: [
        {
          name: 'Data',
          cells: [
            { address: 'A1', row: 0, col: 0, value: 'Revenue' },
            { address: 'B1', row: 0, col: 1, value: 'Value' },
            { address: 'A2', row: 1, col: 0, value: 'North' },
            { address: 'B2', row: 1, col: 1, value: 100 },
          ],
        },
        {
          name: 'Report',
          cells: [
            { address: 'A1', row: 0, col: 0, value: 'Summary' },
            { address: 'A2', row: 1, col: 0, formula: 'SUM(Data!B2:B2)' },
          ],
        },
      ],
    }),
  )
  const sourceWorkbookXml = strFromU8(zip['xl/workbook.xml'] ?? new Uint8Array())
  const workbookProtectionXml = '<workbookProtection lockStructure="1" workbookPassword="AF2B"/>'
  zip['xl/workbook.xml'] = strToU8(insertWorkbookProtection(sourceWorkbookXml, workbookProtectionXml))
  return zipSync(zip)
}

function insertWorkbookProtection(sourceWorkbookXml: string, workbookProtectionXml: string): string {
  const workbookPrMatch = /<workbookPr\b[^>]*(?:\/>|>[\s\S]*?<\/workbookPr>)/u.exec(sourceWorkbookXml)
  if (workbookPrMatch?.index !== undefined) {
    const insertIndex = workbookPrMatch.index + workbookPrMatch[0].length
    return `${sourceWorkbookXml.slice(0, insertIndex)}${workbookProtectionXml}${sourceWorkbookXml.slice(insertIndex)}`
  }
  return sourceWorkbookXml.replace(/<sheets\b/u, `${workbookProtectionXml}<sheets`)
}
