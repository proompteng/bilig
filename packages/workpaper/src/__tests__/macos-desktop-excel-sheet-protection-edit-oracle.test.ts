import { readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'

import { exportXlsx, importXlsx } from '@bilig/excel-import'
import { isMacosExcelInstalled, runMacosExcelStructuralOperationOracle } from '@bilig/excel-fixtures'
import type { WorkbookSnapshot } from '@bilig/protocol'
import { describe, expect, it } from 'vitest'

import { WorkPaper } from '../index.js'
import { createExcelAccessibleTempDir, removeMacosExcelTestDir } from './macos-excel-oracle-test-utils.js'

describe('macOS Desktop Excel sheet protection edit oracle', () => {
  it('allows Bilig content edits to unlocked cells while blocking locked cells on protected sheets', () => {
    const workbook = WorkPaper.buildFromSnapshot(importXlsx(buildProtectedInputWorkbookBytes(), 'protected-input-template.xlsx').snapshot)
    try {
      const inputSheet = workbook.getSheetId('Input')
      if (inputSheet === undefined) {
        throw new Error('Expected Input sheet')
      }

      expect(() => workbook.setCellContents({ sheet: inputSheet, row: 1, col: 1 }, 25)).not.toThrow()
      expect(() => workbook.setCellContents({ sheet: inputSheet, row: 1, col: 2 }, 50)).toThrow(/Workbook protection blocks this change/)

      const exported = importXlsx(exportXlsx(workbook.exportSnapshot()), 'headless-protected-input-template.xlsx')
      expect(exported.snapshot.sheets[0]?.cells).toContainEqual(expect.objectContaining({ address: 'B2', value: 25 }))
      expect(exported.snapshot.sheets[0]?.metadata?.sheetProtection).toEqual({ sheetName: 'Input' })
      expect(exported.snapshot.sheets[0]?.metadata?.styleRanges).toEqual(
        expect.arrayContaining([
          expect.objectContaining({ range: { sheetName: 'Input', startAddress: 'B2', endAddress: 'B2' } }),
          expect.objectContaining({ range: { sheetName: 'Input', startAddress: 'C2', endAddress: 'C2' } }),
        ]),
      )
    } finally {
      workbook.dispose()
    }
  })

  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'matches Desktop Excel unlocked-cell edits and locked-cell metadata on protected sheets',
    () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = createExcelAccessibleTempDir('bilig-headless-excel-sheet-protection-edit-oracle-')
      try {
        const excelAcceptedPath = join(tempDir, 'excel-unlocked-input.xlsx')
        writeFileSync(excelAcceptedPath, buildProtectedInputWorkbookBytes())
        const accepted = runMacosExcelStructuralOperationOracle({
          workbookPath: excelAcceptedPath,
          worksheetName: 'Input',
          operations: [{ kind: 'setCellValue', address: 'B2', value: 25 }],
          formulaCells: [{ address: 'D2', formula: '=B2+1' }],
          inspectCells: ['B2', 'C2', 'D2'],
          saveWorkbook: true,
          timeoutMs: 90_000,
        })
        expect(accepted.excelVersion).toMatch(/^\d+\./u)
        expect(accepted.cells.map((cell) => ({ address: cell.address, value: cell.value }))).toEqual([
          { address: 'B2', value: { kind: 'number', value: 25 } },
          { address: 'C2', value: { kind: 'number', value: 50 } },
          { address: 'D2', value: { kind: 'number', value: 26 } },
        ])

        const importedAccepted = importXlsx(new Uint8Array(readFileSync(excelAcceptedPath)), 'excel-unlocked-input.xlsx')
        expect(importedAccepted.snapshot.sheets[0]?.cells).toContainEqual(expect.objectContaining({ address: 'B2', value: 25 }))
        expect(importedAccepted.snapshot.sheets[0]?.metadata?.sheetProtection).toBeDefined()
        expect(protectionForCell(importedAccepted.snapshot, 'B2')?.locked).toBe(false)
        expect(protectionForCell(importedAccepted.snapshot, 'C2')?.locked).not.toBe(false)

        const workpaper = WorkPaper.buildFromSnapshot(importedAccepted.snapshot)
        try {
          const inputSheet = workpaper.getSheetId('Input')
          if (inputSheet === undefined) {
            throw new Error('Expected Input sheet')
          }
          expect(() => workpaper.setCellContents({ sheet: inputSheet, row: 1, col: 1 }, 30)).not.toThrow()
          expect(() => workpaper.setCellContents({ sheet: inputSheet, row: 1, col: 2 }, 99)).toThrow(
            /Workbook protection blocks this change/,
          )

          const headlessPath = join(tempDir, 'headless-unlocked-input.xlsx')
          writeFileSync(headlessPath, exportXlsx(workpaper.exportSnapshot()))
          const headlessExcel = runMacosExcelStructuralOperationOracle({
            workbookPath: headlessPath,
            worksheetName: 'Input',
            operations: [{ kind: 'setCellValue', address: 'B2', value: 35 }],
            formulaCells: [{ address: 'D2', formula: '=B2+1' }],
            inspectCells: ['B2', 'D2'],
            saveWorkbook: true,
            timeoutMs: 90_000,
          })
          expect(headlessExcel.cells.map((cell) => ({ address: cell.address, value: cell.value }))).toEqual([
            { address: 'B2', value: { kind: 'number', value: 35 } },
            { address: 'D2', value: { kind: 'number', value: 36 } },
          ])
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

function buildProtectedInputWorkbookBytes(): Uint8Array {
  return exportXlsx(protectedInputTemplateSnapshot())
}

function protectionForCell(
  snapshot: WorkbookSnapshot,
  address: string,
): { readonly locked?: boolean; readonly hidden?: boolean } | undefined {
  const styleById = new Map((snapshot.workbook.metadata?.styles ?? []).map((style) => [style.id, style.protection]))
  const styleId = snapshot.sheets[0]?.metadata?.styleRanges?.find(
    (range) => range.range.startAddress === address && range.range.endAddress === address,
  )?.styleId
  return styleId ? styleById.get(styleId) : undefined
}

function protectedInputTemplateSnapshot(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: {
      name: 'Protected input template',
      metadata: {
        styles: [
          {
            id: 'unlocked-input',
            fill: { backgroundColor: '#fff2cc' },
            protection: { locked: false },
          },
          {
            id: 'locked-formula',
            protection: { locked: true, hidden: true },
          },
        ],
      },
    },
    sheets: [
      {
        id: 1,
        name: 'Input',
        order: 0,
        metadata: {
          sheetProtection: { sheetName: 'Input' },
          styleRanges: [
            { range: { sheetName: 'Input', startAddress: 'B2', endAddress: 'B2' }, styleId: 'unlocked-input' },
            { range: { sheetName: 'Input', startAddress: 'C2', endAddress: 'C2' }, styleId: 'locked-formula' },
          ],
        },
        cells: [
          { address: 'A1', value: 'Metric' },
          { address: 'B1', value: 'Input' },
          { address: 'C1', value: 'Locked formula' },
          { address: 'D1', value: 'Check' },
          { address: 'A2', value: 'Revenue' },
          { address: 'B2', value: 10 },
          { address: 'C2', formula: 'B2*2' },
          { address: 'D2', formula: 'B2+1' },
        ],
      },
    ],
  }
}
