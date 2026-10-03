import { readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'

import { exportXlsx, importXlsx } from '@bilig/excel-import'
import { isMacosExcelInstalled, runMacosExcelInspectionOracle, runMacosExcelStructuralOperationOracle } from '@bilig/excel-fixtures'
import type { WorkbookSnapshot } from '@bilig/protocol'
import { describe, expect, it } from 'vitest'

import { WorkPaper } from '../index.js'
import { createExcelAccessibleTempDir, removeMacosExcelTestDir } from './macos-excel-oracle-test-utils.js'

const sharedRichTextXml =
  '<si><r><rPr><b/><color rgb="FF1F4E79"/></rPr><t>Important:</t></r><r><rPr><i/></rPr><t xml:space="preserve"> Before signing off</t></r></si>'
const inlineRichTextXml =
  '<is><r><rPr><u/><color rgb="FF008000"/></rPr><t>Revenue</t></r><r><t xml:space="preserve"> sensitivity</t></r></is>'

describe('macOS Desktop Excel rich text oracle', () => {
  it('preserves and structurally rewrites imported rich text artifacts through WorkPaper export', () => {
    const workpaper = WorkPaper.buildFromSnapshot(importXlsx(buildRichTextWorkbookBytes(), 'rich-text-source.xlsx').snapshot)
    try {
      const sheet = workpaper.getSheetId('Labels')
      if (sheet === undefined) {
        throw new Error('Expected Labels sheet to be available')
      }
      workpaper.addRows(sheet, 0, 1)

      const exported = exportXlsx(workpaper.exportSnapshot())
      const reimported = importXlsx(exported, 'rich-text-headless-roundtrip.xlsx')

      expect(richTextCells(reimported.snapshot)).toEqual([
        { address: 'A2', text: 'Important: Before signing off', storage: 'sharedString' },
        { address: 'B2', text: 'Revenue sensitivity', storage: 'inlineString' },
      ])
      expect(richRunCount(exported)).toBe(4)
    } finally {
      workpaper.dispose()
    }
  })

  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'matches Desktop Excel rich text movement after structural row inserts',
    () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = createExcelAccessibleTempDir('bilig-headless-excel-rich-text-oracle-')
      try {
        const excelWorkbookPath = join(tempDir, 'excel-rich-text-structural-source.xlsx')
        writeFileSync(excelWorkbookPath, buildRichTextWorkbookBytes())

        const excelResult = runMacosExcelStructuralOperationOracle({
          workbookPath: excelWorkbookPath,
          worksheetName: 'Labels',
          operations: [{ kind: 'insertRows', range: '1:1' }],
          inspectCells: ['A1', 'A2', 'B2'],
          saveWorkbook: true,
          timeoutMs: 90_000,
        })

        const excelTruth = importXlsx(new Uint8Array(readFileSync(excelWorkbookPath)), 'excel-rich-text-structural-truth.xlsx')
        const excelRichTextCells = richTextCells(excelTruth.snapshot)
        expect(richTextLabels(excelRichTextCells)).toEqual([
          { address: 'A2', text: 'Important: Before signing off' },
          { address: 'B2', text: 'Revenue sensitivity' },
        ])

        const workpaper = WorkPaper.buildFromSnapshot(importXlsx(buildRichTextWorkbookBytes(), 'headless-rich-text-source.xlsx').snapshot)
        try {
          const sheet = workpaper.getSheetId('Labels')
          if (sheet === undefined) {
            throw new Error('Expected Labels sheet to be available')
          }
          workpaper.addRows(sheet, 0, 1)

          const headlessPath = join(tempDir, 'headless-rich-text-structural.xlsx')
          writeFileSync(headlessPath, exportXlsx(workpaper.exportSnapshot()))
          const headlessExcel = runMacosExcelInspectionOracle({
            workbookPath: headlessPath,
            worksheetName: 'Labels',
            formulaCells: [],
            inspectCells: ['A1', 'A2', 'B2'],
            saveWorkbook: true,
            timeoutMs: 90_000,
          })
          expect(headlessExcel.cells).toEqual(excelResult.cells)

          const headlessTruth = importXlsx(new Uint8Array(readFileSync(headlessPath)), 'headless-rich-text-structural-truth.xlsx')
          expect(richTextLabels(richTextCells(headlessTruth.snapshot))).toEqual(richTextLabels(excelRichTextCells))
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

function buildRichTextWorkbookBytes(): Uint8Array {
  return exportXlsx(richTextSnapshot())
}

function richTextCells(
  snapshot: WorkbookSnapshot,
): Array<{ readonly address: string; readonly text: string; readonly storage: 'sharedString' | 'inlineString' }> {
  return (
    snapshot.sheets[0]?.metadata?.richTextArtifacts?.cells.map((cell) => ({
      address: cell.address,
      text: cell.text,
      storage: cell.storage,
    })) ?? []
  )
}

function richTextLabels(
  cells: readonly { readonly address: string; readonly text: string }[],
): Array<{ readonly address: string; readonly text: string }> {
  return cells.map((cell) => ({ address: cell.address, text: cell.text }))
}

function richRunCount(bytes: Uint8Array): number {
  const exported = importXlsx(bytes, 'rich-text-run-count.xlsx').snapshot
  return (
    exported.sheets[0]?.metadata?.richTextArtifacts?.cells.reduce((count, cell) => count + (cell.xml.match(/<r\b/gu)?.length ?? 0), 0) ?? 0
  )
}

function richTextSnapshot(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: { name: 'Rich text oracle' },
    sheets: [
      {
        id: 1,
        name: 'Labels',
        order: 0,
        metadata: {
          richTextArtifacts: {
            cells: [
              {
                address: 'A1',
                text: 'Important: Before signing off',
                storage: 'sharedString',
                xml: sharedRichTextXml,
              },
              {
                address: 'B1',
                text: 'Revenue sensitivity',
                storage: 'inlineString',
                xml: inlineRichTextXml,
              },
            ],
          },
        },
        cells: [
          { address: 'A1', value: 'Important: Before signing off' },
          { address: 'B1', value: 'Revenue sensitivity' },
        ],
      },
    ],
  }
}
