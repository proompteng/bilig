import { readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'

import { SpreadsheetEngine } from '@bilig/core'
import { exportXlsx, importXlsx } from '@bilig/excel-import'
import { isMacosExcelInstalled, runMacosExcelInspectionOracle, runMacosExcelStructuralOperationOracle } from '@bilig/excel-fixtures'
import type { WorkbookSnapshot } from '@bilig/protocol'
import { describe, expect, it } from 'vitest'

import { createExcelAccessibleTempDir, removeMacosExcelTestDir } from './macos-excel-oracle-test-utils.js'

const printPageSetup = {
  printOptionsXml: '<printOptions horizontalCentered="1" gridLines="1"/>',
  pageMarginsXml: '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>',
  pageSetupXml: '<pageSetup paperSize="9" scale="60" orientation="landscape"/>',
  headerFooterXml: '<headerFooter alignWithMargins="0"><oddFooter>Page &amp;P</oddFooter></headerFooter>',
  rowBreaksXml: '<rowBreaks count="2" manualBreakCount="2"><brk id="3" max="16383" man="1"/><brk id="6" max="16383" man="1"/></rowBreaks>',
  colBreaksXml:
    '<colBreaks count="2" manualBreakCount="2"><brk id="2" max="1048575" man="1"/><brk id="5" max="1048575" man="1"/></colBreaks>',
} as const

describe('macOS Desktop Excel print page setup oracle', () => {
  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'matches Desktop Excel manual page-break ids after structural inserts',
    async () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = createExcelAccessibleTempDir('bilig-headless-excel-print-page-setup-oracle-')
      try {
        await expectStructuralPrintBreakCase({
          tempDir,
          filePrefix: 'row-insert',
          excelOperation: { kind: 'insertRows', range: '1:1' },
          applyHeadlessEdit: (engine) => engine.insertRows('Report', 0, 1),
          expectedRowBreakIds: [4, 7],
          expectedColumnBreakIds: [2, 5],
        })
        await expectStructuralPrintBreakCase({
          tempDir,
          filePrefix: 'column-insert',
          excelOperation: { kind: 'insertColumns', range: 'A:A' },
          applyHeadlessEdit: (engine) => engine.insertColumns('Report', 0, 1),
          expectedRowBreakIds: [3, 6],
          expectedColumnBreakIds: [3, 6],
        })
      } finally {
        removeMacosExcelTestDir(tempDir)
      }
    },
    120_000,
  )
})

async function expectStructuralPrintBreakCase(args: {
  readonly tempDir: string
  readonly filePrefix: string
  readonly excelOperation: { readonly kind: 'insertRows' | 'insertColumns'; readonly range: string }
  readonly applyHeadlessEdit: (engine: SpreadsheetEngine) => void
  readonly expectedRowBreakIds: readonly number[]
  readonly expectedColumnBreakIds: readonly number[]
}): Promise<void> {
  const excelWorkbookPath = join(args.tempDir, `excel-${args.filePrefix}-print-page-setup-oracle.xlsx`)
  writeFileSync(excelWorkbookPath, exportXlsx(printPageSetupSnapshot()))

  runMacosExcelStructuralOperationOracle({
    workbookPath: excelWorkbookPath,
    worksheetName: 'Report',
    operations: [args.excelOperation],
    inspectCells: ['A1', 'A2', 'B1'],
    saveWorkbook: true,
    timeoutMs: 90_000,
  })

  const excelTruth = importXlsx(new Uint8Array(readFileSync(excelWorkbookPath)), `excel-${args.filePrefix}-print-page-setup-oracle.xlsx`)
  expect(printBreakIds(excelTruth.snapshot.sheets[0]?.metadata?.printPageSetup)).toEqual({
    row: args.expectedRowBreakIds,
    column: args.expectedColumnBreakIds,
  })

  const headless = new SpreadsheetEngine({ workbookName: `headless-${args.filePrefix}-print-page-setup-oracle` })
  await headless.ready()
  headless.importSnapshot(printPageSetupSnapshot())
  args.applyHeadlessEdit(headless)
  expect(printBreakIds(headless.exportSnapshot().sheets[0]?.metadata?.printPageSetup)).toEqual({
    row: args.expectedRowBreakIds,
    column: args.expectedColumnBreakIds,
  })

  const headlessWorkbookPath = join(args.tempDir, `headless-${args.filePrefix}-print-page-setup-oracle.xlsx`)
  writeFileSync(headlessWorkbookPath, exportXlsx(headless.exportSnapshot()))
  runMacosExcelInspectionOracle({
    workbookPath: headlessWorkbookPath,
    worksheetName: 'Report',
    formulaCells: [],
    inspectCells: ['A1', 'A2', 'B1'],
    saveWorkbook: true,
    timeoutMs: 90_000,
  })

  const headlessExcelTruth = importXlsx(
    new Uint8Array(readFileSync(headlessWorkbookPath)),
    `headless-${args.filePrefix}-print-page-setup-oracle.xlsx`,
  )
  expect(printBreakIds(headlessExcelTruth.snapshot.sheets[0]?.metadata?.printPageSetup)).toEqual({
    row: args.expectedRowBreakIds,
    column: args.expectedColumnBreakIds,
  })
}

function printBreakIds(printSetup: NonNullable<WorkbookSnapshot['sheets'][number]['metadata']>['printPageSetup'] | undefined): {
  readonly row: number[]
  readonly column: number[]
} {
  return {
    row: breakIds(printSetup?.rowBreaksXml),
    column: breakIds(printSetup?.colBreaksXml),
  }
}

function breakIds(xml: string | undefined): number[] {
  if (!xml) {
    throw new Error('Expected break XML')
  }
  return [...xml.matchAll(/<brk\b[^>]*\bid="(\d+)"/gu)].map((match) => Number(match[1]))
}

function printPageSetupSnapshot(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: { name: 'Print page setup oracle' },
    sheets: [
      {
        id: 1,
        name: 'Report',
        order: 0,
        metadata: {
          printPageSetup,
        },
        cells: Array.from({ length: 8 }, (_value, index) => ({
          address: `A${String(index + 1)}`,
          value: `row-${String(index + 1)}`,
        })),
      },
    ],
  }
}
