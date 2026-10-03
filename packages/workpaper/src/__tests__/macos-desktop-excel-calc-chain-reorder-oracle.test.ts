import { readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'

import { exportXlsx, importXlsx } from '@bilig/excel-import'
import { isMacosExcelInstalled, runMacosExcelStructuralOperationOracle } from '@bilig/excel-fixtures'
import type { WorkbookSnapshot } from '@bilig/protocol'
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { describe, expect, it } from 'vitest'

import { WorkPaper } from '../index.js'
import { createExcelAccessibleTempDir, removeMacosExcelTestDir } from './macos-excel-oracle-test-utils.js'

describe('macOS Desktop Excel calc-chain reorder oracle', () => {
  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'matches Desktop Excel calc-chain sheet ids after moving a sheet tab',
    () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = createExcelAccessibleTempDir('bilig-headless-excel-calc-chain-move-oracle-')
      try {
        const sourceBytes = buildCalcChainSheetMoveXlsx()
        const importedSource = importXlsx(sourceBytes, 'calc-chain-move-source.xlsx').snapshot
        expect(importedSource.sheets.map((sheet) => sheet.name)).toEqual(['Data', 'Inputs', 'Report'])
        expect(calcChainSummary(importedSource)).toEqual(['1:Data:A1', '2:Inputs:A1', '3:Report:A1'])

        const excelWorkbookPath = join(tempDir, 'excel-calc-chain-move-source.xlsx')
        writeFileSync(excelWorkbookPath, sourceBytes)
        const excelResult = runMacosExcelStructuralOperationOracle({
          workbookPath: excelWorkbookPath,
          worksheetName: 'Report',
          operations: [{ kind: 'moveSheet', name: 'Report', before: 'Inputs' }],
          inspectCells: ['A1'],
          saveWorkbook: true,
          timeoutMs: 120_000,
        })
        expect(excelResult.cells[0]?.value).toEqual({ kind: 'number', value: 12 })

        const excelTruth = importXlsx(new Uint8Array(readFileSync(excelWorkbookPath)), 'excel-calc-chain-move-truth.xlsx')
        expect(excelTruth.snapshot.sheets.map((sheet) => sheet.name)).toEqual(['Data', 'Report', 'Inputs'])
        const excelCalcChain = calcChainSummary(excelTruth.snapshot)
        expect(excelCalcChain).toEqual(['1:Data:A1', '2:Inputs:A1', '3:Report:A1'])

        const workpaper = WorkPaper.buildFromSnapshot(importedSource)
        try {
          const reportSheet = workpaper.getSheetId('Report')
          if (reportSheet === undefined) {
            throw new Error('Expected Report sheet')
          }
          workpaper.moveSheet(reportSheet, 1)

          const headlessSnapshot = workpaper.exportSnapshot()
          expect(headlessSnapshot.sheets.map((sheet) => sheet.name)).toEqual(['Data', 'Report', 'Inputs'])
          expect(calcChainSummary(headlessSnapshot)).toEqual(excelCalcChain)
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

function calcChainSheetMoveSnapshot(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: {
      name: 'Desktop Excel calc chain move oracle',
    },
    sheets: [
      {
        id: 1,
        name: 'Data',
        order: 0,
        cells: [{ address: 'A1', formula: '1+1', value: 2 }],
      },
      {
        id: 2,
        name: 'Inputs',
        order: 1,
        cells: [{ address: 'A1', formula: '10+1', value: 11 }],
      },
      {
        id: 3,
        name: 'Report',
        order: 2,
        cells: [{ address: 'A1', formula: 'Inputs!A1+1', value: 12 }],
      },
    ],
  }
}

function buildCalcChainSheetMoveXlsx(): Uint8Array {
  const zip = unzipSync(exportXlsx(calcChainSheetMoveSnapshot()))
  zip['xl/calcChain.xml'] = strToU8(
    [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
      '<c r="A1" i="1"/>',
      '<c r="A1" i="2"/>',
      '<c r="A1" i="3"/>',
      '</calcChain>',
    ].join(''),
  )
  addCalcChainWorkbookRelationship(zip)
  addCalcChainContentType(zip)
  return zipSync(zip)
}

function addCalcChainWorkbookRelationship(zip: Record<string, Uint8Array>): void {
  const path = 'xl/_rels/workbook.xml.rels'
  const xml = strFromU8(zip[path] ?? new Uint8Array())
  if (xml.includes('/calcChain')) {
    return
  }
  zip[path] = strToU8(
    xml.replace(
      '</Relationships>',
      '<Relationship Id="rIdCalcChain1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/></Relationships>',
    ),
  )
}

function addCalcChainContentType(zip: Record<string, Uint8Array>): void {
  const path = '[Content_Types].xml'
  const xml = strFromU8(zip[path] ?? new Uint8Array())
  if (xml.includes('calcChain+xml')) {
    return
  }
  zip[path] = strToU8(
    xml.replace(
      '</Types>',
      '<Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/></Types>',
    ),
  )
}

function calcChainSummary(snapshot: WorkbookSnapshot): string[] {
  return (snapshot.workbook.metadata?.formulaAudit?.calcChain?.cells ?? []).map(
    (cell) => `${String(cell.sheetIndex)}:${cell.sheetName ?? ''}:${cell.address}`,
  )
}
