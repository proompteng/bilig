import { readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'

import { exportXlsx, importXlsx } from '@bilig/excel-import'
import { isMacosExcelInstalled, runMacosExcelInspectionOracle, runMacosExcelStructuralOperationOracle } from '@bilig/excel-fixtures'
import type { WorkbookSnapshot } from '@bilig/protocol'
import { describe, expect, it } from 'vitest'

import { WorkPaper } from '../index.js'
import { createExcelAccessibleTempDir, removeMacosExcelTestDir } from './macos-excel-oracle-test-utils.js'

const cellMetadata = {
  relationshipTarget: 'metadata.xml',
  metadataXml: [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray" xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">',
    '<metadataTypes count="1">',
    '<metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1" cellMeta="1"/>',
    '</metadataTypes>',
    '<futureMetadata name="XLDAPR" count="1"><bk><extLst><ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}"><xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk></futureMetadata>',
    '<cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata>',
    '</metadata>',
  ].join(''),
}

describe('macOS Desktop Excel cell metadata oracle', () => {
  it('preserves and structurally rewrites imported cell metadata refs through WorkPaper export', () => {
    const workpaper = WorkPaper.buildFromSnapshot(importXlsx(buildCellMetadataWorkbookBytes(), 'cell-metadata-source.xlsx').snapshot)
    try {
      const sheet = workpaper.getSheetId('Rich Values')
      if (sheet === undefined) {
        throw new Error('Expected Rich Values sheet to be available')
      }
      workpaper.addRows(sheet, 0, 1)

      const exported = exportXlsx(workpaper.exportSnapshot())
      const reimported = importXlsx(exported, 'cell-metadata-headless-roundtrip.xlsx')

      expect(reimported.snapshot.workbook.metadata?.cellMetadata?.metadataXml).toContain('XLDAPR')
      expect(cellMetadataRefs(reimported.snapshot)).toEqual([
        { address: 'A3', cm: '1' },
        { address: 'B3', cm: '1' },
      ])
    } finally {
      workpaper.dispose()
    }
  })

  it('drops stale cell metadata refs after a headless cell edit changes the imported signature', () => {
    const workpaper = WorkPaper.buildFromSnapshot(importXlsx(buildCellMetadataWorkbookBytes(), 'cell-metadata-source.xlsx').snapshot)
    try {
      const sheet = workpaper.getSheetId('Rich Values')
      if (sheet === undefined) {
        throw new Error('Expected Rich Values sheet to be available')
      }
      workpaper.setCellContents({ sheet, row: 1, col: 0 }, 'CONTOSO')

      const exported = exportXlsx(workpaper.exportSnapshot())
      const reimported = importXlsx(exported, 'cell-metadata-stale-signature-roundtrip.xlsx')

      expect(cellMetadataRefs(reimported.snapshot)).toEqual([{ address: 'B2', cm: '1' }])
    } finally {
      workpaper.dispose()
    }
  })

  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'matches Desktop Excel cell metadata refs after structural row inserts',
    () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = createExcelAccessibleTempDir('bilig-headless-excel-cell-metadata-oracle-')
      try {
        const excelWorkbookPath = join(tempDir, 'excel-cell-metadata-structural-source.xlsx')
        writeFileSync(excelWorkbookPath, buildCellMetadataWorkbookBytes())

        const excelResult = runMacosExcelStructuralOperationOracle({
          workbookPath: excelWorkbookPath,
          worksheetName: 'Rich Values',
          operations: [{ kind: 'insertRows', range: '1:1' }],
          inspectCells: ['A1', 'A3', 'B3'],
          saveWorkbook: true,
          timeoutMs: 90_000,
        })

        const excelTruth = importXlsx(new Uint8Array(readFileSync(excelWorkbookPath)), 'excel-cell-metadata-structural-truth.xlsx')
        const excelRefs = cellMetadataRefs(excelTruth.snapshot)
        expect(excelRefs).toEqual([
          { address: 'A3', cm: '1' },
          { address: 'B3', cm: '1' },
        ])

        const workpaper = WorkPaper.buildFromSnapshot(
          importXlsx(buildCellMetadataWorkbookBytes(), 'headless-cell-metadata-source.xlsx').snapshot,
        )
        try {
          const sheet = workpaper.getSheetId('Rich Values')
          if (sheet === undefined) {
            throw new Error('Expected Rich Values sheet to be available')
          }
          workpaper.addRows(sheet, 0, 1)

          const headlessPath = join(tempDir, 'headless-cell-metadata-structural.xlsx')
          writeFileSync(headlessPath, exportXlsx(workpaper.exportSnapshot()))
          const headlessExcel = runMacosExcelInspectionOracle({
            workbookPath: headlessPath,
            worksheetName: 'Rich Values',
            formulaCells: [],
            inspectCells: ['A1', 'A3', 'B3'],
            saveWorkbook: true,
            timeoutMs: 90_000,
          })
          expect(headlessExcel.cells).toEqual(excelResult.cells)

          const headlessTruth = importXlsx(new Uint8Array(readFileSync(headlessPath)), 'headless-cell-metadata-structural-truth.xlsx')
          expect(cellMetadataRefs(headlessTruth.snapshot)).toEqual(excelRefs)
          expect(headlessTruth.snapshot.workbook.metadata?.cellMetadata?.metadataXml).toContain('XLDAPR')
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

function buildCellMetadataWorkbookBytes(): Uint8Array {
  return exportXlsx(cellMetadataSnapshot())
}

function cellMetadataRefs(snapshot: WorkbookSnapshot): Array<{ readonly address: string; readonly cm?: string; readonly vm?: string }> {
  return (
    snapshot.sheets[0]?.metadata?.cellMetadataRefs?.map((ref) => ({
      address: ref.address,
      ...(ref.cm ? { cm: ref.cm } : {}),
      ...(ref.vm ? { vm: ref.vm } : {}),
    })) ?? []
  )
}

function cellMetadataSnapshot(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: {
      name: 'Cell metadata oracle',
      metadata: { cellMetadata },
    },
    sheets: [
      {
        id: 1,
        name: 'Rich Values',
        order: 0,
        metadata: {
          cellMetadataRefs: [
            {
              address: 'A2',
              cm: '1',
              cellSignature: cellSignature({ value: 'MSFT' }),
            },
            {
              address: 'B2',
              cm: '1',
              cellSignature: cellSignature({ value: 415.32 }),
            },
          ],
        },
        cells: [
          { address: 'A1', value: 'Ticker' },
          { address: 'B1', value: 'Price' },
          { address: 'A2', value: 'MSFT' },
          { address: 'B2', value: 415.32 },
        ],
      },
    ],
  }
}

function cellSignature(cell: { readonly value: string | number }): string {
  return JSON.stringify({
    value: cell.value,
    formula: null,
    format: null,
  })
}
