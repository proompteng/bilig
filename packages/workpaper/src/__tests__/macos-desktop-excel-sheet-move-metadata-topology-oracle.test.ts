import { readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'

import { exportXlsx, importXlsx } from '@bilig/excel-import'
import { isMacosExcelInstalled, runMacosExcelInspectionOracle, runMacosExcelStructuralOperationOracle } from '@bilig/excel-fixtures'
import type { WorkbookSnapshot } from '@bilig/protocol'
import { describe, expect, it } from 'vitest'

import { WorkPaper } from '../index.js'
import { createExcelAccessibleTempDir, removeMacosExcelTestDir } from './macos-excel-oracle-test-utils.js'

describe('macOS Desktop Excel sheet move metadata topology oracle', () => {
  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'matches Desktop Excel worksheet metadata ownership after moving a sheet tab',
    () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = createExcelAccessibleTempDir('bilig-headless-excel-sheet-move-metadata-oracle-')
      try {
        const sourceBytes = exportXlsx(sheetMoveMetadataTopologySnapshot())
        const importedSource = importXlsx(sourceBytes, 'sheet-move-metadata-source.xlsx').snapshot
        expect(metadataCodeNames(importedSource)).toEqual(['Data:DataCode', 'Inputs:InputsCode', 'Report:ReportCode'])
        expect(metadataPolicySummary(importedSource)).toEqual([
          'Data:filters=0;sorts=0;validations=0;protection=none;ranges=;cf=0;sparkline=none',
          'Inputs:filters=1;sorts=1;validations=1;protection=formatCells=0;ranges=InputsLock:Inputs!A2:A2;cf=1;sparkline=Inputs!B2:C2',
          'Report:filters=1;sorts=1;validations=1;protection=formatRows=0;ranges=ReportLock:Report!A2:A2;cf=1;sparkline=Report!B2:C2',
        ])
        expect(metadataObjectSummary(importedSource)).toEqual([
          'Data:tables=;comments=;vml=none',
          'Inputs:tables=InputsTable:Inputs!H1:J3;comments=Inputs!A2:Inputs relationship note;vml=legacy',
          'Report:tables=ReportTable:Report!H1:J3;comments=Report!A2:Report relationship note;vml=legacy',
        ])

        const excelWorkbookPath = join(tempDir, 'excel-sheet-move-metadata-source.xlsx')
        writeFileSync(excelWorkbookPath, sourceBytes)
        const excelResult = runMacosExcelStructuralOperationOracle({
          workbookPath: excelWorkbookPath,
          worksheetName: 'Report',
          operations: [{ kind: 'moveSheet', name: 'Report', before: 'Inputs' }],
          inspectCells: ['A1', 'A2'],
          saveWorkbook: true,
          timeoutMs: 120_000,
        })
        expect(excelResult.cells.map((cell) => cell.value)).toEqual([
          { kind: 'string', value: 'Report' },
          { kind: 'number', value: 33 },
        ])

        const excelTruth = importXlsx(new Uint8Array(readFileSync(excelWorkbookPath)), 'excel-sheet-move-metadata-truth.xlsx')
        expect(excelTruth.snapshot.sheets.map((sheet) => sheet.name)).toEqual(['Data', 'Report', 'Inputs'])
        expect(metadataCodeNames(excelTruth.snapshot)).toEqual(['Data:DataCode', 'Report:ReportCode', 'Inputs:InputsCode'])
        expect(metadataPolicySummary(excelTruth.snapshot)).toEqual([
          'Data:filters=0;sorts=0;validations=0;protection=none;ranges=;cf=0;sparkline=none',
          'Report:filters=1;sorts=1;validations=1;protection=formatRows=0;ranges=ReportLock:Report!A2:A2;cf=1;sparkline=Report!B2:C2',
          'Inputs:filters=1;sorts=1;validations=1;protection=formatCells=0;ranges=InputsLock:Inputs!A2:A2;cf=1;sparkline=Inputs!B2:C2',
        ])
        expect(metadataObjectSummary(excelTruth.snapshot)).toEqual([
          'Data:tables=;comments=;vml=none',
          'Report:tables=ReportTable:Report!H1:J3;comments=Report!A2:Report relationship note;vml=legacy',
          'Inputs:tables=InputsTable:Inputs!H1:J3;comments=Inputs!A2:Inputs relationship note;vml=legacy',
        ])

        const workpaper = WorkPaper.buildFromSnapshot(importedSource)
        try {
          const reportSheet = workpaper.getSheetId('Report')
          if (reportSheet === undefined) {
            throw new Error('Expected Report sheet')
          }
          workpaper.moveSheet(reportSheet, 1)

          const headlessSnapshot = workpaper.exportSnapshot()
          expect(headlessSnapshot.sheets.map((sheet) => sheet.name)).toEqual(['Data', 'Report', 'Inputs'])
          expect(metadataCodeNames(headlessSnapshot)).toEqual(metadataCodeNames(excelTruth.snapshot))
          expect(metadataTabColors(headlessSnapshot)).toEqual(metadataTabColors(excelTruth.snapshot))
          expect(metadataPolicySummary(headlessSnapshot)).toEqual(metadataPolicySummary(excelTruth.snapshot))
          expect(metadataObjectSummary(headlessSnapshot)).toEqual(metadataObjectSummary(excelTruth.snapshot))

          const headlessPath = join(tempDir, 'headless-sheet-move-metadata.xlsx')
          writeFileSync(headlessPath, exportXlsx(headlessSnapshot))
          const headlessExcel = runMacosExcelInspectionOracle({
            workbookPath: headlessPath,
            worksheetName: 'Report',
            formulaCells: [],
            inspectCells: ['A1', 'A2'],
            saveWorkbook: true,
            timeoutMs: 90_000,
          })
          expect(headlessExcel.cells).toEqual(excelResult.cells)

          const headlessTruth = importXlsx(new Uint8Array(readFileSync(headlessPath)), 'headless-sheet-move-metadata-truth.xlsx')
          expect(metadataCodeNames(headlessTruth.snapshot)).toEqual(metadataCodeNames(excelTruth.snapshot))
          expect(metadataPolicySummary(headlessTruth.snapshot)).toEqual(metadataPolicySummary(excelTruth.snapshot))
          expect(metadataObjectSummary(headlessTruth.snapshot)).toEqual(metadataObjectSummary(excelTruth.snapshot))
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

function sheetMoveMetadataTopologySnapshot(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: {
      name: 'Desktop Excel sheet move metadata topology oracle',
      metadata: {
        tables: [
          {
            name: 'InputsTable',
            sheetName: 'Inputs',
            startAddress: 'H1',
            endAddress: 'J3',
            columnNames: ['Kind', 'Metric', 'Trend'],
            headerRow: true,
            totalsRow: false,
          },
          {
            name: 'ReportTable',
            sheetName: 'Report',
            startAddress: 'H1',
            endAddress: 'J3',
            columnNames: ['Kind', 'Metric', 'Trend'],
            headerRow: true,
            totalsRow: false,
          },
        ],
      },
    },
    sheets: [
      {
        id: 1,
        name: 'Data',
        order: 0,
        metadata: {
          sheetPr: { xml: '<sheetPr codeName="DataCode"><outlinePr summaryBelow="0"/></sheetPr>' },
          tabColor: { rgb: 'FF808080' },
        },
        cells: [
          { address: 'A1', value: 'Data' },
          { address: 'A2', value: 11 },
        ],
      },
      {
        id: 2,
        name: 'Inputs',
        order: 1,
        metadata: {
          sheetPr: { xml: '<sheetPr codeName="InputsCode"><outlinePr summaryRight="0"/></sheetPr>' },
          tabColor: { rgb: 'FFFF0000' },
          filters: [
            {
              sheetName: 'Inputs',
              startAddress: 'A1',
              endAddress: 'B3',
              criteria: [{ colId: 0, filters: { values: ['Inputs'] } }],
            },
          ],
          sorts: [
            {
              range: { sheetName: 'Inputs', startAddress: 'A1', endAddress: 'B3' },
              keys: [{ keyAddress: 'B2', direction: 'asc' }],
            },
          ],
          validations: [
            {
              range: { sheetName: 'Inputs', startAddress: 'E2', endAddress: 'E2' },
              rule: { kind: 'list', values: ['red', 'blue'] },
            },
          ],
          sheetProtection: {
            sheetName: 'Inputs',
            xmlAttributes: [
              { name: 'sheet', value: '1' },
              { name: 'formatCells', value: '0' },
            ],
          },
          protectedRanges: [{ id: 'InputsLock', range: { sheetName: 'Inputs', startAddress: 'A2', endAddress: 'A2' } }],
          conditionalFormats: [
            {
              id: 'inputs-cf',
              range: { sheetName: 'Inputs', startAddress: 'B2', endAddress: 'B2' },
              rule: { kind: 'cellIs', operator: 'greaterThan', values: [20] },
              style: {},
            },
          ],
          commentThreads: [
            {
              threadId: 'inputs-note',
              sheetName: 'Inputs',
              address: 'A2',
              comments: [
                {
                  id: 'inputs-note-1',
                  authorDisplayName: 'InputsAudit',
                  body: 'Inputs relationship note',
                },
              ],
            },
          ],
          sparklines: { xml: sparklineExtensionXml('Inputs!B2:C2', 'F2') },
        },
        cells: [
          { address: 'A1', value: 'Inputs' },
          { address: 'A2', value: 22 },
          { address: 'B1', value: 'Metric' },
          { address: 'B2', value: 24 },
          { address: 'C2', value: 25 },
          { address: 'E2', value: 'red' },
          { address: 'F2', value: '' },
          { address: 'H1', value: 'Kind' },
          { address: 'I1', value: 'Metric' },
          { address: 'J1', value: 'Trend' },
          { address: 'H2', value: 'Inputs' },
          { address: 'I2', value: 24 },
          { address: 'J2', value: 25 },
          { address: 'H3', value: 'Inputs' },
          { address: 'I3', value: 28 },
          { address: 'J3', value: 29 },
        ],
      },
      {
        id: 3,
        name: 'Report',
        order: 2,
        metadata: {
          sheetPr: { xml: '<sheetPr codeName="ReportCode"><pageSetUpPr fitToPage="1"/></sheetPr>' },
          tabColor: { rgb: 'FF00AA00' },
          filters: [
            {
              sheetName: 'Report',
              startAddress: 'A1',
              endAddress: 'B3',
              criteria: [{ colId: 0, filters: { values: ['Report'] } }],
            },
          ],
          sorts: [
            {
              range: { sheetName: 'Report', startAddress: 'A1', endAddress: 'B3' },
              keys: [{ keyAddress: 'B2', direction: 'desc' }],
            },
          ],
          validations: [
            {
              range: { sheetName: 'Report', startAddress: 'E2', endAddress: 'E2' },
              rule: { kind: 'whole', operator: 'greaterThan', values: [30] },
            },
          ],
          sheetProtection: {
            sheetName: 'Report',
            xmlAttributes: [
              { name: 'sheet', value: '1' },
              { name: 'formatRows', value: '0' },
            ],
          },
          protectedRanges: [{ id: 'ReportLock', range: { sheetName: 'Report', startAddress: 'A2', endAddress: 'A2' } }],
          conditionalFormats: [
            {
              id: 'report-cf',
              range: { sheetName: 'Report', startAddress: 'B2', endAddress: 'B2' },
              rule: { kind: 'cellIs', operator: 'greaterThan', values: [30] },
              style: {},
            },
          ],
          commentThreads: [
            {
              threadId: 'report-note',
              sheetName: 'Report',
              address: 'A2',
              comments: [
                {
                  id: 'report-note-1',
                  authorDisplayName: 'ReportAudit',
                  body: 'Report relationship note',
                },
              ],
            },
          ],
          sparklines: { xml: sparklineExtensionXml('Report!B2:C2', 'F2') },
        },
        cells: [
          { address: 'A1', value: 'Report' },
          { address: 'A2', value: 33 },
          { address: 'B1', value: 'Metric' },
          { address: 'B2', value: 34 },
          { address: 'C2', value: 38 },
          { address: 'E2', value: 31 },
          { address: 'F2', value: '' },
          { address: 'H1', value: 'Kind' },
          { address: 'I1', value: 'Metric' },
          { address: 'J1', value: 'Trend' },
          { address: 'H2', value: 'Report' },
          { address: 'I2', value: 34 },
          { address: 'J2', value: 38 },
          { address: 'H3', value: 'Report' },
          { address: 'I3', value: 35 },
          { address: 'J3', value: 39 },
        ],
      },
    ],
  }
}

const sparklineExtensionUri = '{05C60535-1F16-4fd2-B633-F4F36F0B64E0}'
const x14Namespace = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main'
const xmNamespace = 'http://schemas.microsoft.com/office/excel/2006/main'

function sparklineExtensionXml(formula: string, sqref: string): string {
  return `<ext uri="${sparklineExtensionUri}" xmlns:x14="${x14Namespace}"><x14:sparklineGroups xmlns:xm="${xmNamespace}"><x14:sparklineGroup type="line"><x14:sparklines><x14:sparkline><xm:f>${formula}</xm:f><xm:sqref>${sqref}</xm:sqref></x14:sparkline></x14:sparklines></x14:sparklineGroup></x14:sparklineGroups></ext>`
}

function metadataCodeNames(snapshot: WorkbookSnapshot): string[] {
  return snapshot.sheets.map((sheet) => `${sheet.name}:${codeName(sheet.metadata?.sheetPr?.xml) ?? ''}`)
}

function metadataTabColors(snapshot: WorkbookSnapshot): string[] {
  return snapshot.sheets.map((sheet) => `${sheet.name}:${JSON.stringify(sheet.metadata?.tabColor ?? null)}`)
}

function metadataPolicySummary(snapshot: WorkbookSnapshot): string[] {
  return snapshot.sheets.map((sheet) => {
    const metadata = sheet.metadata
    const protection = metadata?.sheetProtection?.xmlAttributes
      ?.filter((attribute) => attribute.name !== 'sheet')
      .map((attribute) => `${attribute.name}=${attribute.value}`)
      .join(',')
    const ranges = (metadata?.protectedRanges ?? [])
      .map((range) => `${range.id}:${range.range.sheetName}!${range.range.startAddress}:${range.range.endAddress}`)
      .join(',')
    const sparkline = /<xm:f>([\s\S]*?)<\/xm:f>/u.exec(metadata?.sparklines?.xml ?? '')?.[1] ?? 'none'
    return [
      `${sheet.name}:filters=${String(metadata?.filters?.length ?? 0)}`,
      `sorts=${String(metadata?.sorts?.length ?? 0)}`,
      `validations=${String(metadata?.validations?.length ?? 0)}`,
      `protection=${protection && protection.length > 0 ? protection : 'none'}`,
      `ranges=${ranges}`,
      `cf=${String(metadata?.conditionalFormats?.length ?? 0)}`,
      `sparkline=${sparkline}`,
    ].join(';')
  })
}

function metadataObjectSummary(snapshot: WorkbookSnapshot): string[] {
  const tables = snapshot.workbook.metadata?.tables ?? []
  return snapshot.sheets.map((sheet) => {
    const sheetTables = tables
      .filter((table) => table.sheetName === sheet.name)
      .map((table) => `${table.name}:${table.sheetName}!${table.startAddress}:${table.endAddress}`)
      .join(',')
    const comments = (sheet.metadata?.commentThreads ?? [])
      .map((thread) => `${thread.sheetName}!${thread.address}:${thread.comments.map((comment) => comment.body).join('|')}`)
      .join(',')
    return `${sheet.name}:tables=${sheetTables};comments=${comments};vml=${sheet.metadata?.legacyCommentVml ? 'legacy' : 'none'}`
  })
}

function codeName(sheetPrXml: string | undefined): string | undefined {
  return /\bcodeName="([^"]+)"/u.exec(sheetPrXml ?? '')?.[1]
}
