import { readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'

import { SpreadsheetEngine } from '@bilig/core'
import { exportXlsx, importXlsx } from '@bilig/excel-import'
import { isMacosExcelInstalled, runMacosExcelInspectionOracle, runMacosExcelStructuralOperationOracle } from '@bilig/excel-fixtures'
import { ValueTag, type WorkbookSnapshot } from '@bilig/protocol'
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { describe, expect, it } from 'vitest'

import { WorkPaper } from '../index.js'
import { createExcelAccessibleTempDir, removeMacosExcelTestDir } from './macos-excel-oracle-test-utils.js'

const expectedEastSales = { tag: ValueTag.Number, value: 25 }
const expectedExcelEastSales = { kind: 'number', value: 25 }

describe('macOS Desktop Excel pivot oracle', () => {
  it('refreshes semantic source-backed pivots and exports Desktop Excel-refreshable caches', async () => {
    const engine = await buildEditedPivotEngine()
    const exportedSnapshot = engine.exportSnapshot()

    expect(engine.getCellValue('Pivot', 'E2')).toEqual(expectedEastSales)
    expect(exportedSnapshot.workbook.metadata?.pivots?.[0]).toMatchObject({
      cacheFields: ['Region', 'Notes', 'Product', 'Sales'],
      cachedRecords: [
        ['East', 'priority', 'Widget', 20],
        ['West', 'priority', 'Widget', 7],
        ['East', 'priority', 'Gizmo', 5],
      ],
      rows: 3,
      cols: 2,
    })

    const exportedBytes = exportXlsx(exportedSnapshot)
    const exportedZip = unzipSync(exportedBytes)
    const cacheDefinitionXml = strFromU8(exportedZip['xl/pivotCache/pivotCacheDefinition1.xml'] ?? new Uint8Array())
    expect(cacheDefinitionXml).toContain('invalid="1"')
    expect(cacheDefinitionXml).toContain('refreshOnLoad="1"')
    expect(cacheDefinitionXml).toContain('recordCount="3"')
    const cacheRecordsXml = strFromU8(exportedZip['xl/pivotCache/pivotCacheRecords1.xml'] ?? new Uint8Array())
    expect(cacheRecordsXml).toContain('<pivotCacheRecords')
    expect(cacheRecordsXml).toContain('count="3"')
    expect(cacheRecordsXml).toContain('<r><s v="East"/><s v="priority"/><s v="Widget"/><n v="20"/></r>')

    const roundTrip = importXlsx(exportedBytes, 'headless-pivot-refresh-roundtrip.xlsx')
    expect(roundTrip.snapshot.workbook.metadata?.pivots?.[0]).toMatchObject({
      cacheFields: ['Region', 'Notes', 'Product', 'Sales'],
      cachedRecords: [
        ['East', 'priority', 'Widget', 20],
        ['West', 'priority', 'Widget', 7],
        ['East', 'priority', 'Gizmo', 5],
      ],
      source: { sheetName: 'Data', startAddress: 'A1', endAddress: 'D4' },
    })
  })

  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'matches Desktop Excel GETPIVOTDATA after headless source edits and XLSX export',
    async () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const engine = await buildEditedPivotEngine()
      const tempDir = createExcelAccessibleTempDir('bilig-headless-excel-pivot-oracle-')
      try {
        const workbookPath = join(tempDir, 'headless-pivot-refresh-oracle.xlsx')
        writeFileSync(workbookPath, exportXlsx(engine.exportSnapshot()))

        const excel = runMacosExcelInspectionOracle({
          workbookPath,
          worksheetName: 'Pivot',
          formulaCells: [],
          inspectCells: ['E2'],
          refreshWorkbook: true,
          saveWorkbook: true,
          timeoutMs: 120_000,
        })
        expect(excel.cells.map((cell) => cell.value)).toEqual([expectedExcelEastSales])

        const saved = importXlsx(new Uint8Array(readFileSync(workbookPath)), 'excel-saved-pivot-refresh-oracle.xlsx')
        expect(saved.snapshot.workbook.metadata?.pivots?.[0]?.cachedRecords).toContainEqual(['East', 'priority', 'Widget', 20])
      } finally {
        removeMacosExcelTestDir(tempDir)
      }
    },
    120_000,
  )

  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'matches Desktop Excel pivot package cleanup after deleting a pivot sheet',
    () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = createExcelAccessibleTempDir('bilig-headless-excel-pivot-delete-oracle-')
      try {
        const sourceBytes = sourceBackedPivotWorkbookBytesWithCacheRecordsSidecar()
        const importedSource = importXlsx(sourceBytes, 'pivot-delete-source.xlsx').snapshot
        expect(pivotPackageSummary(importedSource)).toEqual({
          pivotParts: [
            'xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels',
            'xl/pivotCache/pivotCacheDefinition1.xml',
            'xl/pivotCache/pivotCacheRecords1.xml',
            'xl/pivotTables/pivotTable1.xml',
          ],
          workbookCacheIds: ['1'],
          workbookRelationshipTargets: ['pivotCache/pivotCacheDefinition1.xml'],
          pivotSheets: ['Pivot'],
          semanticPivotSheets: ['Pivot'],
        })

        const excelWorkbookPath = join(tempDir, 'excel-pivot-delete-source.xlsx')
        writeFileSync(excelWorkbookPath, sourceBytes)
        runMacosExcelStructuralOperationOracle({
          workbookPath: excelWorkbookPath,
          worksheetName: 'Data',
          operations: [{ kind: 'deleteSheet', name: 'Pivot' }],
          inspectCells: ['A1'],
          saveWorkbook: true,
          timeoutMs: 120_000,
        })

        const excelTruth = importXlsx(new Uint8Array(readFileSync(excelWorkbookPath)), 'excel-pivot-delete-truth.xlsx')
        expect(excelTruth.snapshot.sheets.map((sheet) => sheet.name)).toEqual(['Data'])
        const excelSummary = pivotPackageSummary(excelTruth.snapshot)
        expect(excelSummary).toEqual({
          pivotParts: [],
          workbookCacheIds: [],
          workbookRelationshipTargets: [],
          pivotSheets: [],
          semanticPivotSheets: [],
        })

        const workpaper = WorkPaper.buildFromSnapshot(importedSource)
        try {
          const pivotSheet = workpaper.getSheetId('Pivot')
          if (pivotSheet === undefined) {
            throw new Error('Expected Pivot sheet')
          }
          workpaper.removeSheet(pivotSheet)

          const headlessSnapshot = workpaper.exportSnapshot()
          expect(headlessSnapshot.sheets.map((sheet) => sheet.name)).toEqual(['Data'])
          expect(pivotPackageSummary(headlessSnapshot)).toEqual(excelSummary)

          const headlessWorkbookPath = join(tempDir, 'headless-pivot-delete.xlsx')
          writeFileSync(headlessWorkbookPath, exportXlsx(headlessSnapshot))
          runMacosExcelInspectionOracle({
            workbookPath: headlessWorkbookPath,
            worksheetName: 'Data',
            formulaCells: [],
            inspectCells: ['A1'],
            saveWorkbook: true,
            timeoutMs: 120_000,
          })
          const headlessExcelTruth = importXlsx(
            new Uint8Array(readFileSync(headlessWorkbookPath)),
            'headless-excel-pivot-delete-truth.xlsx',
          )
          expect(pivotPackageSummary(headlessExcelTruth.snapshot)).toEqual(excelSummary)
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

async function buildEditedPivotEngine(): Promise<SpreadsheetEngine> {
  const engine = new SpreadsheetEngine({ workbookName: 'headless-pivot-refresh-oracle' })
  await engine.ready()
  engine.importSnapshot(importedSourceBackedPivotSnapshot())
  engine.setCellValue('Data', 'D2', 20)
  engine.setCellFormula('Pivot', 'E2', 'GETPIVOTDATA("Sum of Sales",$B$2,"Region","East")')
  return engine
}

function importedSourceBackedPivotSnapshot(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: {
      name: 'Headless pivot refresh oracle',
      metadata: {
        pivots: [
          {
            name: 'SalesByRegion',
            sheetName: 'Pivot',
            address: 'B2',
            source: { sheetName: 'Data', startAddress: 'A1', endAddress: 'D4' },
            sourceKind: 'worksheet',
            groupBy: ['Region'],
            values: [{ sourceColumn: 'Sales', summarizeBy: 'sum' }],
            cacheFields: ['Region', 'Notes', 'Product', 'Sales'],
            cachedRecords: [
              ['East', 'priority', 'Widget', 1],
              ['West', 'priority', 'Widget', 2],
            ],
            rows: 1,
            cols: 1,
          },
        ],
      },
    },
    sheets: [
      {
        id: 1,
        name: 'Data',
        order: 0,
        cells: [
          { address: 'A1', value: 'Region' },
          { address: 'B1', value: 'Notes' },
          { address: 'C1', value: 'Product' },
          { address: 'D1', value: 'Sales' },
          { address: 'A2', value: 'East' },
          { address: 'B2', value: 'priority' },
          { address: 'C2', value: 'Widget' },
          { address: 'D2', value: 10 },
          { address: 'A3', value: 'West' },
          { address: 'B3', value: 'priority' },
          { address: 'C3', value: 'Widget' },
          { address: 'D3', value: 7 },
          { address: 'A4', value: 'East' },
          { address: 'B4', value: 'priority' },
          { address: 'C4', value: 'Gizmo' },
          { address: 'D4', value: 5 },
        ],
      },
      {
        id: 2,
        name: 'Pivot',
        order: 1,
        cells: [],
      },
    ],
  }
}

function sourceBackedPivotWorkbookBytesWithCacheRecordsSidecar(): Uint8Array {
  const zip = unzipSync(exportXlsx(importedSourceBackedPivotSnapshot()))
  const cacheDefinitionPath = 'xl/pivotCache/pivotCacheDefinition1.xml'
  const cacheDefinitionBytes = zip[cacheDefinitionPath]
  const contentTypesBytes = zip['[Content_Types].xml']
  if (!cacheDefinitionBytes || !contentTypesBytes) {
    throw new Error('Expected generated pivot cache package parts')
  }

  zip[cacheDefinitionPath] = strToU8(rewritePivotCacheRecordCount(strFromU8(cacheDefinitionBytes), 3))
  zip['xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels'] = strToU8(
    [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
      '<Relationship Id="rIdRecords1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords" Target="pivotCacheRecords1.xml"/>',
      '</Relationships>',
    ].join(''),
  )
  zip['xl/pivotCache/pivotCacheRecords1.xml'] = strToU8(
    [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="3">',
      '<r><s v="East"/><s v="priority"/><s v="Widget"/><n v="10"/></r>',
      '<r><s v="West"/><s v="priority"/><s v="Widget"/><n v="7"/></r>',
      '<r><s v="East"/><s v="priority"/><s v="Gizmo"/><n v="5"/></r>',
      '</pivotCacheRecords>',
    ].join(''),
  )
  zip['[Content_Types].xml'] = strToU8(addPivotCacheRecordsContentType(strFromU8(contentTypesBytes), 1))
  return zipSync(zip)
}

function rewritePivotCacheRecordCount(xml: string, recordCount: number): string {
  const withoutRecordCount = xml.replace(/\srecordCount=(["'])\d+\1/u, '')
  return withoutRecordCount.replace(/<((?:[A-Za-z_][\w.-]*:)?pivotCacheDefinition)\b/u, `<$1 recordCount="${String(recordCount)}"`)
}

function addPivotCacheRecordsContentType(xml: string, cacheNumber: number): string {
  const partName = `/xl/pivotCache/pivotCacheRecords${String(cacheNumber)}.xml`
  if (xml.includes(`PartName="${partName}"`)) {
    return xml
  }
  return xml.replace(
    '</Types>',
    `<Override PartName="${partName}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml"/></Types>`,
  )
}

function pivotPackageSummary(snapshot: WorkbookSnapshot): {
  readonly pivotParts: readonly string[]
  readonly workbookCacheIds: readonly string[]
  readonly workbookRelationshipTargets: readonly string[]
  readonly pivotSheets: readonly string[]
  readonly semanticPivotSheets: readonly string[]
} {
  const metadata = snapshot.workbook.metadata
  const workbookPivotCachesXml = metadata?.pivotArtifacts?.workbookPivotCachesXml ?? ''
  return {
    pivotParts: metadata?.pivotArtifacts?.parts.map((part) => part.path).toSorted() ?? [],
    workbookCacheIds: [...workbookPivotCachesXml.matchAll(/<pivotCache\b[^>]*\bcacheId="([^"]+)"/gu)]
      .map((match) => match[1] ?? '')
      .filter(Boolean),
    workbookRelationshipTargets:
      metadata?.pivotArtifacts?.workbookRelationships?.map((relationship) => relationship.target).toSorted() ?? [],
    pivotSheets: snapshot.sheets.flatMap((sheet) => (sheet.metadata?.pivotArtifacts ? [sheet.name] : [])).toSorted(),
    semanticPivotSheets: metadata?.pivots?.map((pivot) => pivot.sheetName).toSorted() ?? [],
  }
}
