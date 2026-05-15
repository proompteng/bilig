import { describe, expect, it } from 'vitest'
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'

import type { WorkbookSnapshot } from '@bilig/protocol'
import { exportXlsx, externalPivotCachesWarning, importXlsx } from '../index.js'

describe('xlsx pivot import', () => {
  it('preserves existing pivot table package parts when only worksheet relationships reference them', () => {
    const sourceBytes = buildRelationshipOnlyPivotWorkbook()
    const imported = importXlsx(sourceBytes, 'pivot-relationship-only.xlsx')

    const exportedBytes = exportXlsx(imported.snapshot)

    expect(imported.warnings).toEqual([])
    expect(imported.snapshot.workbook.metadata?.pivots).toEqual([
      expect.objectContaining({
        name: 'SalesByRegion',
        sheetName: 'Pivot',
        address: 'A1',
      }),
    ])
    expect(imported.snapshot.workbook.metadata?.pivotArtifacts?.parts.length).toBeGreaterThan(0)
    expect(pivotPackageMetrics(exportedBytes)).toEqual(pivotPackageMetrics(sourceBytes))
    expect(sheetRelationshipsXml(exportedBytes, 2)).toContain(
      'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable"',
    )
  })

  it('imports pivot data fields that rely on the OOXML default subtotal', () => {
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: {
        name: 'pivot-default-subtotal',
        metadata: {
          pivots: [
            {
              name: 'SalesByRegion',
              sheetName: 'Pivot',
              address: 'A1',
              source: { sheetName: 'Data', startAddress: 'A1', endAddress: 'B4' },
              groupBy: ['Region'],
              values: [{ sourceColumn: 'Sales', summarizeBy: 'sum', outputLabel: 'Sales Total' }],
              rows: 4,
              cols: 2,
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
            { address: 'B1', value: 'Sales' },
            { address: 'A2', value: 'East' },
            { address: 'B2', value: 10 },
            { address: 'A3', value: 'West' },
            { address: 'B3', value: 7 },
            { address: 'A4', value: 'East' },
            { address: 'B4', value: 5 },
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
    const zip = unzipSync(exportXlsx(snapshot))
    const pivotPath = 'xl/pivotTables/pivotTable1.xml'
    const pivotXml = strFromU8(zip[pivotPath] ?? new Uint8Array())

    expect(pivotXml).toContain('subtotal="sum"')
    zip[pivotPath] = strToU8(pivotXml.replace(' subtotal="sum"', ''))

    const imported = importXlsx(zipSync(zip), 'pivot-default-subtotal.xlsx')

    expect(imported.snapshot.workbook.metadata?.pivots).toEqual([
      {
        name: 'SalesByRegion',
        sheetName: 'Pivot',
        address: 'A1',
        source: { sheetName: 'Data', startAddress: 'A1', endAddress: 'B4' },
        groupBy: ['Region'],
        values: [{ sourceColumn: 'Sales', summarizeBy: 'sum', outputLabel: 'Sales Total' }],
        rows: 4,
        cols: 2,
      },
    ])
  })

  it('resolves pivot cache table sources to imported table ranges', () => {
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: {
        name: 'pivot-table-source',
        metadata: {
          tables: [
            {
              name: 'InputTable',
              sheetName: 'Data',
              startAddress: 'A1',
              endAddress: 'B4',
              columnNames: ['Region', 'Sales'],
              headerRow: true,
              totalsRow: false,
            },
          ],
          pivots: [
            {
              name: 'SalesByRegion',
              sheetName: 'Pivot',
              address: 'A1',
              source: { sheetName: 'Data', startAddress: 'A1', endAddress: 'B4' },
              groupBy: ['Region'],
              values: [{ sourceColumn: 'Sales', summarizeBy: 'sum', outputLabel: 'Sales Total' }],
              rows: 4,
              cols: 2,
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
            { address: 'B1', value: 'Sales' },
            { address: 'A2', value: 'East' },
            { address: 'B2', value: 10 },
            { address: 'A3', value: 'West' },
            { address: 'B3', value: 7 },
            { address: 'A4', value: 'East' },
            { address: 'B4', value: 5 },
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
    const zip = unzipSync(exportXlsx(snapshot))
    const cachePath = 'xl/pivotCache/pivotCacheDefinition1.xml'
    const cacheXml = strFromU8(zip[cachePath] ?? new Uint8Array())

    expect(cacheXml).toContain('<worksheetSource ref="A1:B4" sheet="Data"/>')
    zip[cachePath] = strToU8(cacheXml.replace('<worksheetSource ref="A1:B4" sheet="Data"/>', '<worksheetSource name="InputTable"/>'))

    const imported = importXlsx(zipSync(zip), 'pivot-table-source.xlsx')

    expect(imported.snapshot.workbook.metadata?.tables).toEqual([
      {
        name: 'InputTable',
        sheetName: 'Data',
        startAddress: 'A1',
        endAddress: 'B4',
        columnNames: ['Region', 'Sales'],
        headerRow: true,
        totalsRow: false,
      },
    ])
    expect(imported.snapshot.workbook.metadata?.pivots).toEqual([
      {
        name: 'SalesByRegion',
        sheetName: 'Pivot',
        address: 'A1',
        source: { sheetName: 'Data', startAddress: 'A1', endAddress: 'B4' },
        groupBy: ['Region'],
        values: [{ sourceColumn: 'Sales', summarizeBy: 'sum', outputLabel: 'Sales Total' }],
        rows: 4,
        cols: 2,
      },
    ])
  })

  it('resolves pivot cache named-range sources to imported defined-name ranges', () => {
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: {
        name: 'pivot-named-range-source',
        metadata: {
          definedNames: [
            {
              name: 'PivotInput',
              value: { kind: 'range-ref', sheetName: 'Data', startAddress: 'A1', endAddress: 'B4' },
            },
          ],
          pivots: [
            {
              name: 'SalesByRegion',
              sheetName: 'Pivot',
              address: 'A1',
              source: { sheetName: 'Data', startAddress: 'A1', endAddress: 'B4' },
              groupBy: ['Region'],
              values: [{ sourceColumn: 'Sales', summarizeBy: 'sum', outputLabel: 'Sales Total' }],
              rows: 4,
              cols: 2,
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
            { address: 'B1', value: 'Sales' },
            { address: 'A2', value: 'East' },
            { address: 'B2', value: 10 },
            { address: 'A3', value: 'West' },
            { address: 'B3', value: 7 },
            { address: 'A4', value: 'East' },
            { address: 'B4', value: 5 },
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
    const zip = unzipSync(exportXlsx(snapshot))
    const cachePath = 'xl/pivotCache/pivotCacheDefinition1.xml'
    const cacheXml = strFromU8(zip[cachePath] ?? new Uint8Array())

    expect(cacheXml).toContain('<worksheetSource ref="A1:B4" sheet="Data"/>')
    zip[cachePath] = strToU8(cacheXml.replace('<worksheetSource ref="A1:B4" sheet="Data"/>', '<worksheetSource name="PivotInput"/>'))

    const imported = importXlsx(zipSync(zip), 'pivot-named-range-source.xlsx')

    expect(imported.snapshot.workbook.metadata?.definedNames).toEqual([
      {
        name: 'PivotInput',
        value: { kind: 'range-ref', sheetName: 'Data', startAddress: 'A1', endAddress: 'B4' },
      },
    ])
    expect(imported.snapshot.workbook.metadata?.pivots).toEqual([
      {
        name: 'SalesByRegion',
        sheetName: 'Pivot',
        address: 'A1',
        source: { sheetName: 'Data', startAddress: 'A1', endAddress: 'B4' },
        groupBy: ['Region'],
        values: [{ sourceColumn: 'Sales', summarizeBy: 'sum', outputLabel: 'Sales Total' }],
        rows: 4,
        cols: 2,
      },
    ])
  })

  it('resolves pivot cache sheet-scoped named-range sources', () => {
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: {
        name: 'pivot-sheet-scoped-name-source',
        metadata: {
          definedNames: [
            {
              name: 'PivotInput',
              scopeSheetName: 'Data',
              value: { kind: 'range-ref', sheetName: 'Data', startAddress: 'A1', endAddress: 'B4' },
            },
          ],
          pivots: [
            {
              name: 'SalesByRegion',
              sheetName: 'Pivot',
              address: 'A1',
              source: { sheetName: 'Data', startAddress: 'A1', endAddress: 'B4' },
              groupBy: ['Region'],
              values: [{ sourceColumn: 'Sales', summarizeBy: 'sum', outputLabel: 'Sales Total' }],
              rows: 4,
              cols: 2,
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
            { address: 'B1', value: 'Sales' },
            { address: 'A2', value: 'East' },
            { address: 'B2', value: 10 },
            { address: 'A3', value: 'West' },
            { address: 'B3', value: 7 },
            { address: 'A4', value: 'East' },
            { address: 'B4', value: 5 },
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
    const zip = unzipSync(exportXlsx(snapshot))
    const cachePath = 'xl/pivotCache/pivotCacheDefinition1.xml'
    const cacheXml = strFromU8(zip[cachePath] ?? new Uint8Array())

    expect(cacheXml).toContain('<worksheetSource ref="A1:B4" sheet="Data"/>')
    zip[cachePath] = strToU8(
      cacheXml.replace('<worksheetSource ref="A1:B4" sheet="Data"/>', '<worksheetSource name="PivotInput" sheet="Data"/>'),
    )

    const imported = importXlsx(zipSync(zip), 'pivot-sheet-scoped-name-source.xlsx')

    expect(imported.snapshot.workbook.metadata?.definedNames).toEqual([
      {
        name: 'PivotInput',
        scopeSheetName: 'Data',
        value: { kind: 'range-ref', sheetName: 'Data', startAddress: 'A1', endAddress: 'B4' },
      },
    ])
    expect(imported.snapshot.workbook.metadata?.pivots).toEqual([
      {
        name: 'SalesByRegion',
        sheetName: 'Pivot',
        address: 'A1',
        source: { sheetName: 'Data', startAddress: 'A1', endAddress: 'B4' },
        groupBy: ['Region'],
        values: [{ sourceColumn: 'Sales', summarizeBy: 'sum', outputLabel: 'Sales Total' }],
        rows: 4,
        cols: 2,
      },
    ])
  })

  it('warns when XLSX pivots are backed by external caches that cannot be semantically imported', () => {
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: {
        name: 'pivot-external-cache',
        metadata: {
          pivots: [
            {
              name: 'SalesByRegion',
              sheetName: 'Pivot',
              address: 'A1',
              source: { sheetName: 'Data', startAddress: 'A1', endAddress: 'B4' },
              groupBy: ['Region'],
              values: [{ sourceColumn: 'Sales', summarizeBy: 'sum', outputLabel: 'Sales Total' }],
              rows: 4,
              cols: 2,
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
            { address: 'B1', value: 'Sales' },
            { address: 'A2', value: 'East' },
            { address: 'B2', value: 10 },
            { address: 'A3', value: 'West' },
            { address: 'B3', value: 7 },
            { address: 'A4', value: 'East' },
            { address: 'B4', value: 5 },
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
    const zip = unzipSync(exportXlsx(snapshot))
    const cachePath = 'xl/pivotCache/pivotCacheDefinition1.xml'
    const cacheXml = strFromU8(zip[cachePath] ?? new Uint8Array())

    expect(cacheXml).toContain('<cacheSource type="worksheet">')
    zip[cachePath] = strToU8(
      cacheXml.replace(
        '<cacheSource type="worksheet"><worksheetSource ref="A1:B4" sheet="Data"/></cacheSource>',
        '<cacheSource type="external" connectionId="1"/>',
      ),
    )

    const imported = importXlsx(zipSync(zip), 'pivot-external-cache.xlsx')

    expect(imported.warnings).toContain(externalPivotCachesWarning)
    expect(imported.snapshot.workbook.metadata?.pivots).toBeUndefined()
    expect(imported.snapshot.workbook.metadata?.unsupportedPivots).toEqual([
      expect.objectContaining({
        kind: 'external-cache',
        cacheId: 1,
        sourceType: 'external',
        sheetName: 'Pivot',
        address: 'A1',
        name: 'SalesByRegion',
      }),
    ])
  })
})

function buildRelationshipOnlyPivotWorkbook(): Uint8Array {
  const zip = unzipSync(exportXlsx(buildBasicPivotWorkbookSnapshot()))
  for (const path of Object.keys(zip)) {
    if (!/^xl\/worksheets\/sheet\d+\.xml$/u.test(path)) {
      continue
    }
    const sheetXml = strFromU8(zip[path] ?? new Uint8Array())
    zip[path] = strToU8(sheetXml.replace(/<pivotTableDefinition\b[^>]*(?:\/>|>[\s\S]*?<\/pivotTableDefinition>)/gu, ''))
  }
  return zipSync(zip)
}

function buildBasicPivotWorkbookSnapshot(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: {
      name: 'pivot-relationship-only',
      metadata: {
        pivots: [
          {
            name: 'SalesByRegion',
            sheetName: 'Pivot',
            address: 'A1',
            source: { sheetName: 'Data', startAddress: 'A1', endAddress: 'B4' },
            groupBy: ['Region'],
            values: [{ sourceColumn: 'Sales', summarizeBy: 'sum', outputLabel: 'Sales Total' }],
            rows: 4,
            cols: 2,
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
          { address: 'B1', value: 'Sales' },
          { address: 'A2', value: 'East' },
          { address: 'B2', value: 10 },
          { address: 'A3', value: 'West' },
          { address: 'B3', value: 7 },
          { address: 'A4', value: 'East' },
          { address: 'B4', value: 5 },
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

function pivotPackageMetrics(bytes: Uint8Array): {
  pivotTables: number
  pivotCacheDefinitions: number
  pivotCacheRecords: number
  workbookPivotCaches: number
} {
  const zip = unzipSync(bytes)
  const workbookXml = strFromU8(zip['xl/workbook.xml'] ?? new Uint8Array())
  return {
    pivotTables: Object.keys(zip).filter((path) => /^xl\/pivotTables\/pivotTable\d+\.xml$/u.test(path)).length,
    pivotCacheDefinitions: Object.keys(zip).filter((path) => /^xl\/pivotCache\/pivotCacheDefinition\d+\.xml$/u.test(path)).length,
    pivotCacheRecords: Object.keys(zip).filter((path) => /^xl\/pivotCache\/pivotCacheRecords\d+\.xml$/u.test(path)).length,
    workbookPivotCaches: workbookXml.match(/<pivotCache\b/gu)?.length ?? 0,
  }
}

function sheetRelationshipsXml(bytes: Uint8Array, sheetIndex: number): string {
  return strFromU8(unzipSync(bytes)[`xl/worksheets/_rels/sheet${String(sheetIndex)}.xml.rels`] ?? new Uint8Array())
}
