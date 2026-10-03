import { readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'

import { SpreadsheetEngine } from '@bilig/core'
import { exportXlsx, importXlsx } from '@bilig/excel-import'
import { isMacosExcelInstalled, runMacosExcelPackageOpenSaveOracle, runMacosExcelStructuralOperationOracle } from '@bilig/excel-fixtures'
import type { WorkbookSnapshot } from '@bilig/protocol'
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { describe, expect, it } from 'vitest'

import { WorkPaper } from '../index.js'
import { createExcelAccessibleTempDir, removeMacosExcelTestDir } from './macos-excel-oracle-test-utils.js'

const officeRelationshipNamespace = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const packageRelationshipNamespace = 'http://schemas.openxmlformats.org/package/2006/relationships'
const queryTableRelationshipType = `${officeRelationshipNamespace}/queryTable`
const connectionsRelationshipType = `${officeRelationshipNamespace}/connections`
const queryTableContentType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml'
const connectionsContentType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml'

describe('macOS Desktop Excel query-table package oracle', () => {
  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'preserves Desktop Excel legacy text query-table topology after open/save and headless export',
    () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = createExcelAccessibleTempDir('bilig-headless-excel-query-table-oracle-')
      try {
        const csvPath = join(tempDir, 'bilig-query-table-revenue.csv')
        writeFileSync(csvPath, 'Region,Amount\nNorth,1200\nSouth,900\n')

        const sourcePath = join(tempDir, 'excel-query-table-source.xlsx')
        const sourceBytes = buildLegacyTextQueryTableSourceXlsx(csvPath)
        writeFileSync(sourcePath, sourceBytes)
        expect(queryTableTopology(sourceBytes)).toMatchObject({
          workbookConnectionRelationships: 1,
          worksheetQueryTableRelationships: 1,
          queryTableConnectionIds: ['1'],
        })

        const excelSource = runMacosExcelPackageOpenSaveOracle({
          workbookPath: sourcePath,
          saveWorkbook: true,
          timeoutMs: 120_000,
        })
        expect(excelSource.excelVersion).toMatch(/^\d+\./u)

        const excelSourceBytes = new Uint8Array(readFileSync(sourcePath))
        const excelSourceTopology = queryTableTopology(excelSourceBytes)
        expect(excelSourceTopology).toEqual(queryTableTopology(sourceBytes))

        const imported = importXlsx(excelSourceBytes, 'excel-query-table-source.xlsx').snapshot
        expect(imported.workbook.metadata?.slicerConnectionArtifacts?.sheetArtifacts).toEqual([
          {
            sheetName: 'Revenue',
            relationships: [
              {
                id: 'rIdQueryTable1',
                type: queryTableRelationshipType,
                target: '../queryTables/queryTable1.xml',
              },
            ],
          },
        ])

        const workpaper = WorkPaper.buildFromSnapshot(imported)
        try {
          const sheet = workpaper.getSheetId('Revenue')
          if (sheet === undefined) {
            throw new Error('Expected Revenue sheet')
          }
          workpaper.setCellContents({ sheet, row: 0, col: 2 }, 'headless reviewed')

          const headlessPath = join(tempDir, 'headless-query-table.xlsx')
          const headlessBytes = exportXlsx(workpaper.exportSnapshot())
          writeFileSync(headlessPath, headlessBytes)
          expect(queryTableTopology(headlessBytes)).toEqual(excelSourceTopology)

          const excelHeadless = runMacosExcelPackageOpenSaveOracle({
            workbookPath: headlessPath,
            saveWorkbook: true,
            timeoutMs: 120_000,
          })
          expect(excelHeadless.excelVersion).toMatch(/^\d+\./u)

          const excelSavedHeadlessBytes = new Uint8Array(readFileSync(headlessPath))
          expect(queryTableTopology(excelSavedHeadlessBytes)).toEqual(excelSourceTopology)
          expect(importXlsx(excelSavedHeadlessBytes, 'excel-saved-headless-query-table.xlsx').snapshot.sheets[0]?.cells).toContainEqual(
            expect.objectContaining({
              address: 'C1',
              value: 'headless reviewed',
            }),
          )
        } finally {
          workpaper.dispose()
        }
      } finally {
        removeMacosExcelTestDir(tempDir)
      }
    },
    180_000,
  )

  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'keeps table-owned query-table relationships attached to their ListObject after headless export',
    () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = createExcelAccessibleTempDir('bilig-headless-excel-table-query-oracle-')
      try {
        const csvPath = join(tempDir, 'bilig-table-query-revenue.csv')
        writeFileSync(csvPath, 'Region,Amount\nNorth,1200\nSouth,900\n')

        const sourcePath = join(tempDir, 'excel-table-query-source.xlsx')
        const sourceBytes = buildTableOwnedQueryTableSourceXlsx(csvPath)
        writeFileSync(sourcePath, sourceBytes)
        expect(queryTableTopology(sourceBytes).tableQueryOwnerNames).toEqual(['ARevenueQuery'])

        const excelSource = runMacosExcelPackageOpenSaveOracle({
          workbookPath: sourcePath,
          saveWorkbook: true,
          timeoutMs: 120_000,
        })
        expect(excelSource.excelVersion).toMatch(/^\d+\./u)

        const excelSourceBytes = new Uint8Array(readFileSync(sourcePath))
        expect(queryTableTopology(excelSourceBytes).tableQueryOwnerNames).toEqual(['ARevenueQuery'])

        const imported = importXlsx(excelSourceBytes, 'excel-table-query-source.xlsx').snapshot
        expect(imported.workbook.metadata?.slicerConnectionArtifacts?.tableArtifacts).toEqual([
          {
            tableName: 'ARevenueQuery',
            sheetName: 'Revenue',
            relationshipPartPath: 'xl/tables/_rels/table2.xml.rels',
            relationships: [
              {
                id: 'rIdQueryTable1',
                type: queryTableRelationshipType,
                target: '../queryTables/queryTable1.xml',
              },
            ],
          },
        ])

        const workpaper = WorkPaper.buildFromSnapshot(imported)
        try {
          const sheet = workpaper.getSheetId('Revenue')
          if (sheet === undefined) {
            throw new Error('Expected Revenue sheet')
          }
          workpaper.setCellContents({ sheet, row: 0, col: 6 }, 'headless reviewed')

          const headlessPath = join(tempDir, 'headless-table-query.xlsx')
          const headlessBytes = exportXlsx(workpaper.exportSnapshot())
          writeFileSync(headlessPath, headlessBytes)
          expect(queryTableTopology(headlessBytes).tableQueryOwnerNames).toEqual(['ARevenueQuery'])

          const excelHeadless = runMacosExcelPackageOpenSaveOracle({
            workbookPath: headlessPath,
            saveWorkbook: true,
            timeoutMs: 120_000,
          })
          expect(excelHeadless.excelVersion).toMatch(/^\d+\./u)

          const excelSavedHeadlessBytes = new Uint8Array(readFileSync(headlessPath))
          expect(queryTableTopology(excelSavedHeadlessBytes).tableQueryOwnerNames).toEqual(['ARevenueQuery'])
          expect(importXlsx(excelSavedHeadlessBytes, 'excel-saved-headless-table-query.xlsx').snapshot.sheets[0]?.cells).toContainEqual(
            expect.objectContaining({
              address: 'G1',
              value: 'headless reviewed',
            }),
          )
        } finally {
          workpaper.dispose()
        }
      } finally {
        removeMacosExcelTestDir(tempDir)
      }
    },
    180_000,
  )

  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'matches Desktop Excel when deleting a ListObject with a table-owned query-table relationship',
    async () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = createExcelAccessibleTempDir('bilig-headless-excel-table-query-delete-oracle-')
      try {
        const csvPath = join(tempDir, 'bilig-table-query-delete-revenue.csv')
        writeFileSync(csvPath, 'Region,Amount\nNorth,1200\nSouth,900\n')

        const sourcePath = join(tempDir, 'excel-table-query-delete-source.xlsx')
        const sourceBytes = buildTableOwnedQueryTableSourceXlsx(csvPath)
        writeFileSync(sourcePath, sourceBytes)
        expect(queryTableTopology(sourceBytes).tableQueryOwnerNames).toEqual(['ARevenueQuery'])

        const excelResult = runMacosExcelStructuralOperationOracle({
          workbookPath: sourcePath,
          worksheetName: 'Revenue',
          operations: [{ kind: 'deleteTable', tableName: 'ARevenueQuery' }],
          inspectCells: ['A1'],
          saveWorkbook: true,
          timeoutMs: 120_000,
        })
        expect(excelResult.excelVersion).toMatch(/^\d+\./u)

        const excelDeletedTopology = queryTableTopology(new Uint8Array(readFileSync(sourcePath)))
        expect(excelDeletedTopology).toMatchObject({
          packageParts: ['xl/connections.xml'],
          workbookConnectionRelationships: 1,
          worksheetQueryTableRelationships: 0,
          tableQueryOwnerNames: [],
          queryTableConnectionIds: [],
          contentTypeOverrides: ['/xl/connections.xml'],
        })

        const headless = new SpreadsheetEngine({ workbookName: 'headless-table-query-delete-oracle' })
        await headless.ready()
        headless.importSnapshot(importXlsx(sourceBytes, 'excel-table-query-delete-source.xlsx').snapshot)
        expect(headless.deleteTable('ARevenueQuery')).toBe(true)

        const headlessBytes = exportXlsx(headless.exportSnapshot())
        expect(queryTableTopology(headlessBytes)).toEqual(excelDeletedTopology)
      } finally {
        removeMacosExcelTestDir(tempDir)
      }
    },
    180_000,
  )
})

function buildLegacyTextQueryTableSourceXlsx(csvPath: string): Uint8Array {
  const zip = unzipSync(exportXlsx(queryTableSourceSnapshot()))
  zip['xl/_rels/workbook.xml.rels'] = strToU8(
    appendRelationship(
      readZipTextFromZip(zip, 'xl/_rels/workbook.xml.rels'),
      `<Relationship Id="rIdConnections" Type="${connectionsRelationshipType}" Target="connections.xml"/>`,
    ),
  )
  zip['xl/worksheets/_rels/sheet1.xml.rels'] = strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdQueryTable1" Type="${queryTableRelationshipType}" Target="../queryTables/queryTable1.xml"/></Relationships>`)
  zip['xl/connections.xml'] = strToU8(
    [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<connections xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
      `<connection id="1" name="bilig-query-table-revenue" type="6" refreshedVersion="8" background="1" saveData="1"><textPr codePage="10000" sourceFile="${escapeXml(csvPath)}" comma="1"><textFields count="2"><textField/><textField/></textFields></textPr></connection>`,
      '</connections>',
    ].join(''),
  )
  zip['xl/queryTables/queryTable1.xml'] = strToU8(
    [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<queryTable xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="bilig-query-table-revenue" connectionId="1" autoFormatId="16" applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="1" applyPatternFormats="1" applyAlignmentFormats="0" applyWidthHeightFormats="0"/>',
    ].join(''),
  )
  zip['[Content_Types].xml'] = strToU8(
    [
      { partName: '/xl/connections.xml', contentType: connectionsContentType },
      { partName: '/xl/queryTables/queryTable1.xml', contentType: queryTableContentType },
    ].reduce((xml, entry) => upsertContentTypeOverride(xml, entry), readZipTextFromZip(zip, '[Content_Types].xml')),
  )
  return zipSync(zip)
}

function buildTableOwnedQueryTableSourceXlsx(csvPath: string): Uint8Array {
  const zip = unzipSync(exportXlsx(tableOwnedQuerySourceSnapshot()))
  zip['xl/_rels/workbook.xml.rels'] = strToU8(
    appendRelationship(
      readZipTextFromZip(zip, 'xl/_rels/workbook.xml.rels'),
      `<Relationship Id="rIdConnections" Type="${connectionsRelationshipType}" Target="connections.xml"/>`,
    ),
  )
  zip['xl/connections.xml'] = strToU8(
    [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<connections xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
      `<connection id="1" name="ARevenueQuery" type="6" refreshedVersion="8" background="1" saveData="1"><textPr codePage="10000" sourceFile="${escapeXml(csvPath)}" comma="1"><textFields count="2"><textField/><textField/></textFields></textPr></connection>`,
      '</connections>',
    ].join(''),
  )
  zip['xl/tables/_rels/table2.xml.rels'] = strToU8(
    [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      `<Relationships xmlns="${packageRelationshipNamespace}">`,
      `<Relationship Id="rIdQueryTable1" Type="${queryTableRelationshipType}" Target="../queryTables/queryTable1.xml"/>`,
      '</Relationships>',
    ].join(''),
  )
  zip['xl/queryTables/queryTable1.xml'] = strToU8(
    [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<queryTable xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="ARevenueQuery" connectionId="1" autoFormatId="16" applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="1" applyPatternFormats="1" applyAlignmentFormats="0" applyWidthHeightFormats="0"/>',
    ].join(''),
  )
  zip['[Content_Types].xml'] = strToU8(
    [
      { partName: '/xl/connections.xml', contentType: connectionsContentType },
      { partName: '/xl/queryTables/queryTable1.xml', contentType: queryTableContentType },
    ].reduce((xml, entry) => upsertContentTypeOverride(xml, entry), readZipTextFromZip(zip, '[Content_Types].xml')),
  )
  return zipSync(zip)
}

function queryTableSourceSnapshot(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: {
      name: 'Legacy text query table source',
    },
    sheets: [
      {
        id: 1,
        name: 'Revenue',
        order: 0,
        cells: [
          { address: 'A1', value: 'Region' },
          { address: 'B1', value: 'Amount' },
          { address: 'A2', value: 'North' },
          { address: 'B2', value: 1200 },
          { address: 'A3', value: 'South' },
          { address: 'B3', value: 900 },
        ],
      },
    ],
  }
}

function tableOwnedQuerySourceSnapshot(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: {
      name: 'Table-owned query table source',
      metadata: {
        tables: [
          {
            name: 'ZPlain',
            sheetName: 'Revenue',
            startAddress: 'A1',
            endAddress: 'B3',
            columnNames: ['Region', 'Amount'],
            headerRow: true,
            totalsRow: false,
          },
          {
            name: 'ARevenueQuery',
            sheetName: 'Revenue',
            startAddress: 'D1',
            endAddress: 'E3',
            columnNames: ['Region', 'Amount'],
            headerRow: true,
            totalsRow: false,
          },
        ],
      },
    },
    sheets: [
      {
        id: 1,
        name: 'Revenue',
        order: 0,
        cells: [
          { address: 'A1', value: 'Region' },
          { address: 'B1', value: 'Amount' },
          { address: 'A2', value: 'West' },
          { address: 'B2', value: 400 },
          { address: 'D1', value: 'Region' },
          { address: 'E1', value: 'Amount' },
          { address: 'D2', value: 'North' },
          { address: 'E2', value: 1200 },
        ],
      },
    ],
  }
}

function queryTableTopology(bytes: Uint8Array): {
  readonly packageParts: readonly string[]
  readonly workbookConnectionRelationships: number
  readonly worksheetQueryTableRelationships: number
  readonly tableQueryOwnerNames: readonly string[]
  readonly queryTableConnectionIds: readonly string[]
  readonly contentTypeOverrides: readonly string[]
} {
  const zip = unzipSync(bytes)
  return {
    packageParts: Object.keys(zip)
      .filter((path) => path === 'xl/connections.xml' || path.startsWith('xl/queryTables/'))
      .toSorted(),
    workbookConnectionRelationships: relationshipsWithType(
      readZipTextFromZip(zip, 'xl/_rels/workbook.xml.rels'),
      connectionsRelationshipType,
    ).length,
    worksheetQueryTableRelationships: Object.entries(zip)
      .filter(([path]) => /^xl\/worksheets\/_rels\/sheet[1-9][0-9]*\.xml\.rels$/u.test(path))
      .flatMap(([, data]) => relationshipsWithType(strFromU8(data), queryTableRelationshipType)).length,
    tableQueryOwnerNames: tableQueryOwnerNames(zip),
    queryTableConnectionIds: Object.entries(zip)
      .filter(([path]) => /^xl\/queryTables\/queryTable[1-9][0-9]*\.xml$/u.test(path))
      .flatMap(([, data]) => readXmlAttribute(strFromU8(data), 'connectionId') ?? [])
      .toSorted(),
    contentTypeOverrides: contentTypeOverridesForQueryTableParts(readZipTextFromZip(zip, '[Content_Types].xml')),
  }
}

function tableQueryOwnerNames(zip: Record<string, Uint8Array>): string[] {
  return Object.entries(zip)
    .filter(([path]) => /^xl\/tables\/table[1-9][0-9]*\.xml$/u.test(path))
    .flatMap(([tablePath, tableBytes]) => {
      const relationshipsXml = readOptionalZipTextFromZip(zip, relationshipPartPath(tablePath)) ?? ''
      if (relationshipsWithType(relationshipsXml, queryTableRelationshipType).length === 0) {
        return []
      }
      return readXmlAttribute(strFromU8(tableBytes), 'displayName') ?? []
    })
    .toSorted()
}

function relationshipPartPath(partPath: string): string {
  const slashIndex = partPath.lastIndexOf('/')
  const directory = slashIndex >= 0 ? partPath.slice(0, slashIndex) : ''
  const fileName = slashIndex >= 0 ? partPath.slice(slashIndex + 1) : partPath
  return directory.length > 0 ? `${directory}/_rels/${fileName}.rels` : `_rels/${fileName}.rels`
}

function relationshipsWithType(relationshipsXml: string, relationshipType: string): string[] {
  return [...relationshipsXml.matchAll(/<Relationship\b([^>]*)\/?>/gu)].flatMap((match) => {
    const attributes = match[1] ?? ''
    return readXmlAttribute(attributes, 'Type') === relationshipType ? [match[0]] : []
  })
}

function contentTypeOverridesForQueryTableParts(contentTypesXml: string): string[] {
  return [...contentTypesXml.matchAll(/<Override\b([^>]*)\/?>/gu)]
    .flatMap((match) => {
      const attributes = match[1] ?? ''
      const partName = readXmlAttribute(attributes, 'PartName')
      return partName === '/xl/connections.xml' || partName?.startsWith('/xl/queryTables/') ? [partName] : []
    })
    .toSorted()
}

function readOptionalZipTextFromZip(zip: Record<string, Uint8Array>, path: string): string | undefined {
  const bytes = zip[path]
  return bytes ? strFromU8(bytes) : undefined
}

function readZipTextFromZip(zip: Record<string, Uint8Array>, path: string): string {
  const bytes = zip[path]
  if (!bytes) {
    throw new Error(`Missing XLSX part: ${path}`)
  }
  return strFromU8(bytes)
}

function appendRelationship(relationshipsXml: string, relationshipXml: string): string {
  return relationshipsXml.replace('</Relationships>', `${relationshipXml}</Relationships>`)
}

function upsertContentTypeOverride(
  contentTypesXml: string,
  input: {
    readonly partName: string
    readonly contentType: string
  },
): string {
  if (contentTypesXml.includes(`PartName="${input.partName}"`)) {
    return contentTypesXml
  }
  return contentTypesXml.replace('</Types>', `<Override PartName="${input.partName}" ContentType="${input.contentType}"/></Types>`)
}

function readXmlAttribute(xml: string, attributeName: string): string | null {
  const match = new RegExp(`\\s${attributeName}=("|')([\\s\\S]*?)\\1`, 'u').exec(xml)
  return match?.[2] ?? null
}

function escapeXml(value: string): string {
  return value.replaceAll('&', '&amp;').replaceAll('"', '&quot;').replaceAll('<', '&lt;').replaceAll('>', '&gt;').replaceAll("'", '&apos;')
}
