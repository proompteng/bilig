import { readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'

import { exportXlsx, importXlsx } from '@bilig/excel-import'
import { isMacosExcelInstalled, runMacosExcelPackageOpenSaveOracle } from '@bilig/excel-fixtures'
import type { WorkbookSnapshot } from '@bilig/protocol'
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { describe, expect, it } from 'vitest'

import { WorkPaper } from '../index.js'
import { createExcelAccessibleTempDir, removeMacosExcelTestDir } from './macos-excel-oracle-test-utils.js'

const powerQueryRelationshipType = 'http://schemas.microsoft.com/office/2014/relationships/query'
const powerQueryGroupRelationshipType = 'http://schemas.microsoft.com/office/2014/relationships/queryGroup'
const powerQueryContentType = 'application/vnd.ms-excel.query+xml'
const powerQueryGroupContentType = 'application/vnd.ms-excel.queryGroup+xml'

describe('macOS Desktop Excel Power Query package oracle', () => {
  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'preserves Power Query package topology after Desktop Excel open/save and headless export',
    () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = createExcelAccessibleTempDir('bilig-headless-excel-power-query-oracle-')
      try {
        const sourcePath = join(tempDir, 'excel-power-query-source.xlsx')
        const sourceBytes = buildPowerQuerySourceXlsx()
        writeFileSync(sourcePath, sourceBytes)
        expect(powerQueryTopology(sourceBytes)).toEqual(expectedPowerQueryTopology())

        const excelSource = runMacosExcelPackageOpenSaveOracle({
          workbookPath: sourcePath,
          saveWorkbook: true,
          timeoutMs: 120_000,
        })
        expect(excelSource.excelVersion).toMatch(/^\d+\./u)

        const excelSourceBytes = new Uint8Array(readFileSync(sourcePath))
        const excelSourceTopology = powerQueryTopology(excelSourceBytes)
        expect(excelSourceTopology).toEqual(expectedPowerQueryTopology())

        const imported = importXlsx(excelSourceBytes, 'excel-power-query-source.xlsx').snapshot
        expect(imported.workbook.metadata?.dataModelArtifacts?.parts.map((part) => part.path).toSorted()).toEqual([
          'xl/queries/query1.xml',
          'xl/queryGroups/queryGroup1.xml',
        ])

        const workpaper = WorkPaper.buildFromSnapshot(imported)
        try {
          const sheet = workpaper.getSheetId('Revenue')
          if (sheet === undefined) {
            throw new Error('Expected Revenue sheet')
          }
          workpaper.setCellContents({ sheet, row: 0, col: 2 }, 'headless reviewed')

          const headlessPath = join(tempDir, 'headless-power-query.xlsx')
          const headlessBytes = exportXlsx(workpaper.exportSnapshot())
          writeFileSync(headlessPath, headlessBytes)
          expect(powerQueryTopology(headlessBytes)).toEqual(excelSourceTopology)

          const excelHeadless = runMacosExcelPackageOpenSaveOracle({
            workbookPath: headlessPath,
            saveWorkbook: true,
            timeoutMs: 120_000,
          })
          expect(excelHeadless.excelVersion).toMatch(/^\d+\./u)

          const excelSavedHeadlessBytes = new Uint8Array(readFileSync(headlessPath))
          expect(powerQueryTopology(excelSavedHeadlessBytes)).toEqual(excelSourceTopology)
          expect(importXlsx(excelSavedHeadlessBytes, 'excel-saved-headless-power-query.xlsx').snapshot.sheets[0]?.cells).toContainEqual(
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
})

function buildPowerQuerySourceXlsx(): Uint8Array {
  const zip = unzipSync(exportXlsx(powerQuerySourceSnapshot()))
  zip['xl/_rels/workbook.xml.rels'] = strToU8(
    readZipTextFromZip(zip, 'xl/_rels/workbook.xml.rels').replace(
      '</Relationships>',
      [
        `<Relationship Id="rIdQuery1" Type="${powerQueryRelationshipType}" Target="queries/query1.xml"/>`,
        `<Relationship Id="rIdQueryGroup1" Type="${powerQueryGroupRelationshipType}" Target="queryGroups/queryGroup1.xml"/>`,
        '</Relationships>',
      ].join(''),
    ),
  )
  zip['xl/queries/query1.xml'] = strToU8(powerQueryXml)
  zip['xl/queryGroups/queryGroup1.xml'] = strToU8(powerQueryGroupXml)
  zip['[Content_Types].xml'] = strToU8(
    [
      { partName: '/xl/queries/query1.xml', contentType: powerQueryContentType },
      { partName: '/xl/queryGroups/queryGroup1.xml', contentType: powerQueryGroupContentType },
    ].reduce((xml, entry) => upsertContentTypeOverride(xml, entry), readZipTextFromZip(zip, '[Content_Types].xml')),
  )
  return zipSync(zip)
}

function powerQuerySourceSnapshot(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: {
      name: 'Power Query package source',
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
        ],
      },
    ],
  }
}

function expectedPowerQueryTopology(): PowerQueryTopology {
  return {
    packageParts: ['xl/queries/query1.xml', 'xl/queryGroups/queryGroup1.xml'],
    workbookQueryGroupRelationships: 1,
    workbookQueryRelationships: 1,
    contentTypeOverrides: ['/xl/queries/query1.xml', '/xl/queryGroups/queryGroup1.xml'],
  }
}

interface PowerQueryTopology {
  readonly packageParts: readonly string[]
  readonly workbookQueryGroupRelationships: number
  readonly workbookQueryRelationships: number
  readonly contentTypeOverrides: readonly string[]
}

function powerQueryTopology(bytes: Uint8Array): PowerQueryTopology {
  const zip = unzipSync(bytes)
  return {
    packageParts: Object.keys(zip)
      .filter((path) => path.startsWith('xl/queries/') || path.startsWith('xl/queryGroups/'))
      .toSorted(),
    workbookQueryGroupRelationships: relationshipsWithType(
      readZipTextFromZip(zip, 'xl/_rels/workbook.xml.rels'),
      powerQueryGroupRelationshipType,
    ).length,
    workbookQueryRelationships: relationshipsWithType(readZipTextFromZip(zip, 'xl/_rels/workbook.xml.rels'), powerQueryRelationshipType)
      .length,
    contentTypeOverrides: contentTypeOverridesForPowerQueryParts(readZipTextFromZip(zip, '[Content_Types].xml')),
  }
}

function relationshipsWithType(relationshipsXml: string, relationshipType: string): string[] {
  return [...relationshipsXml.matchAll(/<Relationship\b([^>]*)\/?>/gu)].flatMap((match) => {
    const attributes = match[1] ?? ''
    return readXmlAttribute(attributes, 'Type') === relationshipType ? [match[0]] : []
  })
}

function contentTypeOverridesForPowerQueryParts(contentTypesXml: string): string[] {
  return [...contentTypesXml.matchAll(/<Override\b([^>]*)\/?>/gu)]
    .flatMap((match) => {
      const attributes = match[1] ?? ''
      const partName = readXmlAttribute(attributes, 'PartName')
      return partName?.startsWith('/xl/queries/') || partName?.startsWith('/xl/queryGroups/') ? [partName] : []
    })
    .toSorted()
}

function readZipTextFromZip(zip: Record<string, Uint8Array>, path: string): string {
  const bytes = zip[path]
  if (!bytes) {
    throw new Error(`Missing XLSX part: ${path}`)
  }
  return strFromU8(bytes)
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

const powerQueryXml = [
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
  '<query xmlns="http://schemas.microsoft.com/office/2014/queries" name="RevenueQuery">',
  '<formula><![CDATA[let Source = #table({"Region","Amount"}, {{"North",1200}}) in Source]]></formula>',
  '</query>',
].join('')

const powerQueryGroupXml = [
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
  '<queryGroup xmlns="http://schemas.microsoft.com/office/2014/queryGroups" name="Finance Queries"/>',
].join('')
