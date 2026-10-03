import { readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'

import { exportXlsx, importXlsx } from '@bilig/excel-import'
import { isMacosExcelInstalled, runMacosExcelInspectionOracle, runMacosExcelStructuralOperationOracle } from '@bilig/excel-fixtures'
import type { WorkbookSnapshot } from '@bilig/protocol'
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { describe, expect, it } from 'vitest'

import { WorkPaper } from '../index.js'
import { createExcelAccessibleTempDir, removeMacosExcelTestDir } from './macos-excel-oracle-test-utils.js'

const relationshipNamespace = 'http://schemas.openxmlformats.org/package/2006/relationships'
const officeRelationshipNamespace = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const slicerCacheRelationshipType = 'http://schemas.microsoft.com/office/2007/relationships/slicerCache'
const slicerRelationshipType = 'http://schemas.microsoft.com/office/2007/relationships/slicer'
const tableRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/table'
const slicerCacheContentType = 'application/vnd.ms-excel.slicerCache+xml'
const slicerContentType = 'application/vnd.ms-excel.slicer+xml'

describe('macOS Desktop Excel slicer connection deletion oracle', () => {
  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'matches Desktop Excel slicer package cleanup after deleting the slicer sheet',
    () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = createExcelAccessibleTempDir('bilig-headless-excel-slicer-delete-oracle-')
      try {
        const sourceBytes = buildSlicerConnectionDeleteSourceXlsx()
        expect(slicerConnectionTopology(sourceBytes)).toMatchObject({
          workbookSlicerCacheRelationships: 1,
          workbookSlicerCacheRefs: 1,
          sheetSlicerRelationships: 1,
          sheetSlicerRefs: 1,
        })

        const excelWorkbookPath = join(tempDir, 'excel-slicer-delete-source.xlsx')
        writeFileSync(excelWorkbookPath, sourceBytes)
        const excelResult = runMacosExcelStructuralOperationOracle({
          workbookPath: excelWorkbookPath,
          worksheetName: 'Keep',
          operations: [{ kind: 'deleteSheet', name: 'Revenue' }],
          inspectCells: ['A1'],
          saveWorkbook: true,
          timeoutMs: 120_000,
        })
        expect(excelResult.cells[0]?.value).toEqual({ kind: 'string', value: 'keep' })

        const excelTruthBytes = new Uint8Array(readFileSync(excelWorkbookPath))
        const excelTruth = importXlsx(excelTruthBytes, 'excel-slicer-delete-truth.xlsx').snapshot
        expect(excelTruth.sheets.map((sheet) => sheet.name)).toEqual(['Keep'])

        const importedSource = importXlsx(sourceBytes, 'slicer-delete-source.xlsx').snapshot
        const workpaper = WorkPaper.buildFromSnapshot(importedSource)
        try {
          const revenueSheet = workpaper.getSheetId('Revenue')
          if (revenueSheet === undefined) {
            throw new Error('Expected Revenue sheet')
          }
          workpaper.removeSheet(revenueSheet)

          const headlessSnapshot = workpaper.exportSnapshot()
          expect(headlessSnapshot.sheets.map((sheet) => sheet.name)).toEqual(['Keep'])
          expect(headlessSnapshot.workbook.metadata?.slicerConnectionArtifacts?.sheetArtifacts).toBeUndefined()

          const headlessPath = join(tempDir, 'headless-slicer-delete.xlsx')
          const headlessBytes = exportXlsx(headlessSnapshot)
          writeFileSync(headlessPath, headlessBytes)

          expect(slicerConnectionTopology(headlessBytes)).toEqual(slicerConnectionTopology(excelTruthBytes))

          const headlessExcel = runMacosExcelInspectionOracle({
            workbookPath: headlessPath,
            worksheetName: 'Keep',
            formulaCells: [],
            inspectCells: ['A1'],
            saveWorkbook: true,
            timeoutMs: 90_000,
          })
          expect(headlessExcel.cells).toEqual(excelResult.cells)
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

function buildSlicerConnectionDeleteSourceXlsx(): Uint8Array {
  const zip = unzipSync(exportXlsx(slicerConnectionDeleteSnapshot()))
  const tablePath = Object.keys(zip).find((path) => /^xl\/tables\/table[1-9][0-9]*\.xml$/u.test(path))
  if (!tablePath) {
    throw new Error('Expected exported workbook to include a table part')
  }
  const tableTarget = `../tables/${tablePath.slice(tablePath.lastIndexOf('/') + 1)}`
  zip['xl/workbook.xml'] = strToU8(
    ensureRelationshipNamespace(readZipTextFromZip(zip, 'xl/workbook.xml')).replace(
      '</workbook>',
      `<extLst>${workbookSlicerCachesExtXml}</extLst></workbook>`,
    ),
  )
  zip['xl/_rels/workbook.xml.rels'] = strToU8(
    readZipTextFromZip(zip, 'xl/_rels/workbook.xml.rels').replace(
      '</Relationships>',
      `<Relationship Id="rId80" Type="${slicerCacheRelationshipType}" Target="slicerCaches/slicerCache1.xml"/></Relationships>`,
    ),
  )
  zip['xl/worksheets/sheet1.xml'] = strToU8(
    ensureRelationshipNamespace(readZipTextFromZip(zip, 'xl/worksheets/sheet1.xml')).replace(
      '</worksheet>',
      `<extLst>${sheetSlicerListExtXml}</extLst></worksheet>`,
    ),
  )
  zip['xl/worksheets/_rels/sheet1.xml.rels'] = strToU8(
    appendRelationship(
      readZipTextFromZip(zip, 'xl/worksheets/_rels/sheet1.xml.rels'),
      `<Relationship Id="rId20" Type="${slicerRelationshipType}" Target="../slicers/slicer1.xml"/>`,
    ),
  )
  zip['xl/slicerCaches/slicerCache1.xml'] = strToU8(slicerCacheXml)
  zip['xl/slicerCaches/_rels/slicerCache1.xml.rels'] = strToU8(slicerCacheRelationshipsXml(tableTarget))
  zip['xl/slicers/slicer1.xml'] = strToU8(slicerXml)
  zip['[Content_Types].xml'] = strToU8(
    [
      { partName: '/xl/slicerCaches/slicerCache1.xml', contentType: slicerCacheContentType },
      { partName: '/xl/slicers/slicer1.xml', contentType: slicerContentType },
    ].reduce((xml, entry) => upsertContentTypeOverride(xml, entry), readZipTextFromZip(zip, '[Content_Types].xml')),
  )
  return zipSync(zip)
}

function slicerConnectionDeleteSnapshot(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: {
      name: 'Slicer connection sheet deletion',
      metadata: {
        tables: [
          {
            name: 'RevenueTable',
            sheetName: 'Revenue',
            startAddress: 'A1',
            endAddress: 'B4',
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
          { address: 'A2', value: 'North' },
          { address: 'B2', value: 1200 },
          { address: 'A3', value: 'South' },
          { address: 'B3', value: 900 },
          { address: 'A4', value: 'North' },
          { address: 'B4', value: 300 },
        ],
      },
      {
        id: 2,
        name: 'Keep',
        order: 1,
        cells: [{ address: 'A1', value: 'keep' }],
      },
    ],
  }
}

function slicerConnectionTopology(bytes: Uint8Array): {
  packageParts: string[]
  workbookSlicerCacheRelationships: number
  workbookSlicerCacheRefs: number
  sheetSlicerRelationships: number
  sheetSlicerRefs: number
  contentTypeOverrides: string[]
} {
  const zip = unzipSync(bytes)
  return {
    packageParts: Object.keys(zip)
      .filter((path) => path.startsWith('xl/slicerCaches/') || path.startsWith('xl/slicers/'))
      .toSorted(),
    workbookSlicerCacheRelationships: relationshipsWithType(
      readZipTextFromZip(zip, 'xl/_rels/workbook.xml.rels'),
      slicerCacheRelationshipType,
    ).length,
    workbookSlicerCacheRefs: countZipXmlMatches(zip, /^xl\/workbook\.xml$/u, /<x15:slicerCache\b/gu),
    sheetSlicerRelationships: Object.entries(zip)
      .filter(([path]) => /^xl\/worksheets\/_rels\/sheet[1-9][0-9]*\.xml\.rels$/u.test(path))
      .flatMap(([, data]) => relationshipsWithType(strFromU8(data), slicerRelationshipType)).length,
    sheetSlicerRefs: countZipXmlMatches(zip, /^xl\/worksheets\/sheet[1-9][0-9]*\.xml$/u, /<x14:slicer\b/gu),
    contentTypeOverrides: contentTypeOverridesForSlicerParts(readZipTextFromZip(zip, '[Content_Types].xml')),
  }
}

function countZipXmlMatches(zip: Record<string, Uint8Array>, pathPattern: RegExp, contentPattern: RegExp): number {
  let count = 0
  for (const [path, data] of Object.entries(zip)) {
    if (!pathPattern.test(path)) {
      continue
    }
    for (const _match of strFromU8(data).matchAll(contentPattern)) {
      count += 1
    }
  }
  return count
}

function relationshipsWithType(relationshipsXml: string, relationshipType: string): string[] {
  return [...relationshipsXml.matchAll(/<Relationship\b([^>]*)\/?>/gu)].flatMap((match) => {
    const attributes = match[1] ?? ''
    return readXmlAttribute(attributes, 'Type') === relationshipType ? [match[0]] : []
  })
}

function contentTypeOverridesForSlicerParts(contentTypesXml: string): string[] {
  return [...contentTypesXml.matchAll(/<Override\b([^>]*)\/?>/gu)]
    .flatMap((match) => {
      const attributes = match[1] ?? ''
      const partName = readXmlAttribute(attributes, 'PartName')
      return partName?.startsWith('/xl/slicer') ? [partName] : []
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

function ensureRelationshipNamespace(xml: string): string {
  if (/xmlns:r=/u.test(xml)) {
    return xml
  }
  return xml.replace(/<([A-Za-z0-9:]+)\b([^>]*)>/u, `<$1$2 xmlns:r="${officeRelationshipNamespace}">`)
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

function readXmlAttribute(attributes: string, name: string): string | null {
  return new RegExp(`\\b${name}=("|')([\\s\\S]*?)\\1`, 'u').exec(attributes)?.[2] ?? null
}

const workbookSlicerCachesExtXml = [
  '<ext uri="{A8765BA9-456A-4DAB-B4F3-ACF838C121DE}" ',
  'xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">',
  '<x15:slicerCaches><x15:slicerCache r:id="rId80"/></x15:slicerCaches>',
  '</ext>',
].join('')

const sheetSlicerListExtXml = [
  '<ext uri="{A8765BA9-456A-4DAB-B4F3-ACF838C121DE}" ',
  'xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">',
  '<x14:slicerList><x14:slicer r:id="rId20"/></x14:slicerList>',
  '</ext>',
].join('')

const slicerCacheXml = [
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
  '<slicerCacheDefinition xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" ',
  `xmlns:r="${officeRelationshipNamespace}" name="Slicer_Region" sourceName="Region" cache="Slicer_Region" r:id="rId1">`,
  '<tableSlicerCache tableId="1" column="1">',
  '<items count="2"><i x="0" s="1"/><i x="1"/></items>',
  '</tableSlicerCache>',
  '</slicerCacheDefinition>',
].join('')

function slicerCacheRelationshipsXml(tableTarget: string): string {
  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    `<Relationships xmlns="${relationshipNamespace}">`,
    `<Relationship Id="rId1" Type="${tableRelationshipType}" Target="${tableTarget}"/>`,
    '</Relationships>',
  ].join('')
}

const slicerXml = [
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
  '<slicer xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" ',
  'name="Slicer_Region" cache="Slicer_Region" caption="Region" startItem="0" columnCount="1"/>',
].join('')
