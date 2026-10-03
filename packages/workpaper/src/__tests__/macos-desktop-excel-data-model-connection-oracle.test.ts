import { readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'

import { exportXlsx, importXlsx } from '@bilig/excel-import'
import { isMacosExcelInstalled, runMacosExcelPackageOpenSaveOracle } from '@bilig/excel-fixtures'
import type { WorkbookExternalConnectionSnapshot, WorkbookSnapshot } from '@bilig/protocol'
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { describe, expect, it } from 'vitest'

import { WorkPaper } from '../index.js'
import { createExcelAccessibleTempDir, removeMacosExcelTestDir } from './macos-excel-oracle-test-utils.js'

const officeRelationshipNamespace = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const connectionsRelationshipType = `${officeRelationshipNamespace}/connections`
const connectionsContentType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml'

describe('macOS Desktop Excel Data Model connection oracle', () => {
  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'classifies Desktop Excel Data Model connections after open/save and headless export',
    () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = createExcelAccessibleTempDir('bilig-headless-excel-data-model-connection-oracle-')
      try {
        const sourcePath = join(tempDir, 'excel-data-model-connection-source.xlsx')
        const sourceBytes = buildDataModelConnectionSourceXlsx()
        writeFileSync(sourcePath, sourceBytes)
        expect(dataModelConnection(sourceBytes)).toMatchObject({
          id: 1,
          name: 'ThisWorkbookDataModel',
          sourceKind: 'model',
          command: 'Model',
        })

        const excelSource = runMacosExcelPackageOpenSaveOracle({
          workbookPath: sourcePath,
          saveWorkbook: true,
          timeoutMs: 120_000,
        })
        expect(excelSource.excelVersion).toMatch(/^\d+\./u)

        const excelSourceBytes = new Uint8Array(readFileSync(sourcePath))
        const excelSourceTopology = dataModelConnectionTopology(excelSourceBytes)
        expect(excelSourceTopology.connectionSourceKinds).toEqual(['model'])
        expect(excelSourceTopology.connectionsPackageParts).toEqual(['xl/connections.xml'])

        const imported = importXlsx(excelSourceBytes, 'excel-data-model-connection-source.xlsx').snapshot
        expect(imported.workbook.metadata?.externalConnections?.connections).toEqual([
          expect.objectContaining({
            id: 1,
            name: 'ThisWorkbookDataModel',
            sourceKind: 'model',
            connection: expect.stringContaining('$Workbook$'),
            command: 'Model',
          }),
        ])
        expect(imported.workbook.metadata?.slicerConnectionArtifacts?.parts.map((part) => part.path)).toEqual(['xl/connections.xml'])

        const workpaper = WorkPaper.buildFromSnapshot(imported)
        try {
          const modelSheet = workpaper.getSheetId('Model')
          if (modelSheet === undefined) {
            throw new Error('Expected Model sheet')
          }
          workpaper.setCellContents({ sheet: modelSheet, row: 0, col: 1 }, 'headless reviewed')

          const headlessPath = join(tempDir, 'headless-data-model-connection.xlsx')
          const headlessBytes = exportXlsx(workpaper.exportSnapshot())
          writeFileSync(headlessPath, headlessBytes)
          expect(dataModelConnectionTopology(headlessBytes)).toEqual(excelSourceTopology)

          const excelHeadless = runMacosExcelPackageOpenSaveOracle({
            workbookPath: headlessPath,
            saveWorkbook: true,
            timeoutMs: 120_000,
          })
          expect(excelHeadless.excelVersion).toMatch(/^\d+\./u)

          const excelSavedHeadlessBytes = new Uint8Array(readFileSync(headlessPath))
          expect(dataModelConnectionTopology(excelSavedHeadlessBytes)).toEqual(excelSourceTopology)
          const reimported = importXlsx(excelSavedHeadlessBytes, 'excel-saved-headless-data-model-connection.xlsx').snapshot
          expect(reimported.workbook.metadata?.externalConnections?.connections?.map((connection) => connection.sourceKind)).toEqual([
            'model',
          ])
          expect(reimported.sheets[0]?.cells).toContainEqual(
            expect.objectContaining({
              address: 'B1',
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

function buildDataModelConnectionSourceXlsx(): Uint8Array {
  const zip = unzipSync(exportXlsx(dataModelConnectionSourceSnapshot()))
  zip['xl/_rels/workbook.xml.rels'] = strToU8(
    readZipTextFromZip(zip, 'xl/_rels/workbook.xml.rels').replace(
      '</Relationships>',
      `<Relationship Id="rIdConnections" Type="${connectionsRelationshipType}" Target="connections.xml"/></Relationships>`,
    ),
  )
  zip['xl/connections.xml'] = strToU8(dataModelConnectionsXml)
  zip['[Content_Types].xml'] = strToU8(
    upsertContentTypeOverride(readZipTextFromZip(zip, '[Content_Types].xml'), {
      contentType: connectionsContentType,
      partName: '/xl/connections.xml',
    }),
  )
  return zipSync(zip)
}

function dataModelConnectionSourceSnapshot(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: { name: 'Data Model connection source' },
    sheets: [
      {
        id: 1,
        name: 'Model',
        order: 0,
        cells: [{ address: 'A1', value: 'Power Pivot fixture' }],
      },
    ],
  }
}

function dataModelConnection(bytes: Uint8Array): WorkbookExternalConnectionSnapshot {
  const connection = importXlsx(bytes, 'data-model-connection.xlsx').snapshot.workbook.metadata?.externalConnections?.connections?.[0]
  if (!connection) {
    throw new Error('Expected imported Data Model connection provenance')
  }
  return connection
}

function dataModelConnectionTopology(bytes: Uint8Array): {
  readonly connectionSourceKinds: readonly string[]
  readonly connectionsPackageParts: readonly string[]
  readonly contentTypeOverrides: readonly string[]
} {
  const zip = unzipSync(bytes)
  return {
    connectionSourceKinds:
      importXlsx(bytes, 'data-model-connection-topology.xlsx')
        .snapshot.workbook.metadata?.externalConnections?.connections?.map((connection) => connection.sourceKind)
        .toSorted() ?? [],
    connectionsPackageParts: Object.keys(zip)
      .filter((path) => path === 'xl/connections.xml')
      .toSorted(),
    contentTypeOverrides: contentTypeOverridesForDataModelParts(readZipTextFromZip(zip, '[Content_Types].xml')),
  }
}

function contentTypeOverridesForDataModelParts(contentTypesXml: string): readonly string[] {
  return [...contentTypesXml.matchAll(/<Override\b([^>]*)\/?>/gu)]
    .flatMap((match) => {
      const partName = readXmlAttribute(match[1] ?? '', 'PartName')
      return partName === '/xl/connections.xml' ? [partName] : []
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
  return contentTypesXml.replace(
    '</Types>',
    `<Override PartName="${escapeXml(input.partName)}" ContentType="${escapeXml(input.contentType)}"/></Types>`,
  )
}

function readXmlAttribute(attributes: string, attributeName: string): string | null {
  return new RegExp(`\\s${attributeName}=("|')([\\s\\S]*?)\\1`, 'u').exec(attributes)?.[2] ?? null
}

function escapeXml(value: string): string {
  return value.replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;').replaceAll('"', '&quot;').replaceAll("'", '&apos;')
}

const dataModelConnectionsXml = [
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
  '<connections xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
  '<connection id="1" name="ThisWorkbookDataModel" description="Embedded Data Model" type="5" refreshedVersion="8" model="1" saveData="1">',
  '<dbPr connection="Provider=MSOLAP.8;Data Source=$Workbook$;Initial Catalog=Model" command="Model" commandType="1"/>',
  '<olapPr sendLocale="1" rowDrillCount="1000"/>',
  '</connection>',
  '</connections>',
].join('')
