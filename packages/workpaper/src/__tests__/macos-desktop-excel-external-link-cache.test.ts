import { mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { pathToFileURL } from 'node:url'

import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'

import { SpreadsheetEngine } from '@bilig/core'
import { exportXlsx, externalWorkbookReferencesWarning, importXlsx } from '@bilig/excel-import'
import { writeSimpleXlsxWorkbook } from '@bilig/xlsx'
import {
  isMacosExcelInstalled,
  runMacosExcelInspectionOracle,
  runMacosExcelPackageOpenSaveOracle,
  runMacosExcelStructuralOperationOracle,
  type NormalizedFormulaValue,
} from '@bilig/excel-fixtures'
import { ValueTag, type CellValue } from '@bilig/protocol'
import { describe, expect, it } from 'vitest'

const officeRelationshipNamespace = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const externalLinkContentType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml'
const externalRangeAddresses = ['C1', 'C2', 'C3'] as const
const initialExternalRangeValues = [
  { address: 'C1', value: { kind: 'number', value: 120 } },
  { address: 'C2', value: { kind: 'number', value: 40 } },
  { address: 'C3', value: { kind: 'number', value: 60 } },
] as const
const changedExternalRangeValues = [
  { address: 'C1', value: { kind: 'number', value: 180 } },
  { address: 'C2', value: { kind: 'number', value: 60 } },
  { address: 'C3', value: { kind: 'number', value: 90 } },
] as const

describe('macOS Desktop Excel external-link cache oracle', () => {
  it('drops orphaned external-link caches when deleting the last formula sheet', async () => {
    const source = buildExternalLinkRangeCacheWorkbook('file:///tmp/rates.xlsx', { includeKeepSheet: true })
    const imported = importXlsx(source, 'external-link-range-cache-delete.xlsx')
    expect(imported.snapshot.sheets.map((sheet) => sheet.name)).toEqual(['Keep', 'Model', '__bilig_ext_1_Rates'])

    const engine = new SpreadsheetEngine({ workbookName: 'external-link-range-cache-delete' })
    await engine.ready()
    engine.importSnapshot(imported.snapshot)
    engine.deleteSheet('Model')

    const snapshot = engine.exportSnapshot()
    expect(snapshot.sheets.map((sheet) => sheet.name)).toEqual(['Keep'])
    expect(snapshot.workbook.metadata?.externalLinkArtifacts).toBeUndefined()
    expect(snapshot.workbook.metadata?.externalWorkbookReferences).toBeUndefined()
    expect(externalLinkPackageMetrics(unzipSync(exportXlsx(snapshot)))).toEqual({
      packageParts: [],
      workbookExternalReferenceTargets: [],
      externalLinkPathRelationshipParts: [],
      contentTypeOverrides: [],
    })
  })

  it('imports cached external ranges as hidden-sheet references', async () => {
    const source = buildExternalLinkRangeCacheWorkbook()
    const imported = importXlsx(source, 'external-link-range-cache.xlsx')
    const cells = new Map(imported.snapshot.sheets[0]?.cells.map((cell) => [cell.address, cell]) ?? [])
    const cacheSheet = imported.snapshot.sheets.find((sheet) => sheet.name === '__bilig_ext_1_Rates')

    expect(imported.snapshot.workbook.metadata?.externalLinkArtifacts?.parts.map((part) => part.path).toSorted()).toEqual([
      'xl/externalLinks/_rels/externalLink5.xml.rels',
      'xl/externalLinks/externalLink5.xml',
    ])
    expect(cacheSheet).toMatchObject({ metadata: { visibility: 'veryHidden' } })
    expect(cells.get('C1')).toMatchObject({ formula: "SUM('__bilig_ext_1_Rates'!$B$2:$B$4)*B1", value: 120 })
    expect(cells.get('C2')).toMatchObject({
      formula: "XLOOKUP(\"B\",'__bilig_ext_1_Rates'!$A$2:$A$4,'__bilig_ext_1_Rates'!$B$2:$B$4)*B1",
      value: 40,
    })
    expect(cells.get('C3')).toMatchObject({
      formula: "SUMIFS('__bilig_ext_1_Rates'!$B$2:$B$4,'__bilig_ext_1_Rates'!$A$2:$A$4,\"C\")*B1",
      value: 60,
    })
    expect(imported.warnings).toEqual([externalWorkbookReferencesWarning])

    const engine = new SpreadsheetEngine({ workbookName: 'external-link-range-cache-import' })
    await engine.ready()
    engine.importSnapshot(imported.snapshot)

    expectEngineValues(engine, initialExternalRangeValues)

    engine.setCellValue('Model', 'B1', 3)

    expectEngineValues(engine, changedExternalRangeValues)

    const exportedZip = unzipSync(exportXlsx(engine.exportSnapshot()))
    expect(externalLinkPackageSummary(exportedZip)).toEqual(externalLinkPackageSummary(unzipSync(source)))

    const modelSheetXml = xmlText(exportedZip, 'xl/worksheets/sheet1.xml')
    const workbookXml = xmlText(exportedZip, 'xl/workbook.xml')

    expect(modelSheetXml).toContain('__bilig_ext_1_Rates')
    expect(modelSheetXml).not.toContain('{')
    expect(workbookXml).toContain('name="__bilig_ext_1_Rates"')
    expect(workbookXml).toContain('state="veryHidden"')
  })

  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'matches Desktop Excel native no-update save for stale external-link caches',
    () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = createExcelAccessibleTempDir('bilig-headless-excel-external-link-native-save-')
      try {
        const sourceWorkbookPath = join(tempDir, 'external-link-native-save.xlsx')
        const sourceBytes = buildExternalLinkRangeCacheWorkbook()
        writeFileSync(sourceWorkbookPath, sourceBytes)

        const sourceExternalLinkSummary = externalLinkPackageSummary(unzipSync(sourceBytes))
        const sourceFormulaCacheValues = worksheetFormulaCacheValues(sourceBytes)
        const excelNativeSave = runMacosExcelPackageOpenSaveOracle({
          workbookPath: sourceWorkbookPath,
          calculationPolicy: 'none',
          saveWorkbook: true,
          timeoutMs: 120_000,
          updateLinks: 'never',
        })
        expect(excelNativeSave.excelVersion).toMatch(/^\d+\./u)

        const excelNativeBytes = new Uint8Array(readFileSync(sourceWorkbookPath))
        expect(externalLinkPackageSummary(unzipSync(excelNativeBytes))).toEqual(sourceExternalLinkSummary)
        expect(worksheetFormulaCacheValues(excelNativeBytes)).toEqual(sourceFormulaCacheValues)

        const exported = exportXlsx(importXlsx(sourceBytes, 'external-link-native-save.xlsx').snapshot)
        expect(externalLinkPackageSummary(unzipSync(exported))).toEqual(sourceExternalLinkSummary)
        expect(worksheetFormulaCacheValues(exported)).toEqual(sourceFormulaCacheValues)
      } finally {
        rmSync(tempDir, { recursive: true, force: true })
      }
    },
    150_000,
  )

  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'round-trips cached external ranges through Desktop Excel and Bilig recalc',
    async () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = createExcelAccessibleTempDir('bilig-headless-excel-external-link-range-cache-')
      try {
        const linkedSourceWorkbookPath = join(tempDir, 'rates.xlsx')
        const sourceWorkbookPath = join(tempDir, 'external-link-range-cache.xlsx')
        writeFileSync(linkedSourceWorkbookPath, buildExternalSourceWorkbook())
        writeFileSync(sourceWorkbookPath, buildExternalLinkRangeCacheWorkbook(pathToFileURL(linkedSourceWorkbookPath).href))

        const excelInitial = runMacosExcelInspectionOracle({
          workbookPath: sourceWorkbookPath,
          worksheetName: 'Model',
          formulaCells: [],
          inspectCells: [...externalRangeAddresses],
          companionWorkbookPaths: [linkedSourceWorkbookPath],
          saveWorkbook: true,
          updateLinks: 'external',
        })
        expect(excelInitial.cells.map(({ address, value }) => ({ address, value }))).toEqual(initialExternalRangeValues)

        const imported = importXlsx(new Uint8Array(readFileSync(sourceWorkbookPath)), 'external-link-range-cache-saved.xlsx')
        const sourceExternalLinkSummary = externalLinkPackageSummary(unzipSync(new Uint8Array(readFileSync(sourceWorkbookPath))))
        const engine = new SpreadsheetEngine({ workbookName: 'external-link-range-cache-oracle' })
        await engine.ready()
        engine.importSnapshot(imported.snapshot)
        engine.setCellValue('Model', 'B1', 3)

        expectEngineValues(engine, changedExternalRangeValues)

        const exportedWorkbookPath = join(tempDir, 'external-link-range-cache-materialized.xlsx')
        writeFileSync(exportedWorkbookPath, exportXlsx(engine.exportSnapshot()))
        expect(externalLinkPackageSummary(unzipSync(new Uint8Array(readFileSync(exportedWorkbookPath))))).toEqual(sourceExternalLinkSummary)

        const excelChanged = runMacosExcelInspectionOracle({
          workbookPath: exportedWorkbookPath,
          worksheetName: 'Model',
          formulaCells: [],
          inspectCells: [...externalRangeAddresses],
          saveWorkbook: false,
        })

        expect(excelChanged.cells.map(({ address, value }) => ({ address, value }))).toEqual(changedExternalRangeValues)
      } finally {
        rmSync(tempDir, { recursive: true, force: true })
      }
    },
    60_000,
  )

  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'matches Desktop Excel external-link cleanup after deleting the formula sheet',
    async () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = createExcelAccessibleTempDir('bilig-headless-excel-external-link-delete-')
      try {
        const linkedSourceWorkbookPath = join(tempDir, 'rates.xlsx')
        const sourceWorkbookPath = join(tempDir, 'external-link-delete-source.xlsx')
        writeFileSync(linkedSourceWorkbookPath, buildExternalSourceWorkbook())
        const sourceBytes = buildExternalLinkRangeCacheWorkbook(pathToFileURL(linkedSourceWorkbookPath).href, {
          includeKeepSheet: true,
        })
        writeFileSync(sourceWorkbookPath, sourceBytes)

        const excelDeleted = runMacosExcelStructuralOperationOracle({
          workbookPath: sourceWorkbookPath,
          worksheetName: 'Keep',
          operations: [{ kind: 'deleteSheet', name: 'Model' }],
          inspectCells: ['A1'],
          companionWorkbookPaths: [linkedSourceWorkbookPath],
          saveWorkbook: true,
          updateLinks: 'external',
          timeoutMs: 120_000,
        })
        expect(excelDeleted.cells[0]?.value).toEqual({ kind: 'string', value: 'keep' })

        const excelTruthBytes = new Uint8Array(readFileSync(sourceWorkbookPath))
        const excelTruthMetrics = externalLinkPackageMetrics(unzipSync(excelTruthBytes))
        expect(excelTruthMetrics).toEqual({
          packageParts: [],
          workbookExternalReferenceTargets: [],
          externalLinkPathRelationshipParts: [],
          contentTypeOverrides: [],
        })

        const imported = importXlsx(sourceBytes, 'external-link-delete-source.xlsx')
        const engine = new SpreadsheetEngine({ workbookName: 'external-link-delete-oracle' })
        await engine.ready()
        engine.importSnapshot(imported.snapshot)
        engine.deleteSheet('Model')

        const headlessSnapshot = engine.exportSnapshot()
        expect(headlessSnapshot.sheets.map((sheet) => sheet.name)).toEqual(['Keep'])

        const exportedWorkbookPath = join(tempDir, 'external-link-delete-headless.xlsx')
        writeFileSync(exportedWorkbookPath, exportXlsx(headlessSnapshot))
        expect(externalLinkPackageMetrics(unzipSync(new Uint8Array(readFileSync(exportedWorkbookPath))))).toEqual(excelTruthMetrics)

        const headlessExcel = runMacosExcelInspectionOracle({
          workbookPath: exportedWorkbookPath,
          worksheetName: 'Keep',
          formulaCells: [],
          inspectCells: ['A1'],
          saveWorkbook: true,
          timeoutMs: 90_000,
        })
        expect(headlessExcel.cells).toEqual(excelDeleted.cells)
      } finally {
        rmSync(tempDir, { recursive: true, force: true })
      }
    },
    150_000,
  )
})

function expectEngineValues(
  engine: SpreadsheetEngine,
  expectedValues: readonly { readonly address: string; readonly value: NormalizedFormulaValue }[],
): void {
  expect(expectedValues.map(({ address }) => ({ address, value: normalizedCellValue(engine.getCellValue('Model', address)) }))).toEqual(
    expectedValues,
  )
}

function normalizedCellValue(value: CellValue): NormalizedFormulaValue {
  switch (value.tag) {
    case ValueTag.Empty:
      return { kind: 'blank' }
    case ValueTag.Boolean:
      return { kind: 'boolean', value: value.value }
    case ValueTag.Error:
      return { kind: 'error', value: String(value.code) }
    case ValueTag.Number:
      return { kind: 'number', value: value.value }
    case ValueTag.String:
      return { kind: 'string', value: value.value }
  }
}

function buildExternalSourceWorkbook(): Uint8Array {
  return writeSimpleXlsxWorkbook({
    sheets: [
      {
        name: 'Rates',
        cells: [
          { address: 'A1', row: 0, col: 0, value: 'SKU' },
          { address: 'B1', row: 0, col: 1, value: 'Rate' },
          { address: 'A2', row: 1, col: 0, value: 'A' },
          { address: 'B2', row: 1, col: 1, value: 10 },
          { address: 'A3', row: 2, col: 0, value: 'B' },
          { address: 'B3', row: 2, col: 1, value: 20 },
          { address: 'A4', row: 3, col: 0, value: 'C' },
          { address: 'B4', row: 3, col: 1, value: 30 },
        ],
      },
    ],
  })
}

function createExcelAccessibleTempDir(prefix: string): string {
  const root = join(tmpdir(), 'bilig-headless-oracle')
  mkdirSync(root, { recursive: true })
  return mkdtempSync(join(root, prefix))
}

function buildExternalLinkRangeCacheWorkbook(
  target = 'file:///tmp/rates.xlsx',
  options: { readonly includeKeepSheet?: boolean } = {},
): Uint8Array {
  const zip = unzipSync(
    writeSimpleXlsxWorkbook({
      sheets: [
        ...(options.includeKeepSheet === true
          ? [
              {
                name: 'Keep',
                cells: [{ address: 'A1', row: 0, col: 0, value: 'keep' }],
              },
            ]
          : []),
        {
          name: 'Model',
          cells: [
            { address: 'B1', row: 0, col: 1, value: 2 },
            { address: 'C1', row: 0, col: 2, formula: "SUM('[1]Rates'!$B$2:$B$4)*B1", value: 120 },
            {
              address: 'C2',
              row: 1,
              col: 2,
              formula: "_xlfn.XLOOKUP(\"B\",'[1]Rates'!$A$2:$A$4,'[1]Rates'!$B$2:$B$4)*B1",
              value: 40,
            },
            { address: 'C3', row: 2, col: 2, formula: "SUMIFS('[1]Rates'!$B$2:$B$4,'[1]Rates'!$A$2:$A$4,\"C\")*B1", value: 60 },
          ],
        },
      ],
    }),
  )
  zip['xl/workbook.xml'] = strToU8(
    ensureRelationshipNamespace(xmlText(zip, 'xl/workbook.xml')).replace(
      '</sheets>',
      '</sheets><externalReferences><externalReference r:id="rId99"/></externalReferences>',
    ),
  )
  zip['xl/_rels/workbook.xml.rels'] = strToU8(
    xmlText(zip, 'xl/_rels/workbook.xml.rels').replace(
      '</Relationships>',
      '<Relationship Id="rId99" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" Target="externalLinks/externalLink5.xml"/></Relationships>',
    ),
  )
  zip['xl/externalLinks/externalLink5.xml'] = strToU8(
    [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      `<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${officeRelationshipNamespace}">`,
      '<externalBook r:id="rId1">',
      '<sheetNames><sheetName val="Rates"/></sheetNames>',
      '<sheetDataSet><sheetData sheetId="0">',
      '<row r="1"><cell r="A1" t="str"><v>SKU</v></cell><cell r="B1" t="str"><v>Rate</v></cell></row>',
      '<row r="2"><cell r="A2" t="str"><v>A</v></cell><cell r="B2"><v>10</v></cell></row>',
      '<row r="3"><cell r="A3" t="str"><v>B</v></cell><cell r="B3"><v>20</v></cell></row>',
      '<row r="4"><cell r="A4" t="str"><v>C</v></cell><cell r="B4"><v>30</v></cell></row>',
      '</sheetData></sheetDataSet>',
      '</externalBook>',
      '</externalLink>',
    ].join(''),
  )
  zip['xl/externalLinks/_rels/externalLink5.xml.rels'] = strToU8(
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath" Target="${target}" TargetMode="External"/>` +
      '</Relationships>',
  )
  zip['[Content_Types].xml'] = strToU8(
    upsertContentTypeOverride(xmlText(zip, '[Content_Types].xml'), {
      partName: '/xl/externalLinks/externalLink5.xml',
      contentType: externalLinkContentType,
    }),
  )
  return zipSync(zip)
}

function externalLinkPackageSummary(zip: Record<string, Uint8Array>): {
  readonly packageParts: readonly (readonly [path: string, xml: string])[]
  readonly workbookExternalReferenceTargets: readonly string[]
  readonly externalLinkPathRelationshipParts: readonly (readonly [path: string, xml: string])[]
  readonly contentTypeOverrides: readonly string[]
} {
  const workbookExternalReferencesXml = extractSingleXml(
    xmlText(zip, 'xl/workbook.xml'),
    /<(?:[A-Za-z_][\w.-]*:)?externalReferences\b[^>]*>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?externalReferences>/u,
  )
  const workbookExternalLinkTargets = workbookExternalLinkTargetsByRelationshipId(xmlText(zip, 'xl/_rels/workbook.xml.rels'))
  return {
    packageParts: Object.entries(zip)
      .filter(([path]) => /^xl\/externalLinks\/externalLink[^/]*\.xml$/u.test(path))
      .map(([path, bytes]) => [path, strFromU8(bytes)] as const)
      .toSorted(([left], [right]) => left.localeCompare(right)),
    workbookExternalReferenceTargets: extractCaptureMatches(workbookExternalReferencesXml, /\br:id="([^"]+)"/gu).map((id) => {
      const target = workbookExternalLinkTargets.get(id)
      if (target === undefined) {
        throw new Error(`Missing workbook external-link relationship target for ${id}`)
      }
      return target
    }),
    externalLinkPathRelationshipParts: Object.entries(zip)
      .filter(([path]) => /^xl\/externalLinks\/_rels\/externalLink[^/]*\.xml\.rels$/u.test(path))
      .map(([path, bytes]) => [path, strFromU8(bytes)] as const)
      .toSorted(([left], [right]) => left.localeCompare(right)),
    contentTypeOverrides: extractXmlMatches(
      xmlText(zip, '[Content_Types].xml'),
      /<Override\b[^>]*ContentType="application\/vnd\.openxmlformats-officedocument\.spreadsheetml\.externalLink\+xml"[^>]*\/>/gu,
    ),
  }
}

function externalLinkPackageMetrics(zip: Record<string, Uint8Array>): {
  readonly packageParts: readonly (readonly [path: string, xml: string])[]
  readonly workbookExternalReferenceTargets: readonly string[]
  readonly externalLinkPathRelationshipParts: readonly (readonly [path: string, xml: string])[]
  readonly contentTypeOverrides: readonly string[]
} {
  const workbookExternalReferencesXml = extractSingleXmlOptional(
    xmlText(zip, 'xl/workbook.xml'),
    /<(?:[A-Za-z_][\w.-]*:)?externalReferences\b[^>]*>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?externalReferences>/u,
  )
  const workbookExternalLinkTargets = workbookExternalLinkTargetsByRelationshipIdOptional(xmlText(zip, 'xl/_rels/workbook.xml.rels'))
  return {
    packageParts: Object.entries(zip)
      .filter(([path]) => /^xl\/externalLinks\/externalLink[^/]*\.xml$/u.test(path))
      .map(([path, bytes]) => [path, strFromU8(bytes)] as const)
      .toSorted(([left], [right]) => left.localeCompare(right)),
    workbookExternalReferenceTargets: workbookExternalReferencesXml
      ? extractCaptureMatchesOptional(workbookExternalReferencesXml, /\br:id="([^"]+)"/gu).flatMap((id) => {
          const target = workbookExternalLinkTargets.get(id)
          return target === undefined ? [] : [target]
        })
      : [],
    externalLinkPathRelationshipParts: Object.entries(zip)
      .filter(([path]) => /^xl\/externalLinks\/_rels\/externalLink[^/]*\.xml\.rels$/u.test(path))
      .map(([path, bytes]) => [path, strFromU8(bytes)] as const)
      .toSorted(([left], [right]) => left.localeCompare(right)),
    contentTypeOverrides: extractXmlMatchesOptional(
      xmlText(zip, '[Content_Types].xml'),
      /<Override\b[^>]*ContentType="application\/vnd\.openxmlformats-officedocument\.spreadsheetml\.externalLink\+xml"[^>]*\/>/gu,
    ),
  }
}

function worksheetFormulaCacheValues(bytes: Uint8Array): Readonly<Record<string, string | null>> {
  const sheetXml = xmlText(unzipSync(bytes), 'xl/worksheets/sheet1.xml')
  return Object.fromEntries(
    externalRangeAddresses.map((address) => {
      const cellXml = new RegExp(`<c\\b(?=[^>]*\\br="${address}")[\\s\\S]*?<\\/c>`, 'u').exec(sheetXml)?.[0]
      if (!cellXml) {
        throw new Error(`Missing formula cell ${address}`)
      }
      return [address, /<v>([\s\S]*?)<\/v>/u.exec(cellXml)?.[1] ?? null]
    }),
  )
}

function extractSingleXmlOptional(xml: string, pattern: RegExp): string | undefined {
  return xml.match(pattern)?.[0]
}

function extractSingleXml(xml: string, pattern: RegExp): string {
  const match = xml.match(pattern)
  if (!match) {
    throw new Error(`Missing expected XLSX XML fragment for pattern ${String(pattern)}`)
  }
  return match[0]
}

function extractXmlMatchesOptional(xml: string, pattern: RegExp): readonly string[] {
  return [...xml.matchAll(pattern)].map((match) => match[0]).toSorted()
}

function extractXmlMatches(xml: string, pattern: RegExp): readonly string[] {
  const matches = [...xml.matchAll(pattern)].map((match) => match[0])
  if (matches.length === 0) {
    throw new Error(`Missing expected XLSX XML fragment for pattern ${String(pattern)}`)
  }
  return matches.toSorted()
}

function extractCaptureMatchesOptional(xml: string, pattern: RegExp): readonly string[] {
  return [...xml.matchAll(pattern)].flatMap((match) => match[1] ?? [])
}

function extractCaptureMatches(xml: string, pattern: RegExp): readonly string[] {
  const matches = [...xml.matchAll(pattern)].flatMap((match) => match[1] ?? [])
  if (matches.length === 0) {
    throw new Error(`Missing expected XLSX XML capture for pattern ${String(pattern)}`)
  }
  return matches
}

function workbookExternalLinkTargetsByRelationshipIdOptional(workbookRelationshipsXml: string): ReadonlyMap<string, string> {
  const targetsById = new Map<string, string>()
  for (const relationship of extractXmlMatchesOptional(workbookRelationshipsXml, /<Relationship\b[^>]*\/>/gu)) {
    if (!relationship.includes('/relationships/externalLink"')) {
      continue
    }
    const id = extractAttribute(relationship, 'Id')
    const target = extractAttribute(relationship, 'Target')
    targetsById.set(id, target)
  }
  return targetsById
}

function workbookExternalLinkTargetsByRelationshipId(workbookRelationshipsXml: string): ReadonlyMap<string, string> {
  const targetsById = new Map<string, string>()
  for (const relationship of extractXmlMatches(workbookRelationshipsXml, /<Relationship\b[^>]*\/>/gu)) {
    if (!relationship.includes('/relationships/externalLink"')) {
      continue
    }
    const id = extractAttribute(relationship, 'Id')
    const target = extractAttribute(relationship, 'Target')
    targetsById.set(id, target)
  }
  if (targetsById.size === 0) {
    throw new Error('Missing workbook external-link relationships')
  }
  return targetsById
}

function extractAttribute(xml: string, attributeName: string): string {
  const match = new RegExp(`\\b${attributeName}="([^"]+)"`, 'u').exec(xml)
  if (!match) {
    throw new Error(`Missing ${attributeName} attribute in ${xml}`)
  }
  return match[1]
}

function ensureRelationshipNamespace(xml: string): string {
  if (/xmlns:r=/u.test(xml)) {
    return xml
  }
  return xml.replace(/<workbook\b([^>]*)>/u, `<workbook$1 xmlns:r="${officeRelationshipNamespace}">`)
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

function xmlText(zip: Record<string, Uint8Array>, path: string): string {
  const bytes = zip[path]
  if (!bytes) {
    throw new Error(`Missing XLSX part: ${path}`)
  }
  return strFromU8(bytes)
}
