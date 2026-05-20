import { mkdirSync, mkdtempSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { strToU8, zipSync } from 'fflate'
import { describe, expect, it } from 'vitest'

import {
  buildPublicWorkbookCorpusScorecard,
  createEmptyPublicWorkbookManifest,
  sha256Hex,
  type PublicWorkbookManifest,
} from '../public-workbook-corpus.ts'
import { publicWorkbookResourceLimitClassifierEvidence } from '../public-workbook-corpus-evidence.ts'
import {
  shouldUseCompactLargeSimpleVerification,
  verifyLargeSimpleWorkbookCompactPreflight,
} from '../public-workbook-corpus-large-simple-compact.ts'
import { startVerificationRuntimeMetrics } from '../public-workbook-corpus-verification-metrics.ts'
import type { PublicWorkbookArtifact, PublicWorkbookFeatureCounts } from '../public-workbook-corpus-types.ts'
import type { WorkbookFootprint } from '../public-workbook-corpus-workbook.ts'

describe('public workbook corpus compact verifier', () => {
  it('uses headless large-simple verification when public arrays are not needed', async () => {
    const workbookBytes = buildLargeSimpleNumericWorkbookBytes(200_001)
    const scorecard = await buildSingleWorkbookScorecard({
      cacheDirPrefix: 'public-workbook-corpus-compact-verifier-',
      fileName: 'large-simple-compact.xlsx',
      sourceId: 'source-large-simple-compact',
      workbookBytes,
    })
    const corpusCase = scorecard.cases[0]

    expect(corpusCase?.passed).toBe(true)
    expect(corpusCase?.status).toBe('unsupported')
    expect(corpusCase?.featureCounts).toMatchObject({
      sheetCount: 1,
      cellCount: 200_001,
      formulaCellCount: 0,
      valueCellCount: 200_001,
    })
    expect(corpusCase?.workbookMetadata.dimensions[0]).toEqual({
      sheetName: 'Data',
      rowCount: 200_001,
      columnCount: 1,
      nonEmptyCellCount: 200_001,
      usedRange: {
        startRow: 0,
        startColumn: 0,
        endRow: 200_000,
        endColumn: 0,
      },
    })
    expect(corpusCase?.phaseTimings?.map((entry) => entry.phase)).toEqual(['read-cache', 'import-xlsx'])
    expect(corpusCase?.unsupportedFeatureClassifications).toEqual([
      'xlsx.publicCorpus.resourceLimit:preflightRoundTripBudget>100000cells',
      'xlsx.publicCorpus.resourceLimit:preflightStructuralSmokeBudget>100000cells',
    ])
    expect(corpusCase?.evidence).toEqual(
      expect.arrayContaining([
        publicWorkbookResourceLimitClassifierEvidence,
        expect.stringContaining('large-simple-import-phase=zip-source-release'),
        expect.stringContaining('Round-trip projection skipped because workbook footprint exceeds verifier resource budget'),
        expect.stringContaining('Structural smoke skipped because workbook footprint exceeds verifier resource budget'),
      ]),
    )
  })

  it('uses headless large-simple verification for formula-heavy workbooks when formula oracle is resource-skipped', async () => {
    const workbookBytes = buildLargeSimpleFormulaHeavyWorkbookBytes({ rowCount: 100_001, formulaRowCount: 2_001 })
    expect(workbookBytes.byteLength).toBeLessThan(1_000_000)
    const scorecard = await buildSingleWorkbookScorecard({
      cacheDirPrefix: 'public-workbook-corpus-compact-formula-verifier-',
      fileName: 'large-simple-formula-compact.xlsx',
      sourceId: 'source-large-simple-formula-compact',
      workbookBytes,
    })
    const corpusCase = scorecard.cases[0]

    expect(corpusCase?.passed).toBe(true)
    expect(corpusCase?.status).toBe('unsupported')
    expect(corpusCase?.featureCounts).toMatchObject({
      sheetCount: 1,
      cellCount: 102_002,
      formulaCellCount: 2_001,
      valueCellCount: 102_002,
    })
    expect(corpusCase?.phaseTimings?.map((entry) => entry.phase)).toEqual(['read-cache', 'import-xlsx'])
    expect(corpusCase?.unsupportedFeatureClassifications).toEqual([
      'xlsx.publicCorpus.resourceLimit:preflightFormulaOracleBudget>2000formulas',
      'xlsx.publicCorpus.resourceLimit:preflightRoundTripBudget>100000cells',
      'xlsx.publicCorpus.resourceLimit:preflightStructuralSmokeBudget>100000cells',
    ])
    expect(corpusCase?.evidence).toEqual(
      expect.arrayContaining([
        publicWorkbookResourceLimitClassifierEvidence,
        expect.stringContaining('large-simple-import-phase=zip-source-release'),
        expect.stringContaining('Formula oracle skipped because workbook has 2001 formulas'),
        expect.stringContaining('Round-trip projection skipped because workbook footprint exceeds verifier resource budget'),
        expect.stringContaining('Structural smoke skipped because workbook footprint exceeds verifier resource budget'),
      ]),
    )
  })

  it('keeps formula-light large-simple workbooks on the full verifier when formula oracle evidence is required', () => {
    expect(
      shouldUseCompactLargeSimpleVerification(
        testArtifact({ byteSize: 1_000 }),
        testFootprint({
          cellCount: 100_001,
          formulaCellCount: 2_000,
          valueCellCount: 100_001,
        }),
        true,
      ),
    ).toBe(false)
  })

  it('fails closed to the regular verifier when compact preflight cannot read a ZIP central directory', () => {
    expect(
      verifyLargeSimpleWorkbookCompactPreflight({
        artifact: testArtifact({ byteSize: 4 }),
        bytes: new Uint8Array([0x50, 0x4b, 0x03, 0x04]),
        baseEvidence: [],
        classifyUnsupportedFeatures: () => [],
        maxCellCount: 1_500_000,
        minByteLength: 0,
        runStructuralSmoke: true,
        runtimeMetrics: startVerificationRuntimeMetrics(),
        workerOptions: {},
      }),
    ).toBeNull()
  })
})

async function buildSingleWorkbookScorecard(args: {
  readonly cacheDirPrefix: string
  readonly fileName: string
  readonly sourceId: string
  readonly workbookBytes: Uint8Array
}) {
  const cacheDir = mkdtempSync(join(tmpdir(), args.cacheDirPrefix))
  mkdirSync(join(cacheDir, 'files'), { recursive: true })
  const sha256 = await sha256Hex(args.workbookBytes)
  writeFileSync(join(cacheDir, 'files', `${sha256}.xlsx`), args.workbookBytes)
  const license = {
    spdxId: 'CC-BY-4.0',
    title: 'Creative Commons Attribution 4.0 International',
    evidenceUrl: 'https://creativecommons.org/licenses/by/4.0/',
  }
  const sourceUrl = `https://example.com/${args.fileName}`
  const manifest: PublicWorkbookManifest = {
    ...createEmptyPublicWorkbookManifest('2026-05-07T00:00:00.000Z'),
    sources: [
      {
        id: args.sourceId,
        kind: 'direct-url',
        sourceUrl,
        downloadUrl: sourceUrl,
        fileName: args.fileName,
        discoveredAt: '2026-05-07T00:00:00.000Z',
        license,
      },
    ],
    artifacts: [
      {
        id: `workbook-${sha256.slice(0, 16)}`,
        sourceId: args.sourceId,
        sourceUrl,
        downloadUrl: sourceUrl,
        fileName: args.fileName,
        cachePath: `files/${sha256}.xlsx`,
        sha256,
        byteSize: args.workbookBytes.byteLength,
        workbookFingerprint: `${args.sourceId}-fingerprint`,
        fetchedAt: '2026-05-07T00:00:00.000Z',
        license,
      },
    ],
  }

  return buildPublicWorkbookCorpusScorecard({
    manifest,
    cacheDir,
    generatedAt: '2026-05-07T01:00:00.000Z',
  })
}

function buildLargeSimpleNumericWorkbookBytes(rowCount: number): Uint8Array {
  const rows: string[] = []
  for (let row = 1; row <= rowCount; row += 1) {
    rows.push(`<row r="${String(row)}"><c r="A${String(row)}"><v>${String(row)}</v></c></row>`)
  }
  return zipSync({
    'xl/workbook.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
</workbook>`),
    'xl/_rels/workbook.xml.rels': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`),
    'xl/worksheets/sheet1.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:A${String(rowCount)}"/>
  <sheetData>${rows.join('')}</sheetData>
</worksheet>`),
  })
}

function buildLargeSimpleFormulaHeavyWorkbookBytes(input: { readonly rowCount: number; readonly formulaRowCount: number }): Uint8Array {
  const rows: string[] = []
  for (let row = 1; row <= input.rowCount; row += 1) {
    if (row <= input.formulaRowCount) {
      rows.push(
        `<row r="${String(row)}"><c r="A${String(row)}"><f>B${String(row)}+1</f><v>${String(row + 1)}</v></c><c r="B${String(
          row,
        )}"><v>${String(row)}</v></c></row>`,
      )
    } else {
      rows.push(`<row r="${String(row)}"><c r="A${String(row)}"><v>${String(row)}</v></c></row>`)
    }
  }
  return zipSync({
    'xl/workbook.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
</workbook>`),
    'xl/_rels/workbook.xml.rels': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`),
    'xl/worksheets/sheet1.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:B${String(input.rowCount)}"/>
  <sheetData>${rows.join('')}</sheetData>
</worksheet>`),
  })
}

function testArtifact(input: { readonly byteSize: number }): PublicWorkbookArtifact {
  const license = {
    spdxId: 'CC-BY-4.0',
    title: 'Creative Commons Attribution 4.0 International',
    evidenceUrl: 'https://creativecommons.org/licenses/by/4.0/',
  }
  return {
    id: 'workbook-test',
    sourceId: 'source-test',
    sourceUrl: 'https://example.com/test.xlsx',
    downloadUrl: 'https://example.com/test.xlsx',
    fileName: 'test.xlsx',
    cachePath: 'files/test.xlsx',
    sha256: 'a'.repeat(64),
    byteSize: input.byteSize,
    workbookFingerprint: 'test-fingerprint',
    fetchedAt: '2026-05-07T00:00:00.000Z',
    license,
  }
}

function testFootprint(overrides: Partial<PublicWorkbookFeatureCounts>): WorkbookFootprint {
  const featureCounts: PublicWorkbookFeatureCounts = {
    sheetCount: 1,
    cellCount: 0,
    formulaCellCount: 0,
    valueCellCount: 0,
    definedNameCount: 0,
    tableCount: 0,
    chartCount: 0,
    pivotCount: 0,
    mergeCount: 0,
    styleRangeCount: 0,
    conditionalFormatCount: 0,
    dataValidationCount: 0,
    macroPayloadCount: 0,
    warningCount: 0,
    ...overrides,
  }
  return {
    featureCounts,
    workbookMetadata: {
      workbookName: 'test',
      sheetNames: ['Data'],
      dimensions: [],
    },
    externalWorkbookReferences: [],
    largeSimpleXlsxImport: { eligible: true, blockers: [] },
  }
}
