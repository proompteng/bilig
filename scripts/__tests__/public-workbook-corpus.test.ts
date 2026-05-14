import { EventEmitter } from 'node:events'
import { mkdtempSync, mkdirSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, describe, expect, it, vi } from 'vitest'
import * as XLSX from 'xlsx'

import {
  externalWorkbookReferencesWarning,
  externalPivotCachesWarning,
  macroExecutionDeclinedWarning,
  manualCalculationModeWarning,
  precisionAsDisplayedCalculationWarning,
  volatileFormulasWarning,
} from '../../packages/excel-import/src/index.js'
import { addExportCalculationSettingsToXlsxBytes } from '../../packages/excel-import/src/xlsx-calculation-settings.js'
import type { WorkbookSnapshot } from '../../packages/protocol/src/types.js'
import {
  buildPublicWorkbookCorpusScorecard,
  createEmptyPublicWorkbookManifest,
  fetchPublicWorkbookArtifacts,
  discoverCkanWorkbookSources,
  parsePublicWorkbookManifestJson,
  sha256Hex,
  validatePublicWorkbookCorpusScorecard,
  validatePublicWorkbookManifest,
  type PublicWorkbookCorpusFetchCheckpointProgress,
  type PublicWorkbookManifest,
  type PublicWorkbookSource,
} from '../public-workbook-corpus.ts'
import {
  publicWorkbookFormulaOracleCacheClassifierEvidence,
  publicWorkbookImportWarningClassifierEvidence,
  publicWorkbookResourceLimitClassifierEvidence,
} from '../public-workbook-corpus-evidence.ts'
import { fingerprintWorkbookFileIsolated } from '../public-workbook-corpus-fetch.ts'
import { roundTripSemanticsDigest } from '../public-workbook-corpus-roundtrip.ts'
import {
  classifyUnsupportedFeatures,
  externalPivotCacheUnsupportedClassification,
  rawPivotPartUnsupportedClassification,
  staleFormulaCacheUnsupportedClassification,
} from '../public-workbook-corpus-verify.ts'

const spawnMock = vi.hoisted(() => vi.fn())

vi.mock('node:child_process', () => ({
  spawn: spawnMock,
}))

describe('public workbook corpus', () => {
  afterEach(() => {
    vi.useRealTimers()
    vi.restoreAllMocks()
    vi.clearAllMocks()
    vi.unstubAllGlobals()
  })

  it('validates source license evidence before a workbook can enter the corpus', () => {
    const manifest = createEmptyPublicWorkbookManifest('2026-05-07T00:00:00.000Z')
    const withMissingLicense: PublicWorkbookManifest = {
      ...manifest,
      sources: [
        {
          id: 'source-missing-license',
          kind: 'direct-url',
          sourceUrl: 'https://example.com/workbook.xlsx',
          downloadUrl: 'https://example.com/workbook.xlsx',
          fileName: 'workbook.xlsx',
          discoveredAt: '2026-05-07T00:00:00.000Z',
          license: {
            spdxId: null,
            title: '',
            evidenceUrl: null,
          },
        },
      ],
    }

    expect(() => validatePublicWorkbookManifest(withMissingLicense)).toThrow(
      'Public workbook source source-missing-license is missing usable license evidence',
    )
  })

  it('validates persisted fetch state against manifest sources', () => {
    const manifest: PublicWorkbookManifest = {
      ...createEmptyPublicWorkbookManifest('2026-05-07T00:00:00.000Z'),
      sources: [
        {
          id: 'source-1',
          kind: 'direct-url',
          sourceUrl: 'https://example.com/workbook.xlsx',
          downloadUrl: 'https://example.com/workbook.xlsx',
          fileName: 'workbook.xlsx',
          discoveredAt: '2026-05-07T00:00:00.000Z',
          license: {
            spdxId: 'CC-BY-4.0',
            title: 'Creative Commons Attribution 4.0 International',
            evidenceUrl: 'https://creativecommons.org/licenses/by/4.0/',
          },
        },
      ],
      fetchState: {
        exhaustedSourceIds: ['source-1'],
      },
    }

    validatePublicWorkbookManifest(manifest)
    expect(parsePublicWorkbookManifestJson(manifest).fetchState?.exhaustedSourceIds).toEqual(['source-1'])
    expect(() =>
      validatePublicWorkbookManifest({
        ...manifest,
        fetchState: { exhaustedSourceIds: ['source-1', 'source-1'] },
      }),
    ).toThrow('Duplicate exhausted public workbook source id: source-1')
    expect(() =>
      validatePublicWorkbookManifest({
        ...manifest,
        fetchState: { exhaustedSourceIds: ['source-missing'] },
      }),
    ).toThrow('Exhausted public workbook source source-missing is not in the manifest')
  })

  it('supports focused corpus slices with a custom target workbook count', async () => {
    const manifest = createEmptyPublicWorkbookManifest('2026-05-07T00:00:00.000Z', 5_000)
    const cacheDir = mkdtempSync(join(tmpdir(), 'public-workbook-corpus-target-'))

    validatePublicWorkbookManifest(manifest)
    const scorecard = await buildPublicWorkbookCorpusScorecard({
      manifest,
      cacheDir,
      generatedAt: '2026-05-07T01:00:00.000Z',
    })

    expect(scorecard.summary).toMatchObject({
      targetWorkbookCount: 5_000,
      cachedWorkbookCount: 0,
      allCachedWorkbooksPassed: true,
      remainingToTarget: 5_000,
    })
    validatePublicWorkbookCorpusScorecard(scorecard)
  })

  it('builds an offline scorecard from cached spreadsheet artifacts', async () => {
    const cacheDir = mkdtempSync(join(tmpdir(), 'public-workbook-corpus-test-'))
    mkdirSync(join(cacheDir, 'files'), { recursive: true })
    const workbookBytes = buildWorkbookBytes()
    const sha256 = await sha256Hex(workbookBytes)
    writeFileSync(join(cacheDir, 'files', `${sha256}.xlsx`), workbookBytes)

    const manifest: PublicWorkbookManifest = {
      ...createEmptyPublicWorkbookManifest('2026-05-07T00:00:00.000Z'),
      sources: [
        {
          id: 'source-1',
          kind: 'direct-url',
          sourceUrl: 'https://example.com/public-budget.xlsx',
          downloadUrl: 'https://example.com/public-budget.xlsx',
          fileName: 'public-budget.xlsx',
          discoveredAt: '2026-05-07T00:00:00.000Z',
          license: {
            spdxId: 'CC-BY-4.0',
            title: 'Creative Commons Attribution 4.0 International',
            evidenceUrl: 'https://creativecommons.org/licenses/by/4.0/',
          },
        },
      ],
      artifacts: [
        {
          id: `workbook-${sha256.slice(0, 16)}`,
          sourceId: 'source-1',
          sourceUrl: 'https://example.com/public-budget.xlsx',
          downloadUrl: 'https://example.com/public-budget.xlsx',
          fileName: 'public-budget.xlsx',
          cachePath: `files/${sha256}.xlsx`,
          sha256,
          byteSize: workbookBytes.byteLength,
          workbookFingerprint: 'test-fingerprint',
          fetchedAt: '2026-05-07T00:00:00.000Z',
          license: {
            spdxId: 'CC-BY-4.0',
            title: 'Creative Commons Attribution 4.0 International',
            evidenceUrl: 'https://creativecommons.org/licenses/by/4.0/',
          },
        },
      ],
    }

    const scorecard = await buildPublicWorkbookCorpusScorecard({
      manifest,
      cacheDir,
      generatedAt: '2026-05-07T01:00:00.000Z',
      structuralSmokeSampleLimit: 1,
    })

    expect(scorecard).toMatchObject({
      schemaVersion: 1,
      suite: 'public-workbook-corpus',
      generatedAt: '2026-05-07T01:00:00.000Z',
      summary: {
        targetWorkbookCount: 10_000,
        sourceCount: 1,
        cachedWorkbookCount: 1,
        importedWorkbookCount: 1,
        formulaOracleComparisonCount: 1,
        formulaOracleMatchCount: 1,
        structuralSmokeRunCount: 1,
        allCachedWorkbooksPassed: true,
        remainingToTarget: 9_999,
      },
    })
    expect(scorecard.cases).toHaveLength(1)
    expect(scorecard.cases[0]).toMatchObject({
      id: `workbook-${sha256.slice(0, 16)}`,
      status: 'passed',
      passed: true,
      featureCounts: {
        sheetCount: 2,
        formulaCellCount: 1,
        definedNameCount: 1,
        mergeCount: 1,
      },
      validation: {
        importPassed: true,
        formulaOraclePassed: true,
        roundTripPassed: true,
        structuralSmokePassed: true,
      },
    })
    expect(scorecard.cases[0]?.unsupportedFeatureClassifications).toEqual([])
    validatePublicWorkbookCorpusScorecard(scorecard)
  })

  it('classifies raw XLSX pivot parts when semantic import cannot project them', () => {
    const classifications = classifyUnsupportedFeatures(
      { version: 1, workbook: { name: 'raw-pivot' }, sheets: [{ id: 1, name: 'Sheet1', order: 0, cells: [] }] },
      [],
      {
        sheetCount: 1,
        cellCount: 0,
        formulaCellCount: 0,
        valueCellCount: 0,
        definedNameCount: 0,
        tableCount: 0,
        chartCount: 0,
        pivotCount: 1,
        mergeCount: 0,
        styleRangeCount: 0,
        conditionalFormatCount: 0,
        dataValidationCount: 0,
        macroPayloadCount: 0,
        warningCount: 0,
      },
    )

    expect(classifications).toEqual([rawPivotPartUnsupportedClassification])
  })

  it('classifies external-cache pivot parts separately from generic raw pivot parts', () => {
    const classifications = classifyUnsupportedFeatures(
      { version: 1, workbook: { name: 'external-cache-pivot' }, sheets: [{ id: 1, name: 'Sheet1', order: 0, cells: [] }] },
      [externalPivotCachesWarning],
      {
        sheetCount: 1,
        cellCount: 0,
        formulaCellCount: 0,
        valueCellCount: 0,
        definedNameCount: 0,
        tableCount: 0,
        chartCount: 0,
        pivotCount: 1,
        mergeCount: 0,
        styleRangeCount: 0,
        conditionalFormatCount: 0,
        dataValidationCount: 0,
        macroPayloadCount: 0,
        warningCount: 1,
      },
    )

    expect(classifications).toEqual([`xlsx.import.warning:${externalPivotCachesWarning}`, externalPivotCacheUnsupportedClassification])
  })

  it('classifies external workbook formula references as unsupported instead of oracle failures', async () => {
    const cacheDir = mkdtempSync(join(tmpdir(), 'public-workbook-corpus-external-formula-'))
    mkdirSync(join(cacheDir, 'files'), { recursive: true })
    const workbookBytes = buildExternalWorkbookReferenceBytes()
    const sha256 = await sha256Hex(workbookBytes)
    writeFileSync(join(cacheDir, 'files', `${sha256}.xlsx`), workbookBytes)
    const license = {
      spdxId: 'CC-BY-4.0',
      title: 'Creative Commons Attribution 4.0 International',
      evidenceUrl: 'https://creativecommons.org/licenses/by/4.0/',
    }
    const manifest: PublicWorkbookManifest = {
      ...createEmptyPublicWorkbookManifest('2026-05-07T00:00:00.000Z'),
      sources: [
        {
          id: 'source-external-formula',
          kind: 'direct-url',
          sourceUrl: 'https://example.com/external-formula.xlsx',
          downloadUrl: 'https://example.com/external-formula.xlsx',
          fileName: 'external-formula.xlsx',
          discoveredAt: '2026-05-07T00:00:00.000Z',
          license,
        },
      ],
      artifacts: [
        {
          id: `workbook-${sha256.slice(0, 16)}`,
          sourceId: 'source-external-formula',
          sourceUrl: 'https://example.com/external-formula.xlsx',
          downloadUrl: 'https://example.com/external-formula.xlsx',
          fileName: 'external-formula.xlsx',
          cachePath: `files/${sha256}.xlsx`,
          sha256,
          byteSize: workbookBytes.byteLength,
          workbookFingerprint: 'external-formula-fingerprint',
          fetchedAt: '2026-05-07T00:00:00.000Z',
          license,
        },
      ],
    }

    const scorecard = await buildPublicWorkbookCorpusScorecard({
      manifest,
      cacheDir,
      generatedAt: '2026-05-07T01:00:00.000Z',
    })

    expect(scorecard.summary.allCachedWorkbooksPassed).toBe(true)
    expect(scorecard.summary.formulaOracleComparisonCount).toBe(0)
    expect(scorecard.cases[0]).toMatchObject({
      status: 'unsupported',
      passed: true,
      featureCounts: { formulaCellCount: 1, warningCount: 1 },
      validation: { formulaOraclePassed: true, formulaOracleComparisons: 0 },
      unsupportedFeatureClassifications: [`xlsx.import.warning:${externalWorkbookReferencesWarning}`],
    })
    expect(scorecard.cases[0]?.evidence).toEqual(
      expect.arrayContaining([
        publicWorkbookImportWarningClassifierEvidence,
        'Round-trip projection skipped because external workbook links are not recalculated during XLSX import.',
      ]),
    )
  })

  it('classifies manual calculation cached formula values as unsupported instead of oracle failures', async () => {
    const cacheDir = mkdtempSync(join(tmpdir(), 'public-workbook-corpus-manual-calculation-'))
    mkdirSync(join(cacheDir, 'files'), { recursive: true })
    const workbookBytes = buildManualCalculationWorkbookBytes()
    const sha256 = await sha256Hex(workbookBytes)
    writeFileSync(join(cacheDir, 'files', `${sha256}.xlsx`), workbookBytes)
    const license = {
      spdxId: 'CC-BY-4.0',
      title: 'Creative Commons Attribution 4.0 International',
      evidenceUrl: 'https://creativecommons.org/licenses/by/4.0/',
    }
    const manifest: PublicWorkbookManifest = {
      ...createEmptyPublicWorkbookManifest('2026-05-07T00:00:00.000Z'),
      sources: [
        {
          id: 'source-manual-calculation',
          kind: 'direct-url',
          sourceUrl: 'https://example.com/manual-calculation.xlsx',
          downloadUrl: 'https://example.com/manual-calculation.xlsx',
          fileName: 'manual-calculation.xlsx',
          discoveredAt: '2026-05-07T00:00:00.000Z',
          license,
        },
      ],
      artifacts: [
        {
          id: `workbook-${sha256.slice(0, 16)}`,
          sourceId: 'source-manual-calculation',
          sourceUrl: 'https://example.com/manual-calculation.xlsx',
          downloadUrl: 'https://example.com/manual-calculation.xlsx',
          fileName: 'manual-calculation.xlsx',
          cachePath: `files/${sha256}.xlsx`,
          sha256,
          byteSize: workbookBytes.byteLength,
          workbookFingerprint: 'manual-calculation-fingerprint',
          fetchedAt: '2026-05-07T00:00:00.000Z',
          license,
        },
      ],
    }

    const scorecard = await buildPublicWorkbookCorpusScorecard({
      manifest,
      cacheDir,
      generatedAt: '2026-05-07T01:00:00.000Z',
    })

    expect(scorecard.summary.allCachedWorkbooksPassed).toBe(true)
    expect(scorecard.summary.formulaOracleComparisonCount).toBe(0)
    expect(scorecard.cases[0]).toMatchObject({
      status: 'unsupported',
      passed: true,
      featureCounts: { formulaCellCount: 1, warningCount: 1 },
      validation: { formulaOraclePassed: true, formulaOracleComparisons: 0, roundTripPassed: true },
      unsupportedFeatureClassifications: [`xlsx.import.warning:${manualCalculationModeWarning}`],
    })
  })

  it('accepts precision-as-displayed formulas when formula oracle validation matches cached Excel values', async () => {
    const cacheDir = mkdtempSync(join(tmpdir(), 'public-workbook-corpus-precision-displayed-'))
    mkdirSync(join(cacheDir, 'files'), { recursive: true })
    const workbookBytes = buildPrecisionAsDisplayedFormulaWorkbookBytes()
    const sha256 = await sha256Hex(workbookBytes)
    writeFileSync(join(cacheDir, 'files', `${sha256}.xlsx`), workbookBytes)
    const license = {
      spdxId: 'CC-BY-4.0',
      title: 'Creative Commons Attribution 4.0 International',
      evidenceUrl: 'https://creativecommons.org/licenses/by/4.0/',
    }
    const manifest: PublicWorkbookManifest = {
      ...createEmptyPublicWorkbookManifest('2026-05-07T00:00:00.000Z'),
      sources: [
        {
          id: 'source-precision-displayed',
          kind: 'direct-url',
          sourceUrl: 'https://example.com/precision-displayed.xlsx',
          downloadUrl: 'https://example.com/precision-displayed.xlsx',
          fileName: 'precision-displayed.xlsx',
          discoveredAt: '2026-05-07T00:00:00.000Z',
          license,
        },
      ],
      artifacts: [
        {
          id: `workbook-${sha256.slice(0, 16)}`,
          sourceId: 'source-precision-displayed',
          sourceUrl: 'https://example.com/precision-displayed.xlsx',
          downloadUrl: 'https://example.com/precision-displayed.xlsx',
          fileName: 'precision-displayed.xlsx',
          cachePath: `files/${sha256}.xlsx`,
          sha256,
          byteSize: workbookBytes.byteLength,
          workbookFingerprint: 'precision-displayed-fingerprint',
          fetchedAt: '2026-05-07T00:00:00.000Z',
          license,
        },
      ],
    }

    const scorecard = await buildPublicWorkbookCorpusScorecard({
      manifest,
      cacheDir,
      generatedAt: '2026-05-07T01:00:00.000Z',
    })

    expect(scorecard.summary.allCachedWorkbooksPassed).toBe(true)
    expect(scorecard.summary.formulaOracleComparisonCount).toBe(1)
    expect(scorecard.cases[0]).toMatchObject({
      status: 'passed',
      passed: true,
      featureCounts: { formulaCellCount: 1, warningCount: 1 },
      validation: { formulaOraclePassed: true, formulaOracleComparisons: 1, roundTripPassed: true },
      unsupportedFeatureClassifications: [],
    })
    expect(scorecard.cases[0]?.evidence).not.toContain(publicWorkbookImportWarningClassifierEvidence)
    expect(scorecard.cases[0]?.unsupportedFeatureClassifications).not.toContain(
      `xlsx.import.warning:${precisionAsDisplayedCalculationWarning}`,
    )
  })

  it('classifies volatile cached formula values as unsupported instead of oracle failures', async () => {
    const cacheDir = mkdtempSync(join(tmpdir(), 'public-workbook-corpus-volatile-formula-'))
    mkdirSync(join(cacheDir, 'files'), { recursive: true })
    const workbookBytes = buildVolatileFormulaWorkbookBytes()
    const sha256 = await sha256Hex(workbookBytes)
    writeFileSync(join(cacheDir, 'files', `${sha256}.xlsx`), workbookBytes)
    const license = {
      spdxId: 'CC-BY-4.0',
      title: 'Creative Commons Attribution 4.0 International',
      evidenceUrl: 'https://creativecommons.org/licenses/by/4.0/',
    }
    const manifest: PublicWorkbookManifest = {
      ...createEmptyPublicWorkbookManifest('2026-05-07T00:00:00.000Z'),
      sources: [
        {
          id: 'source-volatile-formula',
          kind: 'direct-url',
          sourceUrl: 'https://example.com/volatile-formula.xlsx',
          downloadUrl: 'https://example.com/volatile-formula.xlsx',
          fileName: 'volatile-formula.xlsx',
          discoveredAt: '2026-05-07T00:00:00.000Z',
          license,
        },
      ],
      artifacts: [
        {
          id: `workbook-${sha256.slice(0, 16)}`,
          sourceId: 'source-volatile-formula',
          sourceUrl: 'https://example.com/volatile-formula.xlsx',
          downloadUrl: 'https://example.com/volatile-formula.xlsx',
          fileName: 'volatile-formula.xlsx',
          cachePath: `files/${sha256}.xlsx`,
          sha256,
          byteSize: workbookBytes.byteLength,
          workbookFingerprint: 'volatile-formula-fingerprint',
          fetchedAt: '2026-05-07T00:00:00.000Z',
          license,
        },
      ],
    }

    const scorecard = await buildPublicWorkbookCorpusScorecard({
      manifest,
      cacheDir,
      generatedAt: '2026-05-07T01:00:00.000Z',
    })

    expect(scorecard.summary.allCachedWorkbooksPassed).toBe(true)
    expect(scorecard.summary.formulaOracleComparisonCount).toBe(0)
    expect(scorecard.cases[0]).toMatchObject({
      status: 'unsupported',
      passed: true,
      featureCounts: { formulaCellCount: 1, warningCount: 1 },
      validation: { formulaOraclePassed: true, formulaOracleComparisons: 0, roundTripPassed: true },
      unsupportedFeatureClassifications: [`xlsx.import.warning:${volatileFormulasWarning}`],
    })
  })

  it('classifies macro-enabled workbooks as unsupported without formula or round-trip failures', async () => {
    const scorecard = await buildSingleWorkbookScorecard({
      cacheDirPrefix: 'public-workbook-corpus-macro-enabled-',
      fileName: 'macro-enabled.xlsm',
      sourceId: 'source-macro-enabled',
      workbookBytes: buildMacroEnabledWorkbookBytes(),
    })

    expect(scorecard.summary.allCachedWorkbooksPassed).toBe(true)
    expect(scorecard.summary.formulaOracleComparisonCount).toBe(0)
    expect(scorecard.cases[0]).toMatchObject({
      status: 'unsupported',
      passed: true,
      featureCounts: { formulaCellCount: 1, macroPayloadCount: 1, warningCount: 1 },
      validation: { formulaOraclePassed: true, formulaOracleComparisons: 0, roundTripPassed: true },
      unsupportedFeatureClassifications: [`xlsx.import.warning:${macroExecutionDeclinedWarning}`, 'xlsx.macros.execution.declined'],
    })
    expect(scorecard.cases[0]?.evidence).toEqual(
      expect.arrayContaining([
        publicWorkbookImportWarningClassifierEvidence,
        'Round-trip projection skipped because macro execution is intentionally declined during XLSM import.',
      ]),
    )
  })

  it('classifies stale cached formula values after independent recalculation agrees', async () => {
    const scorecard = await buildSingleWorkbookScorecard({
      cacheDirPrefix: 'public-workbook-corpus-stale-cache-',
      fileName: 'stale-cache.xlsx',
      sourceId: 'source-stale-cache',
      workbookBytes: buildStaleFormulaCacheWorkbookBytes(),
    })

    expect(scorecard.summary.allCachedWorkbooksPassed).toBe(true)
    expect(scorecard.summary.formulaOracleComparisonCount).toBe(1)
    expect(scorecard.cases[0]).toMatchObject({
      status: 'unsupported',
      passed: true,
      featureCounts: { formulaCellCount: 1, warningCount: 0 },
      validation: { formulaOraclePassed: true, formulaOracleComparisons: 1, formulaOracleMismatches: [], roundTripPassed: true },
      unsupportedFeatureClassifications: [staleFormulaCacheUnsupportedClassification],
    })
    expect(scorecard.cases[0]?.evidence).toEqual(
      expect.arrayContaining([publicWorkbookFormulaOracleCacheClassifierEvidence, 'independent-recalc=Summary!B2 cached 6 recalculated 9']),
    )
  })

  it('keeps formula oracle failures when independent recalculation cannot confirm a stale cache', async () => {
    const scorecard = await buildSingleWorkbookScorecard({
      cacheDirPrefix: 'public-workbook-corpus-unconfirmed-cache-',
      fileName: 'unconfirmed-cache.xlsx',
      sourceId: 'source-unconfirmed-cache',
      workbookBytes: buildUnconfirmedFormulaCacheWorkbookBytes(),
    })

    expect(scorecard.summary.allCachedWorkbooksPassed).toBe(false)
    expect(scorecard.cases[0]).toMatchObject({
      status: 'failed',
      passed: false,
      validation: { formulaOraclePassed: false, formulaOracleComparisons: 1 },
      unsupportedFeatureClassifications: [],
    })
    expect(scorecard.cases[0]?.validation.formulaOracleMismatches).toHaveLength(1)
    expect(scorecard.cases[0]?.evidence).not.toContain(publicWorkbookFormulaOracleCacheClassifierEvidence)
  })

  it('classifies oversized workbook verification as an explicit unsupported resource case', async () => {
    const cacheDir = mkdtempSync(join(tmpdir(), 'public-workbook-corpus-resource-limit-'))
    mkdirSync(join(cacheDir, 'files'), { recursive: true })
    const workbookBytes = buildWorkbookBytes()
    const sha256 = await sha256Hex(workbookBytes)
    writeFileSync(join(cacheDir, 'files', `${sha256}.xlsx`), workbookBytes)
    const license = {
      spdxId: 'CC-BY-4.0',
      title: 'Creative Commons Attribution 4.0 International',
      evidenceUrl: 'https://creativecommons.org/licenses/by/4.0/',
    }
    const manifest: PublicWorkbookManifest = {
      ...createEmptyPublicWorkbookManifest('2026-05-07T00:00:00.000Z'),
      sources: [
        {
          id: 'source-resource-limit',
          kind: 'direct-url',
          sourceUrl: 'https://example.com/resource-limit.xlsx',
          downloadUrl: 'https://example.com/resource-limit.xlsx',
          fileName: 'resource-limit.xlsx',
          discoveredAt: '2026-05-07T00:00:00.000Z',
          license,
        },
      ],
      artifacts: [
        {
          id: `workbook-${sha256.slice(0, 16)}`,
          sourceId: 'source-resource-limit',
          sourceUrl: 'https://example.com/resource-limit.xlsx',
          downloadUrl: 'https://example.com/resource-limit.xlsx',
          fileName: 'resource-limit.xlsx',
          cachePath: `files/${sha256}.xlsx`,
          sha256,
          byteSize: workbookBytes.byteLength,
          workbookFingerprint: 'resource-limit-fingerprint',
          fetchedAt: '2026-05-07T00:00:00.000Z',
          license,
        },
      ],
    }

    const scorecard = await buildPublicWorkbookCorpusScorecard({
      manifest,
      cacheDir,
      generatedAt: '2026-05-07T01:00:00.000Z',
      verifyMaxCellCount: 2,
    })

    expect(scorecard.summary.allCachedWorkbooksPassed).toBe(true)
    expect(scorecard.summary.importedWorkbookCount).toBe(0)
    expect(scorecard.cases[0]).toMatchObject({
      status: 'unsupported',
      passed: true,
      validation: { importPassed: false },
      unsupportedFeatureClassifications: ['xlsx.publicCorpus.resourceLimit:cellCount>2'],
    })
    expect(scorecard.cases[0]?.evidence).toEqual(
      expect.arrayContaining([
        publicWorkbookResourceLimitClassifierEvidence,
        'Public corpus verification cell-count limit exceeded: 10 > 2',
      ]),
    )
  })

  it('passes corpus round-trip validation for sheet names with trailing spaces', async () => {
    const cacheDir = mkdtempSync(join(tmpdir(), 'public-workbook-corpus-trailing-space-'))
    mkdirSync(join(cacheDir, 'files'), { recursive: true })
    const workbookBytes = buildTrailingSpaceWorkbookBytes()
    const sha256 = await sha256Hex(workbookBytes)
    writeFileSync(join(cacheDir, 'files', `${sha256}.xlsx`), workbookBytes)

    const license = {
      spdxId: 'CC-BY-4.0',
      title: 'Creative Commons Attribution 4.0 International',
      evidenceUrl: 'https://creativecommons.org/licenses/by/4.0/',
    }
    const manifest: PublicWorkbookManifest = {
      ...createEmptyPublicWorkbookManifest('2026-05-07T00:00:00.000Z'),
      sources: [
        {
          id: 'source-trailing-space',
          kind: 'direct-url',
          sourceUrl: 'https://example.com/trailing-space.xlsx',
          downloadUrl: 'https://example.com/trailing-space.xlsx',
          fileName: 'trailing-space.xlsx',
          discoveredAt: '2026-05-07T00:00:00.000Z',
          license,
        },
      ],
      artifacts: [
        {
          id: `workbook-${sha256.slice(0, 16)}`,
          sourceId: 'source-trailing-space',
          sourceUrl: 'https://example.com/trailing-space.xlsx',
          downloadUrl: 'https://example.com/trailing-space.xlsx',
          fileName: 'trailing-space.xlsx',
          cachePath: `files/${sha256}.xlsx`,
          sha256,
          byteSize: workbookBytes.byteLength,
          workbookFingerprint: 'trailing-space-fingerprint',
          fetchedAt: '2026-05-07T00:00:00.000Z',
          license,
        },
      ],
    }

    const scorecard = await buildPublicWorkbookCorpusScorecard({
      manifest,
      cacheDir,
      generatedAt: '2026-05-07T01:00:00.000Z',
      structuralSmokeSampleLimit: 1,
    })

    expect(scorecard.summary.allCachedWorkbooksPassed).toBe(true)
    expect(scorecard.cases[0]).toMatchObject({
      status: 'passed',
      validation: {
        roundTripPassed: true,
        structuralSmokePassed: true,
      },
    })
  })

  it('compares populated-cell styles when round trips shrink blank style ranges', () => {
    const broadStyleRange = buildInteriorStyledSnapshot('A1', 'C3')
    const populatedOnlyStyleRange = buildInteriorStyledSnapshot('B2', 'B2')
    const differentPopulatedStyle = buildInteriorStyledSnapshot('B2', 'B2', '#00ccff')

    expect(roundTripSemanticsDigest(broadStyleRange)).toBe(roundTripSemanticsDigest(populatedOnlyStyleRange))
    expect(roundTripSemanticsDigest(broadStyleRange)).not.toBe(roundTripSemanticsDigest(differentPopulatedStyle))
  })

  it('ignores blank row and column dimensions in round-trip digests', () => {
    const blankAxisDimensions = buildAxisDimensionSnapshot({
      rows: [
        { id: 'row:0', index: 0, size: 22 },
        { id: 'row:1', index: 1, size: 44 },
      ],
      columns: [
        { id: 'col:0', index: 0, size: 72 },
        { id: 'col:1', index: 1, size: 144 },
      ],
    })
    const populatedAxisDimensions = buildAxisDimensionSnapshot({
      rows: [{ id: 'row:1', index: 1, size: 44 }],
      columns: [{ id: 'col:1', index: 1, size: 144 }],
    })
    const differentPopulatedAxisDimensions = buildAxisDimensionSnapshot({
      rows: [],
      columns: [{ id: 'col:1', index: 1, size: 144 }],
    })

    expect(roundTripSemanticsDigest(blankAxisDimensions)).toBe(roundTripSemanticsDigest(populatedAxisDimensions))
    expect(roundTripSemanticsDigest(blankAxisDimensions)).not.toBe(roundTripSemanticsDigest(differentPopulatedAxisDimensions))
  })

  it('normalizes default chart series orientation in round-trip digests', () => {
    const implicitColumnOrientation = buildChartOrientationSnapshot()
    const explicitColumnOrientation = buildChartOrientationSnapshot('columns')
    const rowOrientation = buildChartOrientationSnapshot('rows')

    expect(roundTripSemanticsDigest(implicitColumnOrientation)).toBe(roundTripSemanticsDigest(explicitColumnOrientation))
    expect(roundTripSemanticsDigest(implicitColumnOrientation)).not.toBe(roundTripSemanticsDigest(rowOrientation))
  })

  it('runs structural smoke against a mutable sheet when the first sheet is protected', async () => {
    const cacheDir = mkdtempSync(join(tmpdir(), 'public-workbook-corpus-protected-first-'))
    mkdirSync(join(cacheDir, 'files'), { recursive: true })
    const workbookBytes = buildProtectedFirstWorkbookBytes()
    const sha256 = await sha256Hex(workbookBytes)
    writeFileSync(join(cacheDir, 'files', `${sha256}.xlsx`), workbookBytes)

    const license = {
      spdxId: 'CC-BY-4.0',
      title: 'Creative Commons Attribution 4.0 International',
      evidenceUrl: 'https://creativecommons.org/licenses/by/4.0/',
    }
    const manifest: PublicWorkbookManifest = {
      ...createEmptyPublicWorkbookManifest('2026-05-07T00:00:00.000Z'),
      sources: [
        {
          id: 'source-protected-first',
          kind: 'direct-url',
          sourceUrl: 'https://example.com/protected-first.xlsx',
          downloadUrl: 'https://example.com/protected-first.xlsx',
          fileName: 'protected-first.xlsx',
          discoveredAt: '2026-05-07T00:00:00.000Z',
          license,
        },
      ],
      artifacts: [
        {
          id: `workbook-${sha256.slice(0, 16)}`,
          sourceId: 'source-protected-first',
          sourceUrl: 'https://example.com/protected-first.xlsx',
          downloadUrl: 'https://example.com/protected-first.xlsx',
          fileName: 'protected-first.xlsx',
          cachePath: `files/${sha256}.xlsx`,
          sha256,
          byteSize: workbookBytes.byteLength,
          workbookFingerprint: 'protected-first-fingerprint',
          fetchedAt: '2026-05-07T00:00:00.000Z',
          license,
        },
      ],
    }

    const scorecard = await buildPublicWorkbookCorpusScorecard({
      manifest,
      cacheDir,
      generatedAt: '2026-05-07T01:00:00.000Z',
      structuralSmokeSampleLimit: 1,
    })

    expect(scorecard.summary.allCachedWorkbooksPassed).toBe(true)
    expect(scorecard.cases[0]).toMatchObject({
      status: 'passed',
      validation: {
        roundTripPassed: true,
        structuralSmokePassed: true,
      },
    })
  })

  it('runs structural smoke on an editable sheet when the first sheet is protected', async () => {
    const cacheDir = mkdtempSync(join(tmpdir(), 'public-workbook-corpus-protected-first-sheet-'))
    mkdirSync(join(cacheDir, 'files'), { recursive: true })
    const workbookBytes = buildProtectedFirstSheetWorkbookBytes()
    const sha256 = await sha256Hex(workbookBytes)
    writeFileSync(join(cacheDir, 'files', `${sha256}.xlsx`), workbookBytes)

    const license = {
      spdxId: 'CC-BY-4.0',
      title: 'Creative Commons Attribution 4.0 International',
      evidenceUrl: 'https://creativecommons.org/licenses/by/4.0/',
    }
    const manifest: PublicWorkbookManifest = {
      ...createEmptyPublicWorkbookManifest('2026-05-07T00:00:00.000Z'),
      sources: [
        {
          id: 'source-protected-first-sheet',
          kind: 'direct-url',
          sourceUrl: 'https://example.com/protected-first-sheet.xlsx',
          downloadUrl: 'https://example.com/protected-first-sheet.xlsx',
          fileName: 'protected-first-sheet.xlsx',
          discoveredAt: '2026-05-07T00:00:00.000Z',
          license,
        },
      ],
      artifacts: [
        {
          id: `workbook-${sha256.slice(0, 16)}`,
          sourceId: 'source-protected-first-sheet',
          sourceUrl: 'https://example.com/protected-first-sheet.xlsx',
          downloadUrl: 'https://example.com/protected-first-sheet.xlsx',
          fileName: 'protected-first-sheet.xlsx',
          cachePath: `files/${sha256}.xlsx`,
          sha256,
          byteSize: workbookBytes.byteLength,
          workbookFingerprint: 'protected-first-sheet-fingerprint',
          fetchedAt: '2026-05-07T00:00:00.000Z',
          license,
        },
      ],
    }

    const scorecard = await buildPublicWorkbookCorpusScorecard({
      manifest,
      cacheDir,
      generatedAt: '2026-05-07T01:00:00.000Z',
      structuralSmokeSampleLimit: 1,
    })

    expect(scorecard.summary.allCachedWorkbooksPassed).toBe(true)
    expect(scorecard.cases[0]).toMatchObject({
      status: 'passed',
      validation: {
        roundTripPassed: true,
        structuralSmokePassed: true,
      },
      unsupportedFeatureClassifications: [],
    })
  })

  it('rejects stale scorecards without all cached workbook cases', async () => {
    const scorecard = await buildPublicWorkbookCorpusScorecard({
      manifest: createEmptyPublicWorkbookManifest('2026-05-07T00:00:00.000Z'),
      cacheDir: mkdtempSync(join(tmpdir(), 'public-workbook-corpus-empty-')),
      generatedAt: '2026-05-07T01:00:00.000Z',
    })

    expect(() =>
      validatePublicWorkbookCorpusScorecard({
        ...scorecard,
        summary: {
          ...scorecard.summary,
          cachedWorkbookCount: 1,
        },
      }),
    ).toThrow('Public workbook corpus scorecard case count does not match cached workbook count')
  })

  it('rejects scorecards when any cached workbook fails verification', async () => {
    const cacheDir = mkdtempSync(join(tmpdir(), 'public-workbook-corpus-missing-'))
    const missingHash = 'a'.repeat(64)
    const license = {
      spdxId: 'CC-BY-4.0',
      title: 'Creative Commons Attribution 4.0 International',
      evidenceUrl: 'https://creativecommons.org/licenses/by/4.0/',
    }
    const manifest: PublicWorkbookManifest = {
      ...createEmptyPublicWorkbookManifest('2026-05-07T00:00:00.000Z'),
      sources: [
        {
          id: 'source-missing-file',
          kind: 'direct-url',
          sourceUrl: 'https://example.com/missing.xlsx',
          downloadUrl: 'https://example.com/missing.xlsx',
          fileName: 'missing.xlsx',
          discoveredAt: '2026-05-07T00:00:00.000Z',
          license,
        },
      ],
      artifacts: [
        {
          id: 'workbook-missing-file',
          sourceId: 'source-missing-file',
          sourceUrl: 'https://example.com/missing.xlsx',
          downloadUrl: 'https://example.com/missing.xlsx',
          fileName: 'missing.xlsx',
          cachePath: 'files/missing.xlsx',
          sha256: missingHash,
          byteSize: 1024,
          workbookFingerprint: 'missing-file-fingerprint',
          fetchedAt: '2026-05-07T00:00:00.000Z',
          license,
        },
      ],
    }

    const scorecard = await buildPublicWorkbookCorpusScorecard({
      manifest,
      cacheDir,
      generatedAt: '2026-05-07T01:00:00.000Z',
    })

    expect(scorecard.summary).toMatchObject({
      cachedWorkbookCount: 1,
      importedWorkbookCount: 0,
      errorWorkbookCount: 1,
      formulaOracleMatchCount: 0,
      allCachedWorkbooksPassed: false,
    })
    expect(scorecard.cases[0]?.validation.formulaOracleMismatches).toEqual([])
    expect(() => validatePublicWorkbookCorpusScorecard(scorecard)).toThrow(
      'Public workbook corpus scorecard has cached workbooks that did not pass',
    )
  })

  it('fetches public workbook artifacts in bounded batches', async () => {
    const cacheDir = mkdtempSync(join(tmpdir(), 'public-workbook-corpus-fetch-'))
    const workbookA = buildWorkbookBytes('SummaryA')
    const workbookB = buildWorkbookBytes('SummaryB')
    const fetchMock = vi.fn(async (input: string | URL | Request) => {
      const url = typeof input === 'string' ? input : input instanceof URL ? input.href : input.url
      const sourceIndex = Number(/workbook-(\d+)\.xlsx/u.exec(url)?.[1] ?? '0')
      const bytes = sourceIndex === 1 ? workbookB : workbookA
      return new Response(bytes, {
        headers: {
          'content-length': String(bytes.byteLength),
        },
      })
    })
    vi.stubGlobal('fetch', fetchMock)

    const license = {
      spdxId: 'CC-BY-4.0',
      title: 'Creative Commons Attribution 4.0 International',
      evidenceUrl: 'https://creativecommons.org/licenses/by/4.0/',
    }
    const sources: PublicWorkbookSource[] = Array.from({ length: 50 }, (_, index) => ({
      id: `source-${String(index)}`,
      kind: 'direct-url',
      sourceUrl: `https://example.com/workbook-${String(index)}.xlsx`,
      downloadUrl: `https://example.com/workbook-${String(index)}.xlsx`,
      fileName: `workbook-${String(index)}.xlsx`,
      discoveredAt: '2026-05-07T00:00:00.000Z',
      license,
    }))
    const manifest: PublicWorkbookManifest = {
      ...createEmptyPublicWorkbookManifest('2026-05-07T00:00:00.000Z'),
      sources,
    }

    const fetched = await fetchPublicWorkbookArtifacts({
      manifest,
      cacheDir,
      limit: 2,
      fetchedAt: '2026-05-07T01:00:00.000Z',
    })

    expect(fetched.artifacts).toHaveLength(2)
    expect(fetchMock).toHaveBeenCalledTimes(6)
  })

  it('honors explicit fetch concurrency limits', async () => {
    const cacheDir = mkdtempSync(join(tmpdir(), 'public-workbook-corpus-fetch-concurrency-'))
    const workbookBytes = buildWorkbookBytes('Concurrency')
    let inFlightFetches = 0
    let maxInFlightFetches = 0
    const fetchMock = vi.fn(async () => {
      inFlightFetches += 1
      maxInFlightFetches = Math.max(maxInFlightFetches, inFlightFetches)
      await new Promise((resolve) => setTimeout(resolve, 5))
      inFlightFetches -= 1
      return new Response(workbookBytes, {
        headers: {
          'content-length': String(workbookBytes.byteLength),
        },
      })
    })
    vi.stubGlobal('fetch', fetchMock)

    const license = {
      spdxId: 'CC-BY-4.0',
      title: 'Creative Commons Attribution 4.0 International',
      evidenceUrl: 'https://creativecommons.org/licenses/by/4.0/',
    }
    const manifest: PublicWorkbookManifest = {
      ...createEmptyPublicWorkbookManifest('2026-05-07T00:00:00.000Z'),
      sources: Array.from({ length: 8 }, (_, index) => ({
        id: `source-concurrency-${String(index)}`,
        kind: 'direct-url',
        sourceUrl: `https://example.com/concurrency-${String(index)}.xlsx`,
        downloadUrl: `https://example.com/concurrency-${String(index)}.xlsx`,
        fileName: `concurrency-${String(index)}.xlsx`,
        discoveredAt: '2026-05-07T00:00:00.000Z',
        license,
      })),
    }

    await fetchPublicWorkbookArtifacts({
      manifest,
      cacheDir,
      limit: 1,
      fetchedAt: '2026-05-07T01:00:00.000Z',
      fetchBatchSize: 4,
      fetchConcurrency: 2,
    })

    expect(maxInFlightFetches).toBeLessThanOrEqual(2)
    expect(fetchMock).toHaveBeenCalledTimes(3)
  })

  it('dedupes candidate download URLs before fetching artifacts', async () => {
    const cacheDir = mkdtempSync(join(tmpdir(), 'public-workbook-corpus-fetch-dedupe-'))
    const workbookBytes = buildWorkbookBytes()
    const fetchMock = vi.fn(async () => new Response(workbookBytes, { headers: { 'content-length': String(workbookBytes.byteLength) } }))
    vi.stubGlobal('fetch', fetchMock)

    const license = {
      spdxId: 'CC-BY-4.0',
      title: 'Creative Commons Attribution 4.0 International',
      evidenceUrl: 'https://creativecommons.org/licenses/by/4.0/',
    }
    const manifest: PublicWorkbookManifest = {
      ...createEmptyPublicWorkbookManifest('2026-05-07T00:00:00.000Z'),
      sources: [
        {
          id: 'source-1',
          kind: 'direct-url',
          sourceUrl: 'https://example.com/dataset-1',
          downloadUrl: 'https://example.com/shared-budget.xlsx',
          fileName: 'shared-budget.xlsx',
          discoveredAt: '2026-05-07T00:00:00.000Z',
          license,
        },
        {
          id: 'source-2',
          kind: 'direct-url',
          sourceUrl: 'https://example.com/dataset-2',
          downloadUrl: 'https://example.com/shared-budget.xlsx',
          fileName: 'shared-budget.xlsx',
          discoveredAt: '2026-05-07T00:00:00.000Z',
          license,
        },
      ],
    }

    const fetched = await fetchPublicWorkbookArtifacts({
      manifest,
      cacheDir,
      limit: 2,
      fetchedAt: '2026-05-07T01:00:00.000Z',
    })

    expect(fetchMock).toHaveBeenCalledTimes(1)
    expect(fetched.artifacts).toHaveLength(1)
  })

  it('prioritizes xlsx candidate sources before legacy xls fetches', async () => {
    const cacheDir = mkdtempSync(join(tmpdir(), 'public-workbook-corpus-fetch-xlsx-first-'))
    const workbookBytes = buildWorkbookBytes()
    const fetchMock = vi.fn(async () => new Response(workbookBytes, { headers: { 'content-length': String(workbookBytes.byteLength) } }))
    vi.stubGlobal('fetch', fetchMock)

    const license = {
      spdxId: 'CC-BY-4.0',
      title: 'Creative Commons Attribution 4.0 International',
      evidenceUrl: 'https://creativecommons.org/licenses/by/4.0/',
    }
    const manifest: PublicWorkbookManifest = {
      ...createEmptyPublicWorkbookManifest('2026-05-07T00:00:00.000Z'),
      sources: [
        {
          id: 'source-legacy-xls',
          kind: 'direct-url',
          sourceUrl: 'https://example.com/legacy.xls',
          downloadUrl: 'https://example.com/legacy.xls',
          fileName: 'legacy.xls',
          discoveredAt: '2026-05-07T00:00:00.000Z',
          license,
        },
        {
          id: 'source-modern-xlsx',
          kind: 'direct-url',
          sourceUrl: 'https://example.com/modern.xlsx',
          downloadUrl: 'https://example.com/modern.xlsx',
          fileName: 'modern.xlsx',
          discoveredAt: '2026-05-07T00:00:00.000Z',
          license,
        },
      ],
    }

    const fetched = await fetchPublicWorkbookArtifacts({
      manifest,
      cacheDir,
      limit: 1,
      fetchedAt: '2026-05-07T01:00:00.000Z',
      fetchBatchSize: 1,
    })

    expect(fetchMock).toHaveBeenCalledTimes(1)
    expect(fetchMock.mock.calls[0]?.[0]).toBe('https://example.com/modern.xlsx')
    expect(fetched.artifacts).toHaveLength(1)
    expect(fetched.artifacts[0]?.sourceId).toBe('source-modern-xlsx')
  })

  it('checkpoints the manifest when fetched artifacts are committed', async () => {
    const cacheDir = mkdtempSync(join(tmpdir(), 'public-workbook-corpus-fetch-checkpoint-'))
    const workbookBytes = buildWorkbookBytes()
    vi.stubGlobal(
      'fetch',
      vi.fn(async () => new Response(workbookBytes, { headers: { 'content-length': String(workbookBytes.byteLength) } })),
    )

    const license = {
      spdxId: 'CC-BY-4.0',
      title: 'Creative Commons Attribution 4.0 International',
      evidenceUrl: 'https://creativecommons.org/licenses/by/4.0/',
    }
    const manifest: PublicWorkbookManifest = {
      ...createEmptyPublicWorkbookManifest('2026-05-07T00:00:00.000Z'),
      sources: [
        {
          id: 'source-checkpoint',
          kind: 'direct-url',
          sourceUrl: 'https://example.com/checkpoint.xlsx',
          downloadUrl: 'https://example.com/checkpoint.xlsx',
          fileName: 'checkpoint.xlsx',
          discoveredAt: '2026-05-07T00:00:00.000Z',
          license,
        },
      ],
    }
    const checkpoints: number[] = []

    await fetchPublicWorkbookArtifacts({
      manifest,
      cacheDir,
      limit: 1,
      fetchedAt: '2026-05-07T01:00:00.000Z',
      onArtifactsCommitted: (checkpointManifest) => checkpoints.push(checkpointManifest.artifacts.length),
    })

    expect(checkpoints).toEqual([1])
  })

  it('reports checkpoint progress when fetch sources are exhausted without new artifacts', async () => {
    const cacheDir = mkdtempSync(join(tmpdir(), 'public-workbook-corpus-fetch-progress-'))
    const workbookBytes = buildWorkbookBytes('DuplicateProgress')
    vi.stubGlobal(
      'fetch',
      vi.fn(async () => new Response(workbookBytes, { headers: { 'content-length': String(workbookBytes.byteLength) } })),
    )

    const license = {
      spdxId: 'CC-BY-4.0',
      title: 'Creative Commons Attribution 4.0 International',
      evidenceUrl: 'https://creativecommons.org/licenses/by/4.0/',
    }
    const manifest: PublicWorkbookManifest = {
      ...createEmptyPublicWorkbookManifest('2026-05-07T00:00:00.000Z'),
      sources: [
        {
          id: 'source-progress-a',
          kind: 'direct-url',
          sourceUrl: 'https://example.com/progress-a.xlsx',
          downloadUrl: 'https://example.com/progress-a.xlsx',
          fileName: 'progress-a.xlsx',
          discoveredAt: '2026-05-07T00:00:00.000Z',
          license,
        },
        {
          id: 'source-progress-b',
          kind: 'direct-url',
          sourceUrl: 'https://example.com/progress-b.xlsx',
          downloadUrl: 'https://example.com/progress-b.xlsx',
          fileName: 'progress-b.xlsx',
          discoveredAt: '2026-05-07T00:00:00.000Z',
          license,
        },
      ],
    }
    const progress: PublicWorkbookCorpusFetchCheckpointProgress[] = []

    const fetched = await fetchPublicWorkbookArtifacts({
      manifest,
      cacheDir,
      limit: 2,
      fetchedAt: '2026-05-07T01:00:00.000Z',
      fetchBatchSize: 1,
      onArtifactsCommitted: (_checkpointManifest, checkpointProgress) => progress.push(checkpointProgress),
    })

    expect(fetched.artifacts).toHaveLength(1)
    expect(progress).toEqual([
      {
        artifactCount: 1,
        exhaustedSourceCount: 1,
        committedArtifactCount: 1,
        exhaustedSourceDelta: 1,
        failedSourceCount: 0,
        duplicateHashSourceCount: 0,
        duplicateFingerprintSourceCount: 0,
        failedSourceSamples: [],
      },
      {
        artifactCount: 1,
        exhaustedSourceCount: 2,
        committedArtifactCount: 0,
        exhaustedSourceDelta: 1,
        failedSourceCount: 0,
        duplicateHashSourceCount: 1,
        duplicateFingerprintSourceCount: 0,
        failedSourceSamples: [],
      },
    ])
  })

  it('includes bounded fetch failure samples in checkpoint progress', async () => {
    const cacheDir = mkdtempSync(join(tmpdir(), 'public-workbook-corpus-fetch-failure-progress-'))
    vi.stubGlobal(
      'fetch',
      vi.fn(async () => new Response('missing', { status: 404 })),
    )

    const license = {
      spdxId: 'CC-BY-4.0',
      title: 'Creative Commons Attribution 4.0 International',
      evidenceUrl: 'https://creativecommons.org/licenses/by/4.0/',
    }
    const manifest: PublicWorkbookManifest = {
      ...createEmptyPublicWorkbookManifest('2026-05-07T00:00:00.000Z'),
      sources: [
        {
          id: 'source-failure-progress',
          kind: 'direct-url',
          sourceUrl: 'https://example.com/missing.xlsx',
          downloadUrl: 'https://example.com/missing.xlsx',
          fileName: 'missing.xlsx',
          discoveredAt: '2026-05-07T00:00:00.000Z',
          license,
        },
      ],
    }
    const progress: PublicWorkbookCorpusFetchCheckpointProgress[] = []

    const fetched = await fetchPublicWorkbookArtifacts({
      manifest,
      cacheDir,
      limit: 1,
      fetchedAt: '2026-05-07T01:00:00.000Z',
      fetchBatchSize: 1,
      onArtifactsCommitted: (_checkpointManifest, checkpointProgress) => progress.push(checkpointProgress),
    })

    expect(fetched.artifacts).toHaveLength(0)
    expect(progress).toEqual([
      {
        artifactCount: 0,
        exhaustedSourceCount: 1,
        committedArtifactCount: 0,
        exhaustedSourceDelta: 1,
        failedSourceCount: 1,
        duplicateHashSourceCount: 0,
        duplicateFingerprintSourceCount: 0,
        failedSourceSamples: [
          {
            sourceId: 'source-failure-progress',
            fileName: 'missing.xlsx',
            error: 'Unable to download https://example.com/missing.xlsx: HTTP 404',
          },
        ],
      },
    ])
  })

  it('fetches multi-batch corpus tranches without retaining prior batch results', async () => {
    const cacheDir = mkdtempSync(join(tmpdir(), 'public-workbook-corpus-fetch-iterative-'))
    const gcMock = vi.fn()
    vi.stubGlobal('Bun', { gc: gcMock })
    const fetchMock = vi.fn(async (url: string) => {
      const index = Number(url.match(/iterative-(\d+)/u)?.[1] ?? '0')
      const workbookBytes = buildWorkbookBytes(`Batch${String(index)}`)
      return new Response(workbookBytes, {
        headers: {
          'content-length': String(workbookBytes.byteLength),
        },
      })
    })
    vi.stubGlobal('fetch', fetchMock)

    const license = {
      spdxId: 'CC-BY-4.0',
      title: 'Creative Commons Attribution 4.0 International',
      evidenceUrl: 'https://creativecommons.org/licenses/by/4.0/',
    }
    const manifest: PublicWorkbookManifest = {
      ...createEmptyPublicWorkbookManifest('2026-05-07T00:00:00.000Z'),
      sources: Array.from({ length: 4 }, (_, index) => ({
        id: `source-iterative-${String(index)}`,
        kind: 'direct-url',
        sourceUrl: `https://example.com/iterative-${String(index)}.xlsx`,
        downloadUrl: `https://example.com/iterative-${String(index)}.xlsx`,
        fileName: `iterative-${String(index)}.xlsx`,
        discoveredAt: '2026-05-07T00:00:00.000Z',
        license,
      })),
    }
    const checkpoints: number[] = []

    const fetched = await fetchPublicWorkbookArtifacts({
      manifest,
      cacheDir,
      limit: 4,
      fetchedAt: '2026-05-07T01:00:00.000Z',
      fetchBatchSize: 1,
      onArtifactsCommitted: (checkpointManifest) => checkpoints.push(checkpointManifest.artifacts.length),
    })

    expect(fetched.artifacts).toHaveLength(4)
    expect(fetchMock).toHaveBeenCalledTimes(4)
    expect(checkpoints).toEqual([1, 2, 3, 4])
    expect(gcMock).toHaveBeenCalledTimes(4)
    expect(gcMock).toHaveBeenCalledWith(true)
  })

  it('persists exhausted duplicate sources so resumed fetches do not retry them', async () => {
    const cacheDir = mkdtempSync(join(tmpdir(), 'public-workbook-corpus-fetch-resume-'))
    const workbookA = buildWorkbookBytes('Duplicate')
    const workbookB = buildWorkbookBytes('Fresh')
    const workbookC = buildWorkbookBytes('Later')
    const fetchMock = vi.fn(async (url: string) => {
      const bytes = url.includes('later') ? workbookC : url.includes('fresh') ? workbookB : workbookA
      return new Response(bytes, {
        headers: {
          'content-length': String(bytes.byteLength),
        },
      })
    })
    vi.stubGlobal('fetch', fetchMock)

    const license = {
      spdxId: 'CC-BY-4.0',
      title: 'Creative Commons Attribution 4.0 International',
      evidenceUrl: 'https://creativecommons.org/licenses/by/4.0/',
    }
    const manifest: PublicWorkbookManifest = {
      ...createEmptyPublicWorkbookManifest('2026-05-07T00:00:00.000Z'),
      sources: [
        {
          id: 'source-original',
          kind: 'direct-url',
          sourceUrl: 'https://example.com/original.xlsx',
          downloadUrl: 'https://example.com/original.xlsx',
          fileName: 'original.xlsx',
          discoveredAt: '2026-05-07T00:00:00.000Z',
          license,
        },
        {
          id: 'source-duplicate',
          kind: 'direct-url',
          sourceUrl: 'https://example.com/duplicate.xlsx',
          downloadUrl: 'https://example.com/duplicate.xlsx',
          fileName: 'duplicate.xlsx',
          discoveredAt: '2026-05-07T00:00:00.000Z',
          license,
        },
        {
          id: 'source-fresh',
          kind: 'direct-url',
          sourceUrl: 'https://example.com/fresh.xlsx',
          downloadUrl: 'https://example.com/fresh.xlsx',
          fileName: 'fresh.xlsx',
          discoveredAt: '2026-05-07T00:00:00.000Z',
          license,
        },
        {
          id: 'source-later',
          kind: 'direct-url',
          sourceUrl: 'https://example.com/later.xlsx',
          downloadUrl: 'https://example.com/later.xlsx',
          fileName: 'later.xlsx',
          discoveredAt: '2026-05-07T00:00:00.000Z',
          license,
        },
      ],
    }

    const firstFetch = await fetchPublicWorkbookArtifacts({
      manifest,
      cacheDir,
      limit: 1,
      fetchedAt: '2026-05-07T01:00:00.000Z',
      fetchBatchSize: 1,
    })
    const secondFetch = await fetchPublicWorkbookArtifacts({
      manifest: firstFetch,
      cacheDir,
      limit: 2,
      fetchedAt: '2026-05-07T02:00:00.000Z',
      fetchBatchSize: 1,
    })

    expect(secondFetch.artifacts.map((artifact) => artifact.sourceId)).toEqual(['source-original', 'source-fresh'])
    expect(secondFetch.fetchState?.exhaustedSourceIds).toEqual(['source-original', 'source-duplicate', 'source-fresh'])

    fetchMock.mockClear()
    const thirdFetch = await fetchPublicWorkbookArtifacts({
      manifest: secondFetch,
      cacheDir,
      limit: 3,
      fetchedAt: '2026-05-07T03:00:00.000Z',
      fetchBatchSize: 1,
    })

    expect(fetchMock.mock.calls.map((call) => call[0])).toEqual(['https://example.com/later.xlsx'])
    expect(thirdFetch.artifacts.map((artifact) => artifact.sourceId)).toEqual(['source-original', 'source-fresh', 'source-later'])
  })

  it('times out stalled workbook response bodies during fetch', async () => {
    vi.useFakeTimers()

    const cacheDir = mkdtempSync(join(tmpdir(), 'public-workbook-corpus-stalled-body-'))
    const fetchMock = vi.fn(
      async () =>
        new Response(
          new ReadableStream<Uint8Array>({
            pull: () => new Promise<void>(() => undefined),
          }),
          {
            headers: {
              'content-length': '32',
            },
          },
        ),
    )
    vi.stubGlobal('fetch', fetchMock)

    const license = {
      spdxId: 'CC-BY-4.0',
      title: 'Creative Commons Attribution 4.0 International',
      evidenceUrl: 'https://creativecommons.org/licenses/by/4.0/',
    }
    const manifest: PublicWorkbookManifest = {
      ...createEmptyPublicWorkbookManifest('2026-05-07T00:00:00.000Z'),
      sources: [
        {
          id: 'source-stalled-body',
          kind: 'direct-url',
          sourceUrl: 'https://example.com/stalled.xlsx',
          downloadUrl: 'https://example.com/stalled.xlsx',
          fileName: 'stalled.xlsx',
          discoveredAt: '2026-05-07T00:00:00.000Z',
          license,
        },
      ],
    }

    const fetchedPromise = fetchPublicWorkbookArtifacts({
      manifest,
      cacheDir,
      limit: 1,
      fetchedAt: '2026-05-07T01:00:00.000Z',
      downloadTimeoutMs: 5,
    })

    await vi.advanceTimersByTimeAsync(5)
    const fetched = await fetchedPromise

    expect(fetchMock).toHaveBeenCalledTimes(1)
    expect(fetched.artifacts).toHaveLength(0)
  })

  it('starts isolated fingerprint workers in their own process group for bounded termination', async () => {
    vi.useFakeTimers()

    const child = createMockChildProcess(24_680)
    const killSpy = vi.spyOn(process, 'kill').mockImplementation(() => true)
    spawnMock.mockImplementationOnce(() => child)

    const fingerprintPromise = fingerprintWorkbookFileIsolated('/tmp/public-budget.xlsx', 'public-budget.xlsx', 5, {
      maxRssBytes: 1024 * 1024,
      rssCheckIntervalMs: 100,
    })
    const fingerprintExpectation = expect(fingerprintPromise).rejects.toThrow('Workbook fingerprinting timed out after 5ms')

    await vi.advanceTimersByTimeAsync(5)
    await fingerprintExpectation

    expect(spawnMock.mock.calls[0]?.[2]).toMatchObject({ detached: true })
    expect(killSpy).toHaveBeenCalledWith(-24_680, 'SIGTERM')
    expect(child.kill).not.toHaveBeenCalled()
  })

  it('terminates isolated fingerprint process groups when the parent receives SIGTERM', async () => {
    vi.useFakeTimers()

    const child = createMockChildProcess(24_682)
    const killSpy = vi.spyOn(process, 'kill').mockImplementation(() => true)
    const exitSpy = vi.spyOn(process, 'exit').mockImplementation((code?: string | number | null): never => {
      throw new Error(`process.exit ${String(code)}`)
    })
    spawnMock.mockImplementationOnce(() => child)

    const fingerprintPromise = fingerprintWorkbookFileIsolated('/tmp/public-budget.xlsx', 'public-budget.xlsx', 5_000, {
      maxRssBytes: 1024 * 1024,
      rssCheckIntervalMs: 100,
    })

    expect(() => process.emit('SIGTERM')).toThrow('process.exit 143')
    child.emit('close', 1, null)

    await expect(fingerprintPromise).rejects.toThrow('Workbook fingerprinting subprocess exited with code 1')
    expect(killSpy).toHaveBeenCalledWith(-24_682, 'SIGTERM')
    expect(child.kill).not.toHaveBeenCalled()
    expect(exitSpy).toHaveBeenCalledWith(143)
  })

  it('skips malformed CKAN resource URLs during workbook discovery', async () => {
    const fetchMock = vi.fn(async () =>
      Response.json({
        result: {
          results: [
            {
              id: 'dataset-1',
              name: 'dataset-1',
              license_id: 'CC-BY-4.0',
              license_title: 'Creative Commons Attribution 4.0 International',
              license_url: 'https://creativecommons.org/licenses/by/4.0/',
              resources: [
                {
                  id: 'bad-url',
                  name: '',
                  url: 'http:// https://example.com/not-a-url.xlsx',
                },
                {
                  id: 'good-url',
                  name: 'workbook.xlsx',
                  url: 'https://example.com/workbook.xlsx',
                },
              ],
            },
          ],
        },
      }),
    )
    vi.stubGlobal('fetch', fetchMock)

    const manifest = await discoverCkanWorkbookSources({
      manifest: createEmptyPublicWorkbookManifest('2026-05-07T00:00:00.000Z'),
      portalBases: ['https://example-ckan.test/api/3/action'],
      query: 'xlsx',
      limit: 10,
      rowsPerRequest: 10,
      discoveredAt: '2026-05-07T01:00:00.000Z',
    })

    expect(manifest.sources.map((source) => source.resourceId)).toEqual(['good-url'])
    expect(manifest.sources[0]).toMatchObject({
      downloadUrl: 'https://example.com/workbook.xlsx',
      fileName: 'workbook.xlsx',
    })
  })

  it('skips CKAN resource URLs with malformed percent-encoded filenames', async () => {
    const fetchMock = vi.fn(async () =>
      Response.json({
        result: {
          results: [
            {
              id: 'dataset-malformed-filename',
              name: 'dataset-malformed-filename',
              license_id: 'CC-BY-4.0',
              license_title: 'Creative Commons Attribution 4.0 International',
              license_url: 'https://creativecommons.org/licenses/by/4.0/',
              resources: [
                {
                  id: 'bad-filename',
                  name: '',
                  url: 'https://example.com/public-workbooks/bad%EA.xlsx',
                },
                {
                  id: 'good-filename',
                  name: '',
                  url: 'https://example.com/public-workbooks/good.xlsx',
                },
              ],
            },
          ],
        },
      }),
    )
    vi.stubGlobal('fetch', fetchMock)

    const manifest = await discoverCkanWorkbookSources({
      manifest: createEmptyPublicWorkbookManifest('2026-05-07T00:00:00.000Z'),
      portalBases: ['https://example-ckan.test/api/3/action'],
      query: 'finance',
      limit: 10,
      rowsPerRequest: 10,
      discoveredAt: '2026-05-07T01:00:00.000Z',
    })

    expect(manifest.sources.map((source) => source.resourceId)).toEqual(['good-filename'])
    expect(manifest.sources[0]).toMatchObject({
      downloadUrl: 'https://example.com/public-workbooks/good.xlsx',
      fileName: 'good.xlsx',
    })
  })

  it('resolves relative CKAN resource URLs during workbook discovery', async () => {
    const fetchMock = vi.fn(async () =>
      Response.json({
        result: {
          results: [
            {
              id: 'dataset-relative',
              name: 'dataset-relative',
              license_id: 'CC-BY-4.0',
              license_title: 'Creative Commons Attribution 4.0 International',
              license_url: 'https://creativecommons.org/licenses/by/4.0/',
              resources: [
                {
                  id: 'relative-url',
                  name: '',
                  url: '/data/dataset/dataset-relative/resource/relative-url/download/output.xlsx',
                },
              ],
            },
          ],
        },
      }),
    )
    vi.stubGlobal('fetch', fetchMock)

    const manifest = await discoverCkanWorkbookSources({
      manifest: createEmptyPublicWorkbookManifest('2026-05-07T00:00:00.000Z'),
      portalBases: ['https://example-ckan.test/data/api/3/action'],
      query: 'xlsx',
      limit: 10,
      rowsPerRequest: 10,
      discoveredAt: '2026-05-07T01:00:00.000Z',
    })

    expect(manifest.sources).toHaveLength(1)
    expect(manifest.sources[0]).toMatchObject({
      downloadUrl: 'https://example-ckan.test/data/dataset/dataset-relative/resource/relative-url/download/output.xlsx',
      fileName: 'output.xlsx',
      resourceId: 'relative-url',
    })
  })

  it('filters CKAN discovery to financial workbook topic evidence when requested', async () => {
    const fetchMock = vi.fn(async () =>
      Response.json({
        result: {
          results: [
            {
              id: 'dataset-budget',
              name: 'state-budget',
              title: 'State Budget Financial Tables',
              license_id: 'CC-BY-4.0',
              license_title: 'Creative Commons Attribution 4.0 International',
              license_url: 'https://creativecommons.org/licenses/by/4.0/',
              resources: [{ id: 'budget-resource', name: 'tables.xlsx', url: 'https://example.com/tables.xlsx' }],
            },
            {
              id: 'dataset-population',
              name: 'population',
              title: 'Population estimates',
              license_id: 'CC-BY-4.0',
              license_title: 'Creative Commons Attribution 4.0 International',
              license_url: 'https://creativecommons.org/licenses/by/4.0/',
              resources: [{ id: 'population-resource', name: 'population.xlsx', url: 'https://example.com/population.xlsx' }],
            },
          ],
        },
      }),
    )
    vi.stubGlobal('fetch', fetchMock)

    const manifest = await discoverCkanWorkbookSources({
      manifest: createEmptyPublicWorkbookManifest('2026-05-07T00:00:00.000Z', 5_000),
      portalBases: ['https://example-ckan.test/api/3/action'],
      query: 'budget',
      limit: 5_000,
      rowsPerRequest: 10,
      discoveredAt: '2026-05-07T01:00:00.000Z',
      requiredTopic: 'financial-workpapers',
    })

    expect(manifest.sources.map((source) => source.resourceId)).toEqual(['budget-resource'])
    expect(manifest.sources[0]?.topicEvidence).toEqual(expect.arrayContaining(['budget:dataset.title']))
  })

  it('surfaces isolated verification subprocess failures with stderr evidence', async () => {
    const fixture = createIsolatedVerificationFixture()
    const child = createMockChildProcess()
    spawnMock.mockImplementationOnce(() => child)

    const scorecardPromise = buildPublicWorkbookCorpusScorecard({
      manifest: fixture.manifest,
      cacheDir: fixture.cacheDir,
      manifestPath: fixture.manifestPath,
      generatedAt: '2026-05-07T01:00:00.000Z',
      isolatedVerification: true,
      verifyConcurrency: 1,
      verifyTimeoutMs: 1_000,
    })

    child.stderr.emit('data', 'Error: Cannot find module "./public-workbook-corpus.ts"\n')
    child.emit('close', 1, null)

    const scorecard = await scorecardPromise

    expect(spawnMock).toHaveBeenCalledTimes(1)
    expect(scorecard.cases).toHaveLength(1)
    expect(scorecard.cases[0]).toMatchObject({
      status: 'error',
      passed: false,
    })
    expect(scorecard.cases[0]?.evidence).toEqual(
      expect.arrayContaining(['Verification subprocess exited with code 1', 'Error: Cannot find module "./public-workbook-corpus.ts"']),
    )
  })

  it('starts one isolated verification worker by default', async () => {
    const fixture = createIsolatedVerificationFixture()
    const artifact = fixture.manifest.artifacts[0]
    if (!artifact) {
      throw new Error('expected isolated verification fixture artifact')
    }
    const firstChild = createMockChildProcess()
    const secondChild = createMockChildProcess()
    const children = [firstChild, secondChild]
    spawnMock.mockImplementation(() => children.shift() ?? createMockChildProcess())

    const scorecardPromise = buildPublicWorkbookCorpusScorecard({
      manifest: {
        ...fixture.manifest,
        artifacts: [
          artifact,
          {
            ...artifact,
            id: 'workbook-isolated-verification-2',
            cachePath: 'files/public-budget-2.xlsx',
            sha256: 'b'.repeat(64),
            workbookFingerprint: 'isolated-verification-fingerprint-2',
          },
        ],
      },
      cacheDir: fixture.cacheDir,
      manifestPath: fixture.manifestPath,
      generatedAt: '2026-05-07T01:00:00.000Z',
      isolatedVerification: true,
      verifyTimeoutMs: 1_000,
    })

    expect(spawnMock).toHaveBeenCalledTimes(1)
    firstChild.emit('close', 1, null)
    await Promise.resolve()
    await Promise.resolve()

    expect(spawnMock).toHaveBeenCalledTimes(2)
    secondChild.emit('close', 1, null)

    const scorecard = await scorecardPromise
    expect(scorecard.summary).toMatchObject({
      cachedWorkbookCount: 2,
      errorWorkbookCount: 2,
    })
  })

  it('reports isolated verification timeouts in the evidence trail', async () => {
    vi.useFakeTimers()

    const fixture = createIsolatedVerificationFixture()
    const child = createMockChildProcess(24_681)
    const killSpy = vi.spyOn(process, 'kill').mockImplementation(() => true)
    spawnMock.mockImplementationOnce(() => child)

    const scorecardPromise = buildPublicWorkbookCorpusScorecard({
      manifest: fixture.manifest,
      cacheDir: fixture.cacheDir,
      manifestPath: fixture.manifestPath,
      generatedAt: '2026-05-07T01:00:00.000Z',
      isolatedVerification: true,
      verifyConcurrency: 1,
      verifyTimeoutMs: 5,
    })

    await vi.advanceTimersByTimeAsync(5)
    const scorecard = await scorecardPromise

    expect(spawnMock).toHaveBeenCalledTimes(1)
    expect(spawnMock.mock.calls[0]?.[2]).toMatchObject({ detached: true })
    expect(spawnMock.mock.calls[0]?.[1]).toEqual(expect.arrayContaining(['verify-artifact-worker', '--verify-max-rss-mb', '1536']))
    expect(killSpy).toHaveBeenCalledWith(-24_681, 'SIGTERM')
    expect(child.kill).not.toHaveBeenCalled()
    expect(scorecard.cases[0]?.status).toBe('error')
    expect(scorecard.cases[0]?.evidence).toEqual(
      expect.arrayContaining([
        'Verification timed out after 5ms',
        'The workbook was isolated in a subprocess so the corpus verification run could continue.',
      ]),
    )
  })

  it('kills isolated verification workers that exceed the RSS limit', async () => {
    vi.useFakeTimers()

    const fixture = createIsolatedVerificationFixture()
    const verificationChild = createMockChildProcess(12_345)
    const psChild = createMockChildProcess()
    const killSpy = vi.spyOn(process, 'kill').mockImplementation(() => true)
    spawnMock.mockImplementation((command: string) => (command === '/bin/ps' ? psChild : verificationChild))

    const scorecardPromise = buildPublicWorkbookCorpusScorecard({
      manifest: fixture.manifest,
      cacheDir: fixture.cacheDir,
      manifestPath: fixture.manifestPath,
      generatedAt: '2026-05-07T01:00:00.000Z',
      isolatedVerification: true,
      verifyConcurrency: 1,
      verifyTimeoutMs: 60_000,
      verifyMaxRssBytes: 1024 * 1024,
      verifyRssCheckIntervalMs: 100,
    })

    verificationChild.stderr.emit('data', 'bilig-public-workbook-verify-phase=round-trip\n')
    await vi.advanceTimersByTimeAsync(100)
    psChild.stdout.emit('data', '2048\n')
    psChild.emit('close', 0, null)

    const scorecard = await scorecardPromise

    expect(spawnMock.mock.calls[0]?.[1]).toEqual(expect.arrayContaining(['verify-artifact-worker', '--verify-max-rss-mb', '1']))
    expect(spawnMock.mock.calls[0]?.[2]).toMatchObject({ detached: true })
    expect(killSpy).toHaveBeenCalledWith(-12_345, 'SIGTERM')
    expect(verificationChild.kill).not.toHaveBeenCalled()
    expect(scorecard.cases[0]?.status).toBe('unsupported')
    expect(scorecard.cases[0]?.passed).toBe(true)
    expect(scorecard.cases[0]?.unsupportedFeatureClassifications).toEqual(['xlsx.publicCorpus.resourceLimit:rss>1MiB'])
    expect(scorecard.cases[0]?.evidence).toEqual(
      expect.arrayContaining([
        publicWorkbookResourceLimitClassifierEvidence,
        'Public corpus verification RSS limit exceeded: 2.0 MiB > 1.0 MiB',
        'rss-limit-phase=round-trip',
        'peak-rss=2.0 MiB',
        'The workbook was isolated in a subprocess so the corpus verification run could continue.',
      ]),
    )
  })

  it('rejects high isolated verification RSS overrides unless explicitly enabled', async () => {
    const fixture = createIsolatedVerificationFixture()
    await expect(
      buildPublicWorkbookCorpusScorecard({
        manifest: fixture.manifest,
        cacheDir: fixture.cacheDir,
        manifestPath: fixture.manifestPath,
        generatedAt: '2026-05-07T01:00:00.000Z',
        isolatedVerification: true,
        verifyConcurrency: 1,
        verifyTimeoutMs: 5,
        verifyMaxRssBytes: 12 * 1024 * 1024 * 1024,
      }),
    ).rejects.toThrow('RSS limits above 1536 MiB are disabled because')
    expect(spawnMock).not.toHaveBeenCalled()
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

function buildWorkbookBytes(summarySheetName = 'Summary'): Uint8Array {
  const workbook = XLSX.utils.book_new()
  const summary = XLSX.utils.aoa_to_sheet([
    ['Metric', 'Value'],
    ['Revenue', 12],
    ['Cost', 5],
    ['Profit', null],
  ])
  summary.B4 = { t: 'n', f: 'B2-B3', v: 7 }
  summary['!ref'] = 'A1:B4'
  summary['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 1 } }]
  const assumptions = XLSX.utils.aoa_to_sheet([['TaxRate'], [0.21]])
  XLSX.utils.book_append_sheet(workbook, summary, summarySheetName)
  XLSX.utils.book_append_sheet(workbook, assumptions, 'Assumptions')
  workbook.Workbook = {
    Names: [{ Name: 'ProfitCell', Ref: `${summarySheetName}!$B$4` }],
  }
  return XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' })
}

function buildExternalWorkbookReferenceBytes(): Uint8Array {
  const workbook = XLSX.utils.book_new()
  const sheet = XLSX.utils.aoa_to_sheet([
    ['Key', 'Value'],
    ['A', null],
  ])
  sheet.B2 = { t: 'n', f: "VLOOKUP(A2,'[1]Lookup'!$A$1:$B$2,2,FALSE)", v: 12 }
  sheet['!ref'] = 'A1:B2'
  XLSX.utils.book_append_sheet(workbook, sheet, 'Summary')
  return XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' })
}

function buildManualCalculationWorkbookBytes(): Uint8Array {
  const workbook = XLSX.utils.book_new()
  const sheet = XLSX.utils.aoa_to_sheet([
    ['Input', 'Value'],
    ['A', 1],
    ['B', 2],
    ['Total', null],
  ])
  sheet.B4 = { t: 'n', f: 'B2+B3', v: 99 }
  sheet['!ref'] = 'A1:B4'
  XLSX.utils.book_append_sheet(workbook, sheet, 'Summary')
  return addExportCalculationSettingsToXlsxBytes(XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' }), {
    version: 1,
    workbook: {
      name: 'manual-calculation',
      metadata: { calculationSettings: { mode: 'manual', compatibilityMode: 'excel-modern' } },
    },
    sheets: [],
  })
}

function buildPrecisionAsDisplayedFormulaWorkbookBytes(): Uint8Array {
  const workbook = XLSX.utils.book_new()
  const sheet = XLSX.utils.aoa_to_sheet([
    ['Input', 'Value'],
    ['A', 1],
    ['B', 2],
    ['Total', null],
  ])
  sheet.B4 = { t: 'n', f: 'B2+B3', v: 3 }
  sheet['!ref'] = 'A1:B4'
  XLSX.utils.book_append_sheet(workbook, sheet, 'Summary')
  return addExportCalculationSettingsToXlsxBytes(XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' }), {
    version: 1,
    workbook: {
      name: 'precision-displayed',
      metadata: { calculationSettings: { mode: 'automatic', compatibilityMode: 'excel-modern', fullPrecision: false } },
    },
    sheets: [],
  })
}

function buildVolatileFormulaWorkbookBytes(): Uint8Array {
  const workbook = XLSX.utils.book_new()
  const sheet = XLSX.utils.aoa_to_sheet([
    ['Metric', 'Value'],
    ['Report Date', null],
  ])
  sheet.B2 = { t: 'n', f: 'TODAY()', v: 43073 }
  sheet['!ref'] = 'A1:B2'
  XLSX.utils.book_append_sheet(workbook, sheet, 'Summary')
  return XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' })
}

function buildMacroEnabledWorkbookBytes(): Uint8Array {
  const workbook = XLSX.utils.book_new()
  const sheet = XLSX.utils.aoa_to_sheet([
    ['Input', 'Value'],
    ['A', 1],
    ['B', 2],
    ['Total', null],
  ])
  sheet.B4 = { t: 'n', f: 'B2+B3', v: 99 }
  sheet['!ref'] = 'A1:B4'
  workbook.vbaraw = new Uint8Array([1, 2, 3, 4])
  XLSX.utils.book_append_sheet(workbook, sheet, 'Summary')
  return XLSX.write(workbook, { bookType: 'xlsm', type: 'buffer', bookVBA: true })
}

function buildStaleFormulaCacheWorkbookBytes(): Uint8Array {
  const workbook = XLSX.utils.book_new()
  const sheet = XLSX.utils.aoa_to_sheet([
    ['Text', 'Word Count'],
    ['one two three four five six seven eight nine', null],
  ])
  sheet.B2 = { t: 'n', f: 'LEN(TRIM(A2))-LEN(SUBSTITUTE(A2," ",""))+1', v: 6 }
  sheet['!ref'] = 'A1:B2'
  XLSX.utils.book_append_sheet(workbook, sheet, 'Summary')
  return XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' })
}

function buildUnconfirmedFormulaCacheWorkbookBytes(): Uint8Array {
  const workbook = XLSX.utils.book_new()
  const sheet = XLSX.utils.aoa_to_sheet([
    ['Value', 'Result'],
    [42, null],
  ])
  sheet.B2 = { t: 'n', f: 'UNKNOWNFUNC(A2)', v: 99 }
  sheet['!ref'] = 'A1:B2'
  XLSX.utils.book_append_sheet(workbook, sheet, 'Summary')
  return XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' })
}

function buildTrailingSpaceWorkbookBytes(): Uint8Array {
  const workbook = XLSX.utils.book_new()
  const sheetName = 'Table 2.1.2  '
  const sheet = XLSX.utils.aoa_to_sheet([
    ['Header', 'Value'],
    ['Amount', 12],
  ])
  sheet['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 1 } }]
  XLSX.utils.book_append_sheet(workbook, sheet, sheetName)
  return XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' })
}

function buildInteriorStyledSnapshot(startAddress: string, endAddress: string, backgroundColor = '#ffcc00'): WorkbookSnapshot {
  return {
    version: 1,
    workbook: {
      name: 'interior-styled-range',
      metadata: {
        styles: [
          {
            id: 'accent-style',
            fill: { backgroundColor },
            font: { bold: true },
          },
        ],
      },
    },
    sheets: [
      {
        id: 1,
        name: 'Styled',
        order: 0,
        cells: [{ address: 'B2', value: 'Interior' }],
        metadata: {
          styleRanges: [
            {
              range: { sheetName: 'Styled', startAddress, endAddress },
              styleId: 'accent-style',
            },
          ],
        },
      },
    ],
  }
}

function buildAxisDimensionSnapshot(metadata: NonNullable<WorkbookSnapshot['sheets'][number]['metadata']>): WorkbookSnapshot {
  return {
    version: 1,
    workbook: { name: 'axis-dimensions' },
    sheets: [
      {
        id: 1,
        name: 'Dimensions',
        order: 0,
        cells: [{ address: 'B2', value: 'Visible' }],
        metadata,
      },
    ],
  }
}

function buildChartOrientationSnapshot(seriesOrientation?: 'columns' | 'rows'): WorkbookSnapshot {
  return {
    version: 1,
    workbook: {
      name: 'chart-orientation',
      metadata: {
        charts: [
          {
            id: 'sales-chart',
            sheetName: 'Chart Data',
            address: 'E1',
            source: { sheetName: 'Chart Data', startAddress: 'A1', endAddress: 'B3' },
            chartType: 'column',
            ...(seriesOrientation !== undefined ? { seriesOrientation } : {}),
            firstRowAsHeaders: true,
            firstColumnAsLabels: true,
            rows: 12,
            cols: 6,
          },
        ],
      },
    },
    sheets: [
      {
        id: 1,
        name: 'Chart Data',
        order: 0,
        cells: [
          { address: 'A1', value: 'Month' },
          { address: 'B1', value: 'Revenue' },
          { address: 'A2', value: 'Jan' },
          { address: 'B2', value: 10 },
          { address: 'A3', value: 'Feb' },
          { address: 'B3', value: 12 },
        ],
      },
    ],
  }
}

function buildProtectedFirstWorkbookBytes(): Uint8Array {
  const workbook = XLSX.utils.book_new()
  const protectedSheet = XLSX.utils.aoa_to_sheet([
    ['Locked', 'Value'],
    ['Amount', 12],
  ])
  protectedSheet['!protect'] = {}
  const mutableSheet = XLSX.utils.aoa_to_sheet([
    ['Open', 'Value'],
    ['Amount', 7],
  ])
  XLSX.utils.book_append_sheet(workbook, protectedSheet, 'Locked')
  XLSX.utils.book_append_sheet(workbook, mutableSheet, 'Open')
  return XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' })
}

function buildProtectedFirstSheetWorkbookBytes(): Uint8Array {
  const workbook = XLSX.utils.book_new()
  const locked = XLSX.utils.aoa_to_sheet([['Locked'], [1]])
  locked['!protect'] = {}
  const open = XLSX.utils.aoa_to_sheet([['Open'], [2]])
  XLSX.utils.book_append_sheet(workbook, locked, 'Locked')
  XLSX.utils.book_append_sheet(workbook, open, 'Open')
  return XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' })
}

function createIsolatedVerificationFixture(): {
  readonly cacheDir: string
  readonly manifestPath: string
  readonly manifest: PublicWorkbookManifest
} {
  const cacheDir = mkdtempSync(join(tmpdir(), 'public-workbook-corpus-isolated-'))
  const manifestPath = join(cacheDir, 'manifest.json')
  const manifest = {
    ...createEmptyPublicWorkbookManifest('2026-05-07T00:00:00.000Z'),
    sources: [
      {
        id: 'source-isolated-verification',
        kind: 'direct-url',
        sourceUrl: 'https://example.com/public-budget.xlsx',
        downloadUrl: 'https://example.com/public-budget.xlsx',
        fileName: 'public-budget.xlsx',
        discoveredAt: '2026-05-07T00:00:00.000Z',
        license: {
          spdxId: 'CC-BY-4.0',
          title: 'Creative Commons Attribution 4.0 International',
          evidenceUrl: 'https://creativecommons.org/licenses/by/4.0/',
        },
      },
    ],
    artifacts: [
      {
        id: 'workbook-isolated-verification',
        sourceId: 'source-isolated-verification',
        sourceUrl: 'https://example.com/public-budget.xlsx',
        downloadUrl: 'https://example.com/public-budget.xlsx',
        fileName: 'public-budget.xlsx',
        cachePath: 'files/public-budget.xlsx',
        sha256: 'a'.repeat(64),
        byteSize: 1_024,
        workbookFingerprint: 'isolated-verification-fingerprint',
        fetchedAt: '2026-05-07T00:00:00.000Z',
        license: {
          spdxId: 'CC-BY-4.0',
          title: 'Creative Commons Attribution 4.0 International',
          evidenceUrl: 'https://creativecommons.org/licenses/by/4.0/',
        },
      },
    ],
  } satisfies PublicWorkbookManifest

  return { cacheDir, manifestPath, manifest }
}

class MockProcessStream extends EventEmitter {
  readonly setEncoding = vi.fn()
}

class MockChildProcess extends EventEmitter {
  readonly stdout = new MockProcessStream()
  readonly stderr = new MockProcessStream()
  readonly kill = vi.fn()

  constructor(readonly pid?: number) {
    super()
  }
}

function createMockChildProcess(pid?: number): MockChildProcess {
  return new MockChildProcess(pid)
}
