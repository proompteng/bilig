import { describe, expect, it } from 'vitest'

import {
  buildPublicWorkbookCorpusFeatureWitnessPlan,
  validatePublicWorkbookCorpusFeatureWitnessPlan,
} from '../public-workbook-corpus-feature-witness-plan.ts'
import type { PublicWorkbookCorpusCase, PublicWorkbookFeatureCounts } from '../public-workbook-corpus-types.ts'

describe('public workbook corpus feature witness plan', () => {
  it('reports missing pivot witnesses with a guarded targeted discovery command', () => {
    const plan = buildPublicWorkbookCorpusFeatureWitnessPlan({
      cacheDir: '/repo/.cache/public-workbook-corpus',
      cases: [caseWithFeatures({ pivotCount: 0 })],
      discoveryLimit: 10_000,
      displayRootDir: '/repo',
      generatedAt: '2026-05-08T08:00:00.000Z',
      manifestPath: '/repo/.cache/public-workbook-corpus/manifest.json',
      stopMarkerActive: true,
      stopMarkerPath: '/repo/.agent-coordination/stop.md',
    })
    const pivotCoverage = plan.coverage.find((entry) => entry.id === 'pivots')

    expect(plan).toMatchObject({
      schemaVersion: 1,
      mode: 'feature-witness-plan',
      missingWitnessCount: 1,
      recordedCaseCount: 1,
      stopMarker: {
        active: true,
        path: '.agent-coordination/stop.md',
      },
    })
    expect(pivotCoverage).toMatchObject({
      discoveryQuery: 'pivot table xlsx',
      needsWitness: true,
      totalCount: 0,
      witnessCaseCount: 0,
    })
    expect(pivotCoverage?.commands.discover).toContain('BILIG_ALLOW_PUBLIC_CORPUS_STOP_MARKER_OVERRIDE=1')
    expect(pivotCoverage?.commands.discover).toContain('public-workbook-corpus:discover')
    expect(pivotCoverage?.commands.discover).toContain("--query 'pivot table xlsx'")
    expect(pivotCoverage?.commands.discover).toContain('--allow-active-stop-marker')
    expect(JSON.stringify(plan)).not.toContain('/repo/')
    expect(validatePublicWorkbookCorpusFeatureWitnessPlan(plan)).toEqual([])
  })
})

function caseWithFeatures(featureCounts: Partial<PublicWorkbookFeatureCounts>): PublicWorkbookCorpusCase {
  return {
    id: 'artifact-a',
    sourceId: 'source-a',
    sourceUrl: 'https://example.com/source-a.xlsx',
    fileName: 'source-a.xlsx',
    sha256: 'a'.repeat(64),
    byteSize: 1024,
    license: {
      spdxId: 'CC-BY-4.0',
      title: 'Creative Commons Attribution 4.0 International',
      evidenceUrl: 'https://creativecommons.org/licenses/by/4.0/',
    },
    status: 'passed',
    passed: true,
    featureCounts: {
      sheetCount: 1,
      cellCount: 9,
      formulaCellCount: 1,
      valueCellCount: 1,
      definedNameCount: 1,
      tableCount: 1,
      chartCount: 1,
      pivotCount: 1,
      mergeCount: 1,
      styleRangeCount: 1,
      conditionalFormatCount: 1,
      dataValidationCount: 0,
      macroPayloadCount: 0,
      warningCount: 0,
      ...featureCounts,
    },
    workbookMetadata: {
      workbookName: 'source-a',
      sheetNames: ['Sheet1'],
      dimensions: [
        {
          sheetName: 'Sheet1',
          rowCount: 3,
          columnCount: 3,
          nonEmptyCellCount: 9,
          usedRange: { startRow: 0, startColumn: 0, endRow: 2, endColumn: 2 },
        },
      ],
    },
    validation: {
      importPassed: true,
      formulaOraclePassed: true,
      formulaOracleComparisons: 1,
      formulaOracleMismatches: [],
      roundTripPassed: true,
      structuralSmokePassed: true,
    },
    unsupportedFeatureClassifications: [],
    evidence: [
      'source=https://example.com/source-a.xlsx',
      'license=Creative Commons Attribution 4.0 International',
      `sha256=${'a'.repeat(64)}`,
    ],
  }
}
