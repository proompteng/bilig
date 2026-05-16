import { mkdirSync, mkdtempSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'

import { hasUiResponsivenessSameCorpusTenXGap } from '../bilig-dominance-completion-audit.ts'
import {
  buildSameCorpusProof,
  assertUiResponsivenessLiveBrowserRunAllowed,
  parseUiResponsivenessLiveBrowserCliArgs,
  parseUiResponsivenessLiveBrowserScorecard,
  validateUiResponsivenessLiveBrowserScorecard,
  type SameCorpusCapture,
  type UiResponsivenessLiveBrowserScorecard,
} from '../gen-ui-responsiveness-live-browser-scorecard.ts'
import { readJsonObject } from '../json-scorecard-helpers.ts'
import {
  requiredUiResponsivenessSameCorpusWorkloads,
  uiSameCorpusWorkloadRequiresScrollEventEvidence,
  type UiResponsivenessSameCorpusWorkload,
} from '../ui-responsiveness-same-corpus-workloads.ts'

const repoRoot = resolve(dirname(fileURLToPath(import.meta.url)), '../..')

describe('UI responsiveness live browser scorecard', () => {
  it('validates the checked-in browser timing artifact', () => {
    const scorecard = parseUiResponsivenessLiveBrowserScorecard(
      readJsonObject(resolve(repoRoot, 'packages/benchmarks/baselines/ui-responsiveness-live-browser-scorecard.json')),
    )

    expect(scorecard.summary).toMatchObject({
      directBrowserTimingCaptured: true,
      allRequiredCasesPassed: true,
      requiredVendorCount: 2,
      capturedVendors: ['google-sheets', 'microsoft-excel-web'],
    })
    expect(scorecard.cases.map((entry) => entry.id)).toEqual(['google-sheets-public-grid-scroll', 'microsoft-excel-web-public-xlsx-scroll'])
    expect(scorecard.cases.every((entry) => entry.sampleCount >= 3 && entry.limitations.length > 0)).toBe(true)
    expect(scorecard.sameCorpusProof).toMatchObject({
      captured: true,
      evidenceKind: 'same-corpus-browser-capture',
      requiredProductCount: 2,
      requiredCaseCount: requiredUiResponsivenessSameCorpusWorkloads.length,
      coveredCorpusCaseIds: ['wide-mixed-250k'],
    })
    expect(scorecard.sameCorpusProof.cases.map((entry) => entry.workload)).toEqual(requiredUiResponsivenessSameCorpusWorkloads)
    expect(scorecard.sameCorpusProof.tenXMeanAndP95CaseCount).toBe(
      scorecard.sameCorpusProof.cases.filter((entry) => entry.tenXMeanAndP95AgainstGoogleSheets).length,
    )
    expect(scorecard.sameCorpusProof.tenXMeanAndP95CaseCount).toBeGreaterThan(0)
    validateUiResponsivenessLiveBrowserScorecard(scorecard)
  })

  it('parses live browser scorecard CLI options', () => {
    expect(parseUiResponsivenessLiveBrowserCliArgs(['--check', '--capture', 'tmp/same-corpus-capture.json'])).toEqual({
      isCheckMode: true,
      capturePath: 'tmp/same-corpus-capture.json',
    })
  })

  it('rejects blank live browser capture paths', () => {
    expect(() => parseUiResponsivenessLiveBrowserCliArgs(['--capture', '   '])).toThrow('Missing value after --capture')
  })

  it('rejects live browser capture paths that consume the next flag', () => {
    expect(() => parseUiResponsivenessLiveBrowserCliArgs(['--capture', '--check'])).toThrow('Missing value after --capture')
  })

  it('rejects missing incumbent vendors', () => {
    const scorecard = parseUiResponsivenessLiveBrowserScorecard(
      readJsonObject(resolve(repoRoot, 'packages/benchmarks/baselines/ui-responsiveness-live-browser-scorecard.json')),
    )
    const staleScorecard: UiResponsivenessLiveBrowserScorecard = {
      ...scorecard,
      summary: {
        ...scorecard.summary,
        capturedVendors: ['google-sheets'],
      },
    }

    expect(() => validateUiResponsivenessLiveBrowserScorecard(staleScorecard)).toThrow(
      'UI responsiveness live browser scorecard is missing vendor: microsoft-excel-web',
    )
  })

  it('blocks live browser scorecard generation while the local resource guard is active', () => {
    const rootDir = mkdtempSync(`${tmpdir()}/bilig-ui-browser-live-guard-`)
    const coordinationDir = resolve(rootDir, '.agent-coordination')
    mkdirSync(coordinationDir)
    writeFileSync(
      resolve(coordinationDir, '20260508T092619Z-codex-memory-pressure-stop.md'),
      '# Memory pressure stop\n\nStatus: active on 2026-05-08T09:26:19Z.\n',
    )

    expect(() => assertUiResponsivenessLiveBrowserRunAllowed(rootDir, {})).toThrow(
      /Refusing to start UI responsiveness live browser scorecard generation/u,
    )
    expect(() => assertUiResponsivenessLiveBrowserRunAllowed(rootDir, { BILIG_ALLOW_LOCAL_CI_RESOURCE_GUARD: '1' })).not.toThrow()
  })

  it('derives same-corpus 10x ratios from operation and scroll-event samples', () => {
    const proof = buildSameCorpusProof(buildSameCorpusCapture())

    expect(proof).toMatchObject({
      captured: true,
      evidenceKind: 'same-corpus-browser-capture',
      requiredProductCount: 2,
      requiredCaseCount: requiredUiResponsivenessSameCorpusWorkloads.length,
      tenXMeanAndP95CaseCount: requiredUiResponsivenessSameCorpusWorkloads.length,
      coveredCorpusCaseIds: ['wide-mixed-250k'],
    })
    expect(proof.cases[0]).toMatchObject({
      biligToGoogleSheetsMeanRatio: 0.05,
      biligToGoogleSheetsP95Ratio: 0.06,
      biligToMicrosoftExcelWebMeanRatio: 0.0625,
      biligToMicrosoftExcelWebP95Ratio: 0.06666666666666667,
      tenXMeanAndP95Metric: 'operationResponseMs',
      scenarioProof: {
        biligMeanMs: 5,
        biligP95Ms: 6,
        googleMeanMs: 100,
        googleP95Ms: 100,
        microsoftExcelWebMeanMs: 80,
        microsoftExcelWebP95Ms: 90,
        meanRatio: 0.05,
        p95Ratio: 0.06,
        microsoftExcelWebMeanRatio: 0.0625,
        microsoftExcelWebP95Ratio: 0.06666666666666667,
        screenshotProof: { captured: true, missingProducts: [] },
        pixelGridProof: { captured: true, missingProducts: [] },
      },
      postOperationFrameGuardrailPassed: true,
      passed: true,
    })
    expect(proof.cases.find((entry) => entry.workload === 'scroll-vertical')).toMatchObject({
      biligToGoogleSheetsScrollEventMeanRatio: 0.05,
      biligToGoogleSheetsScrollEventP95Ratio: 0.06,
      biligToMicrosoftExcelWebScrollEventMeanRatio: 0.0625,
      biligToMicrosoftExcelWebScrollEventP95Ratio: 0.06666666666666667,
      tenXMeanAndP95Metric: 'scrollEventResponseMs',
      scrollMovementGuardrailPassed: true,
      passed: true,
    })
  })

  it('rejects legacy operation-only same-corpus captures before generating proof', () => {
    expect(() =>
      buildSameCorpusProof(buildSameCorpusCapture({ includeScrollEventSamples: false, workloads: ['scroll-vertical'] })),
    ).toThrow('UI responsiveness same-corpus capture has too few scroll-event samples for same-corpus-wide-mixed-250k-scroll-vertical')
  })

  it('rejects captured same-corpus proof without scroll-event evidence', () => {
    const scorecard = parseUiResponsivenessLiveBrowserScorecard(
      readJsonObject(resolve(repoRoot, 'packages/benchmarks/baselines/ui-responsiveness-live-browser-scorecard.json')),
    )
    const proof = buildSameCorpusProof(buildSameCorpusCapture())
    const scrollCase = proof.cases.find((entry) => entry.workload === 'scroll-vertical')
    if (!scrollCase) {
      throw new Error('missing scroll-vertical fixture case')
    }
    const {
      scrollEventResponseMs: _biligScrollEventResponseMs,
      scrollMovementPx: _biligScrollMovementPx,
      ...biligWithoutScrollEvidence
    } = scrollCase.bilig
    const {
      scrollEventResponseMs: _googleSheetsScrollEventResponseMs,
      scrollMovementPx: _googleSheetsScrollMovementPx,
      ...googleSheetsWithoutScrollEvidence
    } = scrollCase.googleSheets
    const {
      scrollEventResponseMs: _microsoftExcelWebScrollEventResponseMs,
      scrollMovementPx: _microsoftExcelWebScrollMovementPx,
      ...microsoftExcelWebWithoutScrollEvidence
    } = scrollCase.microsoftExcelWeb
    const staleScorecard: UiResponsivenessLiveBrowserScorecard = {
      ...scorecard,
      sameCorpusProof: {
        ...proof,
        cases: proof.cases.map((entry) =>
          entry.id === scrollCase.id
            ? Object.assign({}, scrollCase, {
                bilig: biligWithoutScrollEvidence,
                googleSheets: googleSheetsWithoutScrollEvidence,
                microsoftExcelWeb: microsoftExcelWebWithoutScrollEvidence,
              })
            : entry,
        ),
      },
    }

    expect(() => validateUiResponsivenessLiveBrowserScorecard(staleScorecard)).toThrow(
      'UI responsiveness same-corpus proof is missing scroll-event evidence for same-corpus-wide-mixed-250k-scroll-vertical',
    )
  })

  it('allows same-corpus proof to clear the public-browser limitation blocker', () => {
    const scorecard = parseUiResponsivenessLiveBrowserScorecard(
      readJsonObject(resolve(repoRoot, 'packages/benchmarks/baselines/ui-responsiveness-live-browser-scorecard.json')),
    )

    expect(
      hasUiResponsivenessSameCorpusTenXGap({
        ...scorecard,
        sameCorpusProof: buildSameCorpusProof(buildSameCorpusCapture()),
      }),
    ).toBe(false)
  })

  it('keeps the same-corpus blocker when required scroll evidence is missing', () => {
    const scorecard = parseUiResponsivenessLiveBrowserScorecard(
      readJsonObject(resolve(repoRoot, 'packages/benchmarks/baselines/ui-responsiveness-live-browser-scorecard.json')),
    )

    expect(() => buildSameCorpusProof(buildSameCorpusCapture({ workloads: ['open-workbook'] }))).toThrow(
      'UI responsiveness same-corpus proof is missing required workload: select-cell',
    )
    expect(
      hasUiResponsivenessSameCorpusTenXGap({
        ...scorecard,
        sameCorpusProof: {
          ...scorecard.sameCorpusProof,
          captured: true,
          evidenceKind: 'same-corpus-browser-capture',
          requiredCaseCount: 1,
          tenXMeanAndP95CaseCount: 1,
          coveredCorpusCaseIds: ['wide-mixed-250k'],
          cases: [],
        },
      }),
    ).toBe(true)
  })

  it('rejects stale same-corpus pass flags and ratios', () => {
    const scorecard = parseUiResponsivenessLiveBrowserScorecard(
      readJsonObject(resolve(repoRoot, 'packages/benchmarks/baselines/ui-responsiveness-live-browser-scorecard.json')),
    )
    const proof = buildSameCorpusProof(buildSameCorpusCapture())
    const staleScorecard: UiResponsivenessLiveBrowserScorecard = {
      ...scorecard,
      sameCorpusProof: {
        ...proof,
        cases: proof.cases.map((entry, index) => (index === 0 ? Object.assign({}, entry, { biligToGoogleSheetsP95Ratio: 0.2 }) : entry)),
      },
    }

    expect(() => validateUiResponsivenessLiveBrowserScorecard(staleScorecard)).toThrow('UI responsiveness same-corpus ratio is stale')
  })

  it('rejects stale same-corpus visual proof required-product metadata', () => {
    const scorecard = parseUiResponsivenessLiveBrowserScorecard(
      readJsonObject(resolve(repoRoot, 'packages/benchmarks/baselines/ui-responsiveness-live-browser-scorecard.json')),
    )
    const proof = buildSameCorpusProof(buildSameCorpusCapture())
    const staleScorecard: UiResponsivenessLiveBrowserScorecard = {
      ...scorecard,
      sameCorpusProof: {
        ...proof,
        cases: proof.cases.map((entry, index) =>
          index === 0
            ? Object.assign({}, entry, {
                scenarioProof: Object.assign({}, entry.scenarioProof, {
                  screenshotProof: Object.assign({}, entry.scenarioProof.screenshotProof, {
                    requiredProducts: ['bilig', 'google-sheets', 'microsoft-excel-web'],
                  }),
                }),
              })
            : entry,
        ),
      },
    }

    expect(() => validateUiResponsivenessLiveBrowserScorecard(staleScorecard)).toThrow(
      'UI responsiveness same-corpus screenshot proof is stale',
    )
  })
})

function buildSameCorpusCapture(
  args: {
    readonly includeScrollEventSamples?: boolean
    readonly workloads?: readonly UiResponsivenessSameCorpusWorkload[]
  } = {},
): SameCorpusCapture {
  const includeScrollEventSamples = args.includeScrollEventSamples ?? true
  const workloads = args.workloads ?? requiredUiResponsivenessSameCorpusWorkloads
  return {
    schemaVersion: 1,
    suite: 'ui-responsiveness-same-corpus-capture',
    sampleCount: 3,
    limitations: [],
    cases: workloads.map((workload) => ({
      id: `same-corpus-wide-mixed-250k-${workload}`,
      corpusCaseId: 'wide-mixed-250k',
      materializedCells: 250000,
      workload,
      scenarioProof: sameCorpusScenarioProof(workload),
      bilig: {
        product: 'bilig',
        source: 'e2e/tests/web-shell-scroll-performance.pw.ts',
        operationResponseMsSamples: [4, 5, 6],
        postOperationFrameMsSamples: [8, 9, 10],
        ...(includeScrollEventSamples && uiSameCorpusWorkloadRequiresScrollEventEvidence(workload)
          ? { scrollEventResponseMsSamples: [4, 5, 6], scrollMovementPxSamples: [720, 720, 720] }
          : {}),
        corpusVerification: corpusVerification('bilig-benchmark-state', []),
        limitations: [],
      },
      googleSheets: {
        product: 'google-sheets',
        source: 'https://docs.google.com/spreadsheets/d/example',
        operationResponseMsSamples: [100, 100, 100],
        postOperationFrameMsSamples: [14, 15, 16],
        ...(includeScrollEventSamples && uiSameCorpusWorkloadRequiresScrollEventEvidence(workload)
          ? { scrollEventResponseMsSamples: [100, 100, 100], scrollMovementPxSamples: [720, 720, 720] }
          : {}),
        corpusVerification: corpusVerification('google-sheets-xlsx-export', verifiedCells()),
        limitations: [],
      },
      microsoftExcelWeb: {
        product: 'microsoft-excel-web',
        source: 'https://view.officeapps.live.com/op/view.aspx?src=example',
        operationResponseMsSamples: [75, 75, 90],
        postOperationFrameMsSamples: [14, 15, 16],
        ...(includeScrollEventSamples && uiSameCorpusWorkloadRequiresScrollEventEvidence(workload)
          ? { scrollEventResponseMsSamples: [75, 75, 90], scrollMovementPxSamples: [720, 720, 720] }
          : {}),
        corpusVerification: corpusVerification('microsoft-excel-web-source-xlsx', verifiedCells()),
        limitations: [],
      },
    })),
  }
}

function sameCorpusScenarioProof(workload: UiResponsivenessSameCorpusWorkload) {
  return {
    biligMeanMs: 5,
    biligP95Ms: 6,
    googleMeanMs: 100,
    googleP95Ms: 100,
    microsoftExcelWebMeanMs: 80,
    microsoftExcelWebP95Ms: 90,
    meanRatio: 0.05,
    p95Ratio: 0.06,
    microsoftExcelWebMeanRatio: 0.0625,
    microsoftExcelWebP95Ratio: 0.06666666666666667,
    screenshotProof: {
      captured: true,
      requiredProducts: ['bilig', 'google-sheets'],
      artifactPaths: [
        `tmp/same-corpus-wide-mixed-250k-${workload}/bilig-sample-1.png`,
        `tmp/same-corpus-wide-mixed-250k-${workload}/google-sheets-sample-1.png`,
        `tmp/same-corpus-wide-mixed-250k-${workload}/microsoft-excel-web-sample-1.png`,
      ],
      missingProducts: [],
    },
    pixelGridProof: {
      captured: true,
      requiredProducts: ['bilig', 'google-sheets'],
      products: [
        {
          product: 'bilig',
          captured: true,
          method: 'typegpu-visible-canvas',
          viewportPixelWidth: 1440,
          viewportPixelHeight: 900,
          evidence: ['mode=typegpu-v3'],
        },
        {
          product: 'google-sheets',
          captured: true,
          method: 'google-sheets-visible-grid',
          viewportPixelWidth: 1440,
          viewportPixelHeight: 900,
          evidence: ['selector=.grid-scrollable-wrapper'],
        },
        {
          product: 'microsoft-excel-web',
          captured: true,
          method: 'excel-web-visible-grid',
          viewportPixelWidth: 1440,
          viewportPixelHeight: 900,
          evidence: ['selector=.ewr-grdcontarea-grid'],
        },
      ],
      missingProducts: [],
    },
  }
}

function corpusVerification(
  method: 'bilig-benchmark-state' | 'google-sheets-xlsx-export' | 'microsoft-excel-web-source-xlsx',
  checkedCells: readonly { address: string; expected: string; actual: string }[],
) {
  return {
    verified: true,
    method,
    sheetName: 'WideGrid',
    materializedCells: 250000,
    checkedCells,
  }
}

function verifiedCells() {
  return [
    { address: 'A1', expected: 'metric-1', actual: 'metric-1' },
    { address: 'B1', expected: 'metric-2', actual: 'metric-2' },
    { address: 'F2', expected: 'note-1-5', actual: 'note-1-5' },
  ]
}
