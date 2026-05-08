import { mkdirSync, mkdtempSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'

import { hasUiResponsivenessSameCorpusTenXGap } from '../bilig-dominance-completion-audit.ts'
import {
  buildSameCorpusProof,
  assertUiResponsivenessLiveBrowserRunAllowed,
  parseUiResponsivenessLiveBrowserScorecard,
  validateUiResponsivenessLiveBrowserScorecard,
  type SameCorpusCapture,
  type UiResponsivenessLiveBrowserScorecard,
} from '../gen-ui-responsiveness-live-browser-scorecard.ts'
import { readJsonObject } from '../json-scorecard-helpers.ts'

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
      captured: false,
      evidenceKind: 'not-captured',
      requiredProductCount: 3,
      requiredCaseCount: 0,
      tenXMeanAndP95CaseCount: 0,
      coveredCorpusCaseIds: [],
    })
    expect(scorecard.sameCorpusProof.cases).toHaveLength(0)
    validateUiResponsivenessLiveBrowserScorecard(scorecard)
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

  it('derives same-corpus 10x ratios from captured scroll-event samples', () => {
    const proof = buildSameCorpusProof(buildSameCorpusCapture({ workload: 'visible-scroll-response' }))

    expect(proof).toMatchObject({
      captured: true,
      evidenceKind: 'same-corpus-browser-capture',
      requiredProductCount: 3,
      requiredCaseCount: 1,
      tenXMeanAndP95CaseCount: 1,
      coveredCorpusCaseIds: ['wide-mixed-250k'],
    })
    expect(proof.cases[0]).toMatchObject({
      biligToGoogleSheetsMeanRatio: 0.05,
      biligToGoogleSheetsP95Ratio: 0.06,
      biligToMicrosoftExcelWebMeanRatio: 0.0625,
      biligToMicrosoftExcelWebP95Ratio: 0.06666666666666667,
      biligToGoogleSheetsScrollEventMeanRatio: 0.05,
      biligToGoogleSheetsScrollEventP95Ratio: 0.06,
      biligToMicrosoftExcelWebScrollEventMeanRatio: 0.0625,
      biligToMicrosoftExcelWebScrollEventP95Ratio: 0.06666666666666667,
      tenXMeanAndP95Metric: 'scrollEventResponseMs',
      postOperationFrameGuardrailPassed: true,
      scrollMovementGuardrailPassed: true,
      passed: true,
    })
  })

  it('rejects legacy operation-only same-corpus captures before generating proof', () => {
    expect(() =>
      buildSameCorpusProof(buildSameCorpusCapture({ includeScrollEventSamples: false, workload: 'visible-scroll-response' })),
    ).toThrow(
      'UI responsiveness same-corpus capture has too few scroll-event samples for same-corpus-wide-mixed-250k-visible-scroll-response',
    )
  })

  it('rejects captured same-corpus proof without scroll-event evidence', () => {
    const scorecard = parseUiResponsivenessLiveBrowserScorecard(
      readJsonObject(resolve(repoRoot, 'packages/benchmarks/baselines/ui-responsiveness-live-browser-scorecard.json')),
    )
    const proof = buildSameCorpusProof(buildSameCorpusCapture({ workload: 'visible-scroll-response' }))
    const {
      scrollEventResponseMs: _biligScrollEventResponseMs,
      scrollMovementPx: _biligScrollMovementPx,
      ...biligWithoutScrollEvidence
    } = proof.cases[0].bilig
    const {
      scrollEventResponseMs: _googleSheetsScrollEventResponseMs,
      scrollMovementPx: _googleSheetsScrollMovementPx,
      ...googleSheetsWithoutScrollEvidence
    } = proof.cases[0].googleSheets
    const {
      scrollEventResponseMs: _microsoftExcelWebScrollEventResponseMs,
      scrollMovementPx: _microsoftExcelWebScrollMovementPx,
      ...microsoftExcelWebWithoutScrollEvidence
    } = proof.cases[0].microsoftExcelWeb
    const staleScorecard: UiResponsivenessLiveBrowserScorecard = {
      ...scorecard,
      sameCorpusProof: {
        ...proof,
        cases: [
          {
            ...proof.cases[0],
            bilig: biligWithoutScrollEvidence,
            googleSheets: googleSheetsWithoutScrollEvidence,
            microsoftExcelWeb: microsoftExcelWebWithoutScrollEvidence,
          },
        ],
      },
    }

    expect(() => validateUiResponsivenessLiveBrowserScorecard(staleScorecard)).toThrow(
      'UI responsiveness same-corpus proof is missing scroll-event evidence for same-corpus-wide-mixed-250k-visible-scroll-response',
    )
  })

  it('allows same-corpus proof to clear the public-browser limitation blocker', () => {
    const scorecard = parseUiResponsivenessLiveBrowserScorecard(
      readJsonObject(resolve(repoRoot, 'packages/benchmarks/baselines/ui-responsiveness-live-browser-scorecard.json')),
    )

    expect(
      hasUiResponsivenessSameCorpusTenXGap({
        ...scorecard,
        sameCorpusProof: buildSameCorpusProof(buildSameCorpusCapture({ workload: 'visible-scroll-response' })),
      }),
    ).toBe(false)
  })

  it('keeps the same-corpus blocker when required scroll evidence is missing', () => {
    const scorecard = parseUiResponsivenessLiveBrowserScorecard(
      readJsonObject(resolve(repoRoot, 'packages/benchmarks/baselines/ui-responsiveness-live-browser-scorecard.json')),
    )

    expect(() => buildSameCorpusProof(buildSameCorpusCapture({ workload: 'visible-edit-commit' }))).toThrow(
      'UI responsiveness same-corpus proof is missing required workload: visible-scroll-response',
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
    const proof = buildSameCorpusProof(buildSameCorpusCapture({ workload: 'visible-scroll-response' }))
    const staleScorecard: UiResponsivenessLiveBrowserScorecard = {
      ...scorecard,
      sameCorpusProof: {
        ...proof,
        cases: [
          {
            ...proof.cases[0],
            biligToGoogleSheetsP95Ratio: 0.2,
          },
        ],
      },
    }

    expect(() => validateUiResponsivenessLiveBrowserScorecard(staleScorecard)).toThrow('UI responsiveness same-corpus ratio is stale')
  })
})

function buildSameCorpusCapture(args: {
  readonly includeScrollEventSamples?: boolean
  readonly workload: 'visible-scroll-response' | 'visible-edit-commit'
}): SameCorpusCapture {
  const includeScrollEventSamples = args.includeScrollEventSamples ?? true
  return {
    schemaVersion: 1,
    suite: 'ui-responsiveness-same-corpus-capture',
    sampleCount: 3,
    limitations: [],
    cases: [
      {
        id: `same-corpus-wide-mixed-250k-${args.workload}`,
        corpusCaseId: 'wide-mixed-250k',
        materializedCells: 250000,
        workload: args.workload,
        bilig: {
          product: 'bilig',
          source: 'e2e/tests/web-shell-scroll-performance.pw.ts',
          operationResponseMsSamples: [4, 5, 6],
          postOperationFrameMsSamples: [8, 9, 10],
          ...(includeScrollEventSamples ? { scrollEventResponseMsSamples: [4, 5, 6], scrollMovementPxSamples: [720, 720, 720] } : {}),
          corpusVerification: corpusVerification('bilig-benchmark-state', []),
          limitations: [],
        },
        googleSheets: {
          product: 'google-sheets',
          source: 'https://docs.google.com/spreadsheets/d/example',
          operationResponseMsSamples: [100, 100, 100],
          postOperationFrameMsSamples: [14, 15, 16],
          ...(includeScrollEventSamples ? { scrollEventResponseMsSamples: [100, 100, 100], scrollMovementPxSamples: [720, 720, 720] } : {}),
          corpusVerification: corpusVerification('google-sheets-xlsx-export', verifiedCells()),
          limitations: [],
        },
        microsoftExcelWeb: {
          product: 'microsoft-excel-web',
          source: 'https://view.officeapps.live.com/op/view.aspx?src=example',
          operationResponseMsSamples: [75, 75, 90],
          postOperationFrameMsSamples: [14, 15, 16],
          ...(includeScrollEventSamples ? { scrollEventResponseMsSamples: [75, 75, 90], scrollMovementPxSamples: [720, 720, 720] } : {}),
          corpusVerification: corpusVerification('microsoft-excel-web-source-xlsx', verifiedCells()),
          limitations: [],
        },
      },
    ],
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
