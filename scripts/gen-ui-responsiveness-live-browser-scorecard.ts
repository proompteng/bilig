#!/usr/bin/env bun

import { existsSync, mkdirSync, writeFileSync } from 'node:fs'
import { dirname, join, resolve } from 'node:path'
import { performance } from 'node:perf_hooks'

import { chromium, type Browser, type Page } from '@playwright/test'
import { summarizeNumbers, type NumericSummary } from '../packages/benchmarks/src/stats.js'
import { assertLocalCiResourceGuardAllowsRun } from './ci-local-resource-guard.ts'
import { readJsonObject } from './json-scorecard-helpers.ts'
import { parseSameCorpusCapture, parseUiResponsivenessLiveBrowserScorecard } from './ui-responsiveness-live-browser-scorecard-parse.ts'

export { parseSameCorpusCapture, parseUiResponsivenessLiveBrowserScorecard } from './ui-responsiveness-live-browser-scorecard-parse.ts'

export type UiResponsivenessLiveBrowserVendor = 'google-sheets' | 'microsoft-excel-web'
export type UiResponsivenessSameCorpusProduct = 'bilig' | 'google-sheets' | 'microsoft-excel-web'
export type UiResponsivenessSameCorpusWorkload = 'visible-scroll-response' | 'visible-edit-commit'

export interface UiResponsivenessLiveBrowserCase {
  readonly id: string
  readonly vendor: UiResponsivenessLiveBrowserVendor
  readonly product: string
  readonly sourceUrl: string
  readonly finalUrl: string
  readonly title: string
  readonly accessMode: 'public-comment-only' | 'public-view-only' | 'public-office-web-viewer'
  readonly workload: 'open-public-workbook-and-scroll-viewport'
  readonly sampleCount: number
  readonly loadToReadyMs: NumericSummary
  readonly scrollResponseMs: NumericSummary
  readonly postScrollFrameMs: NumericSummary
  readonly passed: boolean
  readonly limitations: string[]
}

export interface UiResponsivenessSameCorpusMeasurement {
  readonly product: UiResponsivenessSameCorpusProduct
  readonly source: string
  readonly operationResponseMs: NumericSummary
  readonly postOperationFrameMs: NumericSummary
  readonly scrollEventResponseMs?: NumericSummary
  readonly scrollMovementPx?: NumericSummary
  readonly corpusVerification: SameCorpusCaptureCorpusVerification
  readonly limitations: string[]
}

export interface UiResponsivenessSameCorpusCase {
  readonly id: string
  readonly corpusCaseId: string
  readonly materializedCells: number
  readonly workload: UiResponsivenessSameCorpusWorkload
  readonly sampleCount: number
  readonly bilig: UiResponsivenessSameCorpusMeasurement
  readonly googleSheets: UiResponsivenessSameCorpusMeasurement
  readonly microsoftExcelWeb: UiResponsivenessSameCorpusMeasurement
  readonly biligToGoogleSheetsMeanRatio: number
  readonly biligToGoogleSheetsP95Ratio: number
  readonly biligToMicrosoftExcelWebMeanRatio: number
  readonly biligToMicrosoftExcelWebP95Ratio: number
  readonly biligToGoogleSheetsScrollEventMeanRatio?: number
  readonly biligToGoogleSheetsScrollEventP95Ratio?: number
  readonly biligToMicrosoftExcelWebScrollEventMeanRatio?: number
  readonly biligToMicrosoftExcelWebScrollEventP95Ratio?: number
  readonly tenXMeanAndP95Metric?: 'operationResponseMs' | 'scrollEventResponseMs'
  readonly tenXMeanAndP95AgainstGoogleSheets: boolean
  readonly tenXMeanAndP95AgainstMicrosoftExcelWeb: boolean
  readonly postOperationFrameGuardrailPassed?: boolean
  readonly scrollMovementGuardrailPassed?: boolean
  readonly passed: boolean
}

export interface UiResponsivenessSameCorpusProof {
  readonly captured: boolean
  readonly evidenceKind: 'same-corpus-browser-capture' | 'not-captured'
  readonly requiredProductCount: number
  readonly requiredCaseCount: number
  readonly tenXMeanAndP95CaseCount: number
  readonly coveredCorpusCaseIds: string[]
  readonly limitations: string[]
  readonly cases: UiResponsivenessSameCorpusCase[]
}

export interface UiResponsivenessLiveBrowserScorecard {
  readonly schemaVersion: 1
  readonly suite: 'ui-responsiveness-live-browser-timing'
  readonly generatedAt: string
  readonly host: {
    readonly arch: string
    readonly platform: string
  }
  readonly source: {
    readonly artifactGenerator: 'scripts/gen-ui-responsiveness-live-browser-scorecard.ts'
    readonly evidenceKind: 'live-public-browser-playwright'
    readonly browserEngine: 'chromium'
    readonly measuredOperation: 'public-workbook-load-and-viewport-scroll'
  }
  readonly benchmark: {
    readonly sampleCount: number
    readonly viewport: {
      readonly width: number
      readonly height: number
    }
    readonly samplingOrder: 'google-sheets-then-microsoft-excel-web'
  }
  readonly summary: {
    readonly directBrowserTimingCaptured: boolean
    readonly allRequiredCasesPassed: boolean
    readonly requiredVendorCount: number
    readonly capturedVendors: UiResponsivenessLiveBrowserVendor[]
    readonly limitations: string[]
  }
  readonly cases: UiResponsivenessLiveBrowserCase[]
  readonly sameCorpusProof: UiResponsivenessSameCorpusProof
}

interface BrowserCaseSpec {
  readonly id: string
  readonly vendor: UiResponsivenessLiveBrowserVendor
  readonly product: string
  readonly sourceUrl: string
  readonly expectedTitleIncludes: string
}

interface BrowserCaseSample {
  readonly finalUrl: string
  readonly title: string
  readonly accessMode: UiResponsivenessLiveBrowserCase['accessMode']
  readonly loadToReadyMs: number
  readonly scrollResponseMs: number
  readonly postScrollFrameMs: number
}

export interface SameCorpusCapture {
  readonly schemaVersion: 1
  readonly suite: 'ui-responsiveness-same-corpus-capture'
  readonly sampleCount: number
  readonly limitations: string[]
  readonly cases: SameCorpusCaptureCase[]
}

export interface SameCorpusCaptureCase {
  readonly id: string
  readonly corpusCaseId: string
  readonly materializedCells: number
  readonly workload: UiResponsivenessSameCorpusWorkload
  readonly bilig: SameCorpusCaptureMeasurement
  readonly googleSheets: SameCorpusCaptureMeasurement
  readonly microsoftExcelWeb: SameCorpusCaptureMeasurement
}

export interface SameCorpusCaptureMeasurement {
  readonly product: UiResponsivenessSameCorpusProduct
  readonly source: string
  readonly operationResponseMsSamples: number[]
  readonly postOperationFrameMsSamples: number[]
  readonly scrollEventResponseMsSamples?: number[]
  readonly scrollMovementPxSamples?: number[]
  readonly corpusVerification: SameCorpusCaptureCorpusVerification
  readonly limitations: string[]
}

export interface SameCorpusCaptureVerifiedCell {
  readonly address: string
  readonly expected: string
  readonly actual: string
}

export interface SameCorpusCaptureCorpusVerification {
  readonly verified: boolean
  readonly method: 'bilig-benchmark-state' | 'google-sheets-xlsx-export' | 'microsoft-excel-web-source-xlsx'
  readonly sheetName: string
  readonly materializedCells: number
  readonly checkedCells: readonly SameCorpusCaptureVerifiedCell[]
}

const rootDir = resolve(new URL('..', import.meta.url).pathname)
const outputPath = join(rootDir, 'packages', 'benchmarks', 'baselines', 'ui-responsiveness-live-browser-scorecard.json')
const sampleCount = 3
const requiredSameCorpusWorkloads = ['visible-scroll-response'] as const satisfies readonly UiResponsivenessSameCorpusWorkload[]
const viewport = { width: 1440, height: 900 } as const
const microsoftExcelSourceWorkbook =
  'https://github.com/fileformat-blog-gists/SampleFiles/raw/main/Spreadsheet-File-Formats/XLSX/Pivot-Tables-and-Charts.xlsx'
const caseSpecs = [
  {
    id: 'google-sheets-public-grid-scroll',
    vendor: 'google-sheets',
    product: 'Google Sheets public spreadsheet',
    sourceUrl: 'https://docs.google.com/spreadsheets/d/1Awcx961Qm_cJw7X-7hsKEGCzS-0yMw0TqwbniaxNkeU/edit',
    expectedTitleIncludes: 'Google Sheets',
  },
  {
    id: 'microsoft-excel-web-public-xlsx-scroll',
    vendor: 'microsoft-excel-web',
    product: 'Microsoft Office Web Viewer public XLSX',
    sourceUrl: `https://view.officeapps.live.com/op/view.aspx?src=${encodeURIComponent(microsoftExcelSourceWorkbook)}`,
    expectedTitleIncludes: '.xlsx',
  },
] as const satisfies readonly BrowserCaseSpec[]

async function main(): Promise<void> {
  const isCheckMode = process.argv.includes('--check')
  const capturePath = argumentValue('--capture')
  if (isCheckMode) {
    if (!existsSync(outputPath)) {
      throw new Error(
        `UI responsiveness live browser scorecard is missing. Run: bun scripts/gen-ui-responsiveness-live-browser-scorecard.ts`,
      )
    }
    const scorecard = parseUiResponsivenessLiveBrowserScorecard(readJsonObject(outputPath))
    validateUiResponsivenessLiveBrowserScorecard(scorecard)
    logResult('check', scorecard)
    return
  }

  assertUiResponsivenessLiveBrowserRunAllowed()
  const sameCorpusProof = capturePath
    ? buildSameCorpusProof(parseSameCorpusCapture(readJsonObject(resolve(capturePath))))
    : buildMissingSameCorpusProof()
  const scorecard = await buildUiResponsivenessLiveBrowserScorecard(new Date().toISOString(), sameCorpusProof)
  validateUiResponsivenessLiveBrowserScorecard(scorecard)
  mkdirSync(dirname(outputPath), { recursive: true })
  writeFileSync(outputPath, formatJsonForRepo(`${JSON.stringify(scorecard, null, 2)}\n`))
  logResult('write', scorecard)
}

export async function buildUiResponsivenessLiveBrowserScorecard(
  generatedAt: string,
  sameCorpusProof = buildMissingSameCorpusProof(),
): Promise<UiResponsivenessLiveBrowserScorecard> {
  const browser = await chromium.launch({ headless: true })
  try {
    const cases = await measureBrowserCases(browser)

    return {
      schemaVersion: 1,
      suite: 'ui-responsiveness-live-browser-timing',
      generatedAt,
      host: {
        arch: process.arch,
        platform: process.platform,
      },
      source: {
        artifactGenerator: 'scripts/gen-ui-responsiveness-live-browser-scorecard.ts',
        evidenceKind: 'live-public-browser-playwright',
        browserEngine: 'chromium',
        measuredOperation: 'public-workbook-load-and-viewport-scroll',
      },
      benchmark: {
        sampleCount,
        viewport,
        samplingOrder: 'google-sheets-then-microsoft-excel-web',
      },
      summary: {
        directBrowserTimingCaptured: cases.length === caseSpecs.length,
        allRequiredCasesPassed: cases.every((entry) => entry.passed),
        requiredVendorCount: caseSpecs.length,
        capturedVendors: caseSpecs.map((entry) => entry.vendor),
        limitations: [
          'Public unauthenticated browser timing covers load and viewport scroll only; it does not cover authenticated edit latency.',
          'The incumbent workbooks are public representative workbooks, not bilig-generated benchmark corpuses.',
          'Network, tenant, CDN, and browser-cache conditions can move live public-web measurements between runs.',
        ],
      },
      cases,
      sameCorpusProof,
    }
  } finally {
    await browser.close()
  }
}

export function assertUiResponsivenessLiveBrowserRunAllowed(
  rootDirForGuard: string = rootDir,
  env: Readonly<Record<string, string | undefined>> = process.env,
): void {
  assertLocalCiResourceGuardAllowsRun(rootDirForGuard, env, { runLabel: 'UI responsiveness live browser scorecard generation' })
}

async function measureBrowserCases(
  browser: Browser,
  specIndex = 0,
  cases: UiResponsivenessLiveBrowserCase[] = [],
): Promise<UiResponsivenessLiveBrowserCase[]> {
  const spec = caseSpecs[specIndex]
  if (!spec) {
    return cases
  }
  cases.push(buildBrowserCase(spec, await measureBrowserCaseSamples(browser, spec)))
  return measureBrowserCases(browser, specIndex + 1, cases)
}

async function measureBrowserCaseSamples(
  browser: Browser,
  spec: BrowserCaseSpec,
  sampleIndex = 0,
  samples: BrowserCaseSample[] = [],
): Promise<BrowserCaseSample[]> {
  if (sampleIndex >= sampleCount) {
    return samples
  }
  const page = await browser.newPage({ viewport })
  try {
    samples.push(await measureBrowserCase(page, spec))
  } finally {
    await page.close()
  }
  return measureBrowserCaseSamples(browser, spec, sampleIndex + 1, samples)
}

export function validateUiResponsivenessLiveBrowserScorecard(scorecard: UiResponsivenessLiveBrowserScorecard): void {
  if (scorecard.benchmark.sampleCount < 3) {
    throw new Error('UI responsiveness live browser scorecard must contain at least 3 samples per case')
  }
  if (!scorecard.summary.directBrowserTimingCaptured || !scorecard.summary.allRequiredCasesPassed) {
    throw new Error('UI responsiveness live browser scorecard summary reports missing or failed browser timing evidence')
  }
  for (const spec of caseSpecs) {
    const entry = scorecard.cases.find((candidate) => candidate.id === spec.id)
    if (!entry) {
      throw new Error(`UI responsiveness live browser scorecard is missing required case: ${spec.id}`)
    }
    if (entry.vendor !== spec.vendor) {
      throw new Error(`UI responsiveness live browser scorecard vendor mismatch for case: ${spec.id}`)
    }
    if (!entry.passed) {
      throw new Error(`UI responsiveness live browser scorecard contains a failed case: ${spec.id}`)
    }
    if (entry.sampleCount < scorecard.benchmark.sampleCount) {
      throw new Error(`UI responsiveness live browser scorecard has too few samples for case: ${spec.id}`)
    }
    if (!entry.title.includes(spec.expectedTitleIncludes)) {
      throw new Error(`UI responsiveness live browser scorecard title does not match ${spec.vendor}: ${entry.title}`)
    }
    validateSummary(entry.loadToReadyMs, `${spec.id} loadToReadyMs`)
    validateSummary(entry.scrollResponseMs, `${spec.id} scrollResponseMs`)
    validateSummary(entry.postScrollFrameMs, `${spec.id} postScrollFrameMs`)
  }
  for (const vendor of caseSpecs.map((entry) => entry.vendor)) {
    if (!scorecard.summary.capturedVendors.includes(vendor)) {
      throw new Error(`UI responsiveness live browser scorecard is missing vendor: ${vendor}`)
    }
  }
  if (scorecard.summary.limitations.length === 0 || !scorecard.cases.every((entry) => entry.limitations.length > 0)) {
    throw new Error('UI responsiveness live browser scorecard must disclose benchmark limitations')
  }
  validateSameCorpusProof(scorecard.sameCorpusProof)
}

export function buildMissingSameCorpusProof(): UiResponsivenessSameCorpusProof {
  return {
    captured: false,
    evidenceKind: 'not-captured',
    requiredProductCount: 3,
    requiredCaseCount: 0,
    tenXMeanAndP95CaseCount: 0,
    coveredCorpusCaseIds: [],
    limitations: ['Same-corpus live browser timing against Bilig, Google Sheets, and Microsoft Excel Web has not been captured yet.'],
    cases: [],
  }
}

export function buildSameCorpusProof(capture: SameCorpusCapture): UiResponsivenessSameCorpusProof {
  validateSameCorpusCapture(capture)
  const cases = capture.cases.map(buildSameCorpusCase)
  const proof: UiResponsivenessSameCorpusProof = {
    captured: true,
    evidenceKind: 'same-corpus-browser-capture',
    requiredProductCount: 3,
    requiredCaseCount: capture.cases.length,
    tenXMeanAndP95CaseCount: cases.filter(
      (entry) => entry.tenXMeanAndP95AgainstGoogleSheets && entry.tenXMeanAndP95AgainstMicrosoftExcelWeb,
    ).length,
    coveredCorpusCaseIds: [...new Set(cases.map((entry) => entry.corpusCaseId))].toSorted(),
    limitations: [...capture.limitations],
    cases,
  }
  validateSameCorpusProof(proof)
  return proof
}

function validateSameCorpusCapture(capture: SameCorpusCapture): void {
  if (capture.sampleCount < sampleCountForSameCorpus()) {
    throw new Error('UI responsiveness same-corpus capture must contain at least 3 samples per product')
  }
  if (capture.cases.length === 0) {
    throw new Error('UI responsiveness same-corpus capture must include at least one case')
  }
  for (const entry of capture.cases) {
    const hasAnyScrollEventSamples = [entry.bilig, entry.googleSheets, entry.microsoftExcelWeb].some(
      (measurement) => measurement.scrollEventResponseMsSamples !== undefined || measurement.scrollMovementPxSamples !== undefined,
    )
    const requiresScrollEventSamples = entry.workload === 'visible-scroll-response' || hasAnyScrollEventSamples
    for (const measurement of [entry.bilig, entry.googleSheets, entry.microsoftExcelWeb]) {
      if (
        measurement.operationResponseMsSamples.length < capture.sampleCount ||
        measurement.postOperationFrameMsSamples.length < capture.sampleCount
      ) {
        throw new Error(`UI responsiveness same-corpus capture has too few samples for ${entry.id}`)
      }
      if (
        requiresScrollEventSamples &&
        ((measurement.scrollEventResponseMsSamples?.length ?? 0) < capture.sampleCount ||
          (measurement.scrollMovementPxSamples?.length ?? 0) < capture.sampleCount)
      ) {
        throw new Error(`UI responsiveness same-corpus capture has too few scroll-event samples for ${entry.id}`)
      }
      validateSameCorpusCaptureVerification(measurement.corpusVerification, measurement.product, entry.materializedCells, entry.id)
    }
  }
}

async function measureBrowserCase(page: Page, spec: BrowserCaseSpec): Promise<BrowserCaseSample> {
  const startedAt = performance.now()
  await page.goto(spec.sourceUrl, { waitUntil: 'domcontentloaded', timeout: 45_000 })
  await waitForCaseReady(page, spec)
  const loadToReadyMs = performance.now() - startedAt
  const title = await page.title()
  const accessMode = await detectAccessMode(page, spec.vendor)

  await page.mouse.move(viewport.width / 2, viewport.height / 2)
  const scrollStartedAt = performance.now()
  await page.mouse.wheel(0, 720)
  const frameIntervals = await page.evaluate(async () => {
    const intervals: number[] = []
    let previous = performance.now()
    await new Promise<void>((finish) => {
      const step = (now: number): void => {
        intervals.push(now - previous)
        previous = now
        if (intervals.length >= 12) {
          finish()
          return
        }
        requestAnimationFrame(step)
      }
      requestAnimationFrame(step)
    })
    return intervals
  })
  const scrollResponseMs = performance.now() - scrollStartedAt
  return {
    finalUrl: page.url(),
    title,
    accessMode,
    loadToReadyMs,
    scrollResponseMs,
    postScrollFrameMs: summarizeNumbers(frameIntervals).p95,
  }
}

async function waitForCaseReady(page: Page, spec: BrowserCaseSpec): Promise<void> {
  if (spec.vendor === 'google-sheets') {
    await page.waitForFunction(
      () =>
        !window.location.href.includes('accounts.google.com') &&
        document.title.includes('Google Sheets') &&
        (document.body.innerText.includes('Comment only') || document.body.innerText.includes('View only')),
      { timeout: 45_000 },
    )
    await page.waitForTimeout(2_000)
    return
  }

  await page.waitForFunction(() => document.title.endsWith('.xlsx'), { timeout: 45_000 })
  await page.waitForTimeout(5_000)
}

async function detectAccessMode(
  page: Page,
  vendor: UiResponsivenessLiveBrowserVendor,
): Promise<UiResponsivenessLiveBrowserCase['accessMode']> {
  if (vendor === 'microsoft-excel-web') {
    return 'public-office-web-viewer'
  }
  const bodyText = await page.locator('body').innerText({ timeout: 5_000 })
  return bodyText.includes('Comment only') ? 'public-comment-only' : 'public-view-only'
}

function buildBrowserCase(spec: BrowserCaseSpec, samples: readonly BrowserCaseSample[]): UiResponsivenessLiveBrowserCase {
  const title = samples[0]?.title ?? ''
  const finalUrl = samples[0]?.finalUrl ?? ''
  const accessMode = samples[0]?.accessMode ?? 'public-view-only'
  const loadToReadyMs = summarizeNumbers(samples.map((entry) => entry.loadToReadyMs))
  const scrollResponseMs = summarizeNumbers(samples.map((entry) => entry.scrollResponseMs))
  const postScrollFrameMs = summarizeNumbers(samples.map((entry) => entry.postScrollFrameMs))
  return {
    id: spec.id,
    vendor: spec.vendor,
    product: spec.product,
    sourceUrl: spec.sourceUrl,
    finalUrl,
    title,
    accessMode,
    workload: 'open-public-workbook-and-scroll-viewport',
    sampleCount: samples.length,
    loadToReadyMs,
    scrollResponseMs,
    postScrollFrameMs,
    passed:
      samples.length === sampleCount &&
      title.includes(spec.expectedTitleIncludes) &&
      samples.every(
        (entry) => Number.isFinite(entry.loadToReadyMs) && Number.isFinite(entry.scrollResponseMs) && entry.postScrollFrameMs > 0,
      ),
    limitations: [
      'Public browser timing cannot exercise authenticated workbook editing or tenant-local collaboration paths.',
      'This timing is direct incumbent browser evidence, but it is not a same-corpus 10x proof by itself.',
    ],
  }
}

function buildSameCorpusCase(captureCase: SameCorpusCaptureCase): UiResponsivenessSameCorpusCase {
  const bilig = buildSameCorpusMeasurement(captureCase.bilig)
  const googleSheets = buildSameCorpusMeasurement(captureCase.googleSheets)
  const microsoftExcelWeb = buildSameCorpusMeasurement(captureCase.microsoftExcelWeb)
  const biligToGoogleSheetsMeanRatio = ratio(bilig.operationResponseMs.mean, googleSheets.operationResponseMs.mean)
  const biligToGoogleSheetsP95Ratio = ratio(bilig.operationResponseMs.p95, googleSheets.operationResponseMs.p95)
  const biligToMicrosoftExcelWebMeanRatio = ratio(bilig.operationResponseMs.mean, microsoftExcelWeb.operationResponseMs.mean)
  const biligToMicrosoftExcelWebP95Ratio = ratio(bilig.operationResponseMs.p95, microsoftExcelWeb.operationResponseMs.p95)
  const scrollEventMetrics = sameCorpusScrollEventMetrics(bilig, googleSheets, microsoftExcelWeb)
  const postOperationFrameGuardrailPassed = [bilig, googleSheets, microsoftExcelWeb].every(
    (entry) => entry.postOperationFrameMs.p95 > 0 && entry.postOperationFrameMs.p95 <= 50,
  )
  const scrollMovementGuardrailPassed =
    scrollEventMetrics !== null && [bilig, googleSheets, microsoftExcelWeb].every((entry) => (entry.scrollMovementPx?.min ?? 0) >= 1)
  const tenXMeanAndP95AgainstGoogleSheets =
    scrollEventMetrics !== null &&
    scrollEventMetrics.biligToGoogleSheetsMeanRatio <= 0.1 &&
    scrollEventMetrics.biligToGoogleSheetsP95Ratio <= 0.1 &&
    postOperationFrameGuardrailPassed &&
    scrollMovementGuardrailPassed
  const tenXMeanAndP95AgainstMicrosoftExcelWeb =
    scrollEventMetrics !== null &&
    scrollEventMetrics.biligToMicrosoftExcelWebMeanRatio <= 0.1 &&
    scrollEventMetrics.biligToMicrosoftExcelWebP95Ratio <= 0.1 &&
    postOperationFrameGuardrailPassed &&
    scrollMovementGuardrailPassed
  return {
    id: captureCase.id,
    corpusCaseId: captureCase.corpusCaseId,
    materializedCells: captureCase.materializedCells,
    workload: captureCase.workload,
    sampleCount: Math.min(
      bilig.operationResponseMs.samples.length,
      googleSheets.operationResponseMs.samples.length,
      microsoftExcelWeb.operationResponseMs.samples.length,
    ),
    bilig,
    googleSheets,
    microsoftExcelWeb,
    biligToGoogleSheetsMeanRatio,
    biligToGoogleSheetsP95Ratio,
    biligToMicrosoftExcelWebMeanRatio,
    biligToMicrosoftExcelWebP95Ratio,
    ...(scrollEventMetrics
      ? {
          biligToGoogleSheetsScrollEventMeanRatio: scrollEventMetrics.biligToGoogleSheetsMeanRatio,
          biligToGoogleSheetsScrollEventP95Ratio: scrollEventMetrics.biligToGoogleSheetsP95Ratio,
          biligToMicrosoftExcelWebScrollEventMeanRatio: scrollEventMetrics.biligToMicrosoftExcelWebMeanRatio,
          biligToMicrosoftExcelWebScrollEventP95Ratio: scrollEventMetrics.biligToMicrosoftExcelWebP95Ratio,
          tenXMeanAndP95Metric: 'scrollEventResponseMs' as const,
          postOperationFrameGuardrailPassed,
          scrollMovementGuardrailPassed,
        }
      : { tenXMeanAndP95Metric: 'operationResponseMs' as const }),
    tenXMeanAndP95AgainstGoogleSheets,
    tenXMeanAndP95AgainstMicrosoftExcelWeb,
    passed: tenXMeanAndP95AgainstGoogleSheets && tenXMeanAndP95AgainstMicrosoftExcelWeb,
  }
}

function buildSameCorpusMeasurement(capture: SameCorpusCaptureMeasurement): UiResponsivenessSameCorpusMeasurement {
  return {
    product: capture.product,
    source: capture.source,
    operationResponseMs: summarizeNumbers(capture.operationResponseMsSamples),
    postOperationFrameMs: summarizeNumbers(capture.postOperationFrameMsSamples),
    ...(capture.scrollEventResponseMsSamples ? { scrollEventResponseMs: summarizeNumbers(capture.scrollEventResponseMsSamples) } : {}),
    ...(capture.scrollMovementPxSamples ? { scrollMovementPx: summarizeNumbers(capture.scrollMovementPxSamples) } : {}),
    corpusVerification: cloneSameCorpusVerification(capture.corpusVerification),
    limitations: [...capture.limitations],
  }
}

function sameCorpusScrollEventMetrics(
  bilig: UiResponsivenessSameCorpusMeasurement,
  googleSheets: UiResponsivenessSameCorpusMeasurement,
  microsoftExcelWeb: UiResponsivenessSameCorpusMeasurement,
): {
  readonly biligToGoogleSheetsMeanRatio: number
  readonly biligToGoogleSheetsP95Ratio: number
  readonly biligToMicrosoftExcelWebMeanRatio: number
  readonly biligToMicrosoftExcelWebP95Ratio: number
} | null {
  if (!bilig.scrollEventResponseMs || !googleSheets.scrollEventResponseMs || !microsoftExcelWeb.scrollEventResponseMs) {
    return null
  }
  return {
    biligToGoogleSheetsMeanRatio: ratio(bilig.scrollEventResponseMs.mean, googleSheets.scrollEventResponseMs.mean),
    biligToGoogleSheetsP95Ratio: ratio(bilig.scrollEventResponseMs.p95, googleSheets.scrollEventResponseMs.p95),
    biligToMicrosoftExcelWebMeanRatio: ratio(bilig.scrollEventResponseMs.mean, microsoftExcelWeb.scrollEventResponseMs.mean),
    biligToMicrosoftExcelWebP95Ratio: ratio(bilig.scrollEventResponseMs.p95, microsoftExcelWeb.scrollEventResponseMs.p95),
  }
}

function ratio(numerator: number, denominator: number): number {
  if (denominator <= 0) {
    return Number.POSITIVE_INFINITY
  }
  return numerator / denominator
}

function validateSameCorpusProof(proof: UiResponsivenessSameCorpusProof): void {
  if (proof.requiredProductCount !== 3) {
    throw new Error('UI responsiveness same-corpus proof must compare Bilig, Google Sheets, and Microsoft Excel Web')
  }
  if (!proof.captured) {
    if (proof.evidenceKind !== 'not-captured' || proof.cases.length !== 0 || proof.requiredCaseCount !== 0) {
      throw new Error('UI responsiveness same-corpus proof has stale not-captured metadata')
    }
    if (proof.limitations.length === 0) {
      throw new Error('UI responsiveness same-corpus proof must disclose that capture is missing')
    }
    return
  }
  if (proof.evidenceKind !== 'same-corpus-browser-capture') {
    throw new Error('UI responsiveness same-corpus proof has stale capture metadata')
  }
  if (proof.requiredCaseCount === 0 || proof.cases.length !== proof.requiredCaseCount) {
    throw new Error('UI responsiveness same-corpus proof must include every required captured case')
  }
  for (const workload of requiredSameCorpusWorkloads) {
    if (!proof.cases.some((entry) => entry.workload === workload)) {
      throw new Error(`UI responsiveness same-corpus proof is missing required workload: ${workload}`)
    }
  }
  const tenXCaseCount = proof.cases.filter(
    (entry) => entry.tenXMeanAndP95AgainstGoogleSheets && entry.tenXMeanAndP95AgainstMicrosoftExcelWeb,
  ).length
  if (proof.tenXMeanAndP95CaseCount !== tenXCaseCount) {
    throw new Error('UI responsiveness same-corpus proof 10x case count is stale')
  }
  const coveredCorpusCaseIds = [...new Set(proof.cases.map((entry) => entry.corpusCaseId))].toSorted()
  if (JSON.stringify(proof.coveredCorpusCaseIds) !== JSON.stringify(coveredCorpusCaseIds)) {
    throw new Error('UI responsiveness same-corpus proof covered corpus IDs are stale')
  }
  for (const entry of proof.cases) {
    validateSameCorpusCase(entry)
  }
}

function validateSameCorpusCase(entry: UiResponsivenessSameCorpusCase): void {
  if (entry.materializedCells <= 0 || !Number.isInteger(entry.materializedCells)) {
    throw new Error(`UI responsiveness same-corpus case has invalid materialized cell count: ${entry.id}`)
  }
  validateSameCorpusMeasurement(entry.bilig, 'bilig', entry.id)
  validateSameCorpusMeasurement(entry.googleSheets, 'google-sheets', entry.id)
  validateSameCorpusMeasurement(entry.microsoftExcelWeb, 'microsoft-excel-web', entry.id)
  const comparableSampleCount = Math.min(
    entry.bilig.operationResponseMs.samples.length,
    entry.googleSheets.operationResponseMs.samples.length,
    entry.microsoftExcelWeb.operationResponseMs.samples.length,
  )
  if (entry.sampleCount !== comparableSampleCount || comparableSampleCount < sampleCountForSameCorpus()) {
    throw new Error(`UI responsiveness same-corpus case has too few comparable samples: ${entry.id}`)
  }
  const googleSheetsMeanRatio = ratio(entry.bilig.operationResponseMs.mean, entry.googleSheets.operationResponseMs.mean)
  const googleSheetsP95Ratio = ratio(entry.bilig.operationResponseMs.p95, entry.googleSheets.operationResponseMs.p95)
  const microsoftExcelWebMeanRatio = ratio(entry.bilig.operationResponseMs.mean, entry.microsoftExcelWeb.operationResponseMs.mean)
  const microsoftExcelWebP95Ratio = ratio(entry.bilig.operationResponseMs.p95, entry.microsoftExcelWeb.operationResponseMs.p95)
  if (
    entry.biligToGoogleSheetsMeanRatio !== googleSheetsMeanRatio ||
    entry.biligToGoogleSheetsP95Ratio !== googleSheetsP95Ratio ||
    entry.biligToMicrosoftExcelWebMeanRatio !== microsoftExcelWebMeanRatio ||
    entry.biligToMicrosoftExcelWebP95Ratio !== microsoftExcelWebP95Ratio
  ) {
    throw new Error(`UI responsiveness same-corpus ratio is stale: ${entry.id}`)
  }
  const scrollEventMetrics = sameCorpusScrollEventMetrics(entry.bilig, entry.googleSheets, entry.microsoftExcelWeb)
  const postOperationFrameGuardrailPassed = [entry.bilig, entry.googleSheets, entry.microsoftExcelWeb].every(
    (measurement) => measurement.postOperationFrameMs.p95 > 0 && measurement.postOperationFrameMs.p95 <= 50,
  )
  const scrollMovementGuardrailPassed =
    scrollEventMetrics !== null &&
    [entry.bilig, entry.googleSheets, entry.microsoftExcelWeb].every((measurement) => (measurement.scrollMovementPx?.min ?? 0) >= 1)
  if (scrollEventMetrics) {
    if (
      entry.biligToGoogleSheetsScrollEventMeanRatio !== scrollEventMetrics.biligToGoogleSheetsMeanRatio ||
      entry.biligToGoogleSheetsScrollEventP95Ratio !== scrollEventMetrics.biligToGoogleSheetsP95Ratio ||
      entry.biligToMicrosoftExcelWebScrollEventMeanRatio !== scrollEventMetrics.biligToMicrosoftExcelWebMeanRatio ||
      entry.biligToMicrosoftExcelWebScrollEventP95Ratio !== scrollEventMetrics.biligToMicrosoftExcelWebP95Ratio
    ) {
      throw new Error(`UI responsiveness same-corpus scroll-event ratio is stale: ${entry.id}`)
    }
  }
  if (
    entry.postOperationFrameGuardrailPassed !== undefined &&
    entry.postOperationFrameGuardrailPassed !== postOperationFrameGuardrailPassed
  ) {
    throw new Error(`UI responsiveness same-corpus post-frame guardrail is stale: ${entry.id}`)
  }
  if (entry.scrollMovementGuardrailPassed !== undefined && entry.scrollMovementGuardrailPassed !== scrollMovementGuardrailPassed) {
    throw new Error(`UI responsiveness same-corpus scroll-movement guardrail is stale: ${entry.id}`)
  }
  const usesScrollEventMetric = entry.tenXMeanAndP95Metric === 'scrollEventResponseMs'
  const tenXAgainstGoogleSheets =
    usesScrollEventMetric &&
    scrollEventMetrics !== null &&
    scrollEventMetrics.biligToGoogleSheetsMeanRatio <= 0.1 &&
    scrollEventMetrics.biligToGoogleSheetsP95Ratio <= 0.1 &&
    postOperationFrameGuardrailPassed &&
    scrollMovementGuardrailPassed
  const tenXAgainstMicrosoftExcelWeb =
    usesScrollEventMetric &&
    scrollEventMetrics !== null &&
    scrollEventMetrics.biligToMicrosoftExcelWebMeanRatio <= 0.1 &&
    scrollEventMetrics.biligToMicrosoftExcelWebP95Ratio <= 0.1 &&
    postOperationFrameGuardrailPassed &&
    scrollMovementGuardrailPassed
  if (
    entry.tenXMeanAndP95AgainstGoogleSheets !== tenXAgainstGoogleSheets ||
    entry.tenXMeanAndP95AgainstMicrosoftExcelWeb !== tenXAgainstMicrosoftExcelWeb ||
    entry.passed !== (tenXAgainstGoogleSheets && tenXAgainstMicrosoftExcelWeb)
  ) {
    throw new Error(`UI responsiveness same-corpus pass flag is stale: ${entry.id}`)
  }
}

function validateSameCorpusMeasurement(
  measurement: UiResponsivenessSameCorpusMeasurement,
  product: UiResponsivenessSameCorpusProduct,
  caseId: string,
): void {
  if (measurement.product !== product) {
    throw new Error(`UI responsiveness same-corpus product mismatch for ${caseId}`)
  }
  if (measurement.source.length === 0) {
    throw new Error(`UI responsiveness same-corpus source is missing for ${caseId}`)
  }
  validateSummary(measurement.operationResponseMs, `${caseId} ${product} operationResponseMs`)
  validateSummary(measurement.postOperationFrameMs, `${caseId} ${product} postOperationFrameMs`)
  if (measurement.scrollEventResponseMs) {
    validateSummary(measurement.scrollEventResponseMs, `${caseId} ${product} scrollEventResponseMs`)
  }
  if (measurement.scrollMovementPx) {
    validateSummary(measurement.scrollMovementPx, `${caseId} ${product} scrollMovementPx`)
  }
  validateSameCorpusCaptureVerification(measurement.corpusVerification, product, null, caseId)
}

function validateSameCorpusCaptureVerification(
  verification: SameCorpusCaptureCorpusVerification,
  product: UiResponsivenessSameCorpusProduct,
  expectedMaterializedCells: number | null,
  caseId: string,
): void {
  if (!verification.verified) {
    throw new Error(`UI responsiveness same-corpus verification is not marked verified for ${caseId} ${product}`)
  }
  if (expectedMaterializedCells !== null && verification.materializedCells !== expectedMaterializedCells) {
    throw new Error(`UI responsiveness same-corpus verification materialized cell count mismatch for ${caseId} ${product}`)
  }
  if (product === 'bilig' && verification.method !== 'bilig-benchmark-state') {
    throw new Error(`UI responsiveness same-corpus verification method mismatch for ${caseId} ${product}`)
  }
  if (product === 'google-sheets' && verification.method !== 'google-sheets-xlsx-export') {
    throw new Error(`UI responsiveness same-corpus verification method mismatch for ${caseId} ${product}`)
  }
  if (product === 'microsoft-excel-web' && verification.method !== 'microsoft-excel-web-source-xlsx') {
    throw new Error(`UI responsiveness same-corpus verification method mismatch for ${caseId} ${product}`)
  }
  if (product !== 'bilig' && verification.checkedCells.length < 3) {
    throw new Error(`UI responsiveness same-corpus verification must check at least 3 cells for ${caseId} ${product}`)
  }
  for (const cell of verification.checkedCells) {
    if (cell.address.trim().length === 0 || cell.expected !== cell.actual) {
      throw new Error(`UI responsiveness same-corpus verification cell mismatch for ${caseId} ${product}`)
    }
  }
}

function cloneSameCorpusVerification(verification: SameCorpusCaptureCorpusVerification): SameCorpusCaptureCorpusVerification {
  return {
    verified: verification.verified,
    method: verification.method,
    sheetName: verification.sheetName,
    materializedCells: verification.materializedCells,
    checkedCells: verification.checkedCells.map((cell) => ({ ...cell })),
  }
}

function sampleCountForSameCorpus(): number {
  return 3
}

function validateSummary(summary: NumericSummary, label: string): void {
  if (summary.samples.length < sampleCount) {
    throw new Error(`UI responsiveness live browser scorecard has too few samples for ${label}`)
  }
  for (const value of [summary.min, summary.median, summary.p95, summary.max, summary.mean, ...summary.samples]) {
    if (!Number.isFinite(value) || value < 0) {
      throw new Error(`UI responsiveness live browser scorecard has invalid numeric summary for ${label}`)
    }
  }
}

function argumentValue(name: string): string | null {
  const index = process.argv.indexOf(name)
  if (index === -1) {
    return null
  }
  const value = process.argv[index + 1]
  if (!value) {
    throw new Error(`Missing value after ${name}`)
  }
  return value
}

function logResult(mode: 'check' | 'write', scorecard: UiResponsivenessLiveBrowserScorecard): void {
  console.log(
    JSON.stringify(
      {
        mode,
        outputPath,
        allRequiredCasesPassed: scorecard.summary.allRequiredCasesPassed,
        capturedVendors: scorecard.summary.capturedVendors,
        caseCount: scorecard.cases.length,
        sameCorpusProofCaptured: scorecard.sameCorpusProof.captured,
        sameCorpusTenXMeanAndP95CaseCount: scorecard.sameCorpusProof.tenXMeanAndP95CaseCount,
      },
      null,
      2,
    ),
  )
}

function formatJsonForRepo(value: string): string {
  return `${value.trim()}\n`
}

if (import.meta.url === `file://${process.argv[1]}`) {
  await main()
}
