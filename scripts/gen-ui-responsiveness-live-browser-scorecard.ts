#!/usr/bin/env bun

import { existsSync, mkdirSync, writeFileSync } from 'node:fs'
import { dirname, join, resolve } from 'node:path'
import { performance } from 'node:perf_hooks'

import { chromium, type Browser, type Page } from '@playwright/test'
import { summarizeNumbers, type NumericSummary } from '../packages/benchmarks/src/stats.js'
import {
  arrayField,
  asObject,
  booleanField,
  literalField,
  numberField,
  objectField,
  readJsonObject,
  stringArrayField,
  stringField,
} from './json-scorecard-helpers.ts'

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
  readonly tenXMeanAndP95AgainstGoogleSheets: boolean
  readonly tenXMeanAndP95AgainstMicrosoftExcelWeb: boolean
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

export function parseUiResponsivenessLiveBrowserScorecard(value: Record<string, unknown>): UiResponsivenessLiveBrowserScorecard {
  const host = objectField(value, 'host')
  const source = objectField(value, 'source')
  const benchmark = objectField(value, 'benchmark')
  const benchmarkViewport = objectField(benchmark, 'viewport')
  const summary = objectField(value, 'summary')
  return {
    schemaVersion: literalField(value, 'schemaVersion', 1),
    suite: literalField(value, 'suite', 'ui-responsiveness-live-browser-timing'),
    generatedAt: stringField(value, 'generatedAt'),
    host: {
      arch: stringField(host, 'arch'),
      platform: stringField(host, 'platform'),
    },
    source: {
      artifactGenerator: literalField(source, 'artifactGenerator', 'scripts/gen-ui-responsiveness-live-browser-scorecard.ts'),
      evidenceKind: literalField(source, 'evidenceKind', 'live-public-browser-playwright'),
      browserEngine: literalField(source, 'browserEngine', 'chromium'),
      measuredOperation: literalField(source, 'measuredOperation', 'public-workbook-load-and-viewport-scroll'),
    },
    benchmark: {
      sampleCount: numberField(benchmark, 'sampleCount'),
      viewport: {
        width: numberField(benchmarkViewport, 'width'),
        height: numberField(benchmarkViewport, 'height'),
      },
      samplingOrder: literalField(benchmark, 'samplingOrder', 'google-sheets-then-microsoft-excel-web'),
    },
    summary: {
      directBrowserTimingCaptured: booleanField(summary, 'directBrowserTimingCaptured'),
      allRequiredCasesPassed: booleanField(summary, 'allRequiredCasesPassed'),
      requiredVendorCount: numberField(summary, 'requiredVendorCount'),
      capturedVendors: stringArrayField(summary, 'capturedVendors').map(parseVendor),
      limitations: stringArrayField(summary, 'limitations'),
    },
    cases: arrayField(value, 'cases').map(parseBrowserCase),
    sameCorpusProof: parseSameCorpusProof(objectField(value, 'sameCorpusProof')),
  }
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
  return {
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
}

function validateSameCorpusCapture(capture: SameCorpusCapture): void {
  if (capture.sampleCount < sampleCountForSameCorpus()) {
    throw new Error('UI responsiveness same-corpus capture must contain at least 3 samples per product')
  }
  if (capture.cases.length === 0) {
    throw new Error('UI responsiveness same-corpus capture must include at least one case')
  }
  for (const entry of capture.cases) {
    for (const measurement of [entry.bilig, entry.googleSheets, entry.microsoftExcelWeb]) {
      if (
        measurement.operationResponseMsSamples.length < capture.sampleCount ||
        measurement.postOperationFrameMsSamples.length < capture.sampleCount
      ) {
        throw new Error(`UI responsiveness same-corpus capture has too few samples for ${entry.id}`)
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
  const tenXMeanAndP95AgainstGoogleSheets = biligToGoogleSheetsMeanRatio <= 0.1 && biligToGoogleSheetsP95Ratio <= 0.1
  const tenXMeanAndP95AgainstMicrosoftExcelWeb = biligToMicrosoftExcelWebMeanRatio <= 0.1 && biligToMicrosoftExcelWebP95Ratio <= 0.1
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
    corpusVerification: cloneSameCorpusVerification(capture.corpusVerification),
    limitations: [...capture.limitations],
  }
}

function ratio(numerator: number, denominator: number): number {
  if (denominator <= 0) {
    return Number.POSITIVE_INFINITY
  }
  return numerator / denominator
}

function parseBrowserCase(value: unknown): UiResponsivenessLiveBrowserCase {
  const record = asObject(value, 'UI responsiveness live browser case')
  return {
    id: stringField(record, 'id'),
    vendor: parseVendor(stringField(record, 'vendor')),
    product: stringField(record, 'product'),
    sourceUrl: stringField(record, 'sourceUrl'),
    finalUrl: stringField(record, 'finalUrl'),
    title: stringField(record, 'title'),
    accessMode: parseAccessMode(stringField(record, 'accessMode')),
    workload: literalField(record, 'workload', 'open-public-workbook-and-scroll-viewport'),
    sampleCount: numberField(record, 'sampleCount'),
    loadToReadyMs: parseNumericSummary(objectField(record, 'loadToReadyMs')),
    scrollResponseMs: parseNumericSummary(objectField(record, 'scrollResponseMs')),
    postScrollFrameMs: parseNumericSummary(objectField(record, 'postScrollFrameMs')),
    passed: booleanField(record, 'passed'),
    limitations: stringArrayField(record, 'limitations'),
  }
}

function parseSameCorpusProof(value: Record<string, unknown>): UiResponsivenessSameCorpusProof {
  return {
    captured: booleanField(value, 'captured'),
    evidenceKind: parseSameCorpusEvidenceKind(stringField(value, 'evidenceKind')),
    requiredProductCount: numberField(value, 'requiredProductCount'),
    requiredCaseCount: numberField(value, 'requiredCaseCount'),
    tenXMeanAndP95CaseCount: numberField(value, 'tenXMeanAndP95CaseCount'),
    coveredCorpusCaseIds: stringArrayField(value, 'coveredCorpusCaseIds'),
    limitations: stringArrayField(value, 'limitations'),
    cases: arrayField(value, 'cases').map(parseSameCorpusCase),
  }
}

function parseSameCorpusCase(value: unknown): UiResponsivenessSameCorpusCase {
  const record = asObject(value, 'UI responsiveness same-corpus case')
  return {
    id: stringField(record, 'id'),
    corpusCaseId: stringField(record, 'corpusCaseId'),
    materializedCells: numberField(record, 'materializedCells'),
    workload: parseSameCorpusWorkload(stringField(record, 'workload')),
    sampleCount: numberField(record, 'sampleCount'),
    bilig: parseSameCorpusMeasurement(objectField(record, 'bilig')),
    googleSheets: parseSameCorpusMeasurement(objectField(record, 'googleSheets')),
    microsoftExcelWeb: parseSameCorpusMeasurement(objectField(record, 'microsoftExcelWeb')),
    biligToGoogleSheetsMeanRatio: numberField(record, 'biligToGoogleSheetsMeanRatio'),
    biligToGoogleSheetsP95Ratio: numberField(record, 'biligToGoogleSheetsP95Ratio'),
    biligToMicrosoftExcelWebMeanRatio: numberField(record, 'biligToMicrosoftExcelWebMeanRatio'),
    biligToMicrosoftExcelWebP95Ratio: numberField(record, 'biligToMicrosoftExcelWebP95Ratio'),
    tenXMeanAndP95AgainstGoogleSheets: booleanField(record, 'tenXMeanAndP95AgainstGoogleSheets'),
    tenXMeanAndP95AgainstMicrosoftExcelWeb: booleanField(record, 'tenXMeanAndP95AgainstMicrosoftExcelWeb'),
    passed: booleanField(record, 'passed'),
  }
}

function parseSameCorpusMeasurement(value: Record<string, unknown>): UiResponsivenessSameCorpusMeasurement {
  return {
    product: parseSameCorpusProduct(stringField(value, 'product')),
    source: stringField(value, 'source'),
    operationResponseMs: parseNumericSummary(objectField(value, 'operationResponseMs')),
    postOperationFrameMs: parseNumericSummary(objectField(value, 'postOperationFrameMs')),
    corpusVerification: parseSameCorpusVerification(objectField(value, 'corpusVerification')),
    limitations: stringArrayField(value, 'limitations'),
  }
}

function parseSameCorpusCapture(value: Record<string, unknown>): SameCorpusCapture {
  return {
    schemaVersion: literalField(value, 'schemaVersion', 1),
    suite: literalField(value, 'suite', 'ui-responsiveness-same-corpus-capture'),
    sampleCount: numberField(value, 'sampleCount'),
    limitations: stringArrayField(value, 'limitations'),
    cases: arrayField(value, 'cases').map(parseSameCorpusCaptureCase),
  }
}

function parseSameCorpusCaptureCase(value: unknown): SameCorpusCaptureCase {
  const record = asObject(value, 'UI responsiveness same-corpus capture case')
  return {
    id: stringField(record, 'id'),
    corpusCaseId: stringField(record, 'corpusCaseId'),
    materializedCells: numberField(record, 'materializedCells'),
    workload: parseSameCorpusWorkload(stringField(record, 'workload')),
    bilig: parseSameCorpusCaptureMeasurement(objectField(record, 'bilig'), 'bilig'),
    googleSheets: parseSameCorpusCaptureMeasurement(objectField(record, 'googleSheets'), 'google-sheets'),
    microsoftExcelWeb: parseSameCorpusCaptureMeasurement(objectField(record, 'microsoftExcelWeb'), 'microsoft-excel-web'),
  }
}

function parseSameCorpusCaptureMeasurement(
  value: Record<string, unknown>,
  product: UiResponsivenessSameCorpusProduct,
): SameCorpusCaptureMeasurement {
  const parsedProduct = parseSameCorpusProduct(stringField(value, 'product'))
  if (parsedProduct !== product) {
    throw new Error(`UI responsiveness same-corpus capture product mismatch: expected ${product}, got ${parsedProduct}`)
  }
  return {
    product: parsedProduct,
    source: stringField(value, 'source'),
    operationResponseMsSamples: numericArrayField(value, 'operationResponseMsSamples'),
    postOperationFrameMsSamples: numericArrayField(value, 'postOperationFrameMsSamples'),
    corpusVerification: parseSameCorpusVerification(objectField(value, 'corpusVerification')),
    limitations: stringArrayField(value, 'limitations'),
  }
}

function parseSameCorpusVerification(value: Record<string, unknown>): SameCorpusCaptureCorpusVerification {
  return {
    verified: booleanField(value, 'verified'),
    method: parseSameCorpusVerificationMethod(stringField(value, 'method')),
    sheetName: stringField(value, 'sheetName'),
    materializedCells: numberField(value, 'materializedCells'),
    checkedCells: arrayField(value, 'checkedCells').map(parseSameCorpusVerifiedCell),
  }
}

function parseSameCorpusVerifiedCell(value: unknown): SameCorpusCaptureVerifiedCell {
  const record = asObject(value, 'UI responsiveness same-corpus verified cell')
  return {
    address: stringField(record, 'address'),
    expected: stringField(record, 'expected'),
    actual: stringField(record, 'actual'),
  }
}

function parseNumericSummary(value: Record<string, unknown>): NumericSummary {
  return {
    samples: arrayField(value, 'samples').map((entry) => {
      if (typeof entry !== 'number' || !Number.isFinite(entry)) {
        throw new Error('Expected numeric summary samples to contain finite numbers')
      }
      return entry
    }),
    min: numberField(value, 'min'),
    median: numberField(value, 'median'),
    p95: numberField(value, 'p95'),
    max: numberField(value, 'max'),
    mean: numberField(value, 'mean'),
  }
}

function numericArrayField(value: Record<string, unknown>, key: string): number[] {
  return arrayField(value, key).map((entry) => {
    if (typeof entry !== 'number' || !Number.isFinite(entry) || entry < 0) {
      throw new Error(`Expected ${key} to contain finite non-negative numbers`)
    }
    return entry
  })
}

function parseVendor(value: string): UiResponsivenessLiveBrowserVendor {
  if (value === 'google-sheets' || value === 'microsoft-excel-web') {
    return value
  }
  throw new Error(`Unexpected UI responsiveness live browser vendor: ${value}`)
}

function parseSameCorpusEvidenceKind(value: string): UiResponsivenessSameCorpusProof['evidenceKind'] {
  if (value === 'same-corpus-browser-capture' || value === 'not-captured') {
    return value
  }
  throw new Error(`Unexpected UI responsiveness same-corpus evidence kind: ${value}`)
}

function parseSameCorpusProduct(value: string): UiResponsivenessSameCorpusProduct {
  if (value === 'bilig' || value === 'google-sheets' || value === 'microsoft-excel-web') {
    return value
  }
  throw new Error(`Unexpected UI responsiveness same-corpus product: ${value}`)
}

function parseSameCorpusVerificationMethod(value: string): SameCorpusCaptureCorpusVerification['method'] {
  if (value === 'bilig-benchmark-state' || value === 'google-sheets-xlsx-export' || value === 'microsoft-excel-web-source-xlsx') {
    return value
  }
  throw new Error(`Unexpected UI responsiveness same-corpus verification method: ${value}`)
}

function parseSameCorpusWorkload(value: string): UiResponsivenessSameCorpusWorkload {
  if (value === 'visible-scroll-response' || value === 'visible-edit-commit') {
    return value
  }
  throw new Error(`Unexpected UI responsiveness same-corpus workload: ${value}`)
}

function parseAccessMode(value: string): UiResponsivenessLiveBrowserCase['accessMode'] {
  if (value === 'public-comment-only' || value === 'public-view-only' || value === 'public-office-web-viewer') {
    return value
  }
  throw new Error(`Unexpected UI responsiveness live browser access mode: ${value}`)
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
  const tenXAgainstGoogleSheets = googleSheetsMeanRatio <= 0.1 && googleSheetsP95Ratio <= 0.1
  const tenXAgainstMicrosoftExcelWeb = microsoftExcelWebMeanRatio <= 0.1 && microsoftExcelWebP95Ratio <= 0.1
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
