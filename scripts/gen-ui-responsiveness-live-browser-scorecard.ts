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

  const scorecard = await buildUiResponsivenessLiveBrowserScorecard(new Date().toISOString())
  validateUiResponsivenessLiveBrowserScorecard(scorecard)
  mkdirSync(dirname(outputPath), { recursive: true })
  writeFileSync(outputPath, formatJsonForRepo(`${JSON.stringify(scorecard, null, 2)}\n`))
  logResult('write', scorecard)
}

export async function buildUiResponsivenessLiveBrowserScorecard(generatedAt: string): Promise<UiResponsivenessLiveBrowserScorecard> {
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

function parseVendor(value: string): UiResponsivenessLiveBrowserVendor {
  if (value === 'google-sheets' || value === 'microsoft-excel-web') {
    return value
  }
  throw new Error(`Unexpected UI responsiveness live browser vendor: ${value}`)
}

function parseAccessMode(value: string): UiResponsivenessLiveBrowserCase['accessMode'] {
  if (value === 'public-comment-only' || value === 'public-view-only' || value === 'public-office-web-viewer') {
    return value
  }
  throw new Error(`Unexpected UI responsiveness live browser access mode: ${value}`)
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

function logResult(mode: 'check' | 'write', scorecard: UiResponsivenessLiveBrowserScorecard): void {
  console.log(
    JSON.stringify(
      {
        mode,
        outputPath,
        allRequiredCasesPassed: scorecard.summary.allRequiredCasesPassed,
        capturedVendors: scorecard.summary.capturedVendors,
        caseCount: scorecard.cases.length,
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
