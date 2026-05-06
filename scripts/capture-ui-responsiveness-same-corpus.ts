#!/usr/bin/env bun

import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { dirname, join, resolve } from 'node:path'
import { performance } from 'node:perf_hooks'
import { pathToFileURL } from 'node:url'

import { chromium, type Browser, type BrowserContextOptions, type Page } from '@playwright/test'
import { exportXlsx } from '../packages/excel-import/src/index.js'
import {
  buildWorkbookBenchmarkCorpus,
  getWorkbookBenchmarkCorpusDefinition,
  isWorkbookBenchmarkCorpusId,
  type WorkbookBenchmarkCorpusId,
} from '../packages/benchmarks/src/workbook-corpus.js'
import type {
  SameCorpusCapture,
  SameCorpusCaptureMeasurement,
  UiResponsivenessSameCorpusProduct,
} from './gen-ui-responsiveness-live-browser-scorecard.ts'
import { formatJsonForRepo } from './scorecard-format.ts'

interface CaptureArgs {
  readonly biligUrl: string
  readonly biligStorageStatePath: string | null
  readonly corpusId: WorkbookBenchmarkCorpusId
  readonly deltaX: number
  readonly deltaY: number
  readonly googleSheetsUrl: string
  readonly googleSheetsStorageStatePath: string | null
  readonly headless: boolean
  readonly microsoftExcelWebUrl: string
  readonly microsoftExcelWebStorageStatePath: string | null
  readonly outputPath: string
  readonly readyTimeoutMs: number
  readonly sampleCount: number
  readonly storageStatePath: string | null
}

interface EmitXlsxArgs {
  readonly check: boolean
  readonly corpusId: WorkbookBenchmarkCorpusId
  readonly targetDirectory: string
}

interface SaveStorageStateArgs {
  readonly authUrl: string
  readonly corpusId: WorkbookBenchmarkCorpusId
  readonly headless: boolean
  readonly product: UiResponsivenessSameCorpusProduct
  readonly readyTimeoutMs: number
  readonly targetPath: string
}

interface ScrollSample {
  readonly operationResponseMs: number
  readonly postOperationFrameMs: number
}

const rootDir = resolve(new URL('..', import.meta.url).pathname)
const defaultCorpusId: WorkbookBenchmarkCorpusId = 'wide-mixed-250k'
const defaultViewport = { width: 1440, height: 900 } as const

async function main(): Promise<void> {
  const saveStorageStateArgs = parseSaveStorageStateArgs(process.argv.slice(2))
  if (saveStorageStateArgs) {
    await saveStorageState(saveStorageStateArgs)
    return
  }
  const emitXlsxArgs = parseEmitXlsxArgs(process.argv.slice(2))
  if (emitXlsxArgs) {
    emitSameCorpusXlsx(emitXlsxArgs)
    return
  }
  const args = parseCaptureArgs(process.argv.slice(2))
  const capture = await captureSameCorpusUiResponsiveness(args)
  mkdirSync(dirname(args.outputPath), { recursive: true })
  writeFileSync(
    args.outputPath,
    formatJsonForRepo({
      rootDir,
      serializedJson: `${JSON.stringify(capture, null, 2)}\n`,
      tempPrefix: 'ui-responsiveness-same-corpus-capture',
    }),
  )
  console.log(
    JSON.stringify(
      {
        outputPath: args.outputPath,
        corpusCaseId: args.corpusId,
        sampleCount: args.sampleCount,
        workload: 'visible-scroll-response',
      },
      null,
      2,
    ),
  )
}

export function parseEmitXlsxArgs(argv: readonly string[]): EmitXlsxArgs | null {
  const emitIndex = argv.indexOf('--emit-xlsx')
  if (emitIndex === -1) {
    return null
  }
  const targetDirectory = argv[emitIndex + 1]
  if (!targetDirectory) {
    throw new Error('Missing directory after --emit-xlsx')
  }
  return {
    check: argv.includes('--check'),
    corpusId: parseCorpusId(argumentValue(argv, '--corpus') ?? defaultCorpusId),
    targetDirectory: resolve(targetDirectory),
  }
}

export function parseSaveStorageStateArgs(argv: readonly string[]): SaveStorageStateArgs | null {
  const saveIndex = argv.indexOf('--save-storage-state')
  if (saveIndex === -1) {
    return null
  }
  const targetPath = argv[saveIndex + 1]
  if (!targetPath) {
    throw new Error('Missing file path after --save-storage-state')
  }
  const product = parseSameCorpusProduct(argumentValue(argv, '--auth-product') ?? 'google-sheets')
  const authUrl = argumentValue(argv, '--auth-url') ?? authUrlFromProductArgs(argv, product)
  if (!authUrl) {
    throw new Error('Missing auth URL. Pass --auth-url <url> or the product-specific URL flag.')
  }
  return {
    authUrl,
    corpusId: parseCorpusId(argumentValue(argv, '--corpus') ?? defaultCorpusId),
    headless: argv.includes('--headless'),
    product,
    readyTimeoutMs: parsePositiveInteger(argumentValue(argv, '--ready-timeout-ms') ?? '300000', '--ready-timeout-ms'),
    targetPath: resolve(targetPath),
  }
}

function authUrlFromProductArgs(argv: readonly string[], product: UiResponsivenessSameCorpusProduct): string | null {
  if (product === 'bilig') {
    return argumentValue(argv, '--bilig-url')
  }
  if (product === 'google-sheets') {
    return argumentValue(argv, '--google-sheets-url')
  }
  return argumentValue(argv, '--microsoft-excel-web-url')
}

async function saveStorageState(args: SaveStorageStateArgs): Promise<void> {
  const browser = await chromium.launch({ headless: args.headless })
  const context = await browser.newContext({ viewport: defaultViewport })
  const page = await context.newPage()
  try {
    await page.goto(args.authUrl, { waitUntil: 'domcontentloaded', timeout: 60_000 })
    await waitForProductReady(page, args.product, captureArgsForStorageState(args))
    mkdirSync(dirname(args.targetPath), { recursive: true })
    await context.storageState({ path: args.targetPath })
    console.log(
      JSON.stringify(
        {
          mode: 'save-storage-state',
          product: args.product,
          targetPath: args.targetPath,
          finalUrl: page.url(),
          title: await page.title(),
        },
        null,
        2,
      ),
    )
  } catch (error: unknown) {
    throw new Error(await productReadyFailureMessage(page, args.product, args.authUrl, 0, error), { cause: error })
  } finally {
    await context.close()
    await browser.close()
  }
}

function captureArgsForStorageState(args: SaveStorageStateArgs): CaptureArgs {
  return {
    biligUrl: args.authUrl,
    biligStorageStatePath: null,
    corpusId: args.corpusId,
    deltaX: 0,
    deltaY: 720,
    googleSheetsUrl: args.authUrl,
    googleSheetsStorageStatePath: null,
    headless: args.headless,
    microsoftExcelWebUrl: args.authUrl,
    microsoftExcelWebStorageStatePath: null,
    outputPath: args.targetPath,
    readyTimeoutMs: args.readyTimeoutMs,
    sampleCount: 1,
    storageStatePath: null,
  }
}

export function emitSameCorpusXlsx(args: EmitXlsxArgs): void {
  mkdirSync(args.targetDirectory, { recursive: true })
  const corpus = buildWorkbookBenchmarkCorpus(args.corpusId)
  const outputFile = join(args.targetDirectory, `${args.corpusId}.xlsx`)
  const workbookBytes = Buffer.from(exportXlsx(corpus.snapshot))
  if (args.check) {
    if (!existsSync(outputFile)) {
      throw new Error(`Same-corpus XLSX fixture is missing: ${outputFile}`)
    }
    const existingBytes = readFileSync(outputFile)
    if (!existingBytes.equals(workbookBytes)) {
      throw new Error(`Same-corpus XLSX fixture is stale: ${outputFile}`)
    }
  } else {
    writeFileSync(outputFile, workbookBytes)
  }
  const publicGithubRawUrl = `https://raw.githubusercontent.com/proompteng/bilig/main/packages/benchmarks/baselines/ui-same-corpus/${corpus.id}.xlsx`
  console.log(
    JSON.stringify(
      {
        mode: args.check ? 'check-xlsx' : 'emit-xlsx',
        outputFile,
        corpusCaseId: corpus.id,
        materializedCells: corpus.materializedCellCount,
        googleSheetsUploadMode: 'native_google_sheets',
        publicGithubRawUrl,
        publicForgejoRawUrl: `https://code.proompteng.ai/kalmyk/bilig/raw/branch/main/packages/benchmarks/baselines/ui-same-corpus/${corpus.id}.xlsx`,
        microsoftExcelWebUrl: `https://view.officeapps.live.com/op/view.aspx?src=${encodeURIComponent(publicGithubRawUrl)}`,
        googleSheetsAuthStateCommand:
          'pnpm ui:same-corpus:capture -- --save-storage-state <state.json> --auth-product google-sheets --google-sheets-url <url>',
        captureCommand:
          'pnpm ui:same-corpus:capture -- --output <capture.json> --google-sheets-url <url> --microsoft-excel-web-url <url> [--google-sheets-storage-state <state.json>]',
      },
      null,
      2,
    ),
  )
}

export function parseCaptureArgs(argv: readonly string[]): CaptureArgs {
  const corpusId = parseCorpusId(argumentValue(argv, '--corpus') ?? defaultCorpusId)
  const outputPath = argumentValue(argv, '--output')
  const googleSheetsUrl = argumentValue(argv, '--google-sheets-url')
  const microsoftExcelWebUrl = argumentValue(argv, '--microsoft-excel-web-url')
  if (!outputPath || !googleSheetsUrl || !microsoftExcelWebUrl) {
    throw new Error(
      [
        'Missing required arguments.',
        'Usage: bun scripts/capture-ui-responsiveness-same-corpus.ts',
        '  --output <capture.json>',
        '  --google-sheets-url <same-corpus-google-sheets-url>',
        '  --microsoft-excel-web-url <same-corpus-excel-web-url>',
        '  or: --emit-xlsx <directory>',
        '  [--bilig-url <local-bilig-url>] [--corpus wide-mixed-250k] [--samples 3] [--delta-x 0] [--delta-y 720] [--headed]',
        '  [--storage-state <state.json>]',
        '  [--google-sheets-storage-state <state.json>] [--microsoft-excel-web-storage-state <state.json>] [--bilig-storage-state <state.json>]',
        '  [--ready-timeout-ms 60000]',
      ].join('\n'),
    )
  }
  const sampleCount = parsePositiveInteger(argumentValue(argv, '--samples') ?? '3', '--samples')
  const readyTimeoutMs = parsePositiveInteger(argumentValue(argv, '--ready-timeout-ms') ?? '60000', '--ready-timeout-ms')
  return {
    biligUrl: argumentValue(argv, '--bilig-url') ?? `http://127.0.0.1:5173/?benchmarkCorpus=${encodeURIComponent(corpusId)}`,
    biligStorageStatePath: resolveOptionalPath(argumentValue(argv, '--bilig-storage-state')),
    corpusId,
    deltaX: parseNonNegativeNumber(argumentValue(argv, '--delta-x') ?? '0', '--delta-x'),
    deltaY: parseNonNegativeNumber(argumentValue(argv, '--delta-y') ?? '720', '--delta-y'),
    googleSheetsUrl,
    googleSheetsStorageStatePath: resolveOptionalPath(argumentValue(argv, '--google-sheets-storage-state')),
    headless: !argv.includes('--headed'),
    microsoftExcelWebUrl,
    microsoftExcelWebStorageStatePath: resolveOptionalPath(argumentValue(argv, '--microsoft-excel-web-storage-state')),
    outputPath: resolve(outputPath),
    readyTimeoutMs,
    sampleCount,
    storageStatePath: resolveOptionalPath(argumentValue(argv, '--storage-state')),
  }
}

export async function captureSameCorpusUiResponsiveness(args: CaptureArgs): Promise<SameCorpusCapture> {
  const corpus = getWorkbookBenchmarkCorpusDefinition(args.corpusId)
  const browser = await chromium.launch({ headless: args.headless })
  try {
    const [bilig, googleSheets, microsoftExcelWeb] = await Promise.all([
      measureProduct(browser, 'bilig', args.biligUrl, args),
      measureProduct(browser, 'google-sheets', args.googleSheetsUrl, args),
      measureProduct(browser, 'microsoft-excel-web', args.microsoftExcelWebUrl, args),
    ])
    return {
      schemaVersion: 1,
      suite: 'ui-responsiveness-same-corpus-capture',
      sampleCount: args.sampleCount,
      limitations: [
        'Caller must supply Google Sheets and Microsoft Excel Web URLs for the same exported Bilig benchmark corpus.',
        'This capture measures browser-visible scroll response; edit latency must be captured by a separate same-corpus workload.',
      ],
      cases: [
        {
          id: `same-corpus-${args.corpusId}-visible-scroll-response`,
          corpusCaseId: args.corpusId,
          materializedCells: corpus.materializedCellCount,
          workload: 'visible-scroll-response',
          bilig,
          googleSheets,
          microsoftExcelWeb,
        },
      ],
    }
  } finally {
    await browser.close()
  }
}

async function measureProduct(
  browser: Browser,
  product: UiResponsivenessSameCorpusProduct,
  url: string,
  args: CaptureArgs,
): Promise<SameCorpusCaptureMeasurement> {
  const samples = await measureProductSamples(browser, product, url, args)

  return {
    product,
    source: url,
    operationResponseMsSamples: samples.map((entry) => entry.operationResponseMs),
    postOperationFrameMsSamples: samples.map((entry) => entry.postOperationFrameMs),
    limitations: productLimitations(product, storageStatePathForProduct(product, args)),
  }
}

async function measureProductSamples(
  browser: Browser,
  product: UiResponsivenessSameCorpusProduct,
  url: string,
  args: CaptureArgs,
  sampleIndex = 0,
  samples: ScrollSample[] = [],
): Promise<ScrollSample[]> {
  if (sampleIndex >= args.sampleCount) {
    return samples
  }
  const context = await browser.newContext(browserContextOptionsForProduct(product, args))
  const page = await context.newPage()
  try {
    await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 60_000 })
    await waitForProductReady(page, product, args)
    samples.push(await measureVisibleScrollResponse(page, args.deltaX, args.deltaY))
  } catch (error: unknown) {
    throw new Error(await productReadyFailureMessage(page, product, url, sampleIndex, error), { cause: error })
  } finally {
    await context.close()
  }
  return measureProductSamples(browser, product, url, args, sampleIndex + 1, samples)
}

function browserContextOptionsForProduct(product: UiResponsivenessSameCorpusProduct, args: CaptureArgs): BrowserContextOptions {
  const storageState = storageStatePathForProduct(product, args)
  return {
    viewport: defaultViewport,
    ...(storageState ? { storageState } : {}),
  }
}

function storageStatePathForProduct(product: UiResponsivenessSameCorpusProduct, args: CaptureArgs): string | null {
  if (product === 'bilig') {
    return args.biligStorageStatePath ?? args.storageStatePath
  }
  if (product === 'google-sheets') {
    return args.googleSheetsStorageStatePath ?? args.storageStatePath
  }
  return args.microsoftExcelWebStorageStatePath ?? args.storageStatePath
}

async function waitForProductReady(page: Page, product: UiResponsivenessSameCorpusProduct, args: CaptureArgs): Promise<void> {
  if (product === 'bilig') {
    await page.waitForSelector('[data-testid="sheet-grid"]', { state: 'visible', timeout: args.readyTimeoutMs })
    await page.waitForFunction(
      (expectedCorpusId) => {
        const collector = (
          window as Window & {
            __biligScrollPerf?: {
              getBenchmarkState?: () => {
                state: string
                error: string | null
                fixture: { id: string; materializedCellCount: number; sheetName: string } | null
              }
            }
          }
        ).__biligScrollPerf
        const state = collector?.getBenchmarkState?.()
        return state?.state === 'ready' && state.fixture?.id === expectedCorpusId
      },
      args.corpusId,
      { timeout: args.readyTimeoutMs },
    )
    await settleFrames(page, 12)
    return
  }

  if (product === 'google-sheets') {
    await page.waitForFunction(
      () =>
        !window.location.href.includes('accounts.google.com') &&
        document.title.includes('Google Sheets') &&
        !document.body.innerText.includes('Sign in\nto continue to Google Sheets'),
      { timeout: args.readyTimeoutMs },
    )
    await settleFrames(page, 120)
    return
  }

  await page.waitForFunction(
    () => document.title.toLowerCase().includes('.xlsx') || document.body.innerText.toLowerCase().includes('excel'),
    { timeout: args.readyTimeoutMs },
  )
  await settleFrames(page, 180)
}

async function productReadyFailureMessage(
  page: Page,
  product: UiResponsivenessSameCorpusProduct,
  sourceUrl: string,
  sampleIndex: number,
  cause: unknown,
): Promise<string> {
  const diagnostic = await collectPageDiagnostic(page)
  const causeMessage = cause instanceof Error ? cause.message : String(cause)
  const productHint =
    product === 'google-sheets' && diagnostic.finalUrl.includes('accounts.google.com')
      ? 'Google Sheets redirected to sign-in; provide a public/shareable sheet URL or run with --google-sheets-storage-state from an authenticated Playwright session.'
      : product === 'microsoft-excel-web' && sourceUrl.includes('view.officeapps.live.com')
        ? 'Microsoft Excel Web did not become measurable; confirm the viewer URL wraps a Microsoft-accessible public HTTPS XLSX URL for the same emitted corpus.'
        : 'The same-corpus page did not reach the expected measurable state.'
  return [
    `Failed to prepare ${product} for same-corpus UI capture on sample ${String(sampleIndex + 1)}.`,
    productHint,
    `sourceUrl: ${sourceUrl}`,
    `finalUrl: ${diagnostic.finalUrl}`,
    `title: ${diagnostic.title}`,
    `body: ${diagnostic.bodySnippet}`,
    `cause: ${causeMessage}`,
  ].join('\n')
}

async function collectPageDiagnostic(page: Page): Promise<{ finalUrl: string; title: string; bodySnippet: string }> {
  const [title, bodySnippet] = await Promise.all([
    page.title().catch(() => ''),
    page
      .locator('body')
      .innerText({ timeout: 2_000 })
      .catch(() => ''),
  ])
  return {
    finalUrl: page.url(),
    title,
    bodySnippet: bodySnippet.replace(/\s+/g, ' ').trim().slice(0, 500),
  }
}

async function measureVisibleScrollResponse(page: Page, deltaX: number, deltaY: number): Promise<ScrollSample> {
  await page.mouse.move(defaultViewport.width / 2, defaultViewport.height / 2)
  const startedAt = performance.now()
  await page.mouse.wheel(deltaX, deltaY)
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
  const operationResponseMs = performance.now() - startedAt
  return {
    operationResponseMs,
    postOperationFrameMs: percentile(frameIntervals, 0.95),
  }
}

async function settleFrames(page: Page, frames: number): Promise<void> {
  await page.evaluate(async (frameCount) => {
    await Array.from({ length: frameCount }).reduce<Promise<void>>(async (previous) => {
      await previous
      await new Promise<void>((resolveFrame) => requestAnimationFrame(() => resolveFrame()))
    }, Promise.resolve())
  }, frames)
}

function productLimitations(product: UiResponsivenessSameCorpusProduct, storageStatePath: string | null): string[] {
  const authLimitations = storageStatePath ? ['Browser context used an explicit Playwright storage state for authenticated access.'] : []
  if (product === 'bilig') {
    return ['Bilig timing is captured from the supplied local app URL and benchmarkCorpus route.', ...authLimitations]
  }
  if (product === 'google-sheets') {
    return [
      'Google Sheets timing requires the supplied URL to be browser-accessible and loaded with the same benchmark corpus.',
      ...authLimitations,
    ]
  }
  return [
    'Microsoft Excel Web timing requires the supplied URL to be browser-accessible and loaded with the same benchmark corpus.',
    ...authLimitations,
  ]
}

function percentile(values: readonly number[], percentileValue: number): number {
  if (values.length === 0) {
    throw new Error('Cannot compute percentile for an empty sample set')
  }
  const sorted = [...values].toSorted((left, right) => left - right)
  const index = Math.ceil(percentileValue * sorted.length) - 1
  return sorted[Math.max(0, Math.min(sorted.length - 1, index))]!
}

function parseCorpusId(value: string): WorkbookBenchmarkCorpusId {
  if (!isWorkbookBenchmarkCorpusId(value)) {
    throw new Error(`Unexpected workbook benchmark corpus id: ${value}`)
  }
  return value
}

function parseSameCorpusProduct(value: string): UiResponsivenessSameCorpusProduct {
  if (value === 'bilig' || value === 'google-sheets' || value === 'microsoft-excel-web') {
    return value
  }
  throw new Error(`Unexpected same-corpus product: ${value}`)
}

function parsePositiveInteger(value: string, flag: string): number {
  const parsed = Number(value)
  if (!Number.isInteger(parsed) || parsed < 1) {
    throw new Error(`${flag} must be a positive integer`)
  }
  return parsed
}

function parseNonNegativeNumber(value: string, flag: string): number {
  const parsed = Number(value)
  if (!Number.isFinite(parsed) || parsed < 0) {
    throw new Error(`${flag} must be a non-negative number`)
  }
  return parsed
}

function resolveOptionalPath(value: string | null): string | null {
  return value ? resolve(value) : null
}

function argumentValue(argv: readonly string[], name: string): string | null {
  const index = argv.indexOf(name)
  if (index === -1) {
    return null
  }
  const value = argv[index + 1]
  if (!value) {
    throw new Error(`Missing value after ${name}`)
  }
  return value
}

if (process.argv[1] && pathToFileURL(resolve(process.argv[1])).href === import.meta.url) {
  try {
    await main()
  } catch (error: unknown) {
    console.error(error instanceof Error ? error.message : String(error))
    process.exitCode = 1
  }
}
