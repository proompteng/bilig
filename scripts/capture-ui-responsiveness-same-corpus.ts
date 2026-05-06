#!/usr/bin/env bun

import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { dirname, join, resolve } from 'node:path'
import { performance } from 'node:perf_hooks'
import { pathToFileURL } from 'node:url'

import { chromium, type Browser, type BrowserContextOptions, type Page } from '@playwright/test'
import * as XLSX from 'xlsx'
import { exportXlsx } from '../packages/excel-import/src/index.js'
import {
  buildWorkbookBenchmarkCorpus,
  isWorkbookBenchmarkCorpusId,
  type WorkbookBenchmarkCorpusCase,
  type WorkbookBenchmarkCorpusId,
} from '../packages/benchmarks/src/workbook-corpus.js'
import type {
  SameCorpusCaptureCorpusVerification,
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

interface ProductSampleCollection {
  readonly corpusVerification: SameCorpusCaptureCorpusVerification
  readonly samples: readonly ScrollSample[]
}

interface SameCorpusFingerprint {
  readonly materializedCells: number
  readonly sheetName: string
  readonly checkedCells: readonly {
    readonly address: string
    readonly expected: string
  }[]
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
  const corpus = buildWorkbookBenchmarkCorpus(args.corpusId)
  const browser = await chromium.launch({ headless: args.headless })
  try {
    const [bilig, googleSheets, microsoftExcelWeb] = await Promise.all([
      measureProduct(browser, 'bilig', args.biligUrl, corpus, args),
      measureProduct(browser, 'google-sheets', args.googleSheetsUrl, corpus, args),
      measureProduct(browser, 'microsoft-excel-web', args.microsoftExcelWebUrl, corpus, args),
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
  corpus: WorkbookBenchmarkCorpusCase,
  args: CaptureArgs,
): Promise<SameCorpusCaptureMeasurement> {
  const { corpusVerification, samples } = await measureProductSamples(browser, product, url, corpus, args)

  return {
    product,
    source: url,
    operationResponseMsSamples: samples.map((entry) => entry.operationResponseMs),
    postOperationFrameMsSamples: samples.map((entry) => entry.postOperationFrameMs),
    corpusVerification,
    limitations: productLimitations(product, storageStatePathForProduct(product, args)),
  }
}

async function measureProductSamples(
  browser: Browser,
  product: UiResponsivenessSameCorpusProduct,
  url: string,
  corpus: WorkbookBenchmarkCorpusCase,
  args: CaptureArgs,
  sampleIndex = 0,
  samples: ScrollSample[] = [],
  corpusVerification: SameCorpusCaptureCorpusVerification | null = null,
): Promise<ProductSampleCollection> {
  if (sampleIndex >= args.sampleCount) {
    if (!corpusVerification) {
      throw new Error(`Missing same-corpus fingerprint verification for ${product}`)
    }
    return { corpusVerification, samples }
  }
  const context = await browser.newContext(browserContextOptionsForProduct(product, args))
  const page = await context.newPage()
  let nextCorpusVerification = corpusVerification
  try {
    await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 60_000 })
    await waitForProductReady(page, product, args)
    nextCorpusVerification ??= await verifyProductCorpus(page, product, url, corpus)
    samples.push(await measureVisibleScrollResponse(page, args.deltaX, args.deltaY))
  } catch (error: unknown) {
    throw new Error(await productReadyFailureMessage(page, product, url, sampleIndex, error), { cause: error })
  } finally {
    await context.close()
  }
  return measureProductSamples(browser, product, url, corpus, args, sampleIndex + 1, samples, nextCorpusVerification)
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

async function verifyProductCorpus(
  page: Page,
  product: UiResponsivenessSameCorpusProduct,
  sourceUrl: string,
  corpus: WorkbookBenchmarkCorpusCase,
): Promise<SameCorpusCaptureCorpusVerification> {
  if (product === 'bilig') {
    return verifyBiligBenchmarkState(page, corpus)
  }
  if (product === 'google-sheets') {
    return verifyGoogleSheetsXlsxExport(page, sourceUrl, corpus)
  }
  return verifyMicrosoftExcelWebSourceXlsx(page, sourceUrl, corpus)
}

async function verifyBiligBenchmarkState(page: Page, corpus: WorkbookBenchmarkCorpusCase): Promise<SameCorpusCaptureCorpusVerification> {
  const state = await page.evaluate(() => {
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
    return collector?.getBenchmarkState?.() ?? null
  })
  if (state?.state !== 'ready' || state.fixture?.id !== corpus.id) {
    throw new Error(`Bilig benchmark state does not match same-corpus fixture: expected ${corpus.id}, got ${JSON.stringify(state)}`)
  }
  if (state.fixture.materializedCellCount !== corpus.materializedCellCount) {
    throw new Error(
      `Bilig benchmark state has stale materialized cell count: expected ${String(corpus.materializedCellCount)}, got ${String(
        state.fixture.materializedCellCount,
      )}`,
    )
  }
  return {
    verified: true,
    method: 'bilig-benchmark-state',
    sheetName: state.fixture.sheetName,
    materializedCells: state.fixture.materializedCellCount,
    checkedCells: [],
  }
}

async function verifyGoogleSheetsXlsxExport(
  page: Page,
  sourceUrl: string,
  corpus: WorkbookBenchmarkCorpusCase,
): Promise<SameCorpusCaptureCorpusVerification> {
  const exportUrl = googleSheetsExportUrl(sourceUrl)
  const bytes = await fetchXlsxBytesForPage(page, exportUrl, 'Google Sheets same-corpus XLSX export')
  return verifyXlsxCorpusFingerprint(bytes, corpus, 'google-sheets-xlsx-export')
}

async function verifyMicrosoftExcelWebSourceXlsx(
  page: Page,
  sourceUrl: string,
  corpus: WorkbookBenchmarkCorpusCase,
): Promise<SameCorpusCaptureCorpusVerification> {
  const xlsxUrl = microsoftExcelWebSourceUrl(sourceUrl)
  const bytes = await fetchXlsxBytesForPage(page, xlsxUrl, 'Microsoft Excel Web same-corpus source XLSX')
  return verifyXlsxCorpusFingerprint(bytes, corpus, 'microsoft-excel-web-source-xlsx')
}

async function fetchXlsxBytesForPage(page: Page, url: string, label: string): Promise<Uint8Array> {
  const response = await page.context().request.get(url, { timeout: 60_000 })
  if (!response.ok()) {
    const bodySnippet = (await response.text().catch(() => '')).replace(/\s+/g, ' ').trim().slice(0, 300)
    throw new Error(`${label} returned HTTP ${String(response.status())}: ${bodySnippet}`)
  }
  const bytes = await response.body()
  if (looksLikeHtml(bytes)) {
    throw new Error(`${label} returned HTML instead of XLSX bytes`)
  }
  return bytes
}

function googleSheetsExportUrl(sourceUrl: string): string {
  const spreadsheetId = /\/spreadsheets\/d\/([^/?#]+)/u.exec(sourceUrl)?.[1]
  if (!spreadsheetId) {
    throw new Error(`Unable to extract Google Sheets spreadsheet ID from URL: ${sourceUrl}`)
  }
  return `https://docs.google.com/spreadsheets/d/${encodeURIComponent(spreadsheetId)}/export?format=xlsx`
}

function microsoftExcelWebSourceUrl(sourceUrl: string): string {
  const parsed = new URL(sourceUrl)
  if (!parsed.hostname.includes('view.officeapps.live.com')) {
    return sourceUrl
  }
  const source = parsed.searchParams.get('src')
  if (!source) {
    throw new Error(`Unable to extract Microsoft Excel Web source XLSX URL from viewer URL: ${sourceUrl}`)
  }
  return source
}

function looksLikeHtml(bytes: Uint8Array): boolean {
  const prefix = new TextDecoder()
    .decode(bytes.slice(0, Math.min(bytes.length, 256)))
    .trimStart()
    .toLowerCase()
  return prefix.startsWith('<!doctype html') || prefix.startsWith('<html')
}

export function verifyXlsxCorpusFingerprint(
  bytes: Uint8Array,
  corpus: WorkbookBenchmarkCorpusCase,
  method: SameCorpusCaptureCorpusVerification['method'],
): SameCorpusCaptureCorpusVerification {
  const fingerprint = buildSameCorpusFingerprint(corpus)
  const workbook = XLSX.read(Buffer.from(bytes), { type: 'buffer' })
  const worksheet = workbook.Sheets[fingerprint.sheetName]
  if (!worksheet) {
    throw new Error(`Same-corpus XLSX is missing sheet: ${fingerprint.sheetName}`)
  }
  const checkedCells = fingerprint.checkedCells.map((cell) => {
    const actual = normalizeSpreadsheetValue(worksheet[cell.address]?.v)
    if (actual !== cell.expected) {
      throw new Error(
        `Same-corpus XLSX cell mismatch at ${fingerprint.sheetName}!${cell.address}: expected ${cell.expected}, got ${actual}`,
      )
    }
    return {
      address: cell.address,
      expected: cell.expected,
      actual,
    }
  })
  return {
    verified: true,
    method,
    sheetName: fingerprint.sheetName,
    materializedCells: fingerprint.materializedCells,
    checkedCells,
  }
}

export function buildSameCorpusFingerprint(corpus: WorkbookBenchmarkCorpusCase): SameCorpusFingerprint {
  const sheet = corpus.snapshot.sheets.find((candidate) => candidate.name === corpus.primaryViewport.sheetName)
  if (!sheet) {
    throw new Error(`Same-corpus snapshot is missing primary sheet: ${corpus.primaryViewport.sheetName}`)
  }
  const literalCells = sheet.cells
    .filter((cell) => cell.value !== undefined && cell.value !== null)
    .map((cell) => ({ address: cell.address, expected: normalizeSpreadsheetValue(cell.value) }))
  const checkedCells = selectFingerprintCells(literalCells)
  if (checkedCells.length < 3) {
    throw new Error(`Same-corpus fingerprint needs at least 3 literal cells for ${corpus.id}`)
  }
  return {
    sheetName: sheet.name,
    materializedCells: sheet.cells.length,
    checkedCells,
  }
}

function selectFingerprintCells(
  cells: readonly {
    readonly address: string
    readonly expected: string
  }[],
): readonly {
  readonly address: string
  readonly expected: string
}[] {
  const selected = new Map<string, { address: string; expected: string }>()
  for (const index of [0, 1, Math.floor(cells.length / 2), cells.length - 2, cells.length - 1]) {
    const cell = cells[index]
    if (cell) {
      selected.set(cell.address, cell)
    }
  }
  return [...selected.values()]
}

function normalizeSpreadsheetValue(value: unknown): string {
  if (value === null || value === undefined) {
    return ''
  }
  if (typeof value === 'string' || typeof value === 'number') {
    return String(value)
  }
  if (typeof value === 'boolean') {
    return value ? 'TRUE' : 'FALSE'
  }
  if (value instanceof Date) {
    return value.toISOString()
  }
  const serialized = JSON.stringify(value)
  if (serialized === undefined) {
    throw new Error(`Unable to normalize spreadsheet value of type ${typeof value}`)
  }
  return serialized
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
