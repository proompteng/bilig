import { mkdirSync } from 'node:fs'
import { dirname } from 'node:path'
import { performance } from 'node:perf_hooks'

import { chromium, type Browser, type BrowserContextOptions, type Page } from '@playwright/test'

import { buildWorkbookBenchmarkCorpus, type WorkbookBenchmarkCorpusCase } from '../packages/benchmarks/src/workbook-corpus.js'
import type {
  SameCorpusCapture,
  SameCorpusCaptureCase,
  SameCorpusCaptureCorpusVerification,
  SameCorpusCaptureMeasurement,
  UiResponsivenessSameCorpusProduct,
} from './gen-ui-responsiveness-live-browser-scorecard.ts'
import type { CaptureArgs, PreflightArgs, SaveStorageStateArgs } from './ui-responsiveness-same-corpus-args.ts'
import { defaultViewport } from './ui-responsiveness-same-corpus-args.ts'
import {
  requiredUiResponsivenessSameCorpusWorkloads,
  uiSameCorpusWorkloadRequiresScrollEventEvidence,
  type UiResponsivenessSameCorpusWorkload,
} from './ui-responsiveness-same-corpus-workloads.ts'
import { verifyProductCorpus, waitForVerifiedBiligRenderedSurface } from './ui-responsiveness-same-corpus-verification.ts'
import {
  buildCaptureScenarioProof,
  captureSameCorpusProductVisualProof,
  type SameCorpusProductVisualProof,
} from './ui-responsiveness-same-corpus-proof.ts'
import { productLimitations, sameCorpusChromiumLaunchOptions, settleFrames } from './ui-responsiveness-same-corpus-page-utils.ts'
import {
  measureVisibleScrollResponseWithRetries,
  movePointerToProductViewport,
  resetProductScrollPosition,
} from './ui-responsiveness-same-corpus-scroll-page.ts'
import {
  incumbentEditableWorkloadBlocker,
  measureProductWorkload,
  type ProductOperationSample,
} from './ui-responsiveness-same-corpus-workload-runner.ts'

interface ProductSampleCollection {
  readonly corpusVerification: SameCorpusCaptureCorpusVerification
  readonly samples: readonly ProductOperationSample[]
}

interface PreflightProductResult {
  readonly product: Exclude<UiResponsivenessSameCorpusProduct, 'bilig'>
  readonly source: string
  readonly finalUrl: string
  readonly title: string
  readonly corpusVerification: SameCorpusCaptureCorpusVerification
  readonly limitations: string[]
}

interface SameCorpusPreflight {
  readonly mode: 'preflight'
  readonly corpusCaseId: string
  readonly materializedCells: number
  readonly requiredProductCount: 2
  readonly checkedProductCount: number
  readonly products: readonly PreflightProductResult[]
}

interface SameCorpusProductMeasurementUrls {
  readonly biligUrl: string
  readonly googleSheetsUrl: string
  readonly microsoftExcelWebUrl: string | null
}

interface SameCorpusProductMeasurements {
  readonly bilig: SameCorpusCaptureMeasurement
  readonly googleSheets: SameCorpusCaptureMeasurement
  readonly microsoftExcelWeb?: SameCorpusCaptureMeasurement | undefined
}

type SameCorpusProductMeasure = (
  product: UiResponsivenessSameCorpusProduct,
  url: string,
  workload: UiResponsivenessSameCorpusWorkload,
) => Promise<SameCorpusCaptureMeasurement>

export async function saveStorageState(args: SaveStorageStateArgs): Promise<void> {
  const corpus = buildWorkbookBenchmarkCorpus(args.corpusId)
  const browser = await chromium.launch(sameCorpusChromiumLaunchOptions(args.headless))
  const context = await browser.newContext({ viewport: defaultViewport })
  const page = await context.newPage()
  try {
    await page.goto(args.authUrl, { waitUntil: 'domcontentloaded', timeout: 60_000 })
    await waitForProductReady(page, args.product, captureArgsForStorageState(args))
    const corpusVerification = await verifyProductCorpus(page, args.product, args.authUrl, corpus)
    mkdirSync(dirname(args.targetPath), { recursive: true })
    await context.storageState({ path: args.targetPath })
    console.log(
      JSON.stringify(
        {
          mode: 'save-storage-state',
          product: args.product,
          corpusCaseId: corpus.id,
          targetPath: args.targetPath,
          finalUrl: page.url(),
          title: await page.title(),
          corpusVerification,
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

export async function captureSameCorpusUiResponsiveness(args: CaptureArgs): Promise<SameCorpusCapture> {
  const corpus = buildWorkbookBenchmarkCorpus(args.corpusId)
  const browser = await chromium.launch(sameCorpusChromiumLaunchOptions(args.headless))
  try {
    const cases = await captureSameCorpusWorkloadCases(browser, corpus, args)
    return {
      schemaVersion: 1,
      suite: 'ui-responsiveness-same-corpus-capture',
      sampleCount: args.sampleCount,
      limitations: [
        'Caller must supply a Google Sheets URL for the same exported Bilig benchmark corpus.',
        'Microsoft Excel Web can be supplied as an additional incumbent comparison, but it is not required for the Google Sheets 10x claim.',
        'Edit and format workloads require the supplied incumbent URLs to allow browser-driven editing in the authenticated context.',
      ],
      cases,
    }
  } finally {
    await browser.close()
  }
}

async function captureSameCorpusWorkloadCases(
  browser: Browser,
  corpus: WorkbookBenchmarkCorpusCase,
  args: CaptureArgs,
  workloadIndex = 0,
  cases: SameCorpusCaptureCase[] = [],
): Promise<SameCorpusCaptureCase[]> {
  const workload = requiredUiResponsivenessSameCorpusWorkloads[workloadIndex]
  if (!workload) {
    return cases
  }
  const caseId = `same-corpus-${args.corpusId}-${workload}`
  const visualProofs: SameCorpusProductVisualProof[] = []
  const { bilig, googleSheets, microsoftExcelWeb } = await collectSameCorpusProductMeasurements(
    args,
    (product, url, measuredWorkload) => measureProduct(browser, product, url, corpus, args, measuredWorkload, caseId, visualProofs),
    workload,
  )
  const scenarioProof = buildCaptureScenarioProof({ bilig, googleSheets, microsoftExcelWeb, visualProofs })
  if (!scenarioProof.screenshotProof.captured || !scenarioProof.pixelGridProof.captured) {
    throw new Error(`same-corpus UI capture is missing browser-visible proof for ${caseId}: ${JSON.stringify(scenarioProof)}`)
  }
  cases.push({
    id: caseId,
    corpusCaseId: args.corpusId,
    materializedCells: corpus.materializedCellCount,
    workload,
    scenarioProof,
    bilig,
    googleSheets,
    ...(microsoftExcelWeb ? { microsoftExcelWeb } : {}),
  })
  return await captureSameCorpusWorkloadCases(browser, corpus, args, workloadIndex + 1, cases)
}

export async function collectSameCorpusProductMeasurements(
  urls: SameCorpusProductMeasurementUrls,
  measure: SameCorpusProductMeasure,
  workload: UiResponsivenessSameCorpusWorkload = 'scroll-vertical',
): Promise<SameCorpusProductMeasurements> {
  const bilig = await measure('bilig', urls.biligUrl, workload)
  assertSameCorpusProductMeasurement('bilig', urls.biligUrl, bilig, workload)
  const googleSheets = await measure('google-sheets', urls.googleSheetsUrl, workload)
  assertSameCorpusProductMeasurement('google-sheets', urls.googleSheetsUrl, googleSheets, workload)
  if (!urls.microsoftExcelWebUrl) {
    return { bilig, googleSheets }
  }
  const microsoftExcelWeb = await measure('microsoft-excel-web', urls.microsoftExcelWebUrl, workload)
  assertSameCorpusProductMeasurement('microsoft-excel-web', urls.microsoftExcelWebUrl, microsoftExcelWeb, workload)
  return { bilig, googleSheets, microsoftExcelWeb }
}

function assertSameCorpusProductMeasurement(
  product: UiResponsivenessSameCorpusProduct,
  source: string,
  measurement: SameCorpusCaptureMeasurement,
  workload: UiResponsivenessSameCorpusWorkload,
): void {
  if (measurement.product !== product) {
    throw new Error(`same-corpus UI measurement expected ${product} but received ${measurement.product}`)
  }
  if (measurement.source !== source) {
    throw new Error(`same-corpus UI measurement for ${product} used an unexpected source URL`)
  }
  assertSameCorpusSampleArray(product, 'operation response', measurement.operationResponseMsSamples)
  assertSameCorpusSampleArray(
    product,
    'post-operation frame',
    measurement.postOperationFrameMsSamples,
    measurement.operationResponseMsSamples.length,
  )
  if (uiSameCorpusWorkloadRequiresScrollEventEvidence(workload)) {
    assertSameCorpusSampleArray(
      product,
      'scroll-event response',
      measurement.scrollEventResponseMsSamples,
      measurement.operationResponseMsSamples.length,
    )
    assertSameCorpusSampleArray(
      product,
      'scroll movement',
      measurement.scrollMovementPxSamples,
      measurement.operationResponseMsSamples.length,
    )
  }
}

function assertSameCorpusSampleArray(
  product: UiResponsivenessSameCorpusProduct,
  label: string,
  samples: readonly number[] | undefined,
  expectedLength?: number,
): void {
  if (!samples || samples.length === 0) {
    throw new Error(`same-corpus UI measurement for ${product} is missing ${label} samples`)
  }
  if (expectedLength !== undefined && samples.length !== expectedLength) {
    throw new Error(
      `same-corpus UI measurement for ${product} has ${String(samples.length)} ${label} samples but expected ${String(expectedLength)}`,
    )
  }
  for (const sample of samples) {
    if (!Number.isFinite(sample)) {
      throw new Error(`same-corpus UI measurement for ${product} has a non-finite ${label} sample`)
    }
  }
}

export async function preflightSameCorpusIncumbentAccess(args: PreflightArgs): Promise<SameCorpusPreflight> {
  const corpus = buildWorkbookBenchmarkCorpus(args.corpusId)
  const browser = await chromium.launch(sameCorpusChromiumLaunchOptions(args.headless))
  try {
    const productSpecs = [
      ...(args.googleSheetsUrl ? [{ product: 'google-sheets' as const, url: args.googleSheetsUrl }] : []),
      ...(args.microsoftExcelWebUrl ? [{ product: 'microsoft-excel-web' as const, url: args.microsoftExcelWebUrl }] : []),
    ]
    const products = await Promise.all(productSpecs.map((spec) => preflightIncumbentProduct(browser, spec.product, spec.url, corpus, args)))
    return {
      mode: 'preflight',
      corpusCaseId: corpus.id,
      materializedCells: corpus.materializedCellCount,
      requiredProductCount: 2,
      checkedProductCount: products.length,
      products,
    }
  } finally {
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

async function preflightIncumbentProduct(
  browser: Browser,
  product: Exclude<UiResponsivenessSameCorpusProduct, 'bilig'>,
  url: string,
  corpus: WorkbookBenchmarkCorpusCase,
  args: PreflightArgs,
): Promise<PreflightProductResult> {
  const context = await browser.newContext(browserContextOptionsForPreflightProduct(product, args))
  const page = await context.newPage()
  try {
    await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 60_000 })
    await waitForProductReady(page, product, captureArgsForPreflight(args, product, url))
    await assertIncumbentEditableForPreflight(page, product)
    const corpusVerification = await verifyProductCorpus(page, product, url, corpus)
    return {
      product,
      source: url,
      finalUrl: page.url(),
      title: await page.title(),
      corpusVerification,
      limitations: productLimitations(product, storageStatePathForPreflightProduct(product, args)),
    }
  } catch (error: unknown) {
    throw new Error(await productReadyFailureMessage(page, product, url, 0, error), { cause: error })
  } finally {
    await context.close()
  }
}

async function assertIncumbentEditableForPreflight(
  page: Page,
  product: Exclude<UiResponsivenessSameCorpusProduct, 'bilig'>,
): Promise<void> {
  const bodyText = await page
    .locator('body')
    .innerText({ timeout: 2_000 })
    .catch(() => '')
  const blocker = incumbentEditableWorkloadBlocker(product, page.url(), bodyText)
  if (blocker) {
    throw new Error(`Cannot preflight same-corpus editable workloads on ${product}: ${blocker}`)
  }
}

function browserContextOptionsForPreflightProduct(
  product: Exclude<UiResponsivenessSameCorpusProduct, 'bilig'>,
  args: PreflightArgs,
): BrowserContextOptions {
  const storageState = storageStatePathForPreflightProduct(product, args)
  return {
    viewport: defaultViewport,
    ...(storageState ? { storageState } : {}),
  }
}

function storageStatePathForPreflightProduct(
  product: Exclude<UiResponsivenessSameCorpusProduct, 'bilig'>,
  args: PreflightArgs,
): string | null {
  if (product === 'google-sheets') {
    return args.googleSheetsStorageStatePath ?? args.storageStatePath
  }
  return args.microsoftExcelWebStorageStatePath ?? args.storageStatePath
}

function captureArgsForPreflight(
  args: PreflightArgs,
  product: Exclude<UiResponsivenessSameCorpusProduct, 'bilig'>,
  url: string,
): CaptureArgs {
  return {
    biligUrl: url,
    biligStorageStatePath: null,
    corpusId: args.corpusId,
    deltaX: 0,
    deltaY: 720,
    googleSheetsUrl: product === 'google-sheets' ? url : '',
    googleSheetsStorageStatePath: args.googleSheetsStorageStatePath,
    headless: args.headless,
    microsoftExcelWebUrl: product === 'microsoft-excel-web' ? url : '',
    microsoftExcelWebStorageStatePath: args.microsoftExcelWebStorageStatePath,
    outputPath: args.outputPath ?? '',
    readyTimeoutMs: args.readyTimeoutMs,
    sampleCount: 1,
    storageStatePath: args.storageStatePath,
  }
}

async function measureProduct(
  browser: Browser,
  product: UiResponsivenessSameCorpusProduct,
  url: string,
  corpus: WorkbookBenchmarkCorpusCase,
  args: CaptureArgs,
  workload: UiResponsivenessSameCorpusWorkload,
  caseId?: string,
  visualProofs?: SameCorpusProductVisualProof[],
): Promise<SameCorpusCaptureMeasurement> {
  const { corpusVerification, samples } = await measureProductSamples(browser, product, url, corpus, args, workload, caseId, visualProofs)

  return {
    product,
    source: url,
    operationResponseMsSamples: samples.map((entry) => entry.operationResponseMs),
    postOperationFrameMsSamples: samples.map((entry) => entry.postOperationFrameMs),
    ...(uiSameCorpusWorkloadRequiresScrollEventEvidence(workload)
      ? {
          scrollEventResponseMsSamples: samples.map((entry) => entry.scrollEventResponseMs ?? Number.NaN),
          scrollMovementPxSamples: samples.map((entry) => entry.scrollMovementPx ?? Number.NaN),
        }
      : {}),
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
  workload: UiResponsivenessSameCorpusWorkload,
  caseId: string | undefined = undefined,
  visualProofs: SameCorpusProductVisualProof[] | undefined = undefined,
  sampleIndex = 0,
  samples: ProductOperationSample[] = [],
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
    const loadStartedAt = performance.now()
    await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 60_000 })
    await waitForProductReady(page, product, args)
    const loadToReadyMs = performance.now() - loadStartedAt
    nextCorpusVerification ??= await verifyProductCorpus(page, product, url, corpus)
    if (product !== 'microsoft-excel-web' && uiSameCorpusWorkloadRequiresScrollEventEvidence(workload)) {
      await resetProductScrollPosition(page, product)
      await settleFrames(page, 3)
    }
    samples.push(
      await measureProductWorkload({
        page,
        product,
        captureArgs: args,
        workload,
        sampleIndex,
        loadToReadyMs,
        hooks: {
          measureVisibleScrollResponseWithRetries,
          movePointerToProductViewport,
        },
      }),
    )
    if (caseId && visualProofs && sampleIndex === 0) {
      visualProofs.push(await captureSameCorpusProductVisualProof({ caseId, outputPath: args.outputPath, page, product, sampleIndex }))
    }
  } catch (error: unknown) {
    throw new Error(await productReadyFailureMessage(page, product, url, sampleIndex, error), { cause: error })
  } finally {
    await context.close()
  }
  return measureProductSamples(
    browser,
    product,
    url,
    corpus,
    args,
    workload,
    caseId,
    visualProofs,
    sampleIndex + 1,
    samples,
    nextCorpusVerification,
  )
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
    await settleFrames(page, 180)
    await waitForVerifiedBiligRenderedSurface(page, args.readyTimeoutMs)
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
  await waitForFrameElementReady(page, '.ewr-grdcontarea-grid', args.readyTimeoutMs)
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

async function waitForFrameElementReady(page: Page, selector: string, timeoutMs: number, startedAt = performance.now()): Promise<void> {
  const box = await firstFrameElementBox(page, selector)
  if (box && box.width > 0 && box.height > 0) {
    return
  }
  if (performance.now() - startedAt >= timeoutMs) {
    throw new Error(`Timed out waiting for frame element ${selector}`)
  }
  await page.waitForTimeout(250)
  return waitForFrameElementReady(page, selector, timeoutMs, startedAt)
}

async function firstFrameElementBox(page: Page, selector: string): Promise<{ x: number; y: number; width: number; height: number } | null> {
  const boxes = await Promise.all(
    page.frames().map(async (frame) => {
      const locator = frame.locator(selector).first()
      const count = await locator.count().catch(() => 0)
      if (count === 0) {
        return null
      }
      const box = await locator.boundingBox().catch(() => null)
      if (box && box.width > 0 && box.height > 0) {
        return box
      }
      return null
    }),
  )
  return boxes.find((box): box is { x: number; y: number; width: number; height: number } => box !== null) ?? null
}
