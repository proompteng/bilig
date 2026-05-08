import { mkdirSync } from 'node:fs'
import { dirname } from 'node:path'
import { performance } from 'node:perf_hooks'

import { chromium, type Browser, type BrowserContextOptions, type Frame, type Page } from '@playwright/test'

import { buildWorkbookBenchmarkCorpus, type WorkbookBenchmarkCorpusCase } from '../packages/benchmarks/src/workbook-corpus.js'
import type {
  SameCorpusCapture,
  SameCorpusCaptureCorpusVerification,
  SameCorpusCaptureMeasurement,
  UiResponsivenessSameCorpusProduct,
} from './gen-ui-responsiveness-live-browser-scorecard.ts'
import type { CaptureArgs, PreflightArgs, SaveStorageStateArgs } from './ui-responsiveness-same-corpus-args.ts'
import { defaultViewport } from './ui-responsiveness-same-corpus-args.ts'
import {
  measureVisibleScrollResponseWithHooks,
  ScrollMovementVerificationError,
  type ScrollPositionSnapshot,
  type ScrollSample,
  type ScrollTriggerResult,
} from './ui-responsiveness-same-corpus-scroll.ts'
import { verifyProductCorpus, waitForVerifiedBiligRenderedSurface } from './ui-responsiveness-same-corpus-verification.ts'

interface ProductSampleCollection {
  readonly corpusVerification: SameCorpusCaptureCorpusVerification
  readonly samples: readonly ScrollSample[]
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
  readonly microsoftExcelWebUrl: string
}

interface SameCorpusProductMeasurements {
  readonly bilig: SameCorpusCaptureMeasurement
  readonly googleSheets: SameCorpusCaptureMeasurement
  readonly microsoftExcelWeb: SameCorpusCaptureMeasurement
}

type SameCorpusProductMeasure = (product: UiResponsivenessSameCorpusProduct, url: string) => Promise<SameCorpusCaptureMeasurement>
type ScrollEventResponseProbeContext = Page | Frame

const scrollEventResponseProbeTimeoutMs = 5_000

export async function saveStorageState(args: SaveStorageStateArgs): Promise<void> {
  const corpus = buildWorkbookBenchmarkCorpus(args.corpusId)
  const browser = await chromium.launch({ headless: args.headless })
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
  const browser = await chromium.launch({ headless: args.headless })
  try {
    const { bilig, googleSheets, microsoftExcelWeb } = await collectSameCorpusProductMeasurements(args, (product, url) =>
      measureProduct(browser, product, url, corpus, args),
    )
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

export async function collectSameCorpusProductMeasurements(
  urls: SameCorpusProductMeasurementUrls,
  measure: SameCorpusProductMeasure,
): Promise<SameCorpusProductMeasurements> {
  const bilig = await measure('bilig', urls.biligUrl)
  assertSameCorpusProductMeasurement('bilig', urls.biligUrl, bilig)
  const googleSheets = await measure('google-sheets', urls.googleSheetsUrl)
  assertSameCorpusProductMeasurement('google-sheets', urls.googleSheetsUrl, googleSheets)
  const microsoftExcelWeb = await measure('microsoft-excel-web', urls.microsoftExcelWebUrl)
  assertSameCorpusProductMeasurement('microsoft-excel-web', urls.microsoftExcelWebUrl, microsoftExcelWeb)
  return { bilig, googleSheets, microsoftExcelWeb }
}

function assertSameCorpusProductMeasurement(
  product: UiResponsivenessSameCorpusProduct,
  source: string,
  measurement: SameCorpusCaptureMeasurement,
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
  const browser = await chromium.launch({ headless: args.headless })
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
): Promise<SameCorpusCaptureMeasurement> {
  const { corpusVerification, samples } = await measureProductSamples(browser, product, url, corpus, args)

  return {
    product,
    source: url,
    operationResponseMsSamples: samples.map((entry) => entry.operationResponseMs),
    postOperationFrameMsSamples: samples.map((entry) => entry.postOperationFrameMs),
    scrollEventResponseMsSamples: samples.map((entry) => entry.scrollEventResponseMs),
    scrollMovementPxSamples: samples.map((entry) => entry.scrollMovementPx),
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
    if (product !== 'microsoft-excel-web') {
      await resetProductScrollPosition(page, product)
      await settleFrames(page, 3)
    }
    samples.push(await measureVisibleScrollResponseWithRetries(page, product, args.deltaX, args.deltaY))
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

async function measureVisibleScrollResponseWithRetries(
  page: Page,
  product: UiResponsivenessSameCorpusProduct,
  deltaX: number,
  deltaY: number,
  attempt = 1,
): Promise<ScrollSample> {
  try {
    return await measureVisibleScrollResponse(page, product, deltaX, deltaY)
  } catch (error) {
    if (!(error instanceof ScrollMovementVerificationError) || attempt >= 3) {
      throw error
    }
    await settleFrames(page, 30)
    return measureVisibleScrollResponseWithRetries(page, product, deltaX, deltaY, attempt + 1)
  }
}

async function measureVisibleScrollResponse(
  page: Page,
  product: UiResponsivenessSameCorpusProduct,
  deltaX: number,
  deltaY: number,
): Promise<ScrollSample> {
  return await measureVisibleScrollResponseWithHooks({
    collectFrameIntervals: (frameCount) => collectFrameIntervals(page, frameCount),
    movePointer: async () => {
      await movePointerToProductViewport(page, product)
      if (product === 'microsoft-excel-web') {
        await resetProductScrollPosition(page, product)
        await movePointerToProductViewport(page, product)
      }
    },
    now: () => performance.now(),
    readScrollPosition: () => readProductScrollPosition(page, product),
    scroll: () => triggerProductScrollAndMeasureEvent(page, product, deltaX, deltaY),
    waitForNextFrame: () => waitForNextFrame(page),
  })
}

async function triggerProductScrollAndMeasureEvent(
  page: Page,
  product: UiResponsivenessSameCorpusProduct,
  deltaX: number,
  deltaY: number,
): Promise<ScrollTriggerResult> {
  const probeContext = await installProductScrollEventResponseProbe(page, product)
  try {
    await page.mouse.wheel(deltaX, deltaY)
    const scrollEventResponseMs = await readScrollEventResponseProbe(probeContext)
    if (!Number.isFinite(scrollEventResponseMs)) {
      throw new Error(`Measured ${product} scroll-event response was not finite: ${String(scrollEventResponseMs)}`)
    }
    return { scrollEventResponseMs }
  } finally {
    await clearScrollEventResponseProbe(probeContext)
  }
}

async function installProductScrollEventResponseProbe(
  page: Page,
  product: UiResponsivenessSameCorpusProduct,
): Promise<ScrollEventResponseProbeContext> {
  if (product === 'bilig') {
    await installScrollEventResponseProbe(page, product, ['[data-testid="grid-scroll-viewport"]'])
    return page
  }
  if (product === 'google-sheets') {
    await installScrollEventResponseProbe(page, product, ['.native-scrollbar-y', '.native-scrollbar-x', '.grid-scrollable-wrapper'])
    return page
  }
  const frame = await firstFrameWithElement(page, '.ewr-grdcontarea-grid')
  if (!frame) {
    throw new Error('Unable to locate Microsoft Excel Web grid frame for scroll-event response probe')
  }
  await installScrollEventResponseProbe(frame, product, ['.ewr-grdcontarea-grid'])
  return frame
}

async function installScrollEventResponseProbe(
  context: ScrollEventResponseProbeContext,
  product: UiResponsivenessSameCorpusProduct,
  scrollSelectors: readonly string[],
): Promise<void> {
  const installed = await context.evaluate(
    ({ selectors, timeoutMs }) => {
      interface SameCorpusScrollProbeHost extends Window {
        __biligSameCorpusScrollEventProbe?: {
          readonly result: Promise<number>
          readonly cleanup: () => void
        }
      }

      const host = window as SameCorpusScrollProbeHost
      host.__biligSameCorpusScrollEventProbe?.cleanup()
      delete host.__biligSameCorpusScrollEventProbe

      const elementSet = new Set<HTMLElement>()
      for (const selector of selectors) {
        for (const element of document.querySelectorAll(selector)) {
          if (element instanceof HTMLElement) {
            elementSet.add(element)
          }
        }
      }
      const scrollElements = [...elementSet]
      if (scrollElements.length === 0) {
        delete host.__biligSameCorpusScrollEventProbe
        return false
      }

      const cleanupCallbacks: Array<() => void> = []
      const cleanup = (): void => {
        for (const callback of cleanupCallbacks.splice(0)) {
          callback()
        }
      }
      const result = new Promise<number>((resolve, reject) => {
        let wheelAt: number | null = null
        let finished = false
        const onWheel = (): void => {
          wheelAt ??= performance.now()
        }
        const finish = (settle: (value: number) => void, value: number): void => {
          if (finished) {
            return
          }
          finished = true
          cleanup()
          settle(value)
        }
        const fail = (error: Error): void => {
          if (finished) {
            return
          }
          finished = true
          cleanup()
          reject(error)
        }
        const onScroll = (): void => {
          if (wheelAt === null) {
            return
          }
          finish(resolve, performance.now() - wheelAt)
        }
        const timeoutId = window.setTimeout(
          () => fail(new Error(`Timed out waiting for workbook scroll event after wheel event (${String(timeoutMs)}ms)`)),
          timeoutMs,
        )
        cleanupCallbacks.push(() => {
          window.clearTimeout(timeoutId)
          window.removeEventListener('wheel', onWheel, true)
          for (const element of scrollElements) {
            element.removeEventListener('scroll', onScroll)
          }
        })

        window.addEventListener('wheel', onWheel, { capture: true, passive: true })
        for (const element of scrollElements) {
          element.addEventListener('scroll', onScroll, { passive: true })
        }
      })

      host.__biligSameCorpusScrollEventProbe = { result, cleanup }
      return true
    },
    { selectors: [...scrollSelectors], timeoutMs: scrollEventResponseProbeTimeoutMs },
  )
  if (!installed) {
    throw new Error(`Unable to install ${product} scroll-event response probe; no workbook scroll target matched`)
  }
}

async function readScrollEventResponseProbe(context: ScrollEventResponseProbeContext): Promise<number> {
  return await context.evaluate(async () => {
    interface SameCorpusScrollProbeHost extends Window {
      __biligSameCorpusScrollEventProbe?: {
        readonly result: Promise<number>
      }
    }

    const probe = (window as SameCorpusScrollProbeHost).__biligSameCorpusScrollEventProbe
    if (!probe) {
      throw new Error('Workbook scroll-event response probe was not installed')
    }
    return await probe.result
  })
}

async function clearScrollEventResponseProbe(context: ScrollEventResponseProbeContext): Promise<void> {
  await context
    .evaluate(() => {
      interface SameCorpusScrollProbeHost extends Window {
        __biligSameCorpusScrollEventProbe?: {
          readonly cleanup: () => void
        }
      }

      const host = window as SameCorpusScrollProbeHost
      host.__biligSameCorpusScrollEventProbe?.cleanup()
      delete host.__biligSameCorpusScrollEventProbe
    })
    .catch(() => undefined)
}

async function movePointerToProductViewport(page: Page, product: UiResponsivenessSameCorpusProduct): Promise<void> {
  const box = await productViewportBox(page, product)
  if (box) {
    await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2)
    await page.mouse.click(box.x + box.width / 2, box.y + box.height / 2)
    return
  }
  await page.mouse.move(defaultViewport.width / 2, defaultViewport.height / 2)
}

async function productViewportBox(
  page: Page,
  product: UiResponsivenessSameCorpusProduct,
): Promise<{ x: number; y: number; width: number; height: number } | null> {
  if (product === 'bilig') {
    return await page
      .getByTestId('grid-scroll-viewport')
      .boundingBox()
      .catch(() => null)
  }
  if (product === 'google-sheets') {
    return await page
      .locator('.grid-scrollable-wrapper')
      .first()
      .boundingBox()
      .catch(() => null)
  }
  return await firstFrameElementBox(page, '.ewr-grdcontarea-grid')
}

async function readProductScrollPosition(page: Page, product: UiResponsivenessSameCorpusProduct): Promise<ScrollPositionSnapshot | null> {
  if (product === 'bilig') {
    return await page
      .getByTestId('grid-scroll-viewport')
      .evaluate((element) => ({
        scrollLeft: element.scrollLeft,
        scrollTop: element.scrollTop,
      }))
      .catch(() => null)
  }
  if (product === 'google-sheets') {
    return await page.evaluate(() => {
      const verticalScrollbar = document.querySelector<HTMLElement>('.native-scrollbar-y')
      const horizontalScrollbar = document.querySelector<HTMLElement>('.native-scrollbar-x')
      if (!verticalScrollbar && !horizontalScrollbar) {
        return null
      }
      return {
        scrollLeft: horizontalScrollbar?.scrollLeft ?? 0,
        scrollTop: verticalScrollbar?.scrollTop ?? 0,
      }
    })
  }
  const positions = await Promise.all(
    page.frames().map((frame) =>
      frame
        .evaluate(() => {
          const grid = document.querySelector('.ewr-grdcontarea-grid')
          if (!(grid instanceof HTMLElement)) {
            return null
          }
          return {
            scrollLeft: grid.scrollLeft,
            scrollTop: grid.scrollTop,
          }
        })
        .catch(() => null),
    ),
  )
  return positions.find((position): position is ScrollPositionSnapshot => position !== null) ?? null
}

async function resetProductScrollPosition(page: Page, product: UiResponsivenessSameCorpusProduct): Promise<void> {
  const directReset = await setProductScrollPosition(page, product, { scrollLeft: 0, scrollTop: 0 })
  if (directReset && (await waitForProductScrollPosition(page, product, { scrollLeft: 0, scrollTop: 0 }, 750))) {
    return
  }
  if (await resetProductScrollPositionWithWheel(page, product)) {
    return
  }
  const actual = await readProductScrollPosition(page, product)
  throw new Error(`Unable to reset ${product} workbook viewport before same-corpus scroll timing: ${JSON.stringify(actual)}`)
}

async function setProductScrollPosition(
  page: Page,
  product: UiResponsivenessSameCorpusProduct,
  position: ScrollPositionSnapshot,
): Promise<boolean> {
  if (product === 'bilig') {
    return await page
      .getByTestId('grid-scroll-viewport')
      .evaluate((element, target) => {
        element.scrollLeft = target.scrollLeft
        element.scrollTop = target.scrollTop
        return true
      }, position)
      .catch(() => false)
  }
  if (product === 'google-sheets') {
    return await page.evaluate((target) => {
      const verticalScrollbar = document.querySelector<HTMLElement>('.native-scrollbar-y')
      const horizontalScrollbar = document.querySelector<HTMLElement>('.native-scrollbar-x')
      if (!verticalScrollbar && !horizontalScrollbar) {
        return false
      }
      if (verticalScrollbar) {
        verticalScrollbar.scrollTop = target.scrollTop
      }
      if (horizontalScrollbar) {
        horizontalScrollbar.scrollLeft = target.scrollLeft
      }
      return true
    }, position)
  }
  const results = await Promise.all(
    page.frames().map((frame) =>
      frame
        .evaluate((target) => {
          const grid = document.querySelector('.ewr-grdcontarea-grid')
          if (!(grid instanceof HTMLElement)) {
            return false
          }
          grid.scrollLeft = target.scrollLeft
          grid.scrollTop = target.scrollTop
          grid.scrollTo(target.scrollLeft, target.scrollTop)
          grid.dispatchEvent(new Event('scroll', { bubbles: true }))
          return true
        }, position)
        .catch(() => false),
    ),
  )
  return results.some(Boolean)
}

async function resetProductScrollPositionWithWheel(page: Page, product: UiResponsivenessSameCorpusProduct, attempt = 1): Promise<boolean> {
  if (attempt > 5) {
    return false
  }
  await page.mouse.wheel(0, -10_000)
  await waitForNextFrame(page)
  const position = await readProductScrollPosition(page, product)
  if (position && position.scrollLeft === 0 && position.scrollTop === 0) {
    return true
  }
  return resetProductScrollPositionWithWheel(page, product, attempt + 1)
}

async function waitForProductScrollPosition(
  page: Page,
  product: UiResponsivenessSameCorpusProduct,
  expected: ScrollPositionSnapshot,
  timeoutMs: number,
  startedAt = performance.now(),
): Promise<boolean> {
  const current = await readProductScrollPosition(page, product)
  if (current && Math.abs(current.scrollLeft - expected.scrollLeft) <= 1 && Math.abs(current.scrollTop - expected.scrollTop) <= 1) {
    return true
  }
  if (performance.now() - startedAt >= timeoutMs) {
    return false
  }
  await page.waitForTimeout(50)
  return waitForProductScrollPosition(page, product, expected, timeoutMs, startedAt)
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

async function firstFrameWithElement(page: Page, selector: string): Promise<Frame | null> {
  const frames = await Promise.all(
    page.frames().map(async (frame) => {
      const count = await frame
        .locator(selector)
        .first()
        .count()
        .catch(() => 0)
      return count > 0 ? frame : null
    }),
  )
  return frames.find((frame): frame is Frame => frame !== null) ?? null
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

async function waitForNextFrame(page: Page): Promise<void> {
  await page.evaluate(async () => {
    await new Promise<void>((resolveFrame) => requestAnimationFrame(() => resolveFrame()))
  })
}

async function collectFrameIntervals(page: Page, frameCount: number): Promise<number[]> {
  const frameIntervals = await page.evaluate(async (targetFrameCount) => {
    const intervals: number[] = []
    let previous = performance.now()
    await new Promise<void>((finish) => {
      const step = (now: number): void => {
        intervals.push(now - previous)
        previous = now
        if (intervals.length >= targetFrameCount) {
          finish()
          return
        }
        requestAnimationFrame(step)
      }
      requestAnimationFrame(step)
    })
    return intervals
  }, frameCount)
  return frameIntervals
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
