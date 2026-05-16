import { mkdirSync } from 'node:fs'
import { dirname, relative, resolve } from 'node:path'

import type { Frame, Page } from '@playwright/test'

import { summarizeNumbers, type NumericSummary } from '../packages/benchmarks/src/stats.js'
import type {
  SameCorpusCaptureMeasurement,
  UiResponsivenessSameCorpusMeasurement,
  UiResponsivenessSameCorpusProduct,
} from './ui-responsiveness-same-corpus-scorecard-proof.ts'

const rootDir = resolve(new URL('..', import.meta.url).pathname)

export interface SameCorpusScenarioProof {
  readonly biligMeanMs: number
  readonly biligP95Ms: number
  readonly googleMeanMs: number
  readonly googleP95Ms: number
  readonly microsoftExcelWebMeanMs?: number
  readonly microsoftExcelWebP95Ms?: number
  readonly meanRatio: number
  readonly p95Ratio: number
  readonly microsoftExcelWebMeanRatio?: number
  readonly microsoftExcelWebP95Ratio?: number
  readonly screenshotProof: SameCorpusScreenshotProof
  readonly pixelGridProof: SameCorpusPixelGridProof
}

export interface SameCorpusScreenshotProof {
  readonly captured: boolean
  readonly requiredProducts: readonly UiResponsivenessSameCorpusProduct[]
  readonly artifactPaths: readonly string[]
  readonly missingProducts: readonly UiResponsivenessSameCorpusProduct[]
}

export interface SameCorpusPixelGridProof {
  readonly captured: boolean
  readonly requiredProducts: readonly UiResponsivenessSameCorpusProduct[]
  readonly products: readonly SameCorpusProductPixelGridProof[]
  readonly missingProducts: readonly UiResponsivenessSameCorpusProduct[]
}

export interface SameCorpusProductVisualProof {
  readonly product: UiResponsivenessSameCorpusProduct
  readonly screenshotPath: string | null
  readonly screenshotCaptured: boolean
  readonly pixelGridProof: SameCorpusProductPixelGridProof
}

export interface SameCorpusProductPixelGridProof {
  readonly product: UiResponsivenessSameCorpusProduct
  readonly captured: boolean
  readonly method: 'typegpu-visible-canvas' | 'google-sheets-visible-grid' | 'excel-web-visible-grid'
  readonly viewportPixelWidth: number
  readonly viewportPixelHeight: number
  readonly evidence: readonly string[]
}

export function buildCaptureScenarioProof(args: {
  readonly bilig: SameCorpusCaptureMeasurement
  readonly googleSheets: SameCorpusCaptureMeasurement
  readonly microsoftExcelWeb?: SameCorpusCaptureMeasurement | undefined
  readonly visualProofs: readonly SameCorpusProductVisualProof[]
}): SameCorpusScenarioProof {
  return buildScenarioProof({
    biligTiming: summarizeNumbers(primaryCaptureTimingSamples(args.bilig)),
    googleSheetsTiming: summarizeNumbers(primaryCaptureTimingSamples(args.googleSheets)),
    microsoftExcelWebTiming: args.microsoftExcelWeb ? summarizeNumbers(primaryCaptureTimingSamples(args.microsoftExcelWeb)) : null,
    visualProofs: args.visualProofs,
  })
}

export function buildScorecardScenarioProof(args: {
  readonly bilig: UiResponsivenessSameCorpusMeasurement
  readonly googleSheets: UiResponsivenessSameCorpusMeasurement
  readonly microsoftExcelWeb?: UiResponsivenessSameCorpusMeasurement | undefined
  readonly visualProofs: readonly SameCorpusProductVisualProof[]
}): SameCorpusScenarioProof {
  return buildScenarioProof({
    biligTiming: primaryScorecardTiming(args.bilig),
    googleSheetsTiming: primaryScorecardTiming(args.googleSheets),
    microsoftExcelWebTiming: args.microsoftExcelWeb ? primaryScorecardTiming(args.microsoftExcelWeb) : null,
    visualProofs: args.visualProofs,
  })
}

export async function captureSameCorpusProductVisualProof(args: {
  readonly caseId: string
  readonly outputPath: string
  readonly page: Page
  readonly product: UiResponsivenessSameCorpusProduct
  readonly sampleIndex: number
}): Promise<SameCorpusProductVisualProof> {
  const screenshotPath = screenshotArtifactPath(args.outputPath, args.caseId, args.product, args.sampleIndex)
  mkdirSync(dirname(screenshotPath), { recursive: true })
  const screenshotCaptured = await captureProductScreenshot(args.page, args.product, screenshotPath)
  const pixelGridProof = await readProductPixelGridProof(args.page, args.product)
  return {
    product: args.product,
    screenshotPath: screenshotCaptured ? repoRelativePath(screenshotPath) : null,
    screenshotCaptured,
    pixelGridProof,
  }
}

export function validateSameCorpusScenarioProof(
  proof: SameCorpusScenarioProof,
  caseId: string,
  bilig: UiResponsivenessSameCorpusMeasurement,
  googleSheets: UiResponsivenessSameCorpusMeasurement,
  microsoftExcelWeb?: UiResponsivenessSameCorpusMeasurement,
): void {
  const expected = buildScorecardScenarioProof({
    bilig,
    googleSheets,
    microsoftExcelWeb,
    visualProofs: proof.pixelGridProof.products.map((entry) => ({
      product: entry.product,
      screenshotPath: proof.screenshotProof.artifactPaths.find((artifact) => artifact.includes(`${entry.product}-`)) ?? null,
      screenshotCaptured: !proof.screenshotProof.missingProducts.includes(entry.product),
      pixelGridProof: entry,
    })),
  })
  if (
    proof.biligMeanMs !== expected.biligMeanMs ||
    proof.biligP95Ms !== expected.biligP95Ms ||
    proof.googleMeanMs !== expected.googleMeanMs ||
    proof.googleP95Ms !== expected.googleP95Ms ||
    proof.microsoftExcelWebMeanMs !== expected.microsoftExcelWebMeanMs ||
    proof.microsoftExcelWebP95Ms !== expected.microsoftExcelWebP95Ms ||
    proof.meanRatio !== expected.meanRatio ||
    proof.p95Ratio !== expected.p95Ratio ||
    proof.microsoftExcelWebMeanRatio !== expected.microsoftExcelWebMeanRatio ||
    proof.microsoftExcelWebP95Ratio !== expected.microsoftExcelWebP95Ratio
  ) {
    throw new Error(`UI responsiveness same-corpus scenario proof timing is stale: ${caseId}`)
  }
  if (
    proof.screenshotProof.captured !== expected.screenshotProof.captured ||
    JSON.stringify(proof.screenshotProof.requiredProducts) !== JSON.stringify(expected.screenshotProof.requiredProducts) ||
    JSON.stringify(proof.screenshotProof.missingProducts) !== JSON.stringify(expected.screenshotProof.missingProducts)
  ) {
    throw new Error(`UI responsiveness same-corpus screenshot proof is stale: ${caseId}`)
  }
  if (
    proof.pixelGridProof.captured !== expected.pixelGridProof.captured ||
    JSON.stringify(proof.pixelGridProof.requiredProducts) !== JSON.stringify(expected.pixelGridProof.requiredProducts) ||
    JSON.stringify(proof.pixelGridProof.missingProducts) !== JSON.stringify(expected.pixelGridProof.missingProducts)
  ) {
    throw new Error(`UI responsiveness same-corpus pixel grid proof is stale: ${caseId}`)
  }
  if (!proof.screenshotProof.captured) {
    throw new Error(`UI responsiveness same-corpus scenario proof is missing screenshot proof: ${caseId}`)
  }
  if (!proof.pixelGridProof.captured) {
    throw new Error(`UI responsiveness same-corpus scenario proof is missing pixel grid proof: ${caseId}`)
  }
}

function buildScenarioProof(args: {
  readonly biligTiming: NumericSummary
  readonly googleSheetsTiming: NumericSummary
  readonly microsoftExcelWebTiming: NumericSummary | null
  readonly visualProofs: readonly SameCorpusProductVisualProof[]
}): SameCorpusScenarioProof {
  const requiredProducts = ['bilig', 'google-sheets'] as const satisfies readonly UiResponsivenessSameCorpusProduct[]
  const screenshotProducts = new Set(
    args.visualProofs.filter((entry) => entry.screenshotCaptured && entry.screenshotPath).map((entry) => entry.product),
  )
  const pixelProducts = new Set(args.visualProofs.filter((entry) => entry.pixelGridProof.captured).map((entry) => entry.product))
  return {
    biligMeanMs: args.biligTiming.mean,
    biligP95Ms: args.biligTiming.p95,
    googleMeanMs: args.googleSheetsTiming.mean,
    googleP95Ms: args.googleSheetsTiming.p95,
    ...(args.microsoftExcelWebTiming
      ? {
          microsoftExcelWebMeanMs: args.microsoftExcelWebTiming.mean,
          microsoftExcelWebP95Ms: args.microsoftExcelWebTiming.p95,
          microsoftExcelWebMeanRatio: ratio(args.biligTiming.mean, args.microsoftExcelWebTiming.mean),
          microsoftExcelWebP95Ratio: ratio(args.biligTiming.p95, args.microsoftExcelWebTiming.p95),
        }
      : {}),
    meanRatio: ratio(args.biligTiming.mean, args.googleSheetsTiming.mean),
    p95Ratio: ratio(args.biligTiming.p95, args.googleSheetsTiming.p95),
    screenshotProof: {
      captured: requiredProducts.every((product) => screenshotProducts.has(product)),
      requiredProducts,
      artifactPaths: args.visualProofs.flatMap((entry) => (entry.screenshotPath ? [entry.screenshotPath] : [])),
      missingProducts: requiredProducts.filter((product) => !screenshotProducts.has(product)),
    },
    pixelGridProof: {
      captured: requiredProducts.every((product) => pixelProducts.has(product)),
      requiredProducts,
      products: args.visualProofs.map((entry) => entry.pixelGridProof),
      missingProducts: requiredProducts.filter((product) => !pixelProducts.has(product)),
    },
  }
}

function primaryCaptureTimingSamples(measurement: SameCorpusCaptureMeasurement): readonly number[] {
  return measurement.scrollEventResponseMsSamples ?? measurement.operationResponseMsSamples
}

function primaryScorecardTiming(measurement: UiResponsivenessSameCorpusMeasurement): NumericSummary {
  return measurement.scrollEventResponseMs ?? measurement.operationResponseMs
}

function ratio(numerator: number, denominator: number): number {
  if (denominator <= 0) {
    return Number.POSITIVE_INFINITY
  }
  return numerator / denominator
}

function screenshotArtifactPath(
  outputPath: string,
  caseId: string,
  product: UiResponsivenessSameCorpusProduct,
  sampleIndex: number,
): string {
  return resolve(`${outputPath}.proof`, caseId, `${product}-sample-${String(sampleIndex + 1)}.png`)
}

function repoRelativePath(path: string): string {
  return relative(rootDir, path)
}

async function captureProductScreenshot(page: Page, product: UiResponsivenessSameCorpusProduct, path: string): Promise<boolean> {
  const selector =
    product === 'bilig' ? '[data-testid="sheet-grid"]' : product === 'google-sheets' ? '.grid-scrollable-wrapper' : '.ewr-grdcontarea-grid'
  if (product === 'microsoft-excel-web') {
    const candidates = await Promise.all(
      page.frames().map(async (frame) => {
        const locator = frame.locator(selector).first()
        return { locator, count: await locator.count().catch(() => 0) }
      }),
    )
    const candidate = candidates.find((entry) => entry.count > 0)
    if (candidate) {
      await candidate.locator.screenshot({ path })
      return true
    }
    return false
  }
  const locator = page.locator(selector).first()
  if ((await locator.count().catch(() => 0)) === 0) {
    return false
  }
  await locator.screenshot({ path })
  return true
}

async function readProductPixelGridProof(page: Page, product: UiResponsivenessSameCorpusProduct): Promise<SameCorpusProductPixelGridProof> {
  if (product === 'bilig') {
    return await readBiligPixelGridProof(page)
  }
  if (product === 'google-sheets') {
    return await readDomPixelGridProof(page, product, 'google-sheets-visible-grid', '.grid-scrollable-wrapper')
  }
  const proofs = await Promise.all(
    page
      .frames()
      .map((frame) => readDomPixelGridProof(frame, product, 'excel-web-visible-grid', '.ewr-grdcontarea-grid').catch(() => null)),
  )
  const proof = proofs.find((entry): entry is SameCorpusProductPixelGridProof => Boolean(entry?.captured))
  if (proof) {
    return proof
  }
  return emptyPixelGridProof(product, 'excel-web-visible-grid')
}

async function readBiligPixelGridProof(page: Page): Promise<SameCorpusProductPixelGridProof> {
  return await page.evaluate(() => {
    const grid = document.querySelector('[data-testid="sheet-grid"]')
    const canvas = document.querySelector('[data-testid="grid-pane-renderer"]')
    if (!(grid instanceof HTMLElement) || !(canvas instanceof HTMLCanvasElement)) {
      return {
        product: 'bilig' as const,
        captured: false,
        method: 'typegpu-visible-canvas' as const,
        viewportPixelWidth: 0,
        viewportPixelHeight: 0,
        evidence: ['missing sheet grid or TypeGPU canvas'],
      }
    }
    const dpr = Math.max(1, window.devicePixelRatio || 1)
    const expectedWidth = Math.max(1, Math.floor(grid.clientWidth * dpr))
    const expectedHeight = Math.max(1, Math.floor(grid.clientHeight * dpr))
    const tilePaneCount = Number.parseInt(canvas.getAttribute('data-v3-tile-pane-count') ?? '0', 10) || 0
    const headerPaneCount = Number.parseInt(canvas.getAttribute('data-v3-header-pane-count') ?? '0', 10) || 0
    const mode = canvas.getAttribute('data-renderer-mode') ?? ''
    const captured =
      mode === 'typegpu-v3' &&
      tilePaneCount > 0 &&
      headerPaneCount > 0 &&
      canvas.width >= expectedWidth - 2 &&
      canvas.height >= expectedHeight - 2
    return {
      product: 'bilig' as const,
      captured,
      method: 'typegpu-visible-canvas' as const,
      viewportPixelWidth: canvas.width,
      viewportPixelHeight: canvas.height,
      evidence: [
        `mode=${mode}`,
        `tilePaneCount=${String(tilePaneCount)}`,
        `headerPaneCount=${String(headerPaneCount)}`,
        `expectedPixelWidth=${String(expectedWidth)}`,
        `expectedPixelHeight=${String(expectedHeight)}`,
      ],
    }
  })
}

async function readDomPixelGridProof(
  context: Frame | Page,
  product: UiResponsivenessSameCorpusProduct,
  method: SameCorpusProductPixelGridProof['method'],
  selector: string,
): Promise<SameCorpusProductPixelGridProof> {
  return await context.evaluate(
    ({ method: proofMethod, product: proofProduct, selector: proofSelector }) => {
      const element = document.querySelector(proofSelector)
      if (!(element instanceof HTMLElement)) {
        return {
          product: proofProduct,
          captured: false,
          method: proofMethod,
          viewportPixelWidth: 0,
          viewportPixelHeight: 0,
          evidence: [`missing selector ${proofSelector}`],
        }
      }
      const rect = element.getBoundingClientRect()
      const dpr = Math.max(1, window.devicePixelRatio || 1)
      const width = Math.floor(rect.width * dpr)
      const height = Math.floor(rect.height * dpr)
      return {
        product: proofProduct,
        captured: width > 128 && height > 128,
        method: proofMethod,
        viewportPixelWidth: width,
        viewportPixelHeight: height,
        evidence: [`selector=${proofSelector}`, `cssWidth=${String(rect.width)}`, `cssHeight=${String(rect.height)}`],
      }
    },
    { method, product, selector },
  )
}

function emptyPixelGridProof(
  product: UiResponsivenessSameCorpusProduct,
  method: SameCorpusProductPixelGridProof['method'],
): SameCorpusProductPixelGridProof {
  return {
    product,
    captured: false,
    method,
    viewportPixelWidth: 0,
    viewportPixelHeight: 0,
    evidence: ['visible workbook grid surface was not found'],
  }
}
