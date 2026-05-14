import { performance } from 'node:perf_hooks'

import type { Page } from '@playwright/test'

import type { CaptureArgs } from './ui-responsiveness-same-corpus-args.ts'
import type { UiResponsivenessSameCorpusProduct } from './gen-ui-responsiveness-live-browser-scorecard.ts'
import type { UiResponsivenessSameCorpusWorkload } from './ui-responsiveness-same-corpus-workloads.ts'
import { collectFrameIntervals, waitForNextFrame } from './ui-responsiveness-same-corpus-page-utils.ts'

export interface ProductOperationSample {
  readonly operationResponseMs: number
  readonly postOperationFrameMs: number
  readonly scrollEventResponseMs?: number
  readonly scrollMovementPx?: number
}

export interface SameCorpusWorkloadRunnerHooks {
  readonly measureVisibleScrollResponseWithRetries: (
    page: Page,
    product: UiResponsivenessSameCorpusProduct,
    deltaX: number,
    deltaY: number,
  ) => Promise<ProductOperationSample>
  readonly movePointerToProductViewport: (page: Page, product: UiResponsivenessSameCorpusProduct) => Promise<void>
}

type NonScrollWorkload = Exclude<
  UiResponsivenessSameCorpusWorkload,
  'open-workbook' | 'scroll-vertical' | 'scroll-horizontal' | 'wide-sheet-navigation'
>

type SameCorpusKeyboardOperation = { kind: 'press'; key: string } | { kind: 'type'; text: string }

export async function measureProductWorkload(args: {
  readonly page: Page
  readonly product: UiResponsivenessSameCorpusProduct
  readonly captureArgs: CaptureArgs
  readonly workload: UiResponsivenessSameCorpusWorkload
  readonly sampleIndex: number
  readonly loadToReadyMs: number
  readonly hooks: SameCorpusWorkloadRunnerHooks
}): Promise<ProductOperationSample> {
  if (args.workload === 'open-workbook') {
    return await sampleSettledOperation(args.page, args.loadToReadyMs)
  }
  if (args.workload === 'scroll-vertical') {
    return await args.hooks.measureVisibleScrollResponseWithRetries(args.page, args.product, 0, args.captureArgs.deltaY)
  }
  if (args.workload === 'scroll-horizontal') {
    return await args.hooks.measureVisibleScrollResponseWithRetries(args.page, args.product, Math.max(args.captureArgs.deltaX, 720), 0)
  }
  if (args.workload === 'wide-sheet-navigation') {
    return await args.hooks.measureVisibleScrollResponseWithRetries(args.page, args.product, Math.max(args.captureArgs.deltaX, 1440), 0)
  }
  return await measureNonScrollProductWorkload(args.page, args.product, args.workload, args.sampleIndex, args.hooks)
}

async function measureNonScrollProductWorkload(
  page: Page,
  product: UiResponsivenessSameCorpusProduct,
  workload: NonScrollWorkload,
  sampleIndex: number,
  hooks: SameCorpusWorkloadRunnerHooks,
): Promise<ProductOperationSample> {
  await hooks.movePointerToProductViewport(page, product)
  const startedAt = performance.now()
  await performProductUiOperation(page, product, workload, sampleIndex)
  return await sampleSettledOperation(page, performance.now() - startedAt)
}

async function sampleSettledOperation(page: Page, operationResponseMs: number): Promise<ProductOperationSample> {
  await waitForNextFrame(page)
  const frameIntervals = await collectFrameIntervals(page, 12)
  return {
    operationResponseMs,
    postOperationFrameMs: percentile(frameIntervals, 0.95),
  }
}

async function performProductUiOperation(
  page: Page,
  product: UiResponsivenessSameCorpusProduct,
  workload: NonScrollWorkload,
  sampleIndex: number,
): Promise<void> {
  if (product !== 'bilig') {
    await assertIncumbentEditableForWorkload(page, product, workload)
  }
  await performSameCorpusKeyboardOperations(page, sameCorpusKeyboardOperations(product, workload, sampleIndex))
}

export function sameCorpusKeyboardOperations(
  product: UiResponsivenessSameCorpusProduct,
  workload: NonScrollWorkload,
  sampleIndex: number,
  platform: NodeJS.Platform = process.platform,
): readonly SameCorpusKeyboardOperation[] {
  if (workload === 'select-cell') {
    return [{ kind: 'press', key: 'ArrowRight' }]
  }
  if (workload === 'jump-deep-row') {
    return [{ kind: 'press', key: primaryShortcut('ArrowDown', platform) }]
  }
  if (workload === 'fill-format-change') {
    return [{ kind: 'press', key: primaryShortcut('B', platform) }]
  }
  const value = workload === 'formula-edit' ? `=${String(sampleIndex + 1)}+1` : `${product}-same-corpus-${String(sampleIndex + 1)}`
  return [
    { kind: 'type', text: value },
    { kind: 'press', key: 'Enter' },
  ]
}

function primaryShortcut(key: string, platform: NodeJS.Platform): string {
  return platform === 'darwin' ? `Meta+${key}` : `Control+${key}`
}

async function performSameCorpusKeyboardOperations(
  page: Page,
  operations: readonly SameCorpusKeyboardOperation[],
  index = 0,
): Promise<void> {
  const operation = operations[index]
  if (!operation) {
    return
  }
  if (operation.kind === 'type') {
    await page.keyboard.type(operation.text)
  } else {
    await page.keyboard.press(operation.key)
  }
  await performSameCorpusKeyboardOperations(page, operations, index + 1)
}

async function assertIncumbentEditableForWorkload(
  page: Page,
  product: Exclude<UiResponsivenessSameCorpusProduct, 'bilig'>,
  workload: NonScrollWorkload,
): Promise<void> {
  const bodyText = await page
    .locator('body')
    .innerText({ timeout: 2_000 })
    .catch(() => '')
  const blocker = incumbentEditableWorkloadBlocker(product, page.url(), bodyText)
  if (blocker) {
    throw new Error(`Cannot measure ${workload} on ${product}: ${blocker}`)
  }
}

export function incumbentEditableWorkloadBlocker(
  product: Exclude<UiResponsivenessSameCorpusProduct, 'bilig'>,
  pageUrl: string,
  bodyText: string,
): string | null {
  const normalizedBody = bodyText.replace(/\s+/g, ' ').toLowerCase()
  if (product === 'google-sheets') {
    if (normalizedBody.includes('view only') || normalizedBody.includes('comment only') || normalizedBody.includes('request edit access')) {
      return 'Google Sheets page is read-only; provide an editable same-corpus Google Sheet URL or authenticated storage state.'
    }
    return null
  }
  if (pageUrl.includes('view.officeapps.live.com/op/view.aspx')) {
    return 'Microsoft Excel Web URL is the read-only Office viewer; provide an editable Excel Web workbook URL for edit workloads.'
  }
  if (normalizedBody.includes('view only') || normalizedBody.includes('read-only') || normalizedBody.includes('request edit access')) {
    return 'Microsoft Excel Web page is read-only; provide an editable same-corpus workbook URL or authenticated storage state.'
  }
  return null
}

function percentile(values: readonly number[], percentileValue: number): number {
  if (values.length === 0) {
    throw new Error('Cannot compute percentile for an empty same-corpus UI sample set')
  }
  const sorted = [...values].toSorted((left, right) => left - right)
  const index = Math.ceil(percentileValue * sorted.length) - 1
  return sorted[Math.max(0, Math.min(sorted.length - 1, index))]!
}
