import { performance } from 'node:perf_hooks'

import type { Frame, Page } from '@playwright/test'

import type { UiResponsivenessSameCorpusProduct } from './gen-ui-responsiveness-live-browser-scorecard.ts'
import { defaultViewport } from './ui-responsiveness-same-corpus-args.ts'
import {
  measureVisibleScrollResponseWithHooks,
  ScrollMovementVerificationError,
  type ScrollPositionSnapshot,
  type ScrollSample,
  type ScrollTriggerResult,
} from './ui-responsiveness-same-corpus-scroll.ts'
import { collectFrameIntervals, settleFrames, waitForNextFrame } from './ui-responsiveness-same-corpus-page-utils.ts'

type ScrollEventResponseProbeContext = Page | Frame

const scrollEventResponseProbeTimeoutMs = 5_000
const excelWebGridSelector = '.ewr-grdcontarea-grid'

export function sameCorpusScrollProbeSelectorsForProduct(product: UiResponsivenessSameCorpusProduct): readonly string[] {
  if (product === 'bilig') {
    return ['[data-testid="grid-scroll-viewport"]']
  }
  if (product === 'google-sheets') {
    return ['.native-scrollbar-y', '.native-scrollbar-x', '.grid-scrollable-wrapper']
  }
  return [excelWebGridSelector]
}

export async function measureVisibleScrollResponseWithRetries(
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

export async function movePointerToProductViewport(page: Page, product: UiResponsivenessSameCorpusProduct): Promise<void> {
  const box = await productViewportBox(page, product)
  if (box) {
    await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2)
    await page.mouse.click(box.x + box.width / 2, box.y + box.height / 2)
    return
  }
  await page.mouse.move(defaultViewport.width / 2, defaultViewport.height / 2)
}

export async function resetProductScrollPosition(page: Page, product: UiResponsivenessSameCorpusProduct): Promise<void> {
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
  if (product !== 'microsoft-excel-web') {
    await installScrollEventResponseProbe(page, product, sameCorpusScrollProbeSelectorsForProduct(product))
    return page
  }
  const frame = await firstFrameWithElement(page, excelWebGridSelector)
  if (!frame) {
    throw new Error('Unable to locate Microsoft Excel Web grid frame for scroll-event response probe')
  }
  await installScrollEventResponseProbe(frame, product, sameCorpusScrollProbeSelectorsForProduct(product))
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
  return await firstFrameElementBox(page, excelWebGridSelector)
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
