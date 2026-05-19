import { expect, test, type Page } from '@playwright/test'
import {
  clickProductCell,
  countGreenFillPixelsInCell,
  createTestDocumentId,
  pickToolbarPresetColor,
  waitForWorkbookReady,
} from './web-shell-helpers.js'

test('@browser-ci web app keeps deleted filled cells stable after click-away and viewport churn', async ({ page }) => {
  const documentId = createTestDocumentId('playwright-delete-fill-stability')
  const text = 'delete-fill-no-ghost'
  await page.setViewportSize({ width: 1166, height: 820 })
  await page.goto(`/?document=${encodeURIComponent(documentId)}&persist=0&sheet=Sheet1&cell=A1`)
  await waitForWorkbookReady(page)

  const grid = page.getByTestId('sheet-grid')
  const formulaInput = page.getByTestId('formula-input')

  await clickProductCell(page, 1, 1)
  await formulaInput.fill(text)
  await formulaInput.press('Enter')
  await expect(formulaInput).toHaveValue(text)
  await expect.poll(() => nativeTextRunTextAt(page, 1, 1)).toBe(text)

  await pickToolbarPresetColor(page, 'Fill color', 'green')
  await expect
    .poll(() => countGreenFillPixelsInCell(page, 1, 1), {
      message: 'setup should visibly paint B2 green before deletion',
      timeout: 5_000,
    })
    .toBeGreaterThan(120)

  await clickProductCell(page, 1, 1)
  await grid.press('Delete')
  await expect(formulaInput).toHaveValue('')
  await expect.poll(() => nativeTextRunTextAt(page, 1, 1)).toBe('')
  expect(
    Math.min(...(await sampleGreenFillPixelsAcrossFrames(page, 1, 1, 4))),
    'delete must clear text without flashing the retained fill to default',
  ).toBeGreaterThan(120)

  await clickProductCell(page, 4, 4)
  await expect.poll(() => nativeTextRunsInclude(page, text)).toBe(false)
  await expect.poll(() => countGreenFillPixelsInCell(page, 1, 1)).toBeGreaterThan(120)

  await page.getByTestId('grid-scroll-viewport').evaluate((viewport) => {
    viewport.scrollTop = 900
    viewport.scrollLeft = 220
    viewport.dispatchEvent(new Event('scroll', { bubbles: true }))
  })
  await expect.poll(() => nativeTextRunsInclude(page, text)).toBe(false)

  await page.getByTestId('grid-scroll-viewport').evaluate((viewport) => {
    viewport.scrollTop = 0
    viewport.scrollLeft = 0
    viewport.dispatchEvent(new Event('scroll', { bubbles: true }))
  })
  await expect.poll(() => nativeTextRunsInclude(page, text)).toBe(false)
  await expect.poll(() => countGreenFillPixelsInCell(page, 1, 1)).toBeGreaterThan(120)

  await clickProductCell(page, 1, 1)
  await expect(formulaInput).toHaveValue('')
})

async function nativeTextRunsInclude(page: Page, text: string): Promise<boolean> {
  return await page.evaluate(
    (needle) => Array.from(document.querySelectorAll('[data-native-text-run]')).some((run) => run.textContent?.includes(needle) ?? false),
    text,
  )
}

async function nativeTextRunTextAt(page: Page, columnIndex: number, rowIndex: number): Promise<string> {
  return await page.evaluate(
    ({ col, row }) =>
      document.querySelector(`[data-native-text-run-row="${String(row)}"][data-native-text-run-col="${String(col)}"]`)?.textContent ?? '',
    { col: columnIndex, row: rowIndex },
  )
}

async function sampleGreenFillPixelsAcrossFrames(
  page: Page,
  columnIndex: number,
  rowIndex: number,
  remainingSamples: number,
  samples: readonly number[] = [],
): Promise<readonly number[]> {
  if (remainingSamples <= 0) {
    return samples
  }
  const pixels = await countGreenFillPixelsInCell(page, columnIndex, rowIndex)
  await page.waitForTimeout(50)
  return await sampleGreenFillPixelsAcrossFrames(page, columnIndex, rowIndex, remainingSamples - 1, [...samples, pixels])
}
