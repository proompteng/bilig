import { expect, test, type Page } from '@playwright/test'
import {
  clickProductCell,
  countGreenFillPixelsInCell,
  createTestDocumentId,
  expectToolbarColor,
  getProductColumnLeft,
  getProductColumnWidth,
  getProductRowHeight,
  getProductRowTop,
  getToolbarButton,
  PRIMARY_MODIFIER,
  PRODUCT_HEADER_HEIGHT,
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

test('@browser-ci web app applies fill color after moving text into an empty tile range', async ({ page }) => {
  const documentId = createTestDocumentId('playwright-move-text-fill-range')
  const text = 'moved-fill-stability'
  await page.setViewportSize({ width: 1166, height: 820 })
  await page.goto(`/?document=${encodeURIComponent(documentId)}&persist=0&sheet=Sheet1&cell=A1`)
  await waitForWorkbookReady(page)

  const grid = page.getByTestId('sheet-grid')
  const formulaInput = page.getByTestId('formula-input')

  await clickProductCell(page, 1, 1)
  await formulaInput.fill(text)
  await formulaInput.press('Enter')
  await expect.poll(() => nativeTextRunTextAt(page, 1, 1)).toBe(text)

  await dragProductSelectedContentLane(page, 1, 1, 3, 4)
  await expect.poll(() => nativeTextRunTextAt(page, 1, 1)).toBe('')
  await expect.poll(() => nativeTextRunTextAt(page, 3, 4)).toBe(text)

  await clickProductCell(page, 3, 4)
  await clickProductCell(page, 5, 7, { shift: true })
  await pickToolbarPresetColor(page, 'Fill color', 'green')

  const fillProofCells = [
    [3, 4],
    [4, 4],
    [5, 4],
    [3, 7],
    [4, 7],
    [5, 7],
  ] as const
  await expect
    .poll(
      async () => {
        const pixelCounts = await Promise.all(
          fillProofCells.map(([columnIndex, rowIndex]) => countGreenFillPixelsInCell(page, columnIndex, rowIndex)),
        )
        return Math.min(...pixelCounts)
      },
      {
        message: 'moved text range should paint green fill across visible occupied and empty cells',
        timeout: 5_000,
      },
    )
    .toBeGreaterThan(120)

  await grid.press('Delete')
  await expect(formulaInput).toHaveValue('')
  await expect.poll(() => nativeTextRunsInclude(page, text)).toBe(false)
  await expect.poll(() => countGreenFillPixelsInCell(page, 3, 4)).toBeGreaterThan(120)

  await clickProductCell(page, 7, 9)
  await clickProductCell(page, 3, 4)
  await expect(formulaInput).toHaveValue('')
  await expect.poll(() => nativeTextRunsInclude(page, text)).toBe(false)
  await expect.poll(() => countGreenFillPixelsInCell(page, 3, 4)).toBeGreaterThan(120)
})

test('@browser-ci web app keeps fill undo and redo visually stable from grid keyboard ownership', async ({ page }) => {
  const documentId = createTestDocumentId('playwright-fill-undo-redo-stability')
  const redoShortcut = PRIMARY_MODIFIER === 'Meta' ? 'Meta+Shift+Z' : 'Control+Y'
  await page.setViewportSize({ width: 1166, height: 820 })
  await page.goto(`/?document=${encodeURIComponent(documentId)}&persist=0&sheet=Sheet1&cell=A1`)
  await waitForWorkbookReady(page)

  const grid = page.getByTestId('sheet-grid')
  const formulaInput = page.getByTestId('formula-input')

  await clickProductCell(page, 1, 1)
  await pickToolbarPresetColor(page, 'Fill color', 'green')
  await expectToolbarColor(getToolbarButton(page, 'Fill color'), '#00ff00')
  expect(
    Math.min(...(await sampleGreenFillPixelsAcrossFrames(page, 1, 1, 4))),
    'setup should paint B2 green before undoing the style mutation',
  ).toBeGreaterThan(120)

  await expect
    .poll(async () => page.evaluate(() => document.activeElement?.getAttribute('data-testid') ?? null), {
      message: 'toolbar style command must return keyboard ownership to the grid before undo',
    })
    .toBe('sheet-grid-focus-target')

  await grid.press(`${PRIMARY_MODIFIER}+Z`)
  await expect(formulaInput).toHaveValue('')
  await expect.poll(() => countGreenFillPixelsInCell(page, 1, 1)).toBe(0)
  await expectToolbarColor(getToolbarButton(page, 'Fill color'), '#ffffff')
  await expect(page.getByTestId('name-box')).toHaveValue('B2')

  await grid.press(redoShortcut)
  await expect(formulaInput).toHaveValue('')
  await expectToolbarColor(getToolbarButton(page, 'Fill color'), '#00ff00')
  expect(
    Math.min(...(await sampleGreenFillPixelsAcrossFrames(page, 1, 1, 4))),
    'redo should restore the green fill without flashing through default grid paint',
  ).toBeGreaterThan(120)
  await expect(page.getByTestId('name-box')).toHaveValue('B2')
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

async function dragProductSelectedContentLane(page: Page, startColumn: number, startRow: number, targetColumn: number, targetRow: number) {
  const gridLocator = page.getByTestId('sheet-grid')
  await expect(gridLocator).toBeVisible()
  const grid = await gridLocator.boundingBox()
  if (!grid) {
    throw new Error('sheet grid is not visible')
  }

  const startLeft = await getProductColumnLeft(page, startColumn)
  const startTop = await getProductRowTop(page, startRow)
  const startWidth = await getProductColumnWidth(page, startColumn)
  const startHeight = await getProductRowHeight(page, startRow)
  const targetLeft = await getProductColumnLeft(page, targetColumn)
  const targetTop = await getProductRowTop(page, targetRow)
  const targetWidth = await getProductColumnWidth(page, targetColumn)
  const targetHeight = await getProductRowHeight(page, targetRow)

  await page.mouse.move(
    grid.x + startLeft + Math.min(32, Math.floor(startWidth * 0.35)),
    grid.y + PRODUCT_HEADER_HEIGHT + startTop + Math.floor(startHeight / 2),
  )
  await page.mouse.down()
  await page.mouse.move(
    grid.x + targetLeft + Math.floor(targetWidth / 2),
    grid.y + PRODUCT_HEADER_HEIGHT + targetTop + Math.floor(targetHeight / 2),
    { steps: 12 },
  )
  await page.mouse.up()
}
