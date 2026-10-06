import { expect, test } from '@playwright/test'
import {
  clickProductCell,
  createTestDocumentId,
  getProductFillHandleDragPoints,
  gotoWorkbookShell,
  waitForWorkbookReady,
} from './web-shell-helpers.js'

test('@browser-ci web app keeps the in-cell editor on the same selection chrome', async ({ page }) => {
  await gotoWorkbookShell(page, `/?document=${encodeURIComponent(createTestDocumentId('editor-selection-chrome'))}&persist=0`)
  await waitForWorkbookReady(page)
  await clickProductCell(page, 2, 4)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!C5')
  await page.getByTestId('sheet-grid-focus-target').press('F2')
  await expect(page.getByTestId('cell-editor-overlay')).toBeVisible()
  await expect(page.getByTestId('cell-editor-overlay')).toHaveCSS('border-color', 'rgb(33, 115, 70)')
  await expect(page.locator('[data-grid-selection-visual-role="fill-handle"]')).toBeVisible()
  await expect(page.locator('[data-grid-fill-handle="true"]')).toBeVisible()
})

test('@browser-ci typing in a cell leaves the fill handle available and fills the current draft', async ({ page }) => {
  await gotoWorkbookShell(page, `/?document=${encodeURIComponent(createTestDocumentId('editor-fill-draft'))}&persist=0`)
  await waitForWorkbookReady(page)
  await clickProductCell(page, 1, 1)
  await page.keyboard.type('asdfaf')
  await expect(page.getByTestId('cell-editor-input')).toHaveValue('asdfaf')
  await expect(page.locator('[data-grid-fill-handle="true"]')).toBeVisible()
  await expect(page.locator('[data-grid-selection-visual-role="fill-handle"]')).toBeVisible()

  const { sourceX, sourceY, targetX, targetY } = await getProductFillHandleDragPoints(page, 1, 1, 1, 4)
  await page.mouse.move(sourceX, sourceY)
  await page.mouse.down()
  await page.mouse.move(targetX, targetY, { steps: 8 })
  await page.mouse.up()
  await expect(page.getByTestId('cell-editor-input')).toHaveCount(0)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!B2:B5')
  await [1, 2, 3, 4].reduce(async (previous, row) => {
    await previous
    await clickProductCell(page, 1, row)
    await expect(page.getByTestId('formula-input')).toHaveValue('asdfaf')
  }, Promise.resolve())
})

test('@browser-ci fill dragging commits a formula draft and relocates its references', async ({ page }) => {
  await gotoWorkbookShell(page, `/?document=${encodeURIComponent(createTestDocumentId('editor-fill-formula'))}&persist=0`)
  await waitForWorkbookReady(page)
  const formula = page.getByTestId('formula-input')
  await clickProductCell(page, 1, 0)
  await formula.fill('10')
  await formula.press('Enter')
  await clickProductCell(page, 1, 1)
  await page.keyboard.type('=B1+1')
  await expect(page.getByTestId('cell-editor-input')).toHaveValue('=B1+1')

  const { sourceX, sourceY, targetX, targetY } = await getProductFillHandleDragPoints(page, 1, 1, 1, 4)
  await page.mouse.move(sourceX, sourceY)
  await page.mouse.down()
  await page.mouse.move(targetX, targetY, { steps: 8 })
  await page.mouse.up()
  await expect(page.getByTestId('cell-editor-input')).toHaveCount(0)
  await [1, 2, 3, 4].reduce(async (previous, row) => {
    await previous
    await clickProductCell(page, 1, row)
    await expect(formula).toHaveValue(`=B${row}+1`)
    await expect(page.getByTestId('formula-resolved-value')).toHaveText(String(10 + row))
  }, Promise.resolve())
})
