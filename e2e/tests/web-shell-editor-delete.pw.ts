import { expect, test, type Page } from '@playwright/test'
import { PRIMARY_MODIFIER, clickProductCell, createTestDocumentId, waitForWorkbookReady } from './web-shell-helpers.js'

test('@browser-ci web app keeps an in-cell Delete clear committed after clicking away', async ({ page }) => {
  const staleText = 'editor-delete-clickaway'
  await page.goto(`/?document=${encodeURIComponent(createTestDocumentId('playwright-editor-delete-click-away'))}&persist=0`)
  await waitForWorkbookReady(page)

  const formulaInput = page.getByTestId('formula-input')
  const cellEditor = page.getByTestId('cell-editor-input')

  await clickProductCell(page, 1, 1)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!B2')
  await formulaInput.fill(staleText)
  await formulaInput.press('Enter')
  await expect(formulaInput).toHaveValue(staleText)
  await expect.poll(() => nativeTextRunsInclude(page, staleText)).toBe(true)

  await page.getByTestId('sheet-grid-focus-target').focus()
  await page.keyboard.press('F2')
  await expect(cellEditor).toBeVisible()
  await expect(cellEditor).toHaveValue(staleText)

  await page.keyboard.press(`${PRIMARY_MODIFIER}+A`)
  await page.keyboard.press('Delete')
  await expect(cellEditor).toHaveValue('')

  await clickProductCell(page, 2, 1)
  await expect(cellEditor).toHaveCount(0)
  await expect(formulaInput).toHaveValue('')
  await expect.poll(() => nativeTextRunsInclude(page, staleText)).toBe(false)

  await clickProductCell(page, 1, 1)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!B2')
  await expect(formulaInput).toHaveValue('')
  await expect.poll(() => nativeTextRunsInclude(page, staleText)).toBe(false)
})

async function nativeTextRunsInclude(page: Page, text: string): Promise<boolean> {
  return await page.evaluate(
    (needle) => Array.from(document.querySelectorAll('[data-native-text-run]')).some((run) => run.textContent?.includes(needle) ?? false),
    text,
  )
}
