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

test('@browser-ci web app keeps active in-cell undo and redo local to the draft editor', async ({ page }) => {
  const documentId = createTestDocumentId('playwright-editor-local-undo-redo')
  await page.goto(`/?document=${encodeURIComponent(documentId)}&persist=0&sheet=Sheet1&cell=B2`)
  await waitForWorkbookReady(page)

  const formulaInput = page.getByTestId('formula-input')
  const cellEditor = page.getByTestId('cell-editor-input')
  const redoShortcut = PRIMARY_MODIFIER === 'Meta' ? `${PRIMARY_MODIFIER}+Shift+Z` : `${PRIMARY_MODIFIER}+Y`

  await clickProductCell(page, 3, 3)
  await formulaInput.fill('workbook-history-sentinel')
  await formulaInput.press('Enter')
  await expect(formulaInput).toHaveValue('workbook-history-sentinel')

  await clickProductCell(page, 1, 1)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!B2')
  await page.keyboard.press('a')
  await expect(cellEditor).toBeVisible()
  await expect(cellEditor).toHaveValue('a')
  await cellEditor.press('b')
  await expect(cellEditor).toHaveValue('ab')
  await cellEditor.press('c')
  await expect(cellEditor).toHaveValue('abc')

  await cellEditor.press(`${PRIMARY_MODIFIER}+Z`)
  await expect(cellEditor).toHaveValue('ab')
  await expect(formulaInput).toHaveValue('ab')

  await cellEditor.press(`${PRIMARY_MODIFIER}+Z`)
  await expect(cellEditor).toHaveValue('a')
  await expect(formulaInput).toHaveValue('a')

  await cellEditor.press(redoShortcut)
  await expect(cellEditor).toHaveValue('ab')
  await expect(formulaInput).toHaveValue('ab')

  await page.keyboard.press('Enter')
  await expect(cellEditor).toHaveCount(0)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!B3')
  await expect.poll(() => nativeTextRunsInclude(page, 'abc')).toBe(false)
  await expect.poll(() => nativeTextRunsInclude(page, 'ab')).toBe(true)
  await clickProductCell(page, 1, 1)
  await expect(formulaInput).toHaveValue('ab')
  await clickProductCell(page, 3, 3)
  await expect(formulaInput).toHaveValue('workbook-history-sentinel')
})

async function nativeTextRunsInclude(page: Page, text: string): Promise<boolean> {
  return await page.evaluate(
    (needle) => Array.from(document.querySelectorAll('[data-native-text-run]')).some((run) => run.textContent?.includes(needle) ?? false),
    text,
  )
}
