import { BILIG_CONTENT_TYPE } from '@bilig/agent-api'
import { expect, test } from '@playwright/test'
import { createTestDocumentId, gotoWorkbookShell, waitForWorkbookReady } from './web-shell-helpers.js'

test('root opens the spreadsheet without a document header @browser-ci', async ({ page }) => {
  await page.goto('/')
  await waitForWorkbookReady(page)
  await expect(page.getByTestId('formula-input')).toBeVisible()
  await expect(page.getByTestId('sheet-grid')).toBeVisible()
  await expect(page.locator('header[aria-label="Workbook document"]')).toHaveCount(0)
})

test('whole-sheet clear stays responsive, supports undo, and persists after reload @browser-ci', async ({ page }) => {
  await page.goto(`/?document=${encodeURIComponent(createTestDocumentId('whole-sheet-clear'))}&sheet=Sheet1&cell=A1`)
  await waitForWorkbookReady(page)
  const name = page.getByTestId('name-box')
  const formula = page.getByTestId('formula-input')
  await formula.fill('42')
  await formula.press('Enter')
  await expect(formula).toHaveValue('42')
  await page.getByRole('button', { name: 'Select entire sheet' }).click()
  await page.getByTestId('sheet-grid').press('Delete')
  await name.fill('A1')
  await name.press('Enter')
  await expect(formula).toHaveValue('')
  await page.getByRole('button', { name: 'Undo', exact: true }).click()
  await expect(formula).toHaveValue('42')
  await page.getByRole('button', { name: 'Redo', exact: true }).click()
  await expect(formula).toHaveValue('')
  await page.getByRole('button', { name: 'Undo', exact: true }).click()
  await expect(formula).toHaveValue('42')
  await page.getByRole('button', { name: 'Select entire sheet' }).click()
  await page.getByTestId('sheet-grid').press('Delete')
  await expect(formula).toHaveValue('')
  await expect(page.getByTestId('workbook-save-status')).toHaveText(/Saved|Saved on this device/)
  await page.reload()
  await waitForWorkbookReady(page)
  await name.fill('A1')
  await name.press('Enter')
  await expect(formula).toHaveValue('')
})

test('web app restores a Bilig backup and retains formulas after reload @browser-ci', async ({ page }) => {
  await gotoWorkbookShell(page)
  await waitForWorkbookReady(page)
  await page.getByTestId('workbook-import-toggle').click()
  await page.getByTestId('workbook-import-file').setInputFiles({
    name: 'Restore proof.bilig.json',
    mimeType: BILIG_CONTENT_TYPE,
    buffer: Buffer.from(
      JSON.stringify({
        version: 1,
        workbook: { name: 'Restore proof' },
        sheets: [{ id: 1, name: 'Sheet1', order: 0, cells: [{ address: 'A1', formula: 'SUM(3,4,5)' }] }],
      }),
    ),
  })
  await expect(page.getByTestId('workbook-import-preview-list')).toBeVisible()
  await page.getByTestId('workbook-import-create').click()
  await page.waitForURL(/document=xlsx%3A/)
  await waitForWorkbookReady(page)
  await page.reload()
  await waitForWorkbookReady(page)
  await page.getByTestId('name-box').fill('A1')
  await page.getByTestId('name-box').press('Enter')
  await expect(page.getByTestId('formula-input')).toHaveValue('=SUM(3,4,5)')
  await expect(page.getByTestId('formula-resolved-value')).toHaveText('12')
})
