import { BILIG_CONTENT_TYPE } from '@bilig/agent-api'
import { expect, test } from '@playwright/test'
import { gotoWorkbookShell, waitForWorkbookReady } from './web-shell-helpers.js'

test('web app restores a Bilig backup and retains formulas after reload @browser-ci', async ({ page }) => {
  await gotoWorkbookShell(page)
  await waitForWorkbookReady(page)
  await page.getByRole('button', { name: 'Workbook menu', exact: true }).click()
  await page.getByRole('menuitem', { name: 'Import workbook…', exact: true }).click()
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
  await expect(page.getByTestId('workbook-title')).toHaveText('Restore proof')
  await page.reload()
  await waitForWorkbookReady(page)
  await expect(page.getByTestId('workbook-title')).toHaveText('Restore proof')
  await page.getByTestId('name-box').fill('A1')
  await page.getByTestId('name-box').press('Enter')
  await expect(page.getByTestId('formula-input')).toHaveValue('=SUM(3,4,5)')
})
