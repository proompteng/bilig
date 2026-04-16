import { expect, test } from '@playwright/test'

test('production smoke loads an isolated workbook and commits edits', async ({ page }) => {
  const documentId = `prod-smoke-${Date.now()}`
  const response = await page.goto(`/?document=${documentId}`, {
    waitUntil: 'domcontentloaded',
  })
  expect(response?.ok()).toBeTruthy()

  await expect(page.getByTestId('workbook-shell')).toBeVisible()
  await expect(page.getByTestId('worker-error')).toHaveCount(0)
  await expect(page.getByTestId('name-box')).toHaveValue('A1')
  await expect(page.getByTestId('status-selection')).toContainText('Sheet1!A1')

  const formulaInput = page.getByTestId('formula-input')
  await formulaInput.fill('123')
  await formulaInput.press('Enter')
  await expect.poll(async () => await formulaInput.inputValue()).toBe('123')

  const nameBox = page.getByTestId('name-box')
  await nameBox.fill('B1')
  await nameBox.press('Enter')
  await expect(page.getByTestId('status-selection')).toContainText('Sheet1!B1')

  await formulaInput.fill('=A1*2')
  await formulaInput.press('Enter')
  await expect.poll(async () => (await page.getByTestId('formula-resolved-value').textContent())?.trim()).toBe('246')
  await expect(page.getByTestId('worker-error')).toHaveCount(0)
})
