import { readFile } from 'node:fs/promises'
import { execFileSync } from 'node:child_process'
import { fileURLToPath } from 'node:url'
import { expect, test, type Page } from '@playwright/test'

async function createContribution(page: Page) {
  await page.goto('/models')
  await page.getByRole('button', { name: /What makes this profitable/ }).click()
  await expect(page.getByLabel('Revenue', { exact: true })).toHaveText('$10,000.00')
}

async function editValue(page: Page, label: string, value: string) {
  await page.getByLabel(label, { exact: true }).fill(value)
  await page.getByLabel(label, { exact: true }).press('Enter')
}

test('@browser-ci model storage can recover after an opening failure', async ({ page }) => {
  await page.addInitScript(() => {
    const open = IDBFactory.prototype.open.bind(window.indexedDB)
    let failed = false
    IDBFactory.prototype.open = function (...args: Parameters<IDBFactory['open']>) {
      if (!failed && args[0] === 'bilig-model-workspace') {
        failed = true
        throw new DOMException('Storage is temporarily unavailable', 'InvalidStateError')
      }
      return open(...args)
    }
  })
  await page.goto('/models')
  await expect(page.getByRole('heading', { name: 'Could not open your workspace' })).toBeVisible()
  await page.getByRole('button', { name: 'Retry', exact: true }).click()
  await expect(page.getByRole('heading', { name: 'Your models', exact: true })).toBeVisible()
  await page.getByRole('button', { name: 'New model', exact: true }).click()
  await expect(page.getByLabel('First result', { exact: true })).toHaveText('20')
})

test('@browser-ci model workspace recalculates, compares frozen scenarios, and restores after reload', async ({ page }, testInfo) => {
  await page.goto('/')
  await expect(page.getByRole('heading', { name: 'Your models', exact: true })).toBeVisible()
  await page.screenshot({ path: testInfo.outputPath('models-library.png'), fullPage: true })
  await createContribution(page)
  await page.getByLabel('Scenario name').fill('Baseline')
  await page.getByRole('button', { name: 'Save scenario', exact: true }).click()
  await editValue(page, 'Units sold', '150')
  await expect(page.getByLabel('Revenue', { exact: true })).toHaveText('$15,000.00')
  await expect(page.getByLabel('Operating profit', { exact: true })).toHaveText('$6,000.00')
  await page.screenshot({ path: testInfo.outputPath('model-editor.png'), fullPage: true })
  await page.getByRole('button', { name: /Compare/ }).click()
  const revenue = page.getByRole('row').filter({ has: page.getByRole('rowheader', { name: 'Revenue', exact: true }) })
  await expect(revenue).toContainText('$15,000.00')
  await expect(revenue).toContainText('$10,000.00')
  await page.getByRole('button', { name: 'Restore Baseline' }).click()
  await expect(revenue.getByRole('cell').first()).toHaveText('$10,000.00')
  await expect(page.getByText('Saved in this browser', { exact: true })).toBeVisible()
  await page.reload()
  await expect(page.getByLabel('Units sold', { exact: true })).toHaveValue('100')
  await expect(page.getByLabel('Revenue', { exact: true })).toHaveText('$10,000.00')
  await expect(page.getByRole('button', { name: 'Compare 1' })).toBeVisible()
})

test('@browser-ci model backups and WorkPaper exports contain actual formulas and computed readback', async ({ page }, testInfo) => {
  await createContribution(page)
  await editValue(page, 'Units sold', '150')
  await expect(page.getByLabel('Revenue', { exact: true })).toHaveText('$15,000.00')
  const originalUrl = page.url()
  const downloadEvent = page.waitForEvent('download')
  await page.getByRole('button', { name: 'Export backup', exact: true }).click()
  const backup = await downloadEvent
  const backupPath = testInfo.outputPath('proof.bilig-model.json')
  await backup.saveAs(backupPath)

  const workbookDownloadEvent = page.waitForEvent('download')
  await page.getByRole('button', { name: 'WorkPaper JSON', exact: true }).click()
  const workbookFile = await workbookDownloadEvent
  const workbookPath = testInfo.outputPath('proof.workpaper.json')
  await workbookFile.saveAs(workbookPath)
  const restored = execFileSync(
    process.execPath,
    [
      '--input-type=module',
      '--eval',
      `
    import { readFileSync } from 'node:fs';
    import { createWorkPaperFromDocument, parseWorkPaperDocument } from '@bilig/workpaper/browser';
    const workbook = createWorkPaperFromDocument(parseWorkPaperDocument(readFileSync(process.argv[1], 'utf8')));
    try {
      const sheet = workbook.getSheetId('Results');
      if (sheet === undefined) throw new Error('Missing Results sheet');
      const address = { sheet, row: 1, col: 1 };
      process.stdout.write(JSON.stringify({ value: workbook.getCellValue(address), formula: workbook.getCellFormula(address) }));
    } finally { workbook.dispose(); }
  `,
      workbookPath,
    ],
    { cwd: fileURLToPath(new URL('../../apps/web', import.meta.url)), encoding: 'utf8' },
  )
  const readback: unknown = JSON.parse(restored)
  expect(readback).toMatchObject({ value: { value: 15000 }, formula: '=Inputs!B2*Inputs!B3' })

  await page.getByRole('button', { name: 'Models', exact: true }).last().click()
  await page.getByLabel('Import model backup', { exact: true }).setInputFiles(backupPath)
  await expect(page.getByLabel('Revenue', { exact: true })).toHaveText('$15,000.00')
  expect(page.url()).not.toBe(originalUrl)
  await editValue(page, 'Units sold', '200')
  await expect(page.getByLabel('Revenue', { exact: true })).toHaveText('$20,000.00')
  await page.goto(originalUrl)
  await expect(page.getByLabel('Revenue', { exact: true })).toHaveText('$15,000.00')
})

test('@browser-ci models reject a stale tab save and preserve its unsaved backup', async ({ page, context }, testInfo) => {
  await createContribution(page)
  const second = await context.newPage()
  await second.goto('/models')
  await second.getByRole('button', { name: /^Contribution model/ }).click()
  await expect(second.getByLabel('Revenue', { exact: true })).toHaveText('$10,000.00')
  await editValue(page, 'Units sold', '150')
  await expect(page.getByText('Saved in this browser', { exact: true })).toBeVisible()
  await editValue(second, 'Units sold', '175')
  await expect(second.getByRole('alert')).toContainText('changed in another tab')
  await second.getByRole('button', { name: 'Models', exact: true }).first().click()
  await expect(second.getByLabel('Units sold', { exact: true })).toHaveValue('175')
  await second.goBack()
  await expect(second.getByLabel('Units sold', { exact: true })).toHaveValue('175')
  const downloadEvent = second.waitForEvent('download')
  await second.getByRole('button', { name: 'Export unsaved edits' }).click()
  const backup = await downloadEvent
  const path = testInfo.outputPath('conflict.bilig-model.json')
  await backup.saveAs(path)
  await expect.poll(() => readFile(path, 'utf8')).toContain('175')
  await second.getByRole('button', { name: 'Reload saved version' }).click()
  await expect(second.getByLabel('Units sold', { exact: true })).toHaveValue('150')
  await expect(second.getByRole('alert')).toHaveCount(0)
  await page.reload()
  await expect(page.getByLabel('Units sold', { exact: true })).toHaveValue('150')
  await expect(page.getByLabel('Revenue', { exact: true })).toHaveText('$15,000.00')
})

test('@browser-ci custom models support new assumptions, formula edits, errors, and keyboard correction', async ({ page }) => {
  await page.goto('/models')
  await page.getByRole('button', { name: 'New model', exact: true }).click()
  await expect(page.getByLabel('First result', { exact: true })).toHaveText('20')
  await editValue(page, 'Model name', 'Delivery estimate')
  await page.getByRole('button', { name: 'Add assumption' }).click()
  await editValue(page, 'Input 2', '5')
  await editValue(page, 'First result formula', '=Inputs!B2*Inputs!B3')
  await expect(page.getByLabel('First result', { exact: true })).toHaveText('50')
  await editValue(page, 'First result formula', '=1/0')
  await expect(page.getByLabel('First result', { exact: true })).toHaveText('#DIV/0!')
  await page.getByLabel('First result formula', { exact: true }).fill('bad')
  await page.getByLabel('First result formula', { exact: true }).press('Enter')
  await expect(page.getByRole('alert')).toHaveText('Start formulas with =.')
  await page.getByLabel('First result formula', { exact: true }).press('Escape')
  await expect(page.getByLabel('First result formula', { exact: true })).toHaveValue('=1/0')
  await editValue(page, 'First result formula', '=Inputs!B2*Inputs!B3')
  await expect(page.getByLabel('First result', { exact: true })).toHaveText('50')
  await page.getByRole('button', { name: 'Archive model' }).click()
  await expect(page.getByRole('heading', { name: 'Your models', exact: true })).toBeVisible()
  await page.getByRole('button', { name: 'Archive', exact: true }).click()
  await expect(page.getByRole('button', { name: /Delivery estimate/ })).toBeVisible()
  await page.getByRole('button', { name: 'Restore', exact: true }).click()
  await page.getByRole('button', { name: 'Models', exact: true }).click()
  await page.getByRole('button', { name: /Delivery estimate/ }).click()
  await expect(page.getByLabel('First result', { exact: true })).toHaveText('50')
})

test('@browser-ci model workspace works at a mobile viewport without document overflow', async ({ page }, testInfo) => {
  await page.setViewportSize({ width: 390, height: 844 })
  await createContribution(page)
  await editValue(page, 'Units sold', '150')
  await expect(page.getByLabel('Revenue', { exact: true })).toHaveText('$15,000.00')
  expect(await page.evaluate(() => document.documentElement.scrollWidth <= window.innerWidth)).toBe(true)
  await page.screenshot({ path: testInfo.outputPath('model-mobile.png'), fullPage: true })
})

test('@browser-ci what-if analysis compares five recalculated variations without changing the saved model', async ({ page }, testInfo) => {
  await createContribution(page)
  await page.getByRole('button', { name: 'What if', exact: true }).click()
  const table = page.getByRole('region', { name: 'What-if comparison table' })
  const profit = table.getByRole('row').filter({ has: page.getByRole('rowheader', { name: 'Operating profit', exact: true }) })
  await expect(profit.getByRole('cell')).toHaveText(['$1,800.00', '$2,400.00', '$3,000.00', '$3,600.00', '$4,200.00'])
  await editValue(page, 'Step size', '20')
  await expect(profit.getByRole('cell')).toHaveText(['$600.00', '$1,800.00', '$3,000.00', '$4,200.00', '$5,400.00'])
  await editValue(page, 'Step size', '0')
  await expect(page.getByRole('alert')).toHaveText('Enter a positive number.')
  await page.getByLabel('Step size', { exact: true }).press('Escape')
  await page.getByLabel('Assumption', { exact: true }).selectOption('price')
  await expect(profit.getByRole('cell')).toHaveText(['$1,000.00', '$2,000.00', '$3,000.00', '$4,000.00', '$5,000.00'])
  await page.screenshot({ path: testInfo.outputPath('what-if.png'), fullPage: true })
  await page.setViewportSize({ width: 390, height: 844 })
  expect(await page.evaluate(() => document.documentElement.scrollWidth <= window.innerWidth)).toBe(true)
  await page.getByRole('button', { name: 'Model', exact: true }).click()
  await expect(page.getByLabel('Units sold', { exact: true })).toHaveValue('100')
  await expect(page.getByLabel('Price per unit', { exact: true })).toHaveValue('100')
  await page.reload()
  await expect(page.getByLabel('Revenue', { exact: true })).toHaveText('$10,000.00')
  await expect(page.getByRole('button', { name: 'Compare 0', exact: true })).toBeVisible()
})

test('@browser-ci model import errors leave the library usable', async ({ page }) => {
  await page.goto('/models')
  await page
    .getByLabel('Import model backup', { exact: true })
    .setInputFiles({ name: 'invalid.json', mimeType: 'application/json', buffer: Buffer.from('{"format":"future"}') })
  await expect(page.getByRole('alert')).toContainText('not a supported Bilig model backup')
  await page.getByRole('button', { name: 'New model', exact: true }).click()
  await expect(page.getByLabel('First result', { exact: true })).toHaveText('20')
})
