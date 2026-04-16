import { expect, test } from '@playwright/test'
import {
  TOOLBAR_SYNC_ACTIONS,
  clickProductCell,
  expectMatchingGridRangeScreenshots,
  openZeroWorkbookPage,
  pickToolbarBorderPreset,
  pickToolbarPresetColor,
  remoteSyncEnabled,
  runToolbarSyncActions,
  seedToolbarActionRange,
  selectToolbarActionRange,
  waitForWorkbookReady,
} from './web-shell-helpers.js'

test('web app propagates content and styling changes across live zero tabs', async ({ page }, testInfo) => {
  test.skip(!remoteSyncEnabled, 'requires Zero-backed browser sync')
  test.slow()
  const documentId = `playwright-zero-style-multiplayer-${Date.now()}`
  const mirrorPage = await page.context().newPage()
  const viewport = page.viewportSize()
  if (viewport) {
    await mirrorPage.setViewportSize(viewport)
  }

  try {
    await Promise.all([openZeroWorkbookPage(page, documentId), openZeroWorkbookPage(mirrorPage, documentId)])

    await clickProductCell(page, 1, 1)
    await page.keyboard.type('relay')
    await page.keyboard.press('Enter')
    await selectToolbarActionRange(page)
    await selectToolbarActionRange(mirrorPage)

    const contentElapsed = await expectMatchingGridRangeScreenshots(page, mirrorPage, 'zero-content-relay', testInfo, 1, 1, 2, 2, 1_500, 8)
    expect(contentElapsed).toBeLessThanOrEqual(1_500)

    await page.getByLabel('Bold').click()
    await pickToolbarPresetColor(page, 'Fill color', 'light cornflower blue 3')
    await pickToolbarBorderPreset(page, 'All borders')
    await selectToolbarActionRange(page)
    await selectToolbarActionRange(mirrorPage)

    const styleElapsed = await expectMatchingGridRangeScreenshots(page, mirrorPage, 'zero-style-relay', testInfo, 1, 1, 2, 2, 1_500, 8)
    expect(styleElapsed).toBeLessThanOrEqual(1_500)
  } finally {
    await mirrorPage.close().catch(() => undefined)
  }
})

test('web app keeps two live zero tabs visually converged across toolbar actions', async ({ page }, testInfo) => {
  test.skip(!remoteSyncEnabled, 'requires Zero-backed browser sync')
  test.slow()
  const documentId = `playwright-zero-toolbar-multiplayer-${Date.now()}`
  const mirrorPage = await page.context().newPage()
  const viewport = page.viewportSize()
  if (viewport) {
    await mirrorPage.setViewportSize(viewport)
  }

  try {
    await Promise.all([openZeroWorkbookPage(page, documentId), openZeroWorkbookPage(mirrorPage, documentId)])
    await seedToolbarActionRange(page)
    await selectToolbarActionRange(page)
    await selectToolbarActionRange(mirrorPage)

    const initialElapsed = await expectMatchingGridRangeScreenshots(
      page,
      mirrorPage,
      'zero-toolbar-initial',
      testInfo,
      1,
      1,
      2,
      2,
      5_000,
      8,
    )
    expect(initialElapsed).toBeLessThanOrEqual(5_000)

    await runToolbarSyncActions(page, mirrorPage, TOOLBAR_SYNC_ACTIONS, testInfo)
  } finally {
    await mirrorPage.close().catch(() => undefined)
  }
})

test('web app preserves an in-progress local draft when another tab edits the same cell', async ({ page }) => {
  test.skip(!remoteSyncEnabled, 'requires Zero-backed browser sync')
  test.slow()
  const documentId = `playwright-zero-same-cell-draft-${Date.now()}`
  const mirrorPage = await page.context().newPage()
  const viewport = page.viewportSize()
  if (viewport) {
    await mirrorPage.setViewportSize(viewport)
  }

  try {
    await Promise.all([openZeroWorkbookPage(page, documentId), openZeroWorkbookPage(mirrorPage, documentId)])
    await expect(page.getByTestId('worker-error')).toHaveCount(0)
    await expect(mirrorPage.getByTestId('worker-error')).toHaveCount(0)

    const formulaInput = page.getByTestId('formula-input')
    const mirrorFormulaInput = mirrorPage.getByTestId('formula-input')

    await clickProductCell(page, 0, 0)
    await clickProductCell(mirrorPage, 0, 0)
    await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!A1')
    await expect(mirrorPage.getByTestId('status-selection')).toHaveText('Sheet1!A1')

    await formulaInput.focus()
    await formulaInput.selectText()
    await page.keyboard.type('local-draft')
    await expect(formulaInput).toHaveValue('local-draft')

    await mirrorFormulaInput.focus()
    await mirrorFormulaInput.selectText()
    await mirrorPage.keyboard.type('remote')
    await mirrorFormulaInput.press('Enter')
    await expect(mirrorFormulaInput).toHaveValue('remote')

    await expect(formulaInput).toHaveValue('local-draft')

    await formulaInput.press('Escape')
    await expect(formulaInput).toHaveValue('remote')
  } finally {
    await mirrorPage.close().catch(() => undefined)
  }
})

test('web app compares and applies a stale same-cell draft without losing local work', async ({ page }) => {
  test.skip(!remoteSyncEnabled, 'requires Zero-backed browser sync')
  test.slow()
  const documentId = `playwright-zero-same-cell-conflict-${Date.now()}`
  const mirrorPage = await page.context().newPage()
  const viewport = page.viewportSize()
  if (viewport) {
    await mirrorPage.setViewportSize(viewport)
  }

  try {
    await Promise.all([openZeroWorkbookPage(page, documentId), openZeroWorkbookPage(mirrorPage, documentId)])
    await expect(page.getByTestId('worker-error')).toHaveCount(0)
    await expect(mirrorPage.getByTestId('worker-error')).toHaveCount(0)

    const formulaInput = page.getByTestId('formula-input')
    const mirrorFormulaInput = mirrorPage.getByTestId('formula-input')

    await clickProductCell(page, 0, 0)
    await clickProductCell(mirrorPage, 0, 0)
    await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!A1')
    await expect(mirrorPage.getByTestId('status-selection')).toHaveText('Sheet1!A1')

    await formulaInput.focus()
    await formulaInput.selectText()
    await page.keyboard.type('local-draft')
    await expect(formulaInput).toHaveValue('local-draft')

    await mirrorFormulaInput.focus()
    await mirrorFormulaInput.selectText()
    await mirrorPage.keyboard.type('remote')
    await mirrorFormulaInput.press('Enter')
    await expect(mirrorFormulaInput).toHaveValue('remote')

    await expect(page.getByTestId('editor-conflict-banner')).toContainText('Remote update detected in Sheet1!A1 while you were editing.')
    await expect(formulaInput).toHaveValue('local-draft')

    await formulaInput.press('Enter')
    await expect(page.getByTestId('editor-conflict-apply-mine')).toBeVisible()

    await page.getByTestId('editor-conflict-apply-mine').click()

    await expect(page.getByTestId('editor-conflict-banner')).toHaveCount(0)
    await expect(formulaInput).toHaveValue('local-draft')
    await expect(mirrorFormulaInput).toHaveValue('local-draft')
  } finally {
    await mirrorPage.close().catch(() => undefined)
  }
})

test('web app reverts an authoritative change from the changes pane', async ({ page }) => {
  test.skip(!remoteSyncEnabled, 'requires Zero-backed browser sync')
  const documentId = `playwright-zero-change-revert-${Date.now()}`
  await openZeroWorkbookPage(page, documentId)

  const formulaInput = page.getByTestId('formula-input')
  const changesToggle = page.getByTestId('workbook-side-panel-toggle-changes')
  const changesTab = page.getByTestId('workbook-side-panel-tab-changes')

  await clickProductCell(page, 0, 0)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!A1')
  await formulaInput.fill('seed')
  await formulaInput.press('Enter')
  await expect(formulaInput).toHaveValue('seed')

  await expect(changesToggle).toContainText('1')
  await changesToggle.click()
  await expect(changesTab).toBeVisible()

  const changeRows = page.getByTestId('workbook-change-row')
  await expect(changeRows).toHaveCount(1)
  await expect(changeRows.first()).toContainText('Updated Sheet1!A1')

  await page.getByTestId('workbook-change-revert').click()

  await expect(formulaInput).toHaveValue('')
  await expect(changesToggle).toContainText('2')
  await expect(changeRows).toHaveCount(2)
  await expect(changeRows.first()).toContainText('Reverted r1: Updated Sheet1!A1')
  await expect(changeRows.nth(1)).toContainText('Reverted in r2')
})

test('web app restores persisted workbook state after a full reload', async ({ page }) => {
  test.skip(!remoteSyncEnabled, 'requires Zero-backed browser sync')
  const documentId = `playwright-zero-reload-persist-${Date.now()}`
  const formulaInput = page.getByTestId('formula-input')
  const resolvedValue = page.getByTestId('formula-resolved-value')

  await openZeroWorkbookPage(page, documentId)
  await expect(page.getByTestId('worker-error')).toHaveCount(0)

  await clickProductCell(page, 0, 0)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!A1')
  await formulaInput.fill('17')
  await formulaInput.press('Enter')
  await expect(formulaInput).toHaveValue('17')
  await expect(resolvedValue).toHaveText('17')

  await page.reload({ waitUntil: 'domcontentloaded' })
  await waitForWorkbookReady(page)
  await expect(page.getByTestId('worker-error')).toHaveCount(0)

  await clickProductCell(page, 0, 0)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!A1')
  await expect(formulaInput).toHaveValue('17')
  await expect(resolvedValue).toHaveText('17')
})
