import { expect, test, type Page } from '@playwright/test'
import {
  PRIMARY_MODIFIER,
  PRODUCT_COLUMN_WIDTH,
  PRODUCT_HEADER_HEIGHT,
  PRODUCT_ROW_HEIGHT,
  PRODUCT_ROW_MARKER_WIDTH,
  clickProductBodyOffset,
  clickProductCell,
  clickProductCellUpperHalf,
  createTestDocumentId,
  dragProductBodySelection,
  dragProductColumnResize,
  dragProductHeaderSelection,
  doubleClickProductColumnResizeHandle,
  getProductColumnLeft,
  getProductColumnWidth,
  getProductRowHeight,
  getProductRowTop,
  settleWorkbookScrollPerf,
  startWorkbookScrollPerf,
  stopWorkbookScrollPerf,
  waitForWorkbookReady,
} from './web-shell-helpers.js'
test('web app keeps sheet tabs and status bar visible in a short viewport', async ({ page }) => {
  await page.setViewportSize({ width: 2048, height: 220 })
  await page.goto('/')
  await waitForWorkbookReady(page)

  const sheetTab = page.getByRole('tab', { name: 'Sheet1' })
  const statusSummary = page.getByTestId('workbook-selection-summary')

  await expect(sheetTab).toBeVisible()
  await expect(statusSummary).toBeVisible()

  const tabBox = await sheetTab.boundingBox()
  const statusBox = await statusSummary.boundingBox()
  if (!tabBox || !statusBox) {
    throw new Error('footer controls are not visible')
  }

  expect(tabBox.y + tabBox.height).toBeLessThanOrEqual(220)
  expect(statusBox.y + statusBox.height).toBeLessThanOrEqual(220)
})

test('web app supports column and row header selection', async ({ page }) => {
  await page.goto('/')
  await waitForWorkbookReady(page)

  const grid = page.getByTestId('sheet-grid')

  await grid.click({
    position: {
      x: PRODUCT_ROW_MARKER_WIDTH + PRODUCT_COLUMN_WIDTH + Math.floor(PRODUCT_COLUMN_WIDTH / 2),
      y: Math.floor(PRODUCT_HEADER_HEIGHT / 2),
    },
  })
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!B:B')

  await grid.click({
    position: {
      x: Math.floor(PRODUCT_ROW_MARKER_WIDTH / 2),
      y: PRODUCT_HEADER_HEIGHT + PRODUCT_ROW_HEIGHT + Math.floor(PRODUCT_ROW_HEIGHT / 2),
    },
  })
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!2:2')
})

test('web app supports row and column header drag selection', async ({ page }) => {
  await page.goto('/')
  await waitForWorkbookReady(page)

  await dragProductHeaderSelection(page, 'column', 1, 3)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!B:D')

  await dragProductHeaderSelection(page, 'row', 1, 3)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!2:4')
})

test('@browser-ci web app commits an in-cell edit before applying a header selection', async ({ page }) => {
  const documentId = createTestDocumentId('playwright-header-click-away-edit')
  await page.goto(`/?document=${encodeURIComponent(documentId)}&persist=0`)
  await waitForWorkbookReady(page)

  const draft = 'header-click-away-draft'
  await clickProductCell(page, 1, 1)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!B2')
  await page.getByTestId('sheet-grid-focus-target').focus()
  await page.keyboard.type(draft)
  await expect(page.getByTestId('cell-editor-input')).toHaveValue(draft)

  const grid = page.getByTestId('sheet-grid')
  await grid.click({
    position: {
      x: PRODUCT_ROW_MARKER_WIDTH + PRODUCT_COLUMN_WIDTH * 2 + Math.floor(PRODUCT_COLUMN_WIDTH / 2),
      y: Math.floor(PRODUCT_HEADER_HEIGHT / 2),
    },
  })
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!C:C')

  await selectAddress(page, 'B2')
  await expect(page.getByTestId('formula-input')).toHaveValue(draft)

  await selectAddress(page, 'C1')
  await expect(page.getByTestId('formula-input')).toHaveValue('')
})

test('web app deletes the selected row range from the header context menu', async ({ page }) => {
  const documentId = createTestDocumentId('playwright-delete-selected-rows')
  await page.goto(`/?document=${encodeURIComponent(documentId)}`)
  await waitForWorkbookReady(page)

  await writeCellValue(page, 'A2', 'row-2')
  await writeCellValue(page, 'A3', 'row-3')
  await writeCellValue(page, 'A4', 'row-4')

  await dragProductHeaderSelection(page, 'row', 1, 2)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!2:3')

  await rightClickProductRowHeader(page, 2)
  await page.getByTestId('grid-context-action-delete-row').click()
  await expect(page.getByTestId('grid-context-menu')).toBeHidden({ timeout: 30_000 })

  await selectAddress(page, 'A2')
  await expect.poll(() => readFormulaValue(page)).toBe('row-4')

  await selectAddress(page, 'A3')
  await expect.poll(() => readFormulaValue(page)).toBe('')
})

test('web app opens the selected row context menu with the keyboard shortcut', async ({ page }) => {
  const documentId = createTestDocumentId('playwright-keyboard-context-menu')
  await page.goto(`/?document=${encodeURIComponent(documentId)}`)
  await waitForWorkbookReady(page)

  await dragProductHeaderSelection(page, 'row', 1, 1)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!2:2')

  await page.keyboard.down(PRIMARY_MODIFIER)
  await page.keyboard.down('Shift')
  await page.keyboard.press('Backslash')
  await page.keyboard.up('Shift')
  await page.keyboard.up(PRIMARY_MODIFIER)

  await expect(page.getByTestId('grid-context-menu')).toBeVisible()
  await expect(page.getByTestId('grid-context-action-delete-row')).toBeVisible()
})

test('web app deletes selected rows with the structural delete keyboard shortcut', async ({ page }) => {
  const documentId = createTestDocumentId('playwright-keyboard-delete-selected-rows')
  await page.goto(`/?document=${encodeURIComponent(documentId)}`)
  await waitForWorkbookReady(page)

  await writeCellValue(page, 'A2', 'row-2')
  await writeCellValue(page, 'A3', 'row-3')
  await writeCellValue(page, 'A4', 'row-4')

  await dragProductHeaderSelection(page, 'row', 1, 2)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!2:3')

  await pressStructuralDeleteShortcut(page)

  await selectAddress(page, 'A2')
  await expect.poll(() => readFormulaValue(page)).toBe('row-4')

  await selectAddress(page, 'A3')
  await expect.poll(() => readFormulaValue(page)).toBe('')
})

test('web app deletes the selected column range from the header context menu', async ({ page }) => {
  const documentId = createTestDocumentId('playwright-delete-selected-columns')
  await page.goto(`/?document=${encodeURIComponent(documentId)}`)
  await waitForWorkbookReady(page)

  await writeCellValue(page, 'B1', 'col-b')
  await writeCellValue(page, 'C1', 'col-c')
  await writeCellValue(page, 'D1', 'col-d')

  await dragProductHeaderSelection(page, 'column', 1, 2)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!B:C')

  await rightClickProductColumnHeader(page, 2)
  await page.getByTestId('grid-context-action-delete-column').click()
  await expect(page.getByTestId('grid-context-menu')).toBeHidden({ timeout: 30_000 })

  await selectAddress(page, 'B1')
  await expect.poll(() => readFormulaValue(page)).toBe('col-d')

  await selectAddress(page, 'C1')
  await expect.poll(() => readFormulaValue(page)).toBe('')
})

test('web app deletes selected columns with the structural delete keyboard shortcut', async ({ page }) => {
  const documentId = createTestDocumentId('playwright-keyboard-delete-selected-columns')
  await page.goto(`/?document=${encodeURIComponent(documentId)}`)
  await waitForWorkbookReady(page)

  await writeCellValue(page, 'B1', 'col-b')
  await writeCellValue(page, 'C1', 'col-c')
  await writeCellValue(page, 'D1', 'col-d')

  await dragProductHeaderSelection(page, 'column', 1, 2)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!B:C')

  await pressStructuralDeleteShortcut(page)

  await selectAddress(page, 'B1')
  await expect.poll(() => readFormulaValue(page)).toBe('col-d')

  await selectAddress(page, 'C1')
  await expect.poll(() => readFormulaValue(page)).toBe('')
})

test('web app clears the selected row range with Delete after header selection', async ({ page }) => {
  const documentId = createTestDocumentId('playwright-clear-selected-rows')
  await page.goto(`/?document=${encodeURIComponent(documentId)}`)
  await waitForWorkbookReady(page)

  await writeCellValue(page, 'A2', 'row-2')
  await writeCellValue(page, 'A3', 'row-3')
  await writeCellValue(page, 'A4', 'row-4')

  await dragProductHeaderSelection(page, 'row', 1, 2)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!2:3')

  await page.keyboard.press('Delete')

  await selectAddress(page, 'A2')
  await expect.poll(() => readFormulaValue(page)).toBe('')

  await selectAddress(page, 'A3')
  await expect.poll(() => readFormulaValue(page)).toBe('')

  await selectAddress(page, 'A4')
  await expect.poll(() => readFormulaValue(page)).toBe('row-4')
})

test('web app clears the selected column range with Delete after header selection', async ({ page }) => {
  const documentId = createTestDocumentId('playwright-clear-selected-columns')
  await page.goto(`/?document=${encodeURIComponent(documentId)}`)
  await waitForWorkbookReady(page)

  await writeCellValue(page, 'B1', 'col-b')
  await writeCellValue(page, 'C1', 'col-c')
  await writeCellValue(page, 'D1', 'col-d')

  await dragProductHeaderSelection(page, 'column', 1, 2)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!B:C')

  await page.keyboard.press('Delete')

  await selectAddress(page, 'B1')
  await expect.poll(() => readFormulaValue(page)).toBe('')

  await selectAddress(page, 'C1')
  await expect.poll(() => readFormulaValue(page)).toBe('')

  await selectAddress(page, 'D1')
  await expect.poll(() => readFormulaValue(page)).toBe('col-d')
})

test('web app supports rectangular drag selection', async ({ page }) => {
  await page.goto('/')
  await waitForWorkbookReady(page)

  await dragProductBodySelection(page, 1, 1, 3, 3)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!B2:D4')
})

test('web app preserves the active cell inside a selected area and collapses on body click', async ({ page }) => {
  await page.goto(`/?document=${encodeURIComponent(createTestDocumentId('playwright-range-active-collapse'))}`)
  await waitForWorkbookReady(page)

  await dragProductBodySelection(page, 3, 3, 1, 1)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!B2:D4')
  await expect(page.getByTestId('name-box')).toHaveValue('B2:D4')
  await expect(page.getByTestId('sheet-grid-focus-target')).toHaveAttribute('aria-label', 'Sheet1 D4')
  await expect(page.locator('[data-grid-selection-visual-role="selection-fill"]')).toHaveCount(1)
  await expect(page.locator('[data-grid-selection-visual-role="selection-border"]')).toHaveCount(1)
  await expect(page.locator('[data-grid-selection-visual-role="active-border"]')).toHaveCount(1)
  await expect(page.locator('[data-grid-selection-visual-role="fill-handle"]')).toHaveCount(1)

  await clickProductCell(page, 2, 2)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!C3')
  await expect(page.getByTestId('name-box')).toHaveValue('C3')
  await expect(page.getByTestId('sheet-grid-focus-target')).toHaveAttribute('aria-label', 'Sheet1 C3')
  await expect(page.locator('[data-grid-selection-visual-role="selection-fill"]')).toHaveCount(0)
  await expect(page.locator('[data-grid-selection-visual-role="active-border"]')).toHaveCount(0)
  await expect(page.locator('[data-grid-selection-visual-role="selection-border"]')).toHaveCount(1)
})

test('@browser-ci web app keeps reverse-drag range selection chrome geometrically aligned', async ({ page }) => {
  await page.goto(`/?document=${encodeURIComponent(createTestDocumentId('playwright-range-visual-geometry'))}&persist=0`)
  await waitForWorkbookReady(page)

  await dragProductBodySelection(page, 3, 3, 1, 1)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!B2:D4')
  await expect(page.getByTestId('name-box')).toHaveValue('B2:D4')
  await expect(page.getByTestId('sheet-grid-focus-target')).toHaveAttribute('aria-label', 'Sheet1 D4')

  const expectedRange = await getProductCellRangeBox(page, 1, 1, 3, 3)
  const expectedActiveCell = await getProductCellRangeBox(page, 3, 3, 3, 3)
  const expectedFillHandle = {
    x: expectedRange.x + expectedRange.width - 3.5,
    y: expectedRange.y + expectedRange.height - 3.5,
    width: 7,
    height: 7,
  }

  await expectVisualRectNear(page.locator('[data-grid-selection-visual-role="selection-border"]'), expectedRange, 'selection border')
  await expectVisualRectNear(page.locator('[data-grid-selection-visual-role="active-border"]'), expectedActiveCell, 'active cell border')
  await expectVisualRectNear(page.locator('[data-grid-selection-visual-role="fill-handle"]'), expectedFillHandle, 'fill handle')
})

test('web app clips spilled text before the active selected cell', async ({ page }) => {
  await page.goto(`/?document=${encodeURIComponent(createTestDocumentId('playwright-selection-spill-clip'))}&persist=0`)
  await waitForWorkbookReady(page)

  await writeCellValue(page, 'A5', 'sfasf sfasf sfasf sfasf sfasf sfasf')
  await selectAddress(page, 'C5')
  await expect(page.getByTestId('name-box')).toHaveValue('C5')

  const runBox = await page.locator('[data-native-text-run-row="4"][data-native-text-run-col="0"]').boundingBox()
  const gridBox = await page.getByTestId('sheet-grid').boundingBox()
  if (!runBox || !gridBox) {
    throw new Error('grid text run is not visible')
  }
  const selectedColumnLeft = gridBox.x + (await getProductColumnLeft(page, 2))

  expect(runBox.x + runBox.width).toBeLessThanOrEqual(selectedColumnLeft + 0.5)
})

test('web app clips spilled text before far horizontally scrolled selections', async ({ page }) => {
  await page.setViewportSize({ width: 1166, height: 820 })
  await page.goto(`/?document=${encodeURIComponent(createTestDocumentId('playwright-far-selection-spill-clip'))}&persist=0`)
  await waitForWorkbookReady(page)

  const rowIndex = 4
  const sourceColumnIndex = 128
  const selectedColumnIndex = 130
  await writeCellValue(page, formatGridAddress(rowIndex, sourceColumnIndex), 'far spill text far spill text far spill text')
  await selectAddress(page, formatGridAddress(rowIndex, selectedColumnIndex))

  const runBox = await page
    .locator(`[data-native-text-run-row="${String(rowIndex)}"][data-native-text-run-col="${String(sourceColumnIndex)}"]`)
    .boundingBox()
  const selectionBorderBox = await page.locator('[data-grid-selection-visual-role="selection-border"]').boundingBox()
  if (!runBox || !selectionBorderBox) {
    throw new Error('far grid text run or selection border is not visible')
  }

  expect(runBox.x + runBox.width).toBeLessThanOrEqual(selectionBorderBox.x + 0.5)
})

test('web app clips spilled text before selected whole columns on non-active rows', async ({ page }) => {
  await page.goto(`/?document=${encodeURIComponent(createTestDocumentId('playwright-column-selection-spill-clip'))}&persist=0`)
  await waitForWorkbookReady(page)

  await writeCellValue(page, 'A5', 'column spill text column spill text column spill text')
  await selectAddress(page, 'B1')
  const grid = page.getByTestId('sheet-grid')
  await grid.click({
    position: {
      x: PRODUCT_ROW_MARKER_WIDTH + PRODUCT_COLUMN_WIDTH * 2 + Math.floor(PRODUCT_COLUMN_WIDTH / 2),
      y: Math.floor(PRODUCT_HEADER_HEIGHT / 2),
    },
  })
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!C:C')

  const runBox = await page.locator('[data-native-text-run-row="4"][data-native-text-run-col="0"]').boundingBox()
  const gridBox = await grid.boundingBox()
  if (!runBox || !gridBox) {
    throw new Error('column-selection spill proof target is not visible')
  }
  const selectedColumnLeft = gridBox.x + (await getProductColumnLeft(page, 2))

  expect(runBox.x + runBox.width).toBeLessThanOrEqual(selectedColumnLeft + 0.5)
})

test('web app keeps moved range data visible when border drag reaches the grid edge', async ({ page }) => {
  await page.setViewportSize({ width: 960, height: 420 })
  await page.goto(`/?document=${encodeURIComponent(createTestDocumentId('range-border-edge-drag'))}`)
  await waitForWorkbookReady(page)

  const formulaInput = page.getByTestId('formula-input')

  await selectAddress(page, 'B2')
  await formulaInput.fill('left')
  await formulaInput.press('Enter')

  await selectAddress(page, 'C2')
  await formulaInput.fill('right')
  await formulaInput.press('Enter')

  await selectAddress(page, 'B2')
  await expect(formulaInput).toHaveValue('left')
  await selectAddress(page, 'C2')
  await expect(formulaInput).toHaveValue('right')

  await dragProductBodySelection(page, 1, 1, 2, 1)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!B2:C2')

  await dragSelectedRangeBorderTowardBottom(page)
  await expect.poll(() => getGridScrollTop(page)).toBeGreaterThan(0)
  await page.mouse.up()

  const selection = (await page.getByTestId('status-selection').textContent()) ?? ''
  const match = /^Sheet1!B(\d+):C\1$/.exec(selection)
  expect(match).not.toBeNull()
  const targetRow = Number(match?.[1] ?? 0)
  expect(targetRow).toBeGreaterThan(2)

  await selectAddress(page, `B${targetRow}`)
  await expect(formulaInput).toHaveValue('left')

  await selectAddress(page, `C${targetRow}`)
  await expect(formulaInput).toHaveValue('right')

  await selectAddress(page, 'B2')
  await expect(formulaInput).toHaveValue('')
})

test('web app keeps the active focus inside the sheet grid when clicking a cell', async ({ page }) => {
  await page.goto('/')
  await waitForWorkbookReady(page)

  await clickProductCell(page, 2, 2)
  await expect(page.getByTestId('name-box')).toHaveValue('C3')

  const activeElementState = await page.evaluate(() => {
    const active = document.activeElement
    return {
      testId: active?.getAttribute('data-testid') ?? null,
      insideSheetGrid: Boolean(active?.closest('[data-testid="sheet-grid"]')),
    }
  })

  expect(activeElementState.insideSheetGrid).toBe(true)
  expect(activeElementState.testId).not.toBe('sheet-grid')
})

test('@browser-perf web app keeps normal cell selection out of resident scene invalidation', async ({ page }) => {
  await page.goto(`/?document=${encodeURIComponent(createTestDocumentId('normal-selection-no-resident-refresh'))}`)
  await waitForWorkbookReady(page)
  await settleWorkbookScrollPerf(page, 80)
  await startWorkbookScrollPerf(page, 'normal-selection-no-resident-refresh', { primeRenderer: false })
  await settleWorkbookScrollPerf(page, 24)

  const targets = [
    [1, 1],
    [2, 2],
    [3, 3],
    [4, 4],
    [5, 5],
    [2, 6],
    [6, 3],
    [1, 4],
  ] as const
  await targets.reduce<Promise<void>>(async (previous, [col, row]) => {
    await previous
    await clickProductCell(page, col, row)
    await expect(page.getByTestId('status-selection')).toContainText('!')
  }, Promise.resolve())

  await settleWorkbookScrollPerf(page, 24)
  const report = await stopWorkbookScrollPerf(page)

  expect(report).not.toBeNull()
  expect(report?.counters.rendererTileMisses).toBe(0)
  expect(report?.counters.typeGpuBufferAllocations).toBe(0)
  expect(report?.counters.typeGpuTileMisses).toBe(0)
})

test('@browser-perf web app keeps range-move preview out of resident scene invalidation', async ({ page }) => {
  await page.goto(`/?document=${encodeURIComponent(createTestDocumentId('range-move-no-resident-refresh'))}`)
  await waitForWorkbookReady(page)

  const nameBox = page.getByTestId('name-box')
  const formulaInput = page.getByTestId('formula-input')

  await nameBox.fill('B2')
  await nameBox.press('Enter')
  await formulaInput.fill('left')
  await formulaInput.press('Enter')

  await nameBox.fill('C2')
  await nameBox.press('Enter')
  await formulaInput.fill('right')
  await formulaInput.press('Enter')

  await dragProductBodySelection(page, 1, 1, 2, 1)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!B2:C2')

  await settleWorkbookScrollPerf(page, 80)
  await startWorkbookScrollPerf(page, 'range-move-no-resident-refresh', { primeRenderer: false })
  await settleWorkbookScrollPerf(page, 24)

  let report: Awaited<ReturnType<typeof stopWorkbookScrollPerf>> | null = null
  try {
    await dragSelectedRangeBorderPreview(page, 1, 1, 3, 3)
    await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!D4:E4')
    await settleWorkbookScrollPerf(page, 24)
    report = await stopWorkbookScrollPerf(page)
  } finally {
    await page.mouse.up()
  }

  expect(report).not.toBeNull()
  expect(report?.counters.rendererTileMisses).toBe(0)
  expect(report?.counters.typeGpuBufferAllocations).toBe(0)
  expect(report?.counters.typeGpuTileMisses).toBe(0)
})

test('web app maps clicks in the upper half of a cell to that same visible cell', async ({ page }) => {
  await page.goto('/')
  await waitForWorkbookReady(page)

  await clickProductCellUpperHalf(page, 4, 11)
  await expect(page.getByTestId('name-box')).toHaveValue('E12')
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!E12')

  await clickProductCellUpperHalf(page, 2, 4)
  await expect(page.getByTestId('name-box')).toHaveValue('C5')
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!C5')
})

test('web app maps pointer selection exactly after large vertical and horizontal scroll', async ({ page }) => {
  await page.setViewportSize({ width: 1166, height: 820 })
  await page.goto(`/?document=${encodeURIComponent(createTestDocumentId('large-scroll-selection-hit-test'))}`)
  await waitForWorkbookReady(page)

  await selectAddress(page, formatGridAddress(100_000, 500))
  const scroll = await readProductViewportScroll(page)
  expect(scroll.scrollLeft).toBeGreaterThanOrEqual(PRODUCT_COLUMN_WIDTH * 490)
  expect(scroll.scrollTop).toBeGreaterThanOrEqual(PRODUCT_ROW_HEIGHT * 99_000)

  const expectedAddress = await clickVisibleScrolledBodyCell(page, 2, 3)
  await expect(page.getByTestId('name-box')).toHaveValue(expectedAddress)
  await expect(page.getByTestId('status-selection')).toHaveText(`Sheet1!${expectedAddress}`)
})

test('web app supports column resize without breaking hit testing', async ({ page }) => {
  const documentId = createTestDocumentId('playwright-column-resize-hit-test')
  await page.goto(`/?document=${encodeURIComponent(documentId)}`)
  await waitForWorkbookReady(page)

  await clickProductBodyOffset(page, 82, 0)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!A1')

  await dragProductColumnResize(page, 0, -36)

  await clickProductBodyOffset(page, 82, 0)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!B1')
})

test('web app supports column edge double-click autofit', async ({ page }) => {
  const documentId = createTestDocumentId('playwright-column-autofit')
  await page.goto(`/?document=${encodeURIComponent(documentId)}`)
  await waitForWorkbookReady(page)

  const nameBox = page.getByTestId('name-box')
  const formulaInput = page.getByTestId('formula-input')
  const longValue = 'supercalifragilisticexpialidocious'

  await nameBox.fill('A1')
  await nameBox.press('Enter')
  await formulaInput.fill(longValue)
  await formulaInput.press('Enter')

  await clickProductCell(page, 0, 0)
  await expect(formulaInput).toHaveValue(longValue)

  await clickProductBodyOffset(page, 126, 0)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!B1')

  await clickProductCell(page, 0, 0)
  await expect(formulaInput).toHaveValue(longValue)
  await doubleClickProductColumnResizeHandle(page, 0)
  await expect.poll(async () => await getProductColumnWidth(page, 0), { timeout: 15_000 }).toBeGreaterThan(126)
  const autofitWidth = await getProductColumnWidth(page, 0)

  await clickProductBodyOffset(page, autofitWidth + 8, 0)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!B1')
  await clickProductBodyOffset(page, 126, 0)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!A1')
})

test('web app hit-tests typegpu geometry after hiding rows and columns', async ({ page }) => {
  const documentId = createTestDocumentId('playwright-hidden-axis-hit-test')
  await page.goto(`/?document=${encodeURIComponent(documentId)}`)
  await waitForWorkbookReady(page)

  await clickProductCell(page, 1, 1)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!B2')
  const gridLocator = page.getByTestId('sheet-grid')
  await gridLocator.click({
    button: 'right',
    position: {
      x: PRODUCT_ROW_MARKER_WIDTH + PRODUCT_COLUMN_WIDTH + Math.floor(PRODUCT_COLUMN_WIDTH / 2),
      y: Math.floor(PRODUCT_HEADER_HEIGHT / 2),
    },
  })
  await page.getByTestId('grid-context-action-hide-column').click()
  await expect.poll(() => getProductColumnWidth(page, 1)).toBe(0)
  await settleWorkbookScrollPerf(page, 4)

  await clickProductCell(page, 2, 1)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!C2')

  await gridLocator.click({
    button: 'right',
    position: {
      x: Math.floor(PRODUCT_ROW_MARKER_WIDTH / 2),
      y: PRODUCT_HEADER_HEIGHT + PRODUCT_ROW_HEIGHT + Math.floor(PRODUCT_ROW_HEIGHT / 2),
    },
  })
  await page.getByTestId('grid-context-action-hide-row').click()
  await expect.poll(() => getProductRowHeight(page, 1)).toBe(0)
  await settleWorkbookScrollPerf(page, 4)
  const grid = await gridLocator.boundingBox()
  if (!grid) {
    throw new Error('sheet grid is not visible')
  }
  const columnLeft = await getProductColumnLeft(page, 2)
  const columnWidth = await getProductColumnWidth(page, 2)
  const rowTop = await getProductRowTop(page, 2)
  const rowHeight = await getProductRowHeight(page, 2)
  await page.mouse.click(
    grid.x + columnLeft + Math.floor(columnWidth / 2),
    grid.y + PRODUCT_HEADER_HEIGHT + rowTop + Math.floor(rowHeight / 2),
  )
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!C3')
})

async function dragSelectedRangeBorderTowardBottom(page: Page): Promise<void> {
  const gridLocator = page.getByTestId('sheet-grid')
  await expect(gridLocator).toBeVisible()
  const grid = await gridLocator.boundingBox()
  if (!grid) {
    throw new Error('sheet grid is not visible')
  }

  const startLeft = await getProductColumnLeft(page, 1)
  const targetLeft = await getProductColumnLeft(page, 1)
  const targetWidth = await getProductColumnWidth(page, 1)
  const sourceX = grid.x + startLeft + 3
  const sourceY = grid.y + PRODUCT_HEADER_HEIGHT + PRODUCT_ROW_HEIGHT + 2
  const targetX = grid.x + targetLeft + Math.floor(targetWidth / 2)
  const targetY = grid.y + grid.height - 24

  await page.mouse.move(sourceX, sourceY)
  await page.mouse.down()
  await page.mouse.move(targetX, targetY, { steps: 16 })
}

async function getProductCellRangeBox(
  page: Page,
  startColumn: number,
  startRow: number,
  endColumn: number,
  endRow: number,
): Promise<{
  readonly x: number
  readonly y: number
  readonly width: number
  readonly height: number
}> {
  const gridLocator = page.getByTestId('sheet-grid')
  await expect(gridLocator).toBeVisible()
  const grid = await gridLocator.boundingBox()
  if (!grid) {
    throw new Error('sheet grid is not visible')
  }
  const leftColumn = Math.min(startColumn, endColumn)
  const rightColumn = Math.max(startColumn, endColumn)
  const topRow = Math.min(startRow, endRow)
  const bottomRow = Math.max(startRow, endRow)
  const left = await getProductColumnLeft(page, leftColumn)
  const right = await getProductColumnLeft(page, rightColumn)
  const rightWidth = await getProductColumnWidth(page, rightColumn)
  const top = await getProductRowTop(page, topRow)
  const bottom = await getProductRowTop(page, bottomRow)
  const bottomHeight = await getProductRowHeight(page, bottomRow)
  return {
    x: grid.x + left,
    y: grid.y + PRODUCT_HEADER_HEIGHT + top,
    width: right + rightWidth - left,
    height: bottom + bottomHeight - top,
  }
}

async function expectVisualRectNear(
  locator: ReturnType<Page['locator']>,
  expected: {
    readonly x: number
    readonly y: number
    readonly width: number
    readonly height: number
  },
  label: string,
): Promise<void> {
  await expect(locator, `${label} should be unique`).toHaveCount(1)
  const actual = await locator.boundingBox()
  if (!actual) {
    throw new Error(`${label} is not visible`)
  }
  expect(actual.x, `${label} x`).toBeCloseTo(expected.x, 0)
  expect(actual.y, `${label} y`).toBeCloseTo(expected.y, 0)
  expect(actual.width, `${label} width`).toBeCloseTo(expected.width, 0)
  expect(actual.height, `${label} height`).toBeCloseTo(expected.height, 0)
}

async function dragSelectedRangeBorderPreview(
  page: Page,
  startColumn: number,
  startRow: number,
  targetColumn: number,
  targetRow: number,
): Promise<void> {
  const gridLocator = page.getByTestId('sheet-grid')
  await expect(gridLocator).toBeVisible()
  const grid = await gridLocator.boundingBox()
  if (!grid) {
    throw new Error('sheet grid is not visible')
  }

  const startLeft = await getProductColumnLeft(page, startColumn)
  const sourceX = grid.x + startLeft + 3
  const sourceY = grid.y + PRODUCT_HEADER_HEIGHT + startRow * PRODUCT_ROW_HEIGHT + 2
  const targetLeft = await getProductColumnLeft(page, targetColumn)
  const targetWidth = await getProductColumnWidth(page, targetColumn)
  const targetX = grid.x + targetLeft + Math.floor(targetWidth / 2)
  const targetY = grid.y + PRODUCT_HEADER_HEIGHT + targetRow * PRODUCT_ROW_HEIGHT + Math.floor(PRODUCT_ROW_HEIGHT / 2)

  await page.mouse.move(sourceX, sourceY)
  await page.mouse.down()
  await page.mouse.move(targetX, targetY, { steps: 12 })
}

async function getGridScrollTop(page: Page): Promise<number> {
  return await page.getByTestId('grid-scroll-viewport').evaluate((viewport) => {
    if (!(viewport instanceof HTMLDivElement)) {
      throw new Error('grid scroll viewport is not an HTMLDivElement')
    }
    return viewport.scrollTop
  })
}

async function readProductViewportScroll(page: Page): Promise<{
  readonly scrollLeft: number
  readonly scrollTop: number
}> {
  await settleWorkbookScrollPerf(page, 8)
  return await page.getByTestId('grid-scroll-viewport').evaluate((node) => {
    if (!(node instanceof HTMLDivElement)) {
      throw new Error('grid scroll viewport is not an HTMLDivElement')
    }
    return {
      scrollLeft: node.scrollLeft,
      scrollTop: node.scrollTop,
    }
  })
}

async function clickVisibleScrolledBodyCell(page: Page, visibleColumnOffset: number, visibleRowOffset: number): Promise<string> {
  const gridLocator = page.getByTestId('sheet-grid')
  await expect(gridLocator).toBeVisible()
  const scroll = await page.getByTestId('grid-scroll-viewport').evaluate((node) => {
    if (!(node instanceof HTMLDivElement)) {
      throw new Error('grid scroll viewport is not an HTMLDivElement')
    }
    return {
      scrollLeft: node.scrollLeft,
      scrollTop: node.scrollTop,
    }
  })
  const bodyX = visibleColumnOffset * PRODUCT_COLUMN_WIDTH + Math.floor(PRODUCT_COLUMN_WIDTH / 2)
  const bodyY = visibleRowOffset * PRODUCT_ROW_HEIGHT + Math.floor(PRODUCT_ROW_HEIGHT / 2)
  const expectedColumnIndex = Math.floor((scroll.scrollLeft + bodyX) / PRODUCT_COLUMN_WIDTH)
  const expectedRowIndex = Math.floor((scroll.scrollTop + bodyY) / PRODUCT_ROW_HEIGHT)

  await gridLocator.click({
    position: {
      x: PRODUCT_ROW_MARKER_WIDTH + bodyX,
      y: PRODUCT_HEADER_HEIGHT + bodyY,
    },
  })

  return formatGridAddress(expectedRowIndex, expectedColumnIndex)
}

function formatGridAddress(rowIndex: number, columnIndex: number): string {
  return `${formatColumnLabel(columnIndex)}${String(rowIndex + 1)}`
}

function formatColumnLabel(columnIndex: number): string {
  let remaining = columnIndex + 1
  let label = ''
  while (remaining > 0) {
    const next = (remaining - 1) % 26
    label = String.fromCharCode(65 + next) + label
    remaining = Math.floor((remaining - 1) / 26)
  }
  return label
}

async function rightClickProductRowHeader(page: Page, rowIndex: number): Promise<void> {
  const gridLocator = page.getByTestId('sheet-grid')
  await expect(gridLocator).toBeVisible()
  const grid = await gridLocator.boundingBox()
  if (!grid) {
    throw new Error('sheet grid is not visible')
  }

  const rowTop = await getProductRowTop(page, rowIndex)
  const rowHeight = await getProductRowHeight(page, rowIndex)
  await gridLocator.click({
    button: 'right',
    position: {
      x: Math.floor(PRODUCT_ROW_MARKER_WIDTH / 2),
      y: PRODUCT_HEADER_HEIGHT + rowTop + Math.floor(rowHeight / 2),
    },
  })
}

async function rightClickProductColumnHeader(page: Page, columnIndex: number): Promise<void> {
  const gridLocator = page.getByTestId('sheet-grid')
  await expect(gridLocator).toBeVisible()
  const grid = await gridLocator.boundingBox()
  if (!grid) {
    throw new Error('sheet grid is not visible')
  }

  const columnLeft = await getProductColumnLeft(page, columnIndex)
  const columnWidth = await getProductColumnWidth(page, columnIndex)
  await gridLocator.click({
    button: 'right',
    position: {
      x: columnLeft + Math.floor(columnWidth / 2),
      y: Math.floor(PRODUCT_HEADER_HEIGHT / 2),
    },
  })
}

async function writeCellValue(page: Page, address: string, value: string): Promise<void> {
  await selectAddress(page, address)
  const formulaInput = page.getByTestId('formula-input')
  await formulaInput.fill(value)
  await formulaInput.press('Enter')
}

async function selectAddress(page: Page, address: string): Promise<void> {
  const nameBox = page.getByTestId('name-box')
  await nameBox.fill(address)
  await expect(nameBox).toHaveValue(address)
  await nameBox.press('Enter')
  await expect(page.getByTestId('status-selection')).toHaveText(`Sheet1!${address}`)
}

async function readFormulaValue(page: Page): Promise<string> {
  const formulaInput = page.getByTestId('formula-input')
  return await formulaInput.inputValue()
}

async function pressStructuralDeleteShortcut(page: Page): Promise<void> {
  await page.keyboard.down(PRIMARY_MODIFIER)
  await page.keyboard.down('Alt')
  await page.keyboard.press('Minus')
  await page.keyboard.up('Alt')
  await page.keyboard.up(PRIMARY_MODIFIER)
}
