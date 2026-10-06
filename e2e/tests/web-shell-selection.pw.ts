import { expect, test } from '@playwright/test'
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
  doubleClickProductColumnResizeHandle,
  dragProductBodySelection,
  dragProductColumnResize,
  dragProductHeaderSelection,
  getProductColumnLeft,
  getProductColumnWidth,
  getProductFillHandleDragPoints,
  getProductRowHeight,
  getProductRowTop,
  installTypeGpuCellReadbackHarness,
  settleWorkbookScrollPerf,
  startWorkbookScrollPerf,
  stopWorkbookScrollPerf,
  waitForTypeGpuVisibleFrame,
  waitForWorkbookReady,
} from './web-shell-helpers.js'
import {
  dragProductSelectedInterior,
  dragSelectedRangeBorderPreview,
  dragSelectedRangeBorderTowardBottom,
  expectBorderStyle,
  expectCellTextPixels,
  expectSelectedRangeBodyTint,
  expectSelectionVisualRoles,
  expectVisualRectNear,
  getGridScrollTop,
  getProductCellRangeBox,
  isGridBorderPixel,
  isSelectionAccentPixel,
  isTypeGpuTextReadbackActive,
  maxPixelChannelDistance,
  pressStructuralDeleteShortcut,
  readFormulaValue,
  readProductViewportScroll,
  readSelectedRowHeaderNativeTextProof,
  clickVisibleScrolledBodyCell,
  formatGridAddress,
  rightClickProductColumnHeader,
  rightClickProductRowHeader,
  sampleCompositedViewportPixel,
  selectAddress,
  writeCellValue,
} from './web-shell-selection-helpers.js'

test.beforeEach(async ({ page }) => {
  await installTypeGpuCellReadbackHarness(page)
})

test('web app keeps sheet tabs and status bar visible in a short viewport', async ({ page }) => {
  await page.setViewportSize({ width: 2048, height: 220 })
  await page.goto('/workbook')
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
  await page.goto('/workbook')
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
  await page.goto('/workbook')
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
  await page.goto('/workbook')
  await waitForWorkbookReady(page)

  await dragProductBodySelection(page, 1, 1, 3, 3)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!B2:D4')
})

test('web app paints selected areas as crisp cell interiors and collapses on body click', async ({ page }) => {
  await page.goto(`/?document=${encodeURIComponent(createTestDocumentId('playwright-range-active-collapse'))}`)
  await waitForWorkbookReady(page)

  await dragProductBodySelection(page, 3, 3, 1, 1)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!B2:D4')
  await expect(page.getByTestId('name-box')).toHaveValue('B2:D4')
  await expect(page.getByTestId('sheet-grid-focus-target')).toHaveAttribute('aria-label', 'Sheet1 D4')
  await expect(page.locator('[data-grid-selection-visual-role="selection-fill"]')).toHaveCount(1)
  await expect(page.locator('[data-grid-selection-visual-role="selection-gridline"]')).toHaveCount(4)
  await expect(page.locator('[data-grid-selection-visual-role="selection-border"]')).toHaveCount(1)
  await expect(page.locator('[data-grid-selection-visual-role="active-border"]')).toHaveCount(0)
  await expect(page.locator('[data-grid-selection-visual-role="fill-handle"]')).toHaveCount(1)
  await expectSelectionVisualRoles(page, ['header-fill'], 'visible')
  await expectSelectionVisualRoles(page, ['selection-fill'], 'visible')
  await expectSelectionVisualRoles(page, ['selection-border', 'fill-handle'], 'visible')
  await expect(page.locator('[data-grid-selection-visual-role="selection-fill"]').first()).toHaveCSS(
    'background-color',
    'rgba(33, 115, 70, 0.22)',
  )
  await expectSelectedRangeBodyTint(page, 3, 3)
  await expectSelectedRangeBodyTint(page, 1, 1)
  await expectSelectedRangeBodyTint(page, 1, 2)

  await clickProductCell(page, 2, 2)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!C3')
  await expect(page.getByTestId('name-box')).toHaveValue('C3')
  await expect(page.getByTestId('sheet-grid-focus-target')).toHaveAttribute('aria-label', 'Sheet1 C3')
  await expect(page.locator('[data-grid-selection-visual-role="selection-fill"]')).toHaveCount(0)
  await expect(page.locator('[data-grid-selection-visual-role="active-border"]')).toHaveCount(1)
  await expect(page.locator('[data-grid-selection-visual-role="selection-border"]')).toHaveCount(0)
  await expectSelectionVisualRoles(page, ['header-fill'], 'visible')
  await expectSelectionVisualRoles(page, ['active-border', 'fill-handle'], 'visible')
})

test('@browser-ci web app paints forward-drag ranges without an internal active-cell box', async ({ page }) => {
  await page.goto(`/?document=${encodeURIComponent(createTestDocumentId('playwright-forward-range-visual'))}&persist=0`)
  await waitForWorkbookReady(page)

  await dragProductBodySelection(page, 1, 1, 5, 12)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!B2:F13')
  await expect(page.getByTestId('name-box')).toHaveValue('B2:F13')
  await expect(page.getByTestId('sheet-grid-focus-target')).toHaveAttribute('aria-label', 'Sheet1 B2')
  await expect(page.locator('[data-grid-selection-visual-role="selection-fill"]')).toHaveCount(1)
  await expect(page.locator('[data-grid-selection-visual-role="selection-gridline"]')).toHaveCount(15)
  await expect(page.locator('[data-grid-selection-visual-role="selection-border"]')).toHaveCount(1)
  await expect(page.locator('[data-grid-selection-visual-role="active-border"]')).toHaveCount(0)
  await expectSelectionVisualRoles(page, ['selection-fill', 'selection-border', 'fill-handle'], 'visible')
  await expectSelectedRangeBodyTint(page, 1, 1)
  await expectSelectedRangeBodyTint(page, 1, 2)
  await expectSelectedRangeBodyTint(page, 3, 6)
  await expectSelectedRangeBodyTint(page, 5, 12)
})

test('@browser-ci web app keeps reverse-drag range selection chrome geometrically aligned', async ({ page }) => {
  await page.goto(`/?document=${encodeURIComponent(createTestDocumentId('playwright-range-visual-geometry'))}&persist=0`)
  await waitForWorkbookReady(page)

  await dragProductBodySelection(page, 3, 3, 1, 1)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!B2:D4')
  await expect(page.getByTestId('name-box')).toHaveValue('B2:D4')
  await expect(page.getByTestId('sheet-grid-focus-target')).toHaveAttribute('aria-label', 'Sheet1 D4')

  const expectedRange = await getProductCellRangeBox(page, 1, 1, 3, 3)
  const expectedFillHandle = {
    x: expectedRange.x + expectedRange.width - 4,
    y: expectedRange.y + expectedRange.height - 4,
    width: 8,
    height: 8,
  }

  await expectVisualRectNear(page.locator('[data-grid-selection-visual-role="selection-border"]'), expectedRange, 'selection border')
  await expect(page.locator('[data-grid-selection-visual-role="active-border"]')).toHaveCount(0)
  await expectVisualRectNear(page.locator('[data-grid-selection-visual-role="fill-handle"]'), expectedFillHandle, 'fill handle')
  await expectBorderStyle(page.locator('[data-grid-selection-visual-role="selection-border"]'), {
    boxShadow: 'none',
    color: 'rgb(33, 115, 70)',
    width: '1px',
  })
  await expect(page.locator('[data-grid-selection-visual-role="fill-handle"]')).toHaveCSS('background-color', 'rgb(33, 115, 70)')
})

test('@browser-ci web app paints active-cell seams as one owned stroke over gridlines', async ({ page }) => {
  await page.goto(`/?document=${encodeURIComponent(createTestDocumentId('playwright-active-cell-single-seam'))}&persist=0`)
  await waitForWorkbookReady(page)

  await selectAddress(page, 'A1')
  await waitForTypeGpuVisibleFrame(page)
  const activeCellBox = await getProductCellRangeBox(page, 0, 0, 0, 0)

  const leftStroke = await sampleCompositedViewportPixel(page, activeCellBox.x - 1, activeCellBox.y + 8)
  const afterLeftStroke = await sampleCompositedViewportPixel(page, activeCellBox.x, activeCellBox.y + 8)
  const topStroke = await sampleCompositedViewportPixel(page, activeCellBox.x + 8, activeCellBox.y - 1)
  const afterTopStroke = await sampleCompositedViewportPixel(page, activeCellBox.x + 8, activeCellBox.y)
  const rightStroke = await sampleCompositedViewportPixel(page, activeCellBox.x + activeCellBox.width - 1, activeCellBox.y + 8)
  const afterRightStroke = await sampleCompositedViewportPixel(page, activeCellBox.x + activeCellBox.width, activeCellBox.y + 8)
  const bottomStroke = await sampleCompositedViewportPixel(page, activeCellBox.x + 8, activeCellBox.y + activeCellBox.height - 1)
  const afterBottomStroke = await sampleCompositedViewportPixel(page, activeCellBox.x + 8, activeCellBox.y + activeCellBox.height)

  expect(isSelectionAccentPixel(leftStroke), 'active cell left seam should replace the header/body gridline').toBe(true)
  expect(isSelectionAccentPixel(afterLeftStroke), 'active cell left seam should not leave a second green line inside the body').toBe(false)
  expect(isGridBorderPixel(afterLeftStroke), 'active cell left seam should not leave a second body gridline beside the header seam').toBe(
    false,
  )
  expect(isSelectionAccentPixel(topStroke), 'active cell top seam should replace the header/body gridline').toBe(true)
  expect(isSelectionAccentPixel(afterTopStroke), 'active cell top seam should not leave a second green line inside the body').toBe(false)
  expect(isGridBorderPixel(afterTopStroke), 'active cell top seam should not leave a second body gridline beside the header seam').toBe(
    false,
  )
  expect(isSelectionAccentPixel(rightStroke), 'active cell right seam should replace the gridline with the selection stroke').toBe(true)
  expect(isSelectionAccentPixel(afterRightStroke), 'active cell right seam should not leave a second green line outside the cell').toBe(
    false,
  )
  expect(isGridBorderPixel(afterRightStroke), 'active cell right seam should not leave a second body gridline outside the cell').toBe(false)
  expect(isSelectionAccentPixel(bottomStroke), 'active cell bottom seam should replace the gridline with the selection stroke').toBe(true)
  expect(isSelectionAccentPixel(afterBottomStroke), 'active cell bottom seam should not leave a second green line outside the cell').toBe(
    false,
  )
  expect(isGridBorderPixel(afterBottomStroke), 'active cell bottom seam should not leave a second body gridline outside the cell').toBe(
    false,
  )
})

test('@browser-ci web app keeps selected row headers and body cells on a single seam', async ({ page }) => {
  await page.goto(`/?document=${encodeURIComponent(createTestDocumentId('playwright-row-header-single-seam'))}&persist=0`)
  await waitForWorkbookReady(page)

  await dragProductHeaderSelection(page, 'row', 7, 14)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!8:15')
  await expect(page.locator('[data-grid-selection-visual-role="active-border"]')).toHaveCount(0)
  await waitForTypeGpuVisibleFrame(page)

  const gridLocator = page.getByTestId('sheet-grid')
  await expect(gridLocator).toBeVisible()
  const grid = await gridLocator.boundingBox()
  if (!grid) {
    throw new Error('sheet grid is not visible')
  }
  const rowTop = await getProductRowTop(page, 7)
  const seamY = grid.y + PRODUCT_HEADER_HEIGHT + rowTop + 10
  const internalRowSeamY = grid.y + PRODUCT_HEADER_HEIGHT + (await getProductRowTop(page, 8))
  const headerInteriorPixel = await sampleCompositedViewportPixel(page, grid.x + PRODUCT_ROW_MARKER_WIDTH - 4, seamY)
  const headerLastPixel = await sampleCompositedViewportPixel(page, grid.x + PRODUCT_ROW_MARKER_WIDTH - 1, seamY)
  const bodyFirstPixel = await sampleCompositedViewportPixel(page, grid.x + PRODUCT_ROW_MARKER_WIDTH, seamY)
  const headerInternalSeamPixel = await sampleCompositedViewportPixel(page, grid.x + PRODUCT_ROW_MARKER_WIDTH - 12, internalRowSeamY)
  const selectedHeaderTextProof = await readSelectedRowHeaderNativeTextProof(page, {
    height: PRODUCT_ROW_HEIGHT * 8,
    width: PRODUCT_ROW_MARKER_WIDTH,
    x: grid.x,
    y: grid.y + PRODUCT_HEADER_HEIGHT + rowTop,
  })

  expect(
    isGridBorderPixel(headerLastPixel),
    'selected row header fill should cover the header/body seam instead of exposing a gray line',
  ).toBe(false)
  expect(isGridBorderPixel(bodyFirstPixel), 'row selection should not paint a second leading body gridline after the header seam').toBe(
    false,
  )
  expect(
    isGridBorderPixel(headerLastPixel) && isGridBorderPixel(bodyFirstPixel),
    'row header/body seam should not show adjacent structural gridlines',
  ).toBe(false)
  expect(
    maxPixelChannelDistance(headerInteriorPixel, headerLastPixel),
    'selected row-header fill should fully cover the static header seam instead of translucent double-painting it',
  ).toBeLessThanOrEqual(1)
  expect(
    maxPixelChannelDistance(headerLastPixel, bodyFirstPixel),
    'row header and selected body fill should meet as one continuous selected surface without a visible double seam',
  ).toBeLessThanOrEqual(1)
  expect(
    maxPixelChannelDistance(headerInteriorPixel, headerInternalSeamPixel),
    'selected row-header range should cover internal row separators instead of drawing a double border between selected headers',
  ).toBeLessThanOrEqual(1)
  expect(selectedHeaderTextProof.textLayerZIndex, 'selected row header text should render above the TypeGPU fill layer').toBeGreaterThan(
    selectedHeaderTextProof.typeGpuLayerZIndex,
  )
  expect(selectedHeaderTextProof.labels, 'selected row header numbers should remain mounted above selection fills').toEqual(
    expect.arrayContaining(['8', '15']),
  )
  expect(
    selectedHeaderTextProof.labels.length,
    'all selected row header numbers should remain mounted above selection fills',
  ).toBeGreaterThanOrEqual(8)
  expect(bodyFirstPixel.green, 'first body pixel should remain selected-row fill, not a border').toBeGreaterThan(bodyFirstPixel.red)
})

test('@browser-ci web app leaves the top-left selector transparent so the header seam has one owner', async ({ page }) => {
  await page.goto(`/?document=${encodeURIComponent(createTestDocumentId('playwright-corner-header-single-owner'))}&persist=0`)
  await waitForWorkbookReady(page)
  await waitForTypeGpuVisibleFrame(page)

  const cornerButton = page.getByTestId('grid-select-entire-sheet')
  await expect(cornerButton).toHaveCSS('border-right-width', '0px')
  await expect(cornerButton).toHaveCSS('border-bottom-width', '0px')
  await expect(cornerButton).toHaveCSS('background-color', 'rgba(0, 0, 0, 0)')

  const gridLocator = page.getByTestId('sheet-grid')
  await expect(gridLocator).toBeVisible()
  const grid = await gridLocator.boundingBox()
  if (!grid) {
    throw new Error('sheet grid is not visible')
  }
  const headerMidY = grid.y + Math.floor(PRODUCT_HEADER_HEIGHT / 2)
  const ownedSeamPixel = await sampleCompositedViewportPixel(page, grid.x + PRODUCT_ROW_MARKER_WIDTH - 1, headerMidY)
  const firstColumnHeaderPixel = await sampleCompositedViewportPixel(page, grid.x + PRODUCT_ROW_MARKER_WIDTH, headerMidY)

  expect(isGridBorderPixel(ownedSeamPixel), 'corner/header seam should have exactly one structural gridline').toBe(true)
  expect(isGridBorderPixel(firstColumnHeaderPixel), 'column header should not start with a second adjacent gridline').toBe(false)
})

test('@browser-ci web app keeps unselected header/body pane seams to one structural gridline', async ({ page }) => {
  await page.goto(`/?document=${encodeURIComponent(createTestDocumentId('playwright-unselected-pane-single-seam'))}&persist=0`)
  await waitForWorkbookReady(page)
  await waitForTypeGpuVisibleFrame(page)

  const gridLocator = page.getByTestId('sheet-grid')
  await expect(gridLocator).toBeVisible()
  const grid = await gridLocator.boundingBox()
  if (!grid) {
    throw new Error('sheet grid is not visible')
  }

  const rowTop = await getProductRowTop(page, 9)
  const rowMidY = grid.y + PRODUCT_HEADER_HEIGHT + rowTop + Math.floor(PRODUCT_ROW_HEIGHT / 2)
  const headerSeamPixel = await sampleCompositedViewportPixel(page, grid.x + PRODUCT_ROW_MARKER_WIDTH - 1, rowMidY)
  const bodyFirstPixel = await sampleCompositedViewportPixel(page, grid.x + PRODUCT_ROW_MARKER_WIDTH, rowMidY)

  const columnLeft = await getProductColumnLeft(page, 2)
  const columnWidth = await getProductColumnWidth(page, 2)
  const columnMidX = grid.x + columnLeft + Math.floor(columnWidth / 2)
  const columnHeaderSeamPixel = await sampleCompositedViewportPixel(page, columnMidX, grid.y + PRODUCT_HEADER_HEIGHT - 1)
  const bodyTopPixel = await sampleCompositedViewportPixel(page, columnMidX, grid.y + PRODUCT_HEADER_HEIGHT)

  expect(isGridBorderPixel(headerSeamPixel), 'row-header/body seam should have exactly one structural gridline').toBe(true)
  expect(isGridBorderPixel(bodyFirstPixel), 'first body pixel should not draw a second adjacent gridline').toBe(false)
  expect(isGridBorderPixel(columnHeaderSeamPixel), 'column-header/body seam should have exactly one structural gridline').toBe(true)
  expect(isGridBorderPixel(bodyTopPixel), 'first body row should not draw a second adjacent gridline').toBe(false)
})

test('@browser-ci web app collapses a locally dragged range when the active address is explicitly reselected', async ({ page }) => {
  await page.goto(`/?document=${encodeURIComponent(createTestDocumentId('playwright-range-same-address-collapse'))}&persist=0`)
  await waitForWorkbookReady(page)

  await dragProductBodySelection(page, 1, 1, 3, 3)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!B2:D4')
  await expect(page.getByTestId('name-box')).toHaveValue('B2:D4')
  await expect(page.locator('[data-grid-selection-visual-role="selection-fill"]')).toHaveCount(1)
  await expect(page.locator('[data-grid-selection-visual-role="selection-gridline"]')).toHaveCount(4)
  await expect(page.locator('[data-grid-selection-visual-role="selection-border"]')).toHaveCount(1)

  await selectAddress(page, 'B2')

  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!B2')
  await expect(page.getByTestId('name-box')).toHaveValue('B2')
  await expect(page.locator('[data-grid-selection-visual-role="selection-fill"]')).toHaveCount(0)
  await expect(page.locator('[data-grid-selection-visual-role="selection-gridline"]')).toHaveCount(0)
  await expect(page.locator('[data-grid-selection-visual-role="selection-border"]')).toHaveCount(0)
  await expect(page.locator('[data-grid-selection-visual-role="active-border"]')).toHaveCount(1)
  await expectVisualRectNear(
    page.locator('[data-grid-selection-visual-role="active-border"]'),
    await getProductCellRangeBox(page, 1, 1, 1, 1),
    'collapsed active cell border',
  )
  await expectBorderStyle(page.locator('[data-grid-selection-visual-role="active-border"]'), {
    boxShadow: 'none',
    color: 'rgb(33, 115, 70)',
    width: '1px',
  })
})

test('@browser-ci web app keeps fill-handle hit target aligned and pointer-only', async ({ page }) => {
  await page.goto(`/?document=${encodeURIComponent(createTestDocumentId('playwright-fill-handle-hit-target'))}&persist=0`)
  await waitForWorkbookReady(page)

  await dragProductBodySelection(page, 1, 1, 3, 3)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!B2:D4')

  const expectedRange = await getProductCellRangeBox(page, 1, 1, 3, 3)
  const expectedVisualHandle = {
    x: expectedRange.x + expectedRange.width - 4,
    y: expectedRange.y + expectedRange.height - 4,
    width: 8,
    height: 8,
  }
  const expectedHitTarget = {
    x: expectedRange.x + expectedRange.width - 5,
    y: expectedRange.y + expectedRange.height - 5,
    width: 10,
    height: 10,
  }

  await expectVisualRectNear(page.locator('[data-grid-selection-visual-role="fill-handle"]'), expectedVisualHandle, 'visible fill handle')
  await expectVisualRectNear(page.locator('[data-grid-fill-handle="true"]'), expectedHitTarget, 'fill handle hit target')

  await expect(page.locator('[data-grid-fill-handle="true"]')).toHaveJSProperty('tagName', 'DIV')
  await expect(page.locator('[data-grid-fill-handle="true"]')).toHaveAttribute('aria-hidden', 'true')
  await expect(page.locator('[data-grid-fill-handle="true"]')).toHaveJSProperty('tabIndex', -1)
})

test('@browser-ci web app keeps fill-handle drag source selection and target preview visible', async ({ page }) => {
  await page.goto(`/?document=${encodeURIComponent(createTestDocumentId('playwright-fill-handle-visible-preview'))}&persist=0`)
  await waitForWorkbookReady(page)

  await selectAddress(page, 'B2')
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!B2')

  const downDrag = await getProductFillHandleDragPoints(page, 1, 1, 1, 3)
  await page.mouse.move(downDrag.sourceX, downDrag.sourceY)
  await page.mouse.down()
  await page.mouse.move(downDrag.targetX, downDrag.targetY, { steps: 8 })

  await expectVisualRectNear(
    page.locator('[data-grid-selection-visual-role="active-border"]'),
    await getProductCellRangeBox(page, 1, 1, 1, 1),
    'source active border during downward fill',
  )
  await expectVisualRectNear(
    page.locator('[data-grid-fill-preview="true"]'),
    await getProductCellRangeBox(page, 1, 2, 1, 3),
    'downward fill preview',
  )

  await page.mouse.up()
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!B2:B4')
  await expectVisualRectNear(
    page.locator('[data-grid-selection-visual-role="selection-border"]'),
    await getProductCellRangeBox(page, 1, 1, 1, 3),
    'committed downward fill selection',
  )

  await selectAddress(page, 'D4')
  const rightDrag = await getProductFillHandleDragPoints(page, 3, 3, 5, 3)
  await page.mouse.move(rightDrag.sourceX, rightDrag.sourceY)
  await page.mouse.down()
  await page.mouse.move(rightDrag.targetX, rightDrag.targetY, { steps: 8 })

  await expectVisualRectNear(
    page.locator('[data-grid-selection-visual-role="active-border"]'),
    await getProductCellRangeBox(page, 3, 3, 3, 3),
    'source active border during rightward fill',
  )
  await expectVisualRectNear(
    page.locator('[data-grid-fill-preview="true"]'),
    await getProductCellRangeBox(page, 4, 3, 5, 3),
    'rightward fill preview',
  )

  await page.mouse.up()
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!D4:F4')
  await expectVisualRectNear(
    page.locator('[data-grid-selection-visual-role="selection-border"]'),
    await getProductCellRangeBox(page, 3, 3, 5, 3),
    'committed rightward fill selection',
  )
})

test('@browser-ci web app keeps old fill-handle hit targets from intercepting cell clicks', async ({ page }) => {
  await page.goto(`/?document=${encodeURIComponent(createTestDocumentId('playwright-fill-handle-stale-click-through'))}&persist=0`)
  await waitForWorkbookReady(page)

  await dragProductBodySelection(page, 1, 1, 3, 3)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!B2:D4')

  const oldHitTarget = await page.locator('[data-grid-fill-handle="true"]').boundingBox()
  if (!oldHitTarget) {
    throw new Error('fill handle hit target is not visible')
  }
  const oldCellInteriorPoint = {
    x: oldHitTarget.x + 1,
    y: oldHitTarget.y + 1,
  }

  await clickProductCell(page, 0, 0)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!A1')
  await page.mouse.click(oldCellInteriorPoint.x, oldCellInteriorPoint.y)

  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!D4')
})

test('@browser-ci web app hides fill handles when the selected cell is clipped behind headers', async ({ page }) => {
  await page.goto(`/?document=${encodeURIComponent(createTestDocumentId('playwright-fill-handle-clipped-headers'))}&persist=0`)
  await waitForWorkbookReady(page)

  await clickProductCell(page, 1, 1)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!B2')
  await expect(page.locator('[data-grid-selection-visual-role="fill-handle"]')).toHaveCount(1)
  await expect(page.locator('[data-grid-fill-handle="true"]')).toBeVisible()

  await page.getByTestId('grid-scroll-viewport').evaluate((viewport) => {
    viewport.scrollLeft = 150
    viewport.scrollTop = 60
    viewport.dispatchEvent(new Event('scroll', { bubbles: true }))
  })

  await expect(page.locator('[data-grid-selection-visual-role="fill-handle"]')).toHaveCount(0)
  await expect(page.locator('[data-grid-fill-handle="true"]')).toBeHidden()
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!B2')
})

test('@browser-ci web app keeps scrolled selection chrome and hit targets clipped to the visible pane', async ({ page }) => {
  await page.goto(`/?document=${encodeURIComponent(createTestDocumentId('playwright-scrolled-selection-chrome-clip'))}&persist=0`)
  await waitForWorkbookReady(page)

  await dragProductBodySelection(page, 1, 1, 3, 3)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!B2:D4')

  const scrollTop = 30
  await page.getByTestId('grid-scroll-viewport').evaluate((viewport, nextScrollTop) => {
    viewport.scrollTop = nextScrollTop
    viewport.dispatchEvent(new Event('scroll', { bubbles: true }))
  }, scrollTop)

  const grid = await page.getByTestId('sheet-grid').boundingBox()
  if (!grid) {
    throw new Error('sheet grid is not visible')
  }
  const rangeLeft = await getProductColumnLeft(page, 1)
  const rangeRight = await getProductColumnLeft(page, 3)
  const rangeRightWidth = await getProductColumnWidth(page, 3)
  const rangeTop = await getProductRowTop(page, 1)
  const rangeBottom = await getProductRowTop(page, 3)
  const rangeBottomHeight = await getProductRowHeight(page, 3)
  const unclippedTop = grid.y + PRODUCT_HEADER_HEIGHT + rangeTop - scrollTop
  const unclippedBottom = grid.y + PRODUCT_HEADER_HEIGHT + rangeBottom + rangeBottomHeight - scrollTop
  const clippedTop = Math.max(grid.y + PRODUCT_HEADER_HEIGHT, unclippedTop)
  const clippedBottom = Math.min(grid.y + grid.height, unclippedBottom)
  const expectedRange = {
    x: grid.x + rangeLeft,
    y: clippedTop,
    width: rangeRight + rangeRightWidth - rangeLeft,
    height: clippedBottom - clippedTop,
  }
  const expectedVisualHandle = {
    x: expectedRange.x + expectedRange.width - 4,
    y: expectedRange.y + expectedRange.height - 4,
    width: 8,
    height: 8,
  }
  const expectedHitTarget = {
    x: expectedRange.x + expectedRange.width - 5,
    y: expectedRange.y + expectedRange.height - 5,
    width: 10,
    height: 10,
  }

  await expectVisualRectNear(
    page.locator('[data-grid-selection-visual-role="selection-border"]'),
    expectedRange,
    'clipped selection border',
  )
  await expect(page.locator('[data-grid-selection-visual-role="active-border"]')).toHaveCount(0)
  await expectVisualRectNear(page.locator('[data-grid-selection-visual-role="fill-handle"]'), expectedVisualHandle, 'clipped fill handle')
  await expectVisualRectNear(page.locator('[data-grid-fill-handle="true"]'), expectedHitTarget, 'clipped fill handle hit target')

  await page.getByTestId('sheet-grid').click({
    position: {
      x: rangeLeft + Math.floor((await getProductColumnWidth(page, 1)) / 2),
      y: Math.floor(PRODUCT_HEADER_HEIGHT / 2),
    },
  })
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!B:B')
})

test('@browser-ci web app keeps clean range chrome synchronized while keyboard-cycling within a range', async ({ page }) => {
  await page.goto(`/?document=${encodeURIComponent(createTestDocumentId('playwright-range-keyboard-visual-geometry'))}&persist=0`)
  await waitForWorkbookReady(page)

  await dragProductBodySelection(page, 1, 1, 3, 4)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!B2:D5')
  await expect(page.getByTestId('name-box')).toHaveValue('B2:D5')
  await expect(page.getByTestId('sheet-grid-focus-target')).toHaveAttribute('aria-label', 'Sheet1 B2')

  const expectedRange = await getProductCellRangeBox(page, 1, 1, 3, 4)
  const expectedFillHandle = {
    x: expectedRange.x + expectedRange.width - 4,
    y: expectedRange.y + expectedRange.height - 4,
    width: 8,
    height: 8,
  }

  await page.getByTestId('sheet-grid-focus-target').press('Tab')
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!B2:D5')
  await expect(page.getByTestId('name-box')).toHaveValue('B2:D5')
  await expect(page.getByTestId('sheet-grid-focus-target')).toHaveAttribute('aria-label', 'Sheet1 C2')

  await expectVisualRectNear(page.locator('[data-grid-selection-visual-role="selection-border"]'), expectedRange, 'selection border')
  await expect(page.locator('[data-grid-selection-visual-role="active-border"]')).toHaveCount(0)
  await expectVisualRectNear(page.locator('[data-grid-selection-visual-role="fill-handle"]'), expectedFillHandle, 'fill handle')

  await page.getByTestId('sheet-grid-focus-target').press('Enter')
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!B2:D5')
  await expect(page.getByTestId('name-box')).toHaveValue('B2:D5')
  await expect(page.getByTestId('sheet-grid-focus-target')).toHaveAttribute('aria-label', 'Sheet1 C3')
  await expectVisualRectNear(
    page.locator('[data-grid-selection-visual-role="selection-border"]'),
    expectedRange,
    'selection border after enter',
  )
  await expect(page.locator('[data-grid-selection-visual-role="active-border"]')).toHaveCount(0)
  await expectVisualRectNear(page.locator('[data-grid-selection-visual-role="fill-handle"]'), expectedFillHandle, 'fill handle after enter')

  await page.getByTestId('sheet-grid-focus-target').press('ArrowRight')
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!D3')
  await expect(page.getByTestId('name-box')).toHaveValue('D3')
  await expect(page.getByTestId('sheet-grid-focus-target')).toHaveAttribute('aria-label', 'Sheet1 D3')
  await expect(page.locator('[data-grid-selection-visual-role="selection-border"]')).toHaveCount(0)
  await expectVisualRectNear(
    page.locator('[data-grid-selection-visual-role="active-border"]'),
    await getProductCellRangeBox(page, 3, 2, 3, 2),
    'collapsed active cell border after arrow',
  )
})

test('@browser-ci web app starts a fresh area selection from the selected range interior', async ({ page }) => {
  await page.goto(`/?document=${encodeURIComponent(createTestDocumentId('playwright-interior-drag-selects-area'))}&persist=0`)
  await waitForWorkbookReady(page)

  await writeCellValue(page, 'B2', 'keep-b2')
  await writeCellValue(page, 'D4', 'keep-d4')

  await dragProductBodySelection(page, 1, 1, 3, 3)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!B2:D4')

  await dragProductSelectedInterior(page, 1, 1, 5, 5)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!B2:F6')
  await expect(page.getByTestId('name-box')).toHaveValue('B2:F6')

  await expectVisualRectNear(
    page.locator('[data-grid-selection-visual-role="selection-border"]'),
    await getProductCellRangeBox(page, 1, 1, 5, 5),
    'interior-drag selection border',
  )

  await selectAddress(page, 'B2')
  await expect(page.getByTestId('formula-input')).toHaveValue('keep-b2')
  await selectAddress(page, 'D4')
  await expect(page.getByTestId('formula-input')).toHaveValue('keep-d4')
  await selectAddress(page, 'F6')
  await expect(page.getByTestId('formula-input')).toHaveValue('')
})

test('web app clips spilled text before the active selected cell', async ({ page }) => {
  await installTypeGpuCellReadbackHarness(page)
  await page.goto(`/?document=${encodeURIComponent(createTestDocumentId('playwright-selection-spill-clip'))}&persist=0`)
  await waitForWorkbookReady(page)

  await writeCellValue(page, 'A5', 'sfasf sfasf sfasf sfasf sfasf sfasf')
  await selectAddress(page, 'C5')
  await expect(page.getByTestId('name-box')).toHaveValue('C5')

  if (await isTypeGpuTextReadbackActive(page)) {
    await expectCellTextPixels(page, 0, 4, 'visible')
    await expectCellTextPixels(page, 2, 4, 'hidden')
    return
  }

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
  await installTypeGpuCellReadbackHarness(page)
  await page.goto(`/?document=${encodeURIComponent(createTestDocumentId('playwright-far-selection-spill-clip'))}&persist=0`)
  await waitForWorkbookReady(page)

  const rowIndex = 4
  const sourceColumnIndex = 128
  const selectedColumnIndex = 130
  await writeCellValue(page, formatGridAddress(rowIndex, sourceColumnIndex), 'far spill text far spill text far spill text')
  await selectAddress(page, formatGridAddress(rowIndex, selectedColumnIndex))

  if (await isTypeGpuTextReadbackActive(page)) {
    await expectCellTextPixels(page, sourceColumnIndex, rowIndex, 'visible')
    await expectCellTextPixels(page, selectedColumnIndex, rowIndex, 'hidden')
    return
  }

  const runBox = await page
    .locator(`[data-native-text-run-row="${String(rowIndex)}"][data-native-text-run-col="${String(sourceColumnIndex)}"]`)
    .boundingBox()
  const activeBorderBox = await page.locator('[data-grid-selection-visual-role="active-border"]').boundingBox()
  if (!runBox || !activeBorderBox) {
    throw new Error('far grid text run or selection border is not visible')
  }

  expect(runBox.x + runBox.width).toBeLessThanOrEqual(activeBorderBox.x + 0.5)
})

test('web app clips spilled text before selected whole columns on non-active rows', async ({ page }) => {
  await installTypeGpuCellReadbackHarness(page)
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

  if (await isTypeGpuTextReadbackActive(page)) {
    await expectCellTextPixels(page, 0, 4, 'visible')
    await expectCellTextPixels(page, 2, 4, 'hidden')
    return
  }

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
  await page.goto('/workbook')
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
  await page.goto('/workbook')
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
