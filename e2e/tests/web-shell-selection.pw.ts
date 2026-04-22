import { expect, test } from '@playwright/test'
import {
  PRODUCT_COLUMN_WIDTH,
  PRODUCT_HEADER_HEIGHT,
  PRODUCT_ROW_HEIGHT,
  PRODUCT_ROW_MARKER_WIDTH,
  clickProductBodyOffset,
  clickProductCell,
  clickProductCellUpperHalf,
  dragProductBodySelection,
  dragProductColumnResize,
  dragProductHeaderSelection,
  doubleClickProductColumnResizeHandle,
  getProductColumnLeft,
  getProductColumnWidth,
  waitForWorkbookReady,
} from './web-shell-helpers.js'

test('web app keeps sheet tabs and status bar visible in a short viewport', async ({ page }) => {
  await page.setViewportSize({ width: 2048, height: 220 })
  await page.goto('/')
  await waitForWorkbookReady(page)

  const sheetTab = page.getByRole('tab', { name: 'Sheet1' })
  const statusSync = page.getByTestId('status-sync')

  await expect(sheetTab).toBeVisible()
  await expect(statusSync).toBeVisible()

  const tabBox = await sheetTab.boundingBox()
  const statusBox = await statusSync.boundingBox()
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

test('web app supports rectangular drag selection', async ({ page }) => {
  await page.goto('/')
  await waitForWorkbookReady(page)

  await dragProductBodySelection(page, 1, 1, 3, 3)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!B2:D4')
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

test('web app supports column resize without breaking hit testing', async ({ page }) => {
  const documentId = `playwright-column-resize-hit-test-${Date.now()}`
  await page.goto(`/?document=${encodeURIComponent(documentId)}`)
  await waitForWorkbookReady(page)

  await clickProductBodyOffset(page, 82, 0)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!A1')

  await dragProductColumnResize(page, 0, -36)

  await clickProductBodyOffset(page, 82, 0)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!B1')
})

test('web app supports column edge double-click autofit', async ({ page }) => {
  const documentId = `playwright-column-autofit-${Date.now()}`
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
  await expect.poll(async () => await getProductColumnWidth(page, 0)).toBeGreaterThan(126)
  const autofitWidth = await getProductColumnWidth(page, 0)

  await clickProductBodyOffset(page, autofitWidth + 8, 0)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!B1')
  await clickProductBodyOffset(page, 126, 0)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!A1')
})

test('web app hit-tests typegpu geometry after hiding rows and columns', async ({ page }) => {
  const documentId = `playwright-hidden-axis-hit-test-${Date.now()}`
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
  const grid = await gridLocator.boundingBox()
  if (!grid) {
    throw new Error('sheet grid is not visible')
  }
  const columnLeft = await getProductColumnLeft(page, 2)
  const columnWidth = await getProductColumnWidth(page, 2)
  await page.mouse.click(
    grid.x + columnLeft + Math.floor(columnWidth / 2),
    grid.y + PRODUCT_HEADER_HEIGHT + PRODUCT_ROW_HEIGHT + Math.floor(PRODUCT_ROW_HEIGHT / 2),
  )
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!C3')
})
