import { expect, test, type Page } from '@playwright/test'
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

test('web app keeps moved range data visible when border drag reaches the grid edge', async ({ page }) => {
  await page.setViewportSize({ width: 960, height: 420 })
  await page.goto(`/?document=range-border-edge-drag-${Date.now()}`)
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

  await nameBox.fill('B2')
  await nameBox.press('Enter')
  await expect(formulaInput).toHaveValue('left')
  await nameBox.fill('C2')
  await nameBox.press('Enter')
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

  await nameBox.fill(`B${targetRow}`)
  await nameBox.press('Enter')
  await expect(formulaInput).toHaveValue('left')

  await nameBox.fill(`C${targetRow}`)
  await nameBox.press('Enter')
  await expect(formulaInput).toHaveValue('right')

  await nameBox.fill('B2')
  await nameBox.press('Enter')
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
  await page.goto(`/?document=normal-selection-no-resident-refresh-${Date.now()}`)
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
  expect(report?.counters.scenePacketRefreshes).toBe(0)
  expect(report?.counters.typeGpuScenePacketsApplied).toBe(0)
  expect(report?.counters.typeGpuBufferAllocations).toBe(0)
  expect(report?.counters.typeGpuTileMisses).toBe(0)
})

test('@browser-perf web app keeps range-move preview out of resident scene invalidation', async ({ page }) => {
  await page.goto(`/?document=range-move-no-resident-refresh-${Date.now()}`)
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
  expect(report?.counters.scenePacketRefreshes).toBe(0)
  expect(report?.counters.typeGpuScenePacketsApplied).toBe(0)
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
