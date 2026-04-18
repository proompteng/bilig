import { expect, test, type Locator } from '@playwright/test'
import * as fc from 'fast-check'
import { runProperty, shouldRunFuzzSuite } from '../../packages/test-fuzz/src/index.ts'
import {
  PRIMARY_MODIFIER,
  PRODUCT_COLUMN_WIDTH,
  PRODUCT_HEADER_HEIGHT,
  PRODUCT_ROW_HEIGHT,
  clickProductCell,
  clickGridRightEdge,
  dragProductBodySelection,
  dragProductColumnResize,
  getProductColumnLeft,
  getProductColumnWidth,
  gotoWorkbookShell,
  remoteSyncEnabled,
  waitForWorkbookReady,
} from './web-shell-helpers.js'
const fuzzBrowserEnabled = process.env['BILIG_FUZZ_BROWSER'] === '1'

type BrowserSelectionAction =
  | { kind: 'click'; row: number; col: number }
  | { kind: 'shiftClick'; row: number; col: number }
  | { kind: 'key'; key: 'ArrowLeft' | 'ArrowRight' | 'ArrowUp' | 'ArrowDown'; shift: boolean }

async function dragProductFillHandle(
  page: Parameters<typeof test>[0]['page'],
  sourceCol: number,
  sourceRow: number,
  targetCol: number,
  targetRow: number,
) {
  const { sourceX, sourceY, targetX, targetY } = await getProductFillHandleDragPoints(page, sourceCol, sourceRow, targetCol, targetRow)

  await page.mouse.move(sourceX, sourceY)
  await page.mouse.down()
  await page.mouse.move(targetX, targetY, {
    steps: 10,
  })
  await page.mouse.up()
}

async function getProductFillHandleDragPoints(
  page: Parameters<typeof test>[0]['page'],
  sourceCol: number,
  sourceRow: number,
  targetCol: number,
  targetRow: number,
) {
  const gridLocator = page.getByTestId('sheet-grid')
  await expect(gridLocator).toBeVisible()
  const grid = await gridLocator.boundingBox()
  if (!grid) {
    throw new Error('sheet grid is not visible')
  }

  const sourceLeft = grid.x + (await getProductColumnLeft(page, sourceCol))
  const sourceTop = grid.y + PRODUCT_HEADER_HEIGHT + sourceRow * PRODUCT_ROW_HEIGHT
  const targetLeft = grid.x + (await getProductColumnLeft(page, targetCol))
  const targetTop = grid.y + PRODUCT_HEADER_HEIGHT + targetRow * PRODUCT_ROW_HEIGHT
  const sourceWidth = await getProductColumnWidth(page, sourceCol)
  const targetWidth = await getProductColumnWidth(page, targetCol)

  return {
    sourceX: sourceLeft + sourceWidth - 3,
    sourceY: sourceTop + PRODUCT_ROW_HEIGHT - 3,
    targetX: targetLeft + targetWidth - 3,
    targetY: targetTop + PRODUCT_ROW_HEIGHT - 3,
  }
}

async function clickSelectionFuzzCell(page: Parameters<typeof test>[0]['page'], columnIndex: number, rowIndex: number, shift = false) {
  const gridLocator = page.getByTestId('sheet-grid')
  await expect(gridLocator).toBeVisible()
  const grid = await gridLocator.boundingBox()
  if (!grid) {
    throw new Error('sheet grid is not visible')
  }
  const columnLeft = await getProductColumnLeft(page, columnIndex)
  const columnWidth = await getProductColumnWidth(page, columnIndex)
  const scrollLeft = await gridLocator.evaluate(
    (node, target) => {
      const scrollViewport = node.querySelector('[aria-hidden="true"]')
      if (!(scrollViewport instanceof HTMLElement)) {
        return 0
      }
      const targetCenter = target.columnLeft + target.columnWidth / 2
      const visibleStart = scrollViewport.scrollLeft
      const visibleEnd = visibleStart + scrollViewport.clientWidth
      if (targetCenter < visibleStart || targetCenter > visibleEnd) {
        scrollViewport.scrollLeft = Math.max(0, targetCenter - scrollViewport.clientWidth / 2)
      }
      return scrollViewport.scrollLeft
    },
    {
      columnLeft,
      columnWidth,
    },
  )
  const point = {
    x: grid.x + columnLeft - scrollLeft + columnWidth / 2,
    y: grid.y + PRODUCT_HEADER_HEIGHT + rowIndex * PRODUCT_ROW_HEIGHT + PRODUCT_ROW_HEIGHT / 2,
  }
  if (shift) {
    await page.keyboard.down('Shift')
  }
  try {
    await page.mouse.click(point.x, point.y)
  } finally {
    if (shift) {
      await page.keyboard.up('Shift')
    }
  }
}

async function runSelectionFuzzActions(
  page: Parameters<typeof test>[0]['page'],
  grid: Locator,
  actions: readonly BrowserSelectionAction[],
  index = 0,
): Promise<void> {
  const action = actions[index]
  if (!action) {
    return
  }

  if (action.kind === 'click') {
    await clickSelectionFuzzCell(page, action.col, action.row)
  } else if (action.kind === 'shiftClick') {
    await clickSelectionFuzzCell(page, action.col, action.row, true)
  } else {
    await grid.press(action.shift ? `Shift+${action.key}` : action.key)
  }

  const selection = await page.getByTestId('status-selection').textContent()
  expect(selection).toMatch(/^Sheet1!(?:[A-Z]+[0-9]+(?::[A-Z]+[0-9]+)?|[A-Z]+:[A-Z]+|[0-9]+:[0-9]+|All)$/)

  const focusInsideShell = await page.evaluate(() => {
    const active = document.activeElement
    return Boolean(
      active?.closest('[data-testid="sheet-grid"]') ||
      active?.closest('[data-testid="formula-bar"]') ||
      active?.closest('[role="toolbar"]'),
    )
  })
  expect(focusInsideShell).toBe(true)

  await runSelectionFuzzActions(page, grid, actions, index + 1)
}

async function clickProductSelectedCellTopBorder(page: Parameters<typeof test>[0]['page'], columnIndex: number, rowIndex: number) {
  const gridLocator = page.getByTestId('sheet-grid')
  await expect(gridLocator).toBeVisible()
  const grid = await gridLocator.boundingBox()
  if (!grid) {
    throw new Error('sheet grid is not visible')
  }

  const columnLeft = await getProductColumnLeft(page, columnIndex)
  const columnWidth = await getProductColumnWidth(page, columnIndex)
  await page.mouse.click(
    grid.x + columnLeft + Math.floor(columnWidth / 2),
    grid.y + PRODUCT_HEADER_HEIGHT + rowIndex * PRODUCT_ROW_HEIGHT - 1,
  )
}

async function dragProductSelectionBorder(
  page: Parameters<typeof test>[0]['page'],
  startColumn: number,
  startRow: number,
  endColumn: number,
  endRow: number,
  targetColumn: number,
  targetRow: number,
) {
  const gridLocator = page.getByTestId('sheet-grid')
  await expect(gridLocator).toBeVisible()
  const grid = await gridLocator.boundingBox()
  if (!grid) {
    throw new Error('sheet grid is not visible')
  }

  const startLeft = await getProductColumnLeft(page, startColumn)
  const rangeTop = grid.y + PRODUCT_HEADER_HEIGHT + startRow * PRODUCT_ROW_HEIGHT
  const sourceX = grid.x + startLeft + 3
  const sourceY = rangeTop + 2
  const targetLeft = await getProductColumnLeft(page, targetColumn)
  const targetWidth = await getProductColumnWidth(page, targetColumn)
  const targetX = grid.x + targetLeft + Math.floor(targetWidth / 2)
  const targetY = grid.y + PRODUCT_HEADER_HEIGHT + targetRow * PRODUCT_ROW_HEIGHT + Math.floor(PRODUCT_ROW_HEIGHT / 2)

  await page.mouse.move(sourceX, sourceY)
  await page.mouse.down()
  await page.mouse.move(targetX, targetY, { steps: 12 })
  await page.mouse.up()
}

test('web app accepts string values and string comparison formulas', async ({ page }) => {
  await page.goto('/')
  await waitForWorkbookReady(page)

  const nameBox = page.getByTestId('name-box')
  const formulaInput = page.getByTestId('formula-input')
  const resolvedValue = page.getByTestId('formula-resolved-value')

  await clickProductCell(page, 0, 0)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!A1')
  await formulaInput.fill('hello')
  await formulaInput.press('Enter')
  await expect(nameBox).toHaveValue('A1')
  await expect(formulaInput).toHaveValue('hello')
  await clickProductCell(page, 0, 0)
  await expect(resolvedValue).toHaveText('hello')

  await nameBox.fill('A2')
  await nameBox.press('Enter')
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!A2')
  await formulaInput.fill('=A1="HELLO"')
  await formulaInput.press('Enter')
  await clickProductCell(page, 0, 1)
  await expect(resolvedValue).toHaveText('TRUE')
})

test('web app supports type-to-replace and Enter or Tab commit movement', async ({ page }) => {
  await page.goto('/')
  await waitForWorkbookReady(page)

  const grid = page.getByTestId('sheet-grid')
  const nameBox = page.getByTestId('name-box')
  const formulaInput = page.getByTestId('formula-input')
  const cellEditor = page.getByTestId('cell-editor-input')

  await clickProductCell(page, 0, 0)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!A1')
  await grid.press('h')
  await expect(cellEditor).toBeVisible()
  await expect(cellEditor).toHaveValue('h')
  await page.keyboard.press('Enter')
  await expect(cellEditor).toBeHidden()

  await expect(nameBox).toHaveValue('A2')
  await nameBox.fill('A1')
  await nameBox.press('Enter')
  await expect(formulaInput).toHaveValue('h')

  await clickProductCell(page, 0, 1)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!A2')
  await grid.press('w')
  await expect(cellEditor).toBeVisible()
  await expect(cellEditor).toHaveValue('w')
  await page.keyboard.press('Tab')
  await expect(cellEditor).toBeHidden()

  await expect(nameBox).toHaveValue('B2')
  await nameBox.fill('A2')
  await nameBox.press('Enter')
  await expect(formulaInput).toHaveValue('w')

  await grid.press('Enter')
  await expect(nameBox).toHaveValue('A3')
  await grid.press('Shift+Enter')
  await expect(nameBox).toHaveValue('A2')
})

test('web app preserves multi-digit numeric type-to-replace input', async ({ page }) => {
  await page.goto('/')
  await waitForWorkbookReady(page)

  const grid = page.getByTestId('sheet-grid')
  const nameBox = page.getByTestId('name-box')
  const formulaInput = page.getByTestId('formula-input')
  const cellEditor = page.getByTestId('cell-editor-input')

  await clickProductCell(page, 0, 0)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!A1')

  await page.keyboard.type('123')
  await expect(cellEditor).toBeVisible()
  await expect(cellEditor).toHaveValue('123')
  await page.keyboard.press('Enter')

  await expect(nameBox).toHaveValue('A2')
  await clickProductCell(page, 0, 0)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!A1')
  await expect(formulaInput).toHaveValue('123')

  await clickProductCell(page, 1, 0)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!B1')
  await grid.press('4')
  await expect(cellEditor).toBeVisible()
  await expect(cellEditor).toHaveValue('4')
})

test('web app right-aligns numeric in-cell editing like numeric view state', async ({ page }) => {
  await page.goto('/')
  await waitForWorkbookReady(page)

  const grid = page.getByTestId('sheet-grid')
  const cellEditor = page.getByTestId('cell-editor-input')

  await clickProductCell(page, 0, 0)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!A1')
  await page.keyboard.type('123')
  await expect(cellEditor).toBeVisible()
  await expect(cellEditor).toHaveValue('123')
  await expect(cellEditor).toHaveCSS('text-align', 'right')

  await page.keyboard.press('Escape')
  await clickProductCell(page, 1, 0)
  await grid.press('h')
  await expect(cellEditor).toBeVisible()
  await expect(cellEditor).toHaveValue('h')
  await expect(cellEditor).toHaveCSS('text-align', 'left')
})

test('web app accepts numpad digits for in-cell numeric entry', async ({ page }) => {
  await page.goto('/')
  await waitForWorkbookReady(page)

  const formulaInput = page.getByTestId('formula-input')
  const cellEditor = page.getByTestId('cell-editor-input')

  await clickProductCell(page, 0, 0)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!A1')

  await page.keyboard.press('Numpad1')
  await page.keyboard.press('Numpad2')
  await page.keyboard.press('Numpad3')
  await expect(cellEditor).toBeVisible()
  await expect(cellEditor).toHaveValue('123')
  await page.keyboard.press('Enter')

  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!A2')
  await clickProductCell(page, 0, 0)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!A1')
  await expect(formulaInput).toHaveValue('123')
})

test('web app supports F2 edit in the product shell', async ({ page }) => {
  await page.goto('/')
  await waitForWorkbookReady(page)
  await waitForWorkbookReady(page)

  const grid = page.getByTestId('sheet-grid')
  const nameBox = page.getByTestId('name-box')
  const formulaInput = page.getByTestId('formula-input')
  const cellEditor = page.getByTestId('cell-editor-input')

  await nameBox.fill('C3')
  await nameBox.press('Enter')
  await formulaInput.fill('seed')
  await formulaInput.press('Enter')

  await clickProductCell(page, 2, 2)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!C3')
  await grid.press('F2')
  await expect(cellEditor).toBeVisible()
  await expect(cellEditor).toHaveValue('seed')
  await cellEditor.press('!')
  await expect(cellEditor).toHaveValue('seed!')
  await clickProductCell(page, 3, 2)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!D3')

  await clickProductCell(page, 2, 2)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!C3')
  await expect(formulaInput).toHaveValue('seed!')
})

test('web app offers formula autocomplete and inserts a function with Tab', async ({ page }) => {
  await page.goto('/')
  await waitForWorkbookReady(page)

  const nameBox = page.getByTestId('name-box')
  const formulaInput = page.getByTestId('formula-input')
  const autocomplete = page.getByTestId('formula-autocomplete')
  const argHint = page.getByTestId('formula-arg-hint')
  const resolvedValue = page.getByTestId('formula-resolved-value')

  await clickProductCell(page, 0, 0)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!A1')

  await formulaInput.focus()
  await page.keyboard.type('=su')
  await expect(autocomplete).toBeVisible()
  await expect(autocomplete).toContainText('SUM')

  await page.keyboard.press('Tab')
  await expect(formulaInput).toHaveValue('=SUM()')
  await expect(argHint).toContainText('number1')

  await page.keyboard.type('7')
  await expect(formulaInput).toHaveValue('=SUM(7)')
  await page.keyboard.press('Enter')

  await nameBox.fill('A1')
  await nameBox.press('Enter')
  await expect(formulaInput).toHaveValue('=SUM(7)')
  await expect(resolvedValue).toHaveText('7')
})

test('web app shows formula argument hints while typing', async ({ page }) => {
  await page.goto('/')
  await waitForWorkbookReady(page)

  const formulaInput = page.getByTestId('formula-input')
  const argHint = page.getByTestId('formula-arg-hint')

  await clickProductCell(page, 0, 0)
  await formulaInput.focus()
  await page.keyboard.type('=IF(A1,')

  await expect(argHint).toBeVisible()
  await expect(argHint).toContainText('value_if_true')
})

test('web app double-click edits the exact clicked cell', async ({ page }) => {
  await page.goto('/')
  await waitForWorkbookReady(page)

  const nameBox = page.getByTestId('name-box')
  const formulaInput = page.getByTestId('formula-input')
  const cellEditor = page.getByTestId('cell-editor-input')
  const gridLocator = page.getByTestId('sheet-grid')

  await nameBox.fill('C4')
  await nameBox.press('Enter')
  await formulaInput.fill('above')
  await formulaInput.press('Enter')

  await nameBox.fill('C5')
  await nameBox.press('Enter')
  await formulaInput.fill('target')
  await formulaInput.press('Enter')

  await expect(gridLocator).toBeVisible()
  const grid = await gridLocator.boundingBox()
  if (!grid) {
    throw new Error('sheet grid is not visible')
  }

  const columnLeft = await getProductColumnLeft(page, 2)
  const columnWidth = await getProductColumnWidth(page, 2)
  const targetX = grid.x + columnLeft + Math.floor(columnWidth / 2)
  const targetY = grid.y + PRODUCT_HEADER_HEIGHT + 4 * PRODUCT_ROW_HEIGHT + Math.floor(PRODUCT_ROW_HEIGHT / 2)
  await page.mouse.dblclick(targetX, targetY)

  await expect(nameBox).toHaveValue('C5')
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!C5')
  await expect(cellEditor).toBeVisible()
  await expect(cellEditor).toHaveValue('target')
  await expect(cellEditor).toHaveAttribute('aria-label', 'Sheet1!C5 editor')
})

test('web app keeps the selected cell when clicking its top border', async ({ page }) => {
  await page.goto('/')
  await waitForWorkbookReady(page)

  const nameBox = page.getByTestId('name-box')

  await nameBox.fill('C5')
  await nameBox.press('Enter')
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!C5')

  await clickProductSelectedCellTopBorder(page, 2, 4)
  await expect(nameBox).toHaveValue('C5')
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!C5')
})

test('web app keeps selected text cells visible when clicked', async ({ page }) => {
  await page.goto('/')
  await waitForWorkbookReady(page)

  const nameBox = page.getByTestId('name-box')
  const formulaInput = page.getByTestId('formula-input')
  const textOverlay = page.getByTestId('grid-text-overlay')
  const sampleText = 'visible text sample'

  await nameBox.fill('C5')
  await nameBox.press('Enter')
  await formulaInput.fill(sampleText)
  await formulaInput.press('Enter')

  await clickProductCell(page, 0, 0)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!A1')

  await clickProductCell(page, 2, 4)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!C5')
  const spilledText = textOverlay.getByText(sampleText, { exact: true })
  await expect(spilledText).toBeVisible()
  await expect(formulaInput).toHaveValue(sampleText)
  await expect
    .poll(() =>
      spilledText.evaluate((element) => {
        if (!(element instanceof HTMLElement)) {
          return 0
        }
        return Math.round(element.getBoundingClientRect().width)
      }),
    )
    .toBeGreaterThan(PRODUCT_COLUMN_WIDTH)
})

test('web app supports fill-handle propagation', async ({ page }) => {
  await page.goto('/')
  await waitForWorkbookReady(page)

  const nameBox = page.getByTestId('name-box')
  const formulaInput = page.getByTestId('formula-input')
  const resolvedValue = page.getByTestId('formula-resolved-value')

  await nameBox.fill('F6')
  await nameBox.press('Enter')
  await formulaInput.fill('7')
  await formulaInput.press('Enter')

  await dragProductFillHandle(page, 5, 5, 5, 7)

  await nameBox.fill('F8')
  await nameBox.press('Enter')
  await expect(formulaInput).toHaveValue('7')
  await expect(resolvedValue).toHaveText('7')
})

test('web app enables undo and redo for a normal edit', async ({ page }) => {
  test.skip(!remoteSyncEnabled, 'requires authoritative remote sync history')
  const documentId = `playwright-undo-redo-basic-${Date.now()}`
  await page.goto(`/?document=${encodeURIComponent(documentId)}`)
  await waitForWorkbookReady(page)
  await expect(page.getByTestId('status-sync')).toHaveText('Saved', { timeout: 30_000 })

  const undoButton = page.getByRole('button', { name: 'Undo', exact: true })
  const redoButton = page.getByRole('button', { name: 'Redo', exact: true })
  const nameBox = page.getByTestId('name-box')
  const formulaInput = page.getByTestId('formula-input')
  const resolvedValue = page.getByTestId('formula-resolved-value')

  await expect(undoButton).toBeDisabled()
  await expect(redoButton).toBeDisabled()

  await nameBox.fill('A1')
  await nameBox.press('Enter')
  await formulaInput.fill('undo-check')
  await formulaInput.press('Enter')

  await expect(undoButton).toBeEnabled()
  await expect(redoButton).toBeDisabled()
  await expect(formulaInput).toHaveValue('undo-check')
  await expect(resolvedValue).toHaveText('undo-check')

  await undoButton.click()
  await expect(redoButton).toBeEnabled()
  await expect(formulaInput).toHaveValue('')
  await expect(resolvedValue).toHaveText('∅')

  await redoButton.click()
  await expect(undoButton).toBeEnabled()
  await expect(formulaInput).toHaveValue('undo-check')
  await expect(resolvedValue).toHaveText('undo-check')
})

test('web app preserves redo across a longer undo history', async ({ page }) => {
  test.skip(!remoteSyncEnabled, 'requires authoritative remote sync history')
  const documentId = `playwright-undo-redo-long-${Date.now()}`
  await page.goto(`/?document=${encodeURIComponent(documentId)}`)
  await waitForWorkbookReady(page)
  await expect(page.getByTestId('status-sync')).toHaveText('Saved', { timeout: 30_000 })

  const undoButton = page.getByRole('button', { name: 'Undo', exact: true })
  const redoButton = page.getByRole('button', { name: 'Redo', exact: true })
  const nameBox = page.getByTestId('name-box')
  const formulaInput = page.getByTestId('formula-input')

  await nameBox.fill('A1')
  await nameBox.press('Enter')
  await formulaInput.fill('alpha')
  await formulaInput.press('Enter')

  await nameBox.fill('B1')
  await nameBox.press('Enter')
  await formulaInput.fill('beta')
  await formulaInput.press('Enter')

  await nameBox.fill('C1')
  await nameBox.press('Enter')
  await formulaInput.fill('gamma')
  await formulaInput.press('Enter')

  await expect(undoButton).toBeEnabled()
  await undoButton.click()
  await expect(redoButton).toBeEnabled()
  await expect(undoButton).toBeEnabled()
  await undoButton.click()
  await expect(redoButton).toBeEnabled()
  await expect(undoButton).toBeEnabled()
  await undoButton.click()
  await expect(redoButton).toBeEnabled()

  await redoButton.click()
  await expect(redoButton).toBeEnabled()
  await expect(undoButton).toBeEnabled()

  await redoButton.click()
  await expect(redoButton).toBeEnabled()
  await expect(undoButton).toBeEnabled()

  await redoButton.click()
  await expect(redoButton).toBeDisabled()

  await nameBox.fill('A1')
  await nameBox.press('Enter')
  await expect(formulaInput).toHaveValue('alpha')
  await nameBox.fill('B1')
  await nameBox.press('Enter')
  await expect(formulaInput).toHaveValue('beta')
  await nameBox.fill('C1')
  await nameBox.press('Enter')
  await expect(formulaInput).toHaveValue('gamma')
})

test('web app clears redo after a fresh edit branches history', async ({ page }) => {
  test.skip(!remoteSyncEnabled, 'requires authoritative remote sync history')
  const documentId = `playwright-undo-redo-branch-${Date.now()}`
  await page.goto(`/?document=${encodeURIComponent(documentId)}`)
  await waitForWorkbookReady(page)
  await expect(page.getByTestId('status-sync')).toHaveText('Saved', { timeout: 30_000 })

  const undoButton = page.getByRole('button', { name: 'Undo', exact: true })
  const redoButton = page.getByRole('button', { name: 'Redo', exact: true })
  const nameBox = page.getByTestId('name-box')
  const formulaInput = page.getByTestId('formula-input')

  await nameBox.fill('A1')
  await nameBox.press('Enter')
  await formulaInput.fill('seed')
  await formulaInput.press('Enter')

  await undoButton.click()
  await expect(redoButton).toBeEnabled()

  await nameBox.fill('D1')
  await nameBox.press('Enter')
  await formulaInput.fill('branch')
  await formulaInput.press('Enter')

  await expect(redoButton).toBeDisabled()
})

test('web app previews and fills rightward autofill like Sheets', async ({ page }) => {
  await page.goto('/')
  await waitForWorkbookReady(page)

  const nameBox = page.getByTestId('name-box')
  const formulaInput = page.getByTestId('formula-input')
  const resolvedValue = page.getByTestId('formula-resolved-value')
  const selectionStatus = page.getByTestId('status-selection')
  const fillPreview = page.locator("[data-grid-fill-preview='true']")

  await nameBox.fill('F6')
  await nameBox.press('Enter')
  await formulaInput.fill('7')
  await formulaInput.press('Enter')

  const { sourceX, sourceY, targetX, targetY } = await getProductFillHandleDragPoints(page, 5, 5, 7, 5)
  await page.mouse.move(sourceX, sourceY)
  await page.mouse.down()
  await page.mouse.move(targetX, targetY, { steps: 10 })

  await expect(fillPreview).toBeVisible()
  await expect(fillPreview).toHaveCSS('border-top-style', 'dashed')

  await page.mouse.up()

  await expect(selectionStatus).toContainText('!F6:H6')

  await nameBox.fill('H6')
  await nameBox.press('Enter')
  await expect(formulaInput).toHaveValue('7')
  await expect(resolvedValue).toHaveText('7')
})

test('web app supports rectangular clipboard copy and external paste', async ({ page, context }) => {
  await context.grantPermissions(['clipboard-read', 'clipboard-write'])
  await page.goto('/')
  await waitForWorkbookReady(page)

  const grid = page.getByTestId('sheet-grid')
  const nameBox = page.getByTestId('name-box')
  const formulaInput = page.getByTestId('formula-input')
  const resolvedValue = page.getByTestId('formula-resolved-value')

  await nameBox.fill('B2')
  await nameBox.press('Enter')
  await formulaInput.fill('11')
  await formulaInput.press('Enter')

  await nameBox.fill('C2')
  await nameBox.press('Enter')
  await formulaInput.fill('12')
  await formulaInput.press('Enter')

  await nameBox.fill('B3')
  await nameBox.press('Enter')
  await formulaInput.fill('13')
  await formulaInput.press('Enter')

  await nameBox.fill('C3')
  await nameBox.press('Enter')
  await formulaInput.fill('14')
  await formulaInput.press('Enter')

  await dragProductBodySelection(page, 1, 1, 2, 2)
  await grid.press(`${PRIMARY_MODIFIER}+C`)

  await expect.poll(() => page.evaluate(() => navigator.clipboard.readText())).toBe('11\t12\n13\t14')

  await page.evaluate(() => navigator.clipboard.writeText('21\t22\n23\t24'))
  await clickProductCell(page, 4, 4)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!E5')
  await grid.press(`${PRIMARY_MODIFIER}+V`)

  await nameBox.fill('E5')
  await nameBox.press('Enter')
  await expect(formulaInput).toHaveValue('21')
  await expect(resolvedValue).toHaveText('21')

  await nameBox.fill('F5')
  await nameBox.press('Enter')
  await expect(formulaInput).toHaveValue('22')
  await expect(resolvedValue).toHaveText('22')

  await nameBox.fill('E6')
  await nameBox.press('Enter')
  await expect(formulaInput).toHaveValue('23')
  await expect(resolvedValue).toHaveText('23')

  await nameBox.fill('F6')
  await nameBox.press('Enter')
  await expect(formulaInput).toHaveValue('24')
  await expect(resolvedValue).toHaveText('24')
})

test('web app relocates formulas when using rectangular clipboard paste', async ({ page, context }) => {
  await context.grantPermissions(['clipboard-read', 'clipboard-write'])
  await page.goto('/')
  await waitForWorkbookReady(page)

  const grid = page.getByTestId('sheet-grid')
  const nameBox = page.getByTestId('name-box')
  const formulaInput = page.getByTestId('formula-input')
  const resolvedValue = page.getByTestId('formula-resolved-value')

  await nameBox.fill('B2')
  await nameBox.press('Enter')
  await formulaInput.fill('3')
  await formulaInput.press('Enter')

  await nameBox.fill('B3')
  await nameBox.press('Enter')
  await formulaInput.fill('4')
  await formulaInput.press('Enter')

  await nameBox.fill('C2')
  await nameBox.press('Enter')
  await formulaInput.fill('=B2*2')
  await formulaInput.press('Enter')

  await nameBox.fill('C3')
  await nameBox.press('Enter')
  await formulaInput.fill('=B3*2')
  await formulaInput.press('Enter')

  await dragProductBodySelection(page, 1, 1, 2, 2)
  await grid.press(`${PRIMARY_MODIFIER}+C`)

  await clickProductCell(page, 3, 1)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!D2')
  await grid.press(`${PRIMARY_MODIFIER}+V`)

  await nameBox.fill('D2')
  await nameBox.press('Enter')
  await expect(formulaInput).toHaveValue('3')
  await expect(resolvedValue).toHaveText('3')

  await nameBox.fill('E2')
  await nameBox.press('Enter')
  await expect(formulaInput).toHaveValue('=D2*2')
  await expect(resolvedValue).toHaveText('6')

  await nameBox.fill('E3')
  await nameBox.press('Enter')
  await expect(formulaInput).toHaveValue('=D3*2')
  await expect(resolvedValue).toHaveText('8')
})

test('web app supports product-shell column resize', async ({ page }) => {
  await page.goto('/')
  await waitForWorkbookReady(page)

  const baselineWidth = await getProductColumnWidth(page, 0)
  await dragProductColumnResize(page, 0, 48)
  await expect.poll(() => getProductColumnWidth(page, 0)).toBeGreaterThan(baselineWidth + 30)
})

test('web app shows #VALUE! for invalid formulas', async ({ page }) => {
  const documentId = `playwright-invalid-formula-${Date.now()}`
  await page.goto(`/?document=${encodeURIComponent(documentId)}`)
  await waitForWorkbookReady(page)

  const nameBox = page.getByTestId('name-box')
  const formulaInput = page.getByTestId('formula-input')
  const resolvedValue = page.getByTestId('formula-resolved-value')

  await nameBox.fill('A1')
  await nameBox.press('Enter')
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!A1')

  await formulaInput.fill('=1+')
  await expect(formulaInput).toHaveValue('=1+')
  await formulaInput.press('Enter')

  await expect(formulaInput).toHaveValue('#VALUE!')
  await expect(resolvedValue).toHaveText('#VALUE!')
})

test('web app commits in-cell string edits when clicking away', async ({ page }) => {
  await page.goto('/')
  await waitForWorkbookReady(page)

  const grid = page.getByTestId('sheet-grid')
  const nameBox = page.getByTestId('name-box')
  const formulaInput = page.getByTestId('formula-input')
  const resolvedValue = page.getByTestId('formula-resolved-value')
  const cellEditor = page.getByTestId('cell-editor-input')

  await clickProductCell(page, 1, 0)
  await expect(nameBox).toHaveValue('B1')
  await grid.press('h')
  await expect(cellEditor).toBeVisible()
  await expect(cellEditor).toHaveValue('h')
  await clickProductCell(page, 2, 0)

  await expect(nameBox).toHaveValue('C1')
  await clickProductCell(page, 1, 0)
  await expect(nameBox).toHaveValue('B1')
  await expect(formulaInput).toHaveValue('h')
  await expect(resolvedValue).toHaveText('h')
})

test('web app drags a selected range by its border with a grab cursor', async ({ page }) => {
  await page.goto('/')
  await waitForWorkbookReady(page)

  const nameBox = page.getByTestId('name-box')
  const formulaInput = page.getByTestId('formula-input')
  const resolvedValue = page.getByTestId('formula-resolved-value')

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

  await dragProductSelectionBorder(page, 1, 1, 2, 1, 3, 3)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!D4:E4')

  await nameBox.fill('B2')
  await nameBox.press('Enter')
  await expect(formulaInput).toHaveValue('')
  await expect(resolvedValue).toHaveText('∅')

  await nameBox.fill('C2')
  await nameBox.press('Enter')
  await expect(formulaInput).toHaveValue('')
  await expect(resolvedValue).toHaveText('∅')

  await nameBox.fill('D4')
  await nameBox.press('Enter')
  await expect(formulaInput).toHaveValue('left')
  await expect(resolvedValue).toHaveText('left')

  await nameBox.fill('E4')
  await nameBox.press('Enter')
  await expect(formulaInput).toHaveValue('right')
  await expect(resolvedValue).toHaveText('right')
})

test('web app applies core formatting shortcuts from the keyboard', async ({ page }) => {
  await page.goto('/')
  await waitForWorkbookReady(page)

  const grid = page.getByTestId('sheet-grid')
  await clickProductCell(page, 0, 0)
  await grid.press(`${PRIMARY_MODIFIER}+B`)
  await expect(page.getByLabel('Bold')).toHaveClass(/bg-\[var\(--wb-accent-soft\)\]/)
  await grid.press(`${PRIMARY_MODIFIER}+I`)
  await expect(page.getByLabel('Italic')).toHaveClass(/bg-\[var\(--wb-accent-soft\)\]/)
  await grid.press(`${PRIMARY_MODIFIER}+U`)
  await expect(page.getByLabel('Underline')).toHaveClass(/bg-\[var\(--wb-accent-soft\)\]/)
  await grid.press(`${PRIMARY_MODIFIER}+Shift+E`)
  await expect(page.getByLabel('Align center')).toHaveClass(/bg-\[var\(--wb-accent-soft\)\]/)
  await grid.press(`${PRIMARY_MODIFIER}+Shift+R`)
  await expect(page.getByLabel('Align right')).toHaveClass(/bg-\[var\(--wb-accent-soft\)\]/)
  await grid.press(`${PRIMARY_MODIFIER}+Shift+L`)
  await expect(page.getByLabel('Align left')).toHaveClass(/bg-\[var\(--wb-accent-soft\)\]/)
})

test('web app supports row, column, and full-sheet selection shortcuts', async ({ page }) => {
  await page.goto('/')
  await waitForWorkbookReady(page)

  const grid = page.getByTestId('sheet-grid')
  await clickProductCell(page, 2, 4)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!C5')

  await grid.press('Shift+Space')
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!5:5')

  await grid.press(`${PRIMARY_MODIFIER}+Space`)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!C:C')

  await grid.press(`${PRIMARY_MODIFIER}+Shift+Space`)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!All')

  await grid.press(`${PRIMARY_MODIFIER}+A`)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!All')
})

test('web app expands the active range with repeated shift arrows', async ({ page }) => {
  await page.goto('/')
  await waitForWorkbookReady(page)

  const grid = page.getByTestId('sheet-grid')
  await clickProductCell(page, 2, 4)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!C5')

  await grid.press('Shift+ArrowRight')
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!C5:D5')

  await grid.press('Shift+ArrowRight')
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!C5:E5')

  await grid.press('Shift+ArrowDown')
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!C5:E6')
})

test('web app expands the active range with shift-click', async ({ page }) => {
  await page.goto('/')
  await waitForWorkbookReady(page)

  await clickProductCell(page, 1, 1)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!B2')

  await clickProductCell(page, 4, 5, { shift: true })
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!B2:E6')
})

for (const key of ['Delete', 'Backspace'] as const) {
  test(`web app clears the full selected range with ${key.toLowerCase()}`, async ({ page, context }) => {
    await context.grantPermissions(['clipboard-read', 'clipboard-write'])
    await page.goto('/')
    await waitForWorkbookReady(page)

    const grid = page.getByTestId('sheet-grid')
    const formulaInput = page.getByTestId('formula-input')

    await clickProductCell(page, 1, 1)
    await page.evaluate(() => navigator.clipboard.writeText('11\t12\n13\t14'))
    await grid.press(`${PRIMARY_MODIFIER}+V`)
    await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!B2')

    await dragProductBodySelection(page, 1, 1, 2, 2)
    await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!B2:C3')

    await grid.press(key)

    await clickProductCell(page, 1, 1)
    await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!B2')
    await expect(formulaInput).toHaveValue('')
    await clickProductCell(page, 2, 1)
    await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!C2')
    await expect(formulaInput).toHaveValue('')
    await clickProductCell(page, 1, 2)
    await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!B3')
    await expect(formulaInput).toHaveValue('')
    await clickProductCell(page, 2, 2)
    await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!C3')
    await expect(formulaInput).toHaveValue('')
  })
}

test('web app ignores right gutter clicks', async ({ page }) => {
  await page.goto('/')
  await waitForWorkbookReady(page)

  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!A1')
  await clickGridRightEdge(page, 3)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!A1')
})

test('@fuzz-browser web app preserves valid selection geometry and focus under generated selection actions', async ({ page }) => {
  test.skip(!fuzzBrowserEnabled || !shouldRunFuzzSuite('browser/grid-selection-focus', 'browser'), 'browser fuzz runs only in fuzz mode')

  await runProperty({
    suite: 'browser/grid-selection-focus',
    kind: 'browser',
    arbitrary: fc.array(
      fc.oneof<BrowserSelectionAction>(
        fc.record({
          kind: fc.constant<'click'>('click'),
          row: fc.integer({ min: 0, max: 8 }),
          col: fc.integer({ min: 0, max: 8 }),
        }),
        fc.record({
          kind: fc.constant<'shiftClick'>('shiftClick'),
          row: fc.integer({ min: 0, max: 8 }),
          col: fc.integer({ min: 0, max: 8 }),
        }),
        fc.record({
          kind: fc.constant<'key'>('key'),
          key: fc.constantFrom('ArrowLeft', 'ArrowRight', 'ArrowUp', 'ArrowDown'),
          shift: fc.boolean(),
        }),
      ),
      { minLength: 6, maxLength: 10 },
    ),
    parameters: {
      interruptAfterTimeLimit: 40_000,
    },
    predicate: async (actions) => {
      await gotoWorkbookShell(page)
      await waitForWorkbookReady(page)
      const grid = page.getByTestId('sheet-grid')
      const nameBox = page.getByTestId('name-box')
      await expect(grid).toBeVisible({ timeout: 15_000 })
      await nameBox.fill('C5')
      await nameBox.press('Enter')
      await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!C5')
      await runSelectionFuzzActions(page, grid, actions)
    },
  })
})
