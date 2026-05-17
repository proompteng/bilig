import { expect, test, type Locator } from '@playwright/test'
import * as fc from 'fast-check'
import { runProperty, shouldRunFuzzSuite } from '../../packages/test-fuzz/src/index.ts'
import {
  PRIMARY_MODIFIER,
  PRODUCT_HEADER_HEIGHT,
  PRODUCT_ROW_HEIGHT,
  clickProductCell,
  clickGridRightEdge,
  createTestDocumentId,
  dragProductBodySelection,
  dragProductColumnResize,
  getProductFillHandleDragPoints,
  getProductColumnLeft,
  getProductColumnWidth,
  gotoWorkbookShell,
  remoteSyncEnabled,
  waitForProductColumnWidthChange,
  waitForWorkbookReady,
} from './web-shell-helpers.js'
const fuzzBrowserEnabled = process.env['BILIG_FUZZ_BROWSER'] === '1'
const remoteSyncTest = remoteSyncEnabled ? test : test.skip.bind(test)
const fuzzBrowserTest = fuzzBrowserEnabled && shouldRunFuzzSuite('browser/grid-selection-focus', 'browser') ? test : test.skip.bind(test)

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

async function nativeTextRunsInclude(page: Parameters<typeof test>[0]['page'], text: string): Promise<boolean> {
  return await page.evaluate(
    (needle) => Array.from(document.querySelectorAll('[data-native-text-run]')).some((run) => run.textContent?.includes(needle) ?? false),
    text,
  )
}

async function textControlValue(locator: Locator): Promise<string> {
  return await locator.evaluate((control) =>
    control instanceof HTMLInputElement || control instanceof HTMLTextAreaElement ? control.value : '',
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

async function dragProductSelectedContentLane(
  page: Parameters<typeof test>[0]['page'],
  startColumn: number,
  startRow: number,
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
  const startWidth = await getProductColumnWidth(page, startColumn)
  const sourceX = grid.x + startLeft + Math.min(32, Math.floor(startWidth * 0.35))
  const sourceY = grid.y + PRODUCT_HEADER_HEIGHT + startRow * PRODUCT_ROW_HEIGHT + Math.floor(PRODUCT_ROW_HEIGHT / 2)
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
  const documentId = createTestDocumentId('playwright-string-comparison')
  await page.goto(`/?document=${encodeURIComponent(documentId)}`)
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

  const nameBox = page.getByTestId('name-box')
  const formulaInput = page.getByTestId('formula-input')
  const cellEditor = page.getByTestId('cell-editor-input')

  await clickProductCell(page, 0, 0)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!A1')
  await page.keyboard.press('h')
  await expect(cellEditor).toBeVisible()
  await expect(cellEditor).toHaveValue('h')
  await page.keyboard.press('Enter')
  await expect(cellEditor).toBeHidden()

  await expect(nameBox).toHaveValue('A2', { timeout: 15_000 })
  await clickProductCell(page, 0, 0)
  await expect(formulaInput).toHaveValue('h')

  await clickProductCell(page, 0, 1)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!A2')
  await page.keyboard.press('w')
  await expect(cellEditor).toBeVisible()
  await expect(cellEditor).toHaveValue('w')
  await page.keyboard.press('Tab')
  await expect(cellEditor).toBeHidden()

  await expect(nameBox).toHaveValue('B2', { timeout: 15_000 })
  await clickProductCell(page, 0, 1)
  await expect(formulaInput).toHaveValue('w')

  await page.keyboard.press('Enter')
  await expect(nameBox).toHaveValue('A3', { timeout: 15_000 })
  await page.keyboard.press('Shift+Enter')
  await expect(nameBox).toHaveValue('A2', { timeout: 15_000 })
})

test('@browser-ci web app keeps click-away commits and keyboard clears stable', async ({ page }) => {
  const documentId = createTestDocumentId('playwright-click-away-clear')
  await page.goto(`/?document=${encodeURIComponent(documentId)}&persist=0&sheet=Sheet1&cell=D7`)
  await waitForWorkbookReady(page)

  const formulaInput = page.getByTestId('formula-input')
  const resolvedValue = page.getByTestId('formula-resolved-value')
  const cellEditor = page.getByTestId('cell-editor-input')

  await clickProductCell(page, 3, 6)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!D7')
  await page.keyboard.type('stable-proof')
  await expect(cellEditor).toBeVisible()
  await expect(cellEditor).toHaveValue('stable-proof')
  await clickProductCell(page, 4, 6)
  await expect(cellEditor).toBeHidden()

  await clickProductCell(page, 3, 6)
  await expect(formulaInput).toHaveValue('stable-proof')
  await page.keyboard.press('Delete')
  await expect(formulaInput).toHaveValue('')
  await clickProductCell(page, 4, 6)
  await clickProductCell(page, 3, 6)
  await expect(formulaInput).toHaveValue('')
  await expect(resolvedValue).toHaveText('∅')

  await page.keyboard.type('backspace-proof')
  await expect(cellEditor).toBeVisible()
  await expect(cellEditor).toHaveValue('backspace-proof')
  await clickProductCell(page, 4, 6)
  await expect(cellEditor).toBeHidden()

  await clickProductCell(page, 3, 6)
  await expect(formulaInput).toHaveValue('backspace-proof')
  await page.keyboard.press('Backspace')
  await expect(formulaInput).toHaveValue('')
  await clickProductCell(page, 4, 6)
  await clickProductCell(page, 3, 6)
  await expect(formulaInput).toHaveValue('')
  await expect(resolvedValue).toHaveText('∅')
})

test('@browser-ci web app recovers after runtime config failures outlive the fast retry window', async ({ page }) => {
  let runtimeConfigAttempts = 0
  await page.route('**/runtime-config.json', async (route) => {
    runtimeConfigAttempts += 1
    if (runtimeConfigAttempts <= 5) {
      await route.fulfill({
        body: 'temporary runtime config failure',
        contentType: 'text/plain',
        status: 502,
      })
      return
    }
    await route.continue()
  })

  await page.goto(`/?document=${encodeURIComponent(createTestDocumentId('playwright-runtime-config-recovery'))}&persist=0`)

  await expect(page.getByTestId('worker-error')).toContainText('Failed to load runtime config (502)', { timeout: 6_000 })
  await expect(page.getByTestId('formula-bar')).toBeVisible({ timeout: 20_000 })
  await expect(page.getByTestId('sheet-grid')).toBeVisible({ timeout: 15_000 })
  await expect(page.getByTestId('worker-error')).toHaveCount(0)
  expect(runtimeConfigAttempts).toBeGreaterThan(5)
})

test('@browser-ci web app keeps an editor clear after click-away selection', async ({ page }) => {
  const documentId = createTestDocumentId('playwright-editor-clear-click-away')
  await page.goto(`/?document=${encodeURIComponent(documentId)}&persist=0&sheet=Sheet1&cell=A1`)
  await waitForWorkbookReady(page)

  const gridLocator = page.getByTestId('sheet-grid')
  const nameBox = page.getByTestId('name-box')
  const formulaInput = page.getByTestId('formula-input')
  const cellEditor = page.getByTestId('cell-editor-input')

  await clickProductCell(page, 0, 0)
  await page.keyboard.type('ghost-value')
  await page.keyboard.press('Enter')
  await expect(nameBox).toHaveValue('A2', { timeout: 15_000 })

  const grid = await gridLocator.boundingBox()
  if (!grid) {
    throw new Error('sheet grid is not visible')
  }
  const columnLeft = await getProductColumnLeft(page, 0)
  const columnWidth = await getProductColumnWidth(page, 0)
  await page.mouse.dblclick(
    grid.x + columnLeft + Math.floor(columnWidth / 2),
    grid.y + PRODUCT_HEADER_HEIGHT + Math.floor(PRODUCT_ROW_HEIGHT / 2),
  )

  await expect(cellEditor).toBeVisible()
  await expect(cellEditor).toBeFocused()
  await page.keyboard.press(`${PRIMARY_MODIFIER}+A`)
  await page.keyboard.press('Backspace')
  await expect(cellEditor).toHaveValue('')

  await clickProductCell(page, 2, 3)
  await expect(cellEditor).toBeHidden()
  await expect(nameBox).toHaveValue('C4', { timeout: 15_000 })
  await expect(formulaInput).toHaveValue('')
  await expect.poll(() => nativeTextRunsInclude(page, 'ghost-value')).toBe(false)

  await clickProductCell(page, 0, 0)
  await expect(nameBox).toHaveValue('A1', { timeout: 15_000 })
  await expect(formulaInput).toHaveValue('')
  await expect.poll(() => nativeTextRunsInclude(page, 'ghost-value')).toBe(false)
})

test('@browser-ci web app gates unmerge on real merged-cell state', async ({ page }) => {
  const documentId = createTestDocumentId('playwright-structure-unmerge-availability')
  await page.goto(`/?document=${encodeURIComponent(documentId)}&persist=0&sheet=Sheet1&cell=A1`)
  await waitForWorkbookReady(page)

  const structureButton = page.getByRole('button', { name: 'Structure' })
  const unmergeButton = page.getByRole('button', { exact: true, name: 'Unmerge cells' })

  await clickProductCell(page, 0, 0)
  await structureButton.click()
  await expect(unmergeButton).toBeDisabled()
  await page.keyboard.press('Escape')

  await dragProductBodySelection(page, 1, 1, 2, 2)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!B2:C3')
  await structureButton.click()
  await page.getByRole('button', { exact: true, name: 'Merge cells' }).click()

  await structureButton.click()
  await expect(unmergeButton).toBeEnabled()
  await unmergeButton.click()

  await structureButton.click()
  await expect(unmergeButton).toBeDisabled()
})

test('web app preserves Alt+Enter multiline edits across commit, formula bar, and reopen', async ({ page }) => {
  const documentId = createTestDocumentId('playwright-alt-enter-multiline-edit')
  await page.goto(`/?document=${encodeURIComponent(documentId)}&sheet=Sheet1&cell=A1`)
  await waitForWorkbookReady(page)

  const formulaInput = page.getByTestId('formula-input')
  const cellEditor = page.getByTestId('cell-editor-input')

  await clickProductCell(page, 0, 0)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!A1')
  await page.keyboard.type('alpha')
  await expect(cellEditor).toBeVisible()
  await expect(cellEditor).toHaveValue('alpha')

  await cellEditor.press('Alt+Enter')
  await page.keyboard.type('beta')
  await expect(cellEditor).toHaveValue('alpha\nbeta')

  await cellEditor.press('Enter')
  await expect(cellEditor).toBeHidden()
  await clickProductCell(page, 0, 0)
  await expect(formulaInput).toHaveValue('alpha\nbeta')

  await page.getByTestId('sheet-grid').press('F2')
  await expect(cellEditor).toBeVisible()
  await expect(cellEditor).toHaveValue('alpha\nbeta')
})

test('web app preserves multi-digit numeric type-to-replace input', async ({ page }) => {
  await page.goto('/')
  await waitForWorkbookReady(page)

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

test('@browser-serial web app supports F2 edit in the product shell', async ({ page }) => {
  await page.goto('/')
  await waitForWorkbookReady(page)
  await waitForWorkbookReady(page)

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

test('@browser-ci web app offers formula autocomplete and inserts a function with Tab', async ({ page }) => {
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
  const sampleText = 'visible text sample'

  await nameBox.fill('C5')
  await nameBox.press('Enter')
  await formulaInput.fill(sampleText)
  await formulaInput.press('Enter')

  await clickProductCell(page, 0, 0)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!A1')

  await clickProductCell(page, 2, 4)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!C5')
  await expect(formulaInput).toHaveValue(sampleText)
  const supportsWebGpu = await page.evaluate(() => 'gpu' in navigator)
  if (supportsWebGpu) {
    await expect(page.getByTestId('grid-pane-renderer')).toHaveJSProperty('tagName', 'CANVAS')
    await expect(page.getByTestId('grid-text-pane-body')).toHaveCount(0)
  } else {
    await expect(page.getByTestId('grid-text-pane-body')).toHaveJSProperty('tagName', 'CANVAS')
    await expect(page.getByTestId('grid-text-pane-top-body')).toHaveJSProperty('tagName', 'CANVAS')
    await expect(page.getByTestId('grid-text-pane-left-body')).toHaveJSProperty('tagName', 'CANVAS')
  }
})

test('@browser-ci web app supports fill-handle propagation', async ({ page }) => {
  await gotoWorkbookShell(page, `/?document=${encodeURIComponent(createTestDocumentId('fill-handle-propagation'))}`)
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
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!F8')
  await expect(formulaInput).toHaveValue('7')
  await expect(resolvedValue).toHaveText('7')
})

remoteSyncTest('web app enables undo and redo for a normal edit', async ({ page }) => {
  const documentId = createTestDocumentId('playwright-undo-redo-basic')
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

remoteSyncTest('web app preserves redo across a longer undo history', async ({ page }) => {
  const documentId = createTestDocumentId('playwright-undo-redo-long')
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

remoteSyncTest('web app clears redo after a fresh edit branches history', async ({ page }) => {
  const documentId = createTestDocumentId('playwright-undo-redo-branch')
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
  await gotoWorkbookShell(page, `/?document=${encodeURIComponent(createTestDocumentId('rightward-autofill'))}`)
  await waitForWorkbookReady(page)

  const nameBox = page.getByTestId('name-box')
  const formulaInput = page.getByTestId('formula-input')
  const resolvedValue = page.getByTestId('formula-resolved-value')
  const selectionStatus = page.getByTestId('status-selection')
  const fillPreview = page.locator("[data-grid-fill-preview='true']")
  const renderer = page.getByTestId('grid-pane-renderer')

  await nameBox.fill('F6')
  await nameBox.press('Enter')
  await formulaInput.fill('7')
  await formulaInput.press('Enter')

  const { sourceX, sourceY, targetX, targetY } = await getProductFillHandleDragPoints(page, 5, 5, 7, 5)
  await page.mouse.move(sourceX, sourceY)
  await page.mouse.down()
  await page.mouse.move(targetX, targetY, { steps: 10 })

  await expect(renderer).toBeVisible()
  await expect(fillPreview).toHaveCount(0)

  await page.mouse.up()

  await expect(selectionStatus).toContainText('!F6:H6')

  await nameBox.fill('H6')
  await nameBox.press('Enter')
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!H6')
  await expect(formulaInput).toHaveValue('7')
  await expect(resolvedValue).toHaveText('7')
})

test.describe('@clipboard-global web app clipboard flows', () => {
  test.describe.configure({ mode: 'serial' })

  test('web app supports rectangular clipboard copy and external paste', async ({ page, context }) => {
    await context.grantPermissions(['clipboard-read', 'clipboard-write'])
    await page.goto('/')
    await waitForWorkbookReady(page)

    const grid = page.getByTestId('sheet-grid')
    const nameBox = page.getByTestId('name-box')
    const formulaInput = page.getByTestId('formula-input')
    const resolvedValue = page.getByTestId('formula-resolved-value')

    const writeFormulaBarCell = async (address: string, value: string) => {
      await nameBox.fill(address)
      await expect(nameBox).toHaveValue(address)
      await nameBox.press('Enter')
      await expect(page.getByTestId('status-selection')).toHaveText(`Sheet1!${address}`)
      await formulaInput.fill(value)
      await expect(formulaInput).toHaveValue(value)
      await formulaInput.press('Enter')
      await nameBox.fill(address)
      await expect(nameBox).toHaveValue(address)
      await nameBox.press('Enter')
      await expect(page.getByTestId('status-selection')).toHaveText(`Sheet1!${address}`)
      await expect(formulaInput).toHaveValue(value)
      await expect(resolvedValue).toHaveText(value)
    }

    await writeFormulaBarCell('B2', '11')
    await writeFormulaBarCell('C2', '12')
    await writeFormulaBarCell('B3', '13')
    await writeFormulaBarCell('C3', '14')

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

    const writeFormulaBarCell = async (address: string, value: string, resolved = value) => {
      await nameBox.fill(address)
      await nameBox.press('Enter')
      await formulaInput.fill(value)
      await formulaInput.press('Enter')
      await nameBox.fill(address)
      await nameBox.press('Enter')
      await expect(formulaInput).toHaveValue(value)
      await expect(resolvedValue).toHaveText(resolved)
    }

    await writeFormulaBarCell('B2', '3')
    await writeFormulaBarCell('B3', '4')
    await writeFormulaBarCell('C2', '=B2*2', '6')
    await writeFormulaBarCell('C3', '=B3*2', '8')

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

  test('web app moves rectangular ranges with the cut keyboard shortcut', async ({ page, context }) => {
    await context.grantPermissions(['clipboard-read', 'clipboard-write'])
    const documentId = createTestDocumentId('playwright-clipboard-cut-move')
    await page.goto(`/?document=${encodeURIComponent(documentId)}&persist=0`)
    await waitForWorkbookReady(page)

    const grid = page.getByTestId('sheet-grid')
    const nameBox = page.getByTestId('name-box')
    const formulaInput = page.getByTestId('formula-input')
    const resolvedValue = page.getByTestId('formula-resolved-value')

    const writeFormulaBarCell = async (address: string, value: string) => {
      await nameBox.fill(address)
      await expect(nameBox).toHaveValue(address)
      await nameBox.press('Enter')
      await expect(page.getByTestId('status-selection')).toHaveText(`Sheet1!${address}`)
      await formulaInput.fill(value)
      await expect(formulaInput).toHaveValue(value)
      await formulaInput.press('Enter')
      await nameBox.fill(address)
      await expect(nameBox).toHaveValue(address)
      await nameBox.press('Enter')
      await expect(page.getByTestId('status-selection')).toHaveText(`Sheet1!${address}`)
      await expect(formulaInput).toHaveValue(value)
      await expect(resolvedValue).toHaveText(value)
    }

    const expectCellValue = async (address: string, value: string, resolved = value) => {
      await nameBox.fill(address)
      await nameBox.press('Enter')
      await expect(formulaInput).toHaveValue(value)
      await expect(resolvedValue).toHaveText(resolved)
    }

    await writeFormulaBarCell('B2', 'cut-a')
    await writeFormulaBarCell('C2', 'cut-b')
    await writeFormulaBarCell('B3', 'cut-c')
    await writeFormulaBarCell('C3', 'cut-d')

    await dragProductBodySelection(page, 1, 1, 2, 2)
    await grid.press(`${PRIMARY_MODIFIER}+X`)
    await expect.poll(() => page.evaluate(() => navigator.clipboard.readText())).toBe('cut-a\tcut-b\ncut-c\tcut-d')

    await clickProductCell(page, 4, 4)
    await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!E5')
    await grid.press(`${PRIMARY_MODIFIER}+V`)

    await expectCellValue('E5', 'cut-a')
    await expectCellValue('F5', 'cut-b')
    await expectCellValue('E6', 'cut-c')
    await expectCellValue('F6', 'cut-d')
    await expectCellValue('B2', '', '∅')
    await expectCellValue('C2', '', '∅')
    await expectCellValue('B3', '', '∅')
    await expectCellValue('C3', '', '∅')
  })
})

test('web app supports product-shell column resize', async ({ page }) => {
  const documentId = createTestDocumentId('playwright-product-shell-column-resize')
  await page.goto(`/?document=${encodeURIComponent(documentId)}`)
  await waitForWorkbookReady(page)

  const baselineWidth = await getProductColumnWidth(page, 0)
  const committedWidthPromise = waitForProductColumnWidthChange(page, 0, baselineWidth)
  await dragProductColumnResize(page, 0, 48)
  await expect(committedWidthPromise).resolves.toBeGreaterThan(baselineWidth + 30)
})

test('web app shows #VALUE! for invalid formulas', async ({ page }) => {
  const documentId = createTestDocumentId('playwright-invalid-formula')
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

test('web app focuses the name box from the Go To keyboard shortcut', async ({ page }) => {
  const documentId = createTestDocumentId('playwright-goto-shortcut')
  await page.goto(`/?document=${encodeURIComponent(documentId)}`)
  await waitForWorkbookReady(page)

  const nameBox = page.getByTestId('name-box')
  await clickProductCell(page, 0, 0)
  await page.keyboard.press(`${PRIMARY_MODIFIER}+G`)
  await page.keyboard.up(PRIMARY_MODIFIER)

  await expect(nameBox).toBeFocused()
  await nameBox.fill('C12')
  await nameBox.press('Enter')

  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!C12')
  await expect(nameBox).toHaveValue('C12')
})

test('web app supports Google Sheets-style shortcut help and sheet switching keys', async ({ page }) => {
  const documentId = createTestDocumentId('playwright-google-sheets-shortcut-parity')
  await page.goto(`/?document=${encodeURIComponent(documentId)}&sheet=Sheet1&cell=C22`)
  await waitForWorkbookReady(page)

  const grid = page.getByTestId('sheet-grid')
  await page.getByTestId('workbook-sheet-add').click()
  await expect(page.getByTestId('workbook-sheet-tab-Sheet2')).toBeVisible()

  await page.goto(`/?document=${encodeURIComponent(documentId)}&sheet=Sheet1&cell=C22`)
  await waitForWorkbookReady(page)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!C22')

  await grid.press('Alt+ArrowDown')
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet2!C22')

  await grid.press('Alt+ArrowUp')
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!C22')

  await grid.press(`${PRIMARY_MODIFIER}+/`)
  await expect(page.getByTestId('workbook-shortcut-dialog')).toBeVisible()
  await expect(page.getByTestId('workbook-shortcut-search')).toBeFocused()
})

test('web app commits in-cell string edits when clicking away', async ({ page }) => {
  await page.keyboard.up(PRIMARY_MODIFIER)
  await page.goto(`/?document=${encodeURIComponent(createTestDocumentId('playwright-click-away-edit'))}`)
  await waitForWorkbookReady(page)

  const nameBox = page.getByTestId('name-box')
  const formulaInput = page.getByTestId('formula-input')
  const resolvedValue = page.getByTestId('formula-resolved-value')
  const cellEditor = page.getByTestId('cell-editor-input')

  await clickProductCell(page, 1, 0)
  await expect(nameBox).toHaveValue('B1')
  await page.getByTestId('sheet-grid-focus-target').focus()
  await page.keyboard.press('a')
  await expect(cellEditor).toBeVisible()
  await expect(cellEditor).toHaveValue('a')
  expect(await textControlValue(formulaInput)).toBe('a')
  await expect
    .poll(async () => await cellEditor.evaluate((input) => (input instanceof HTMLTextAreaElement ? input.selectionStart : -1)))
    .toBe(1)
  const pressRemainingText = async (remainingCharacters: readonly string[], previousText: string): Promise<void> => {
    const [character, ...rest] = remainingCharacters
    if (!character) {
      return
    }
    const nextText = `${previousText}${character}`
    await cellEditor.press(character)
    await expect(cellEditor).toHaveValue(nextText)
    expect(await textControlValue(formulaInput)).toBe(nextText)
    await expect
      .poll(async () => await cellEditor.evaluate((input) => (input instanceof HTMLTextAreaElement ? input.selectionStart : -1)))
      .toBe(nextText.length)
    await pressRemainingText(rest, nextText)
  }
  await pressRemainingText(['b', 'c', 'd', 'e', 'f'], 'a')
  await clickProductCell(page, 2, 0)

  await expect(nameBox).toHaveValue('C1')
  await expect(page.getByTestId('grid-pane-text-overlay')).toHaveCount(0)
  await clickProductCell(page, 1, 0)
  await expect(nameBox).toHaveValue('B1')
  await expect(formulaInput).toHaveValue('abcdef')
  await expect(resolvedValue).toHaveText('abcdef')
})

test('web app commits a cleared formula bar draft when clicking away', async ({ page }) => {
  const staleText = 'formula-clear-click-away'
  await page.goto(`/?document=${encodeURIComponent(createTestDocumentId('playwright-formula-clear-click-away'))}`)
  await waitForWorkbookReady(page)

  const formulaInput = page.getByTestId('formula-input')

  await clickProductCell(page, 3, 6)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!D7')
  await formulaInput.fill(staleText)
  await formulaInput.press('Enter')
  await expect(formulaInput).toHaveValue(staleText)
  await expect.poll(() => nativeTextRunsInclude(page, staleText)).toBe(true)

  await formulaInput.fill('')
  await expect(formulaInput).toHaveValue('')
  await clickProductCell(page, 4, 6)

  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!E7')
  await expect(formulaInput).toHaveValue('')
  await expect.poll(() => nativeTextRunsInclude(page, staleText)).toBe(false)

  await clickProductCell(page, 3, 6)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!D7')
  await expect(formulaInput).toHaveValue('')
})

test('web app commits a cleared formula bar draft with Enter', async ({ page }) => {
  const staleText = 'formula-clear-enter'
  await page.goto(`/?document=${encodeURIComponent(createTestDocumentId('playwright-formula-clear-enter'))}`)
  await waitForWorkbookReady(page)

  const formulaInput = page.getByTestId('formula-input')

  await clickProductCell(page, 3, 6)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!D7')
  await formulaInput.fill(staleText)
  await formulaInput.press('Enter')
  await expect(formulaInput).toHaveValue(staleText)
  await expect.poll(() => nativeTextRunsInclude(page, staleText)).toBe(true)

  await formulaInput.fill('')
  await expect(formulaInput).toHaveValue('')
  await formulaInput.press('Enter')

  await expect(formulaInput).toHaveValue('')
  await expect.poll(() => nativeTextRunsInclude(page, staleText)).toBe(false)

  await clickProductCell(page, 4, 6)
  await clickProductCell(page, 3, 6)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!D7')
  await expect(formulaInput).toHaveValue('')
  await expect.poll(() => nativeTextRunsInclude(page, staleText)).toBe(false)
})

test('web app does not resurrect a keyboard-cleared cell after click-away', async ({ page }) => {
  const staleText = 'delete-clear-click-away'
  await page.goto(`/?document=${encodeURIComponent(createTestDocumentId('playwright-delete-clear-click-away'))}`)
  await waitForWorkbookReady(page)

  const formulaInput = page.getByTestId('formula-input')

  await clickProductCell(page, 3, 6)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!D7')
  await formulaInput.fill(staleText)
  await formulaInput.press('Enter')
  await expect(formulaInput).toHaveValue(staleText)
  await expect.poll(() => nativeTextRunsInclude(page, staleText)).toBe(true)

  await page.keyboard.press('Delete')
  await expect(formulaInput).toHaveValue('')
  await expect.poll(() => nativeTextRunsInclude(page, staleText)).toBe(false)

  await clickProductCell(page, 4, 6)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!E7')
  await expect(formulaInput).toHaveValue('')
  await expect.poll(() => nativeTextRunsInclude(page, staleText)).toBe(false)

  await clickProductCell(page, 3, 6)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!D7')
  await expect(formulaInput).toHaveValue('')
  await expect.poll(() => nativeTextRunsInclude(page, staleText)).toBe(false)
})

test('@browser-ci web app keeps deleted content cleared through viewport churn and reload', async ({ page }) => {
  const staleText = 'delete-clear-viewport-reload'
  const documentId = createTestDocumentId('playwright-delete-clear-viewport-reload')
  await page.goto(`/?document=${encodeURIComponent(documentId)}&sheet=Sheet1&cell=D10`)
  await waitForWorkbookReady(page)

  const formulaInput = page.getByTestId('formula-input')

  await clickProductCell(page, 3, 9)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!D10')
  await formulaInput.fill(staleText)
  await formulaInput.press('Enter')
  await expect(formulaInput).toHaveValue(staleText)
  await expect.poll(() => nativeTextRunsInclude(page, staleText)).toBe(true)

  await page.keyboard.press('Delete')
  await expect(formulaInput).toHaveValue('')
  await expect.poll(() => nativeTextRunsInclude(page, staleText)).toBe(false)

  await page.getByTestId('grid-scroll-viewport').evaluate((viewport) => {
    viewport.scrollTop = 900
    viewport.scrollLeft = 220
    viewport.dispatchEvent(new Event('scroll', { bubbles: true }))
  })
  await expect.poll(() => nativeTextRunsInclude(page, staleText)).toBe(false)

  await page.getByTestId('grid-scroll-viewport').evaluate((viewport) => {
    viewport.scrollTop = 0
    viewport.scrollLeft = 0
    viewport.dispatchEvent(new Event('scroll', { bubbles: true }))
  })
  await expect.poll(() => nativeTextRunsInclude(page, staleText)).toBe(false)

  await page.reload({ waitUntil: 'domcontentloaded' })
  await waitForWorkbookReady(page)
  await clickProductCell(page, 3, 9)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!D10')
  await expect(formulaInput).toHaveValue('')
  await expect.poll(() => nativeTextRunsInclude(page, staleText)).toBe(false)
})

test('web app keeps delayed in-cell typing anchored and exits cleanly on click-away', async ({ page }) => {
  await page.keyboard.up(PRIMARY_MODIFIER)
  await page.goto(`/?document=${encodeURIComponent(createTestDocumentId('playwright-delayed-click-away-edit'))}`)
  await waitForWorkbookReady(page)

  const formulaInput = page.getByTestId('formula-input')
  const cellEditor = page.getByTestId('cell-editor-input')
  const renderer = page.getByTestId('grid-pane-renderer')

  await clickProductCell(page, 2, 11)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!C12')
  await page.keyboard.press('a')
  await expect(cellEditor).toBeVisible()
  await expect(cellEditor).toHaveValue('a')
  expect(await textControlValue(formulaInput)).toBe('a')
  await expect
    .poll(async () => await cellEditor.evaluate((input) => (input instanceof HTMLTextAreaElement ? input.selectionStart : -1)))
    .toBe(1)

  await page.waitForTimeout(300)
  await cellEditor.press('s')
  await expect(cellEditor).toHaveValue('as')
  expect(await textControlValue(formulaInput)).toBe('as')
  await expect
    .poll(async () => await cellEditor.evaluate((input) => (input instanceof HTMLTextAreaElement ? input.selectionStart : -1)))
    .toBe(2)

  await page.waitForTimeout(300)
  await cellEditor.press('d')
  await cellEditor.press('f')
  await expect(cellEditor).toHaveValue('asdf')
  expect(await textControlValue(formulaInput)).toBe('asdf')
  await expect
    .poll(async () => await cellEditor.evaluate((input) => (input instanceof HTMLTextAreaElement ? input.selectionStart : -1)))
    .toBe(4)
  await expect.poll(async () => Number((await renderer.getAttribute('data-v3-header-pane-count')) ?? '0')).toBeGreaterThan(0)
  await expect.poll(async () => Number((await renderer.getAttribute('data-v3-header-text-run-count')) ?? '0')).toBeGreaterThan(10)

  await clickProductCell(page, 3, 11)

  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!D12')
  await expect(cellEditor).toHaveCount(0)
  await expect.poll(async () => Number((await renderer.getAttribute('data-v3-header-pane-count')) ?? '0')).toBeGreaterThan(0)
  await expect.poll(async () => Number((await renderer.getAttribute('data-v3-header-text-run-count')) ?? '0')).toBeGreaterThan(10)

  await clickProductCell(page, 2, 11)
  await expect(formulaInput).toHaveValue('asdf')
})

test('web app drags a selected range by its border with a grab cursor', async ({ page }) => {
  await gotoWorkbookShell(page, `/?document=${encodeURIComponent(createTestDocumentId('range-border-drag'))}`)
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
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!B2')
  await expect(formulaInput).toHaveValue('')
  await expect(resolvedValue).toHaveText('∅')

  await nameBox.fill('C2')
  await nameBox.press('Enter')
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!C2')
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

test('web app moves selected cell content from the content drag lane', async ({ page }) => {
  await gotoWorkbookShell(page, `/?document=${encodeURIComponent(createTestDocumentId('range-content-drag'))}`)
  await waitForWorkbookReady(page)

  const nameBox = page.getByTestId('name-box')
  const formulaInput = page.getByTestId('formula-input')
  const resolvedValue = page.getByTestId('formula-resolved-value')

  await nameBox.fill('B2')
  await nameBox.press('Enter')
  await formulaInput.fill('move-me')
  await formulaInput.press('Enter')
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!B2')

  await dragProductSelectedContentLane(page, 1, 1, 3, 3)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!D4')

  await nameBox.fill('B2')
  await nameBox.press('Enter')
  await expect(formulaInput).toHaveValue('')
  await expect(resolvedValue).toHaveText('∅')

  await nameBox.fill('D4')
  await nameBox.press('Enter')
  await expect(formulaInput).toHaveValue('move-me')
  await expect(resolvedValue).toHaveText('move-me')
})

test('web app applies core formatting shortcuts from the keyboard', async ({ page }) => {
  const documentId = createTestDocumentId('playwright-core-formatting-shortcuts')
  await page.goto(`/?document=${encodeURIComponent(documentId)}`)
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
  await page.keyboard.down(PRIMARY_MODIFIER)
  await page.keyboard.press('Backslash')
  await page.keyboard.up(PRIMARY_MODIFIER)
  await expect(page.getByLabel('Bold')).not.toHaveClass(/bg-\[var\(--wb-accent-soft\)\]/)
  await expect(page.getByLabel('Italic')).not.toHaveClass(/bg-\[var\(--wb-accent-soft\)\]/)
  await expect(page.getByLabel('Underline')).not.toHaveClass(/bg-\[var\(--wb-accent-soft\)\]/)
})

test('web app applies advertised number and border formatting shortcuts from the keyboard', async ({ page }) => {
  const documentId = createTestDocumentId('playwright-advanced-formatting-shortcuts')
  await page.goto(`/?document=${encodeURIComponent(documentId)}`)
  await waitForWorkbookReady(page)

  const grid = page.getByTestId('sheet-grid')
  const formulaInput = page.getByTestId('formula-input')
  const numberFormat = page.getByRole('combobox', { name: 'Number format', exact: true })
  const borders = page.getByRole('button', { name: 'Borders', exact: true })

  await clickProductCell(page, 0, 0)
  await formulaInput.fill('1234')
  await formulaInput.press('Enter')
  await clickProductCell(page, 0, 0)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!A1')
  await expect(page.getByTestId('formula-resolved-value')).toHaveText('1234')

  await grid.press(`${PRIMARY_MODIFIER}+Shift+1`)
  await expect(numberFormat).toHaveAttribute('data-current-value', 'number')

  await grid.press(`${PRIMARY_MODIFIER}+Shift+4`)
  await expect(numberFormat).toHaveAttribute('data-current-value', 'currency')

  await grid.press(`${PRIMARY_MODIFIER}+Shift+5`)
  await expect(numberFormat).toHaveAttribute('data-current-value', 'percent')

  await grid.press(`${PRIMARY_MODIFIER}+Shift+7`)
  await expect(borders).toHaveAttribute('aria-pressed', 'true')
  await expect(borders).toHaveClass(/bg-\[var\(--wb-accent-soft\)\]/)

  await page.keyboard.down(PRIMARY_MODIFIER)
  await page.keyboard.press('Backslash')
  await page.keyboard.up(PRIMARY_MODIFIER)
  await expect(numberFormat).toHaveAttribute('data-current-value', 'general')
  await expect(borders).toHaveAttribute('aria-pressed', 'false')
})

test('web app keeps formatting shortcuts active after toolbar focus without letting delete keys clear cells', async ({ page }) => {
  const documentId = createTestDocumentId('playwright-toolbar-focus-shortcut-scope')
  await page.goto(`/?document=${encodeURIComponent(documentId)}`)
  await waitForWorkbookReady(page)

  const formulaInput = page.getByTestId('formula-input')
  const boldButton = page.getByLabel('Bold')

  await clickProductCell(page, 2, 2)
  await formulaInput.fill('clear-after-toolbar-focus')
  await formulaInput.press('Enter')
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!C3')
  await expect(formulaInput).toHaveValue('clear-after-toolbar-focus')

  await boldButton.click()
  await expect(boldButton).toBeFocused()
  await expect(page.getByLabel('Bold')).toHaveClass(/bg-\[var\(--wb-accent-soft\)\]/)
  await page.keyboard.press('Delete')
  await expect(formulaInput).toHaveValue('clear-after-toolbar-focus')
  await expect(boldButton).toBeFocused()

  await page.keyboard.press('Backspace')
  await expect(formulaInput).toHaveValue('clear-after-toolbar-focus')
  await expect(boldButton).toBeFocused()

  await page.keyboard.press(`${PRIMARY_MODIFIER}+B`)
  await expect(page.getByLabel('Bold')).not.toHaveClass(/bg-\[var\(--wb-accent-soft\)\]/)

  await clickProductCell(page, 2, 2)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!C3')
  await page.keyboard.press('Delete')
  await expect(formulaInput).toHaveValue('')

  await boldButton.click()
  await expect(boldButton).toBeFocused()
  await page.keyboard.press('Enter')
  await expect(formulaInput).toHaveValue('')
  await expect(boldButton).toBeFocused()
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

test('web app collapses the selected range before typing into the cell editor', async ({ page }) => {
  await page.goto('/')
  await waitForWorkbookReady(page)

  const grid = page.getByTestId('sheet-grid')
  const nameBox = page.getByTestId('name-box')
  await clickProductCell(page, 2, 4)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!C5')

  await grid.press('Shift+ArrowRight')
  await grid.press('Shift+ArrowRight')
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!C5:E5')

  await page.getByTestId('sheet-grid-focus-target').focus()
  await page.keyboard.press('a')

  await expect(page.getByTestId('cell-editor-input')).toHaveValue('a')
  await expect(nameBox).toHaveValue('C5')
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!C5')
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
    const documentId = createTestDocumentId(`playwright-clear-selected-range-${key.toLowerCase()}`)
    await page.goto(`/?document=${encodeURIComponent(documentId)}`)
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

for (const key of ['Delete', 'Backspace'] as const) {
  test(`web app clears the selected cell with ${key.toLowerCase()} after name-box navigation`, async ({ page }) => {
    const documentId = createTestDocumentId(`playwright-${key.toLowerCase()}-after-name-box`)
    await page.goto(`/?document=${encodeURIComponent(documentId)}`)
    await waitForWorkbookReady(page)

    const formulaInput = page.getByTestId('formula-input')
    const nameBox = page.getByTestId('name-box')

    await clickProductCell(page, 1, 1)
    await formulaInput.fill(`${key.toLowerCase()}-after-name-box`)
    await formulaInput.press('Enter')

    await nameBox.fill('B2')
    await nameBox.press('Enter')
    await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!B2')

    await page.keyboard.press(key)
    await expect(formulaInput).toHaveValue('')
  })
}

for (const key of ['Delete', 'Backspace'] as const) {
  test(`web app keeps ${key.toLowerCase()} scoped to the formula input while editing`, async ({ page }) => {
    const documentId = createTestDocumentId(`playwright-${key.toLowerCase()}-formula-input-scope`)
    await page.goto(`/?document=${encodeURIComponent(documentId)}`)
    await waitForWorkbookReady(page)

    const formulaInput = page.getByTestId('formula-input')
    const nameBox = page.getByTestId('name-box')

    await nameBox.fill('B2')
    await nameBox.press('Enter')
    await formulaInput.fill(`protected-${key.toLowerCase()}-value`)
    await formulaInput.press('Enter')
    await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!B2')
    await expect(formulaInput).toHaveValue(`protected-${key.toLowerCase()}-value`)

    await formulaInput.focus()
    await formulaInput.selectText()
    await formulaInput.press(key)
    await expect(formulaInput).toHaveValue('')
    await formulaInput.press('Escape')

    await nameBox.fill('B2')
    await nameBox.press('Enter')
    await expect(formulaInput).toHaveValue(`protected-${key.toLowerCase()}-value`)
  })
}

for (const key of ['Delete', 'Backspace'] as const) {
  test(`web app routes ${key.toLowerCase()} to the grid after committing the formula input with enter`, async ({ page }) => {
    const documentId = createTestDocumentId(`playwright-${key.toLowerCase()}-after-formula-enter`)
    await page.goto(`/?document=${encodeURIComponent(documentId)}`)
    await waitForWorkbookReady(page)

    const formulaInput = page.getByTestId('formula-input')

    await clickProductCell(page, 2, 11)
    await formulaInput.fill(`${key.toLowerCase()}-after-formula-enter`)
    await formulaInput.press('Enter')
    await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!C12')
    await expect(formulaInput).toHaveValue(`${key.toLowerCase()}-after-formula-enter`)

    await page.keyboard.press(key)
    await expect(formulaInput).toHaveValue('')

    await formulaInput.fill('second-edit-still-works')
    await formulaInput.press('Enter')
    await expect(formulaInput).toHaveValue('second-edit-still-works')
  })
}

test('@browser-ci web app restores a keyboard clear through undo and redo history controls', async ({ page }) => {
  const documentId = createTestDocumentId('playwright-clear-undo-redo-shortcuts')
  await page.goto(`/?document=${encodeURIComponent(documentId)}`)
  await waitForWorkbookReady(page)

  const formulaInput = page.getByTestId('formula-input')

  await clickProductCell(page, 3, 11)
  await formulaInput.fill('delete-undo-redo')
  await formulaInput.press('Enter')
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!D12')
  await expect(formulaInput).toHaveValue('delete-undo-redo')

  await page.keyboard.press('Delete')
  await expect(formulaInput).toHaveValue('')

  await page.getByRole('button', { name: 'Undo' }).click()
  await expect(formulaInput).toHaveValue('delete-undo-redo')

  await page.getByRole('button', { name: 'Redo' }).click()
  await expect(formulaInput).toHaveValue('')
})

test('@browser-ci web app routes workbook undo and redo keyboard shortcuts from the grid', async ({ page }) => {
  const documentId = createTestDocumentId('playwright-grid-history-shortcuts')
  await page.goto(`/?document=${encodeURIComponent(documentId)}`)
  await waitForWorkbookReady(page)

  const formulaInput = page.getByTestId('formula-input')
  const resolvedValue = page.getByTestId('formula-resolved-value')
  const grid = page.getByTestId('sheet-grid')
  const redoShortcut = PRIMARY_MODIFIER === 'Meta' ? `${PRIMARY_MODIFIER}+Shift+Z` : `${PRIMARY_MODIFIER}+Y`

  await clickProductCell(page, 3, 11)
  await formulaInput.fill('keyboard-history-check')
  await formulaInput.press('Enter')
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!D12')
  await expect(formulaInput).toHaveValue('keyboard-history-check')
  await expect(resolvedValue).toHaveText('keyboard-history-check')

  await grid.press(`${PRIMARY_MODIFIER}+Z`)
  await expect(formulaInput).toHaveValue('')
  await expect(resolvedValue).toHaveText('∅')

  await grid.press(redoShortcut)
  await expect(formulaInput).toHaveValue('keyboard-history-check')
  await expect(resolvedValue).toHaveText('keyboard-history-check')
})

test('web app ignores modified delete keys instead of clearing the grid selection', async ({ page }) => {
  const documentId = createTestDocumentId('playwright-modified-delete-ignored')
  await page.goto(`/?document=${encodeURIComponent(documentId)}`)
  await waitForWorkbookReady(page)

  const grid = page.getByTestId('sheet-grid')
  const formulaInput = page.getByTestId('formula-input')

  await clickProductCell(page, 2, 2)
  await formulaInput.fill('keep-modified-delete')
  await formulaInput.press('Enter')
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!C3')
  await expect(formulaInput).toHaveValue('keep-modified-delete')
  await clickProductCell(page, 2, 2)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!C3')
  await expect(formulaInput).toHaveValue('keep-modified-delete')

  const assertModifiedDeleteIgnored = async (key: string) => {
    await grid.press(key)
    await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!C3')
    await expect(formulaInput).toHaveValue('keep-modified-delete')
  }

  await assertModifiedDeleteIgnored(`${PRIMARY_MODIFIER}+Backspace`)
  await assertModifiedDeleteIgnored(`${PRIMARY_MODIFIER}+Delete`)
  await assertModifiedDeleteIgnored('Alt+Backspace')
  await assertModifiedDeleteIgnored('Shift+Delete')
})

for (const key of ['Delete', 'Backspace'] as const) {
  test(`web app clears the querystring-selected cell with ${key.toLowerCase()} after page load`, async ({ page }) => {
    const documentId = createTestDocumentId(`playwright-${key.toLowerCase()}-querystring-selection`)
    await page.goto(`/?document=${encodeURIComponent(documentId)}`)
    await waitForWorkbookReady(page)

    const formulaInput = page.getByTestId('formula-input')
    const nameBox = page.getByTestId('name-box')

    await nameBox.fill('F39')
    await nameBox.press('Enter')
    await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!F39')
    await formulaInput.fill('stale-persisted-querystring-selection')
    await formulaInput.press('Enter')

    await page.goto(`/?document=${encodeURIComponent(documentId)}&sheet=Sheet1&cell=C12`)
    await waitForWorkbookReady(page)
    await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!C12')
    await expect(page.getByTestId('name-box')).toHaveValue('C12')

    await formulaInput.fill(`${key.toLowerCase()}-after-querystring-load`)
    await formulaInput.press('Enter')

    await page.goto(`/?document=${encodeURIComponent(documentId)}&sheet=Sheet1&cell=C12`)
    await waitForWorkbookReady(page)
    await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!C12')
    await expect(page.getByTestId('name-box')).toHaveValue('C12')

    await page.keyboard.press(key)
    await expect(page.getByTestId('formula-input')).toHaveValue('')
  })
}

test('web app keeps delete keys scoped to the in-cell editor while editing', async ({ page }) => {
  const documentId = createTestDocumentId('playwright-delete-keys-cell-editor-scope')
  await page.goto(`/?document=${encodeURIComponent(documentId)}`)
  await waitForWorkbookReady(page)

  const formulaInput = page.getByTestId('formula-input')

  await clickProductCell(page, 2, 2)
  await page.getByTestId('sheet-grid-focus-target').focus()
  await page.keyboard.press('a')
  const editor = page.getByTestId('cell-editor-input')
  await expect(editor).toBeVisible()
  await page.keyboard.type('bcd')
  await expect(editor).toHaveValue('abcd')

  await editor.press('Backspace')
  await expect(editor).toHaveValue('abc')
  await expect(formulaInput).toHaveValue('abc')

  await editor.press('Home')
  await editor.press('Delete')
  await expect(editor).toHaveValue('bc')
  await expect(formulaInput).toHaveValue('bc')

  await editor.press('Enter')
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!C4')

  await clickProductCell(page, 2, 2)
  await expect(formulaInput).toHaveValue('bc')
})

test('web app clears the clicked cell after a prior name-box selection changes pending app selection', async ({ page }) => {
  await page.goto('/')
  await waitForWorkbookReady(page)

  const formulaInput = page.getByTestId('formula-input')
  const nameBox = page.getByTestId('name-box')

  await clickProductCell(page, 1, 1)
  await formulaInput.fill('keep-b2')
  await formulaInput.press('Enter')

  await clickProductCell(page, 2, 2)
  await formulaInput.fill('delete-c3')
  await formulaInput.press('Enter')

  await nameBox.fill('B2')
  await nameBox.press('Enter')
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!B2')

  await clickProductCell(page, 2, 2)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!C3')
  await page.keyboard.press('Delete')

  await clickProductCell(page, 2, 2)
  await expect(formulaInput).toHaveValue('')

  await clickProductCell(page, 1, 1)
  await expect(formulaInput).toHaveValue('keep-b2')
})

test('web app ignores right gutter clicks', async ({ page }) => {
  await page.goto('/')
  await waitForWorkbookReady(page)

  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!A1')
  await clickGridRightEdge(page, 3)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!A1')
})

fuzzBrowserTest(
  '@fuzz-browser web app preserves valid selection geometry and focus under generated selection actions',
  async ({ page }) => {
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
  },
)
