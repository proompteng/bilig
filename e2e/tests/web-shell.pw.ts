import { expect, test } from '@playwright/test'
import {
  PRIMARY_MODIFIER,
  PRODUCT_HEADER_HEIGHT,
  PRODUCT_ROW_HEIGHT,
  clickProductCell,
  createTestDocumentId,
  dragProductBodySelection,
  dragProductColumnResize,
  getProductColumnLeft,
  getProductColumnWidth,
  getProductFillHandleDragPoints,
  gotoWorkbookShell,
  installTypeGpuCellReadbackHarness,
  waitForProductColumnWidthChange,
  waitForWorkbookReady,
} from './web-shell-helpers.js'
import {
  clickProductSelectedCellTopBorder,
  dragProductFillHandle,
  dragProductSelectionBorder,
  expectCellRenderedText,
  textControlValue,
} from './web-shell-main-helpers.js'

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
  await expect(formulaInput).toHaveValue('=A1="HELLO"')
  await expect(resolvedValue).toHaveText('TRUE')
})

test('web app supports type-to-replace and Enter or Tab commit movement', async ({ page }) => {
  await page.goto('/workbook')
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

test('@browser-ci web app keeps in-cell caret movement stable during rapid typing', async ({ page }) => {
  const documentId = createTestDocumentId('playwright-editor-caret-rapid-typing')
  await page.goto(`/?document=${encodeURIComponent(documentId)}&persist=0`)
  await waitForWorkbookReady(page)

  await clickProductCell(page, 0, 0)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!A1')
  await page.getByTestId('sheet-grid-focus-target').focus()

  const editor = page.getByTestId('cell-editor-input')
  await page.keyboard.press('a')
  await expect(editor).toBeVisible()
  const readEditorSelection = () =>
    editor.evaluate((element) => {
      if (!(element instanceof HTMLTextAreaElement)) {
        throw new Error('Expected in-cell editor textarea')
      }
      const start = element.selectionStart
      const end = element.selectionEnd
      return {
        direction: start === end ? 'none' : element.selectionDirection,
        end,
        start,
      }
    })

  await page.keyboard.press('b')
  await page.keyboard.press('c')
  await page.keyboard.press('ArrowLeft')
  await page.keyboard.press('x')

  await expect(editor).toHaveValue('abxc')
  await expect.poll(readEditorSelection).toEqual({
    direction: 'none',
    end: 3,
    start: 3,
  })

  await page.keyboard.press('Shift+ArrowLeft')
  await expect.poll(readEditorSelection).toEqual({
    direction: 'backward',
    end: 3,
    start: 2,
  })

  await page.keyboard.press('Home')
  await page.keyboard.press('z')
  await expect(editor).toHaveValue('zabxc')
  await expect.poll(readEditorSelection).toEqual({
    direction: 'none',
    end: 1,
    start: 1,
  })

  await page.keyboard.press('End')
  await page.keyboard.press('!')
  await expect(editor).toHaveValue('zabxc!')
  await expect.poll(readEditorSelection).toEqual({
    direction: 'none',
    end: 6,
    start: 6,
  })
})

test('@browser-ci web app preserves a rapid type-to-edit burst during editor focus handoff', async ({ page }) => {
  const documentId = createTestDocumentId('playwright-editor-focus-handoff-burst')
  await page.goto(`/?document=${encodeURIComponent(documentId)}&persist=0`)
  await waitForWorkbookReady(page)

  await clickProductCell(page, 1, 1)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!B2')
  await page.getByTestId('sheet-grid-focus-target').focus()
  await page.keyboard.type('abcdef')

  const editor = page.getByTestId('cell-editor-input')
  await expect(editor).toBeVisible()
  await expect(editor).toHaveValue('abcdef')
  await expect
    .poll(() =>
      editor.evaluate((element) => {
        if (!(element instanceof HTMLTextAreaElement)) {
          throw new Error('Expected in-cell editor textarea')
        }
        return {
          end: element.selectionEnd,
          start: element.selectionStart,
        }
      }),
    )
    .toEqual({ end: 6, start: 6 })
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

test('@browser-sync @browser-ci web app recovers after runtime config failures outlive the fast retry window', async ({ page }) => {
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

  await page.goto(`/?document=${encodeURIComponent(createTestDocumentId('playwright-runtime-config-recovery'))}`)

  await expect(page.getByTestId('formula-bar')).toBeVisible({ timeout: 20_000 })
  await expect(page.getByTestId('sheet-grid')).toBeVisible({ timeout: 15_000 })
  await expect(page.getByTestId('worker-error')).toHaveCount(0)
  expect(runtimeConfigAttempts).toBeGreaterThan(5)
})

test('@browser-ci web app keeps an editor clear after click-away selection', async ({ page }) => {
  const documentId = createTestDocumentId('playwright-editor-clear-click-away')
  await installTypeGpuCellReadbackHarness(page)
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
  await expectCellRenderedText(page, 0, 0, 'ghost-value', 'hidden')

  await clickProductCell(page, 0, 0)
  await expect(nameBox).toHaveValue('A1', { timeout: 15_000 })
  await expect(formulaInput).toHaveValue('')
  await expectCellRenderedText(page, 0, 0, 'ghost-value', 'hidden')
})

test('web app keeps formula bar focus when clicking it from an active cell editor', async ({ page }) => {
  const documentId = createTestDocumentId('playwright-editor-formula-focus-handoff')
  await page.goto(`/?document=${encodeURIComponent(documentId)}&persist=0&sheet=Sheet1&cell=A1`)
  await waitForWorkbookReady(page)

  const gridLocator = page.getByTestId('sheet-grid')
  const formulaInput = page.getByTestId('formula-input')
  const cellEditor = page.getByTestId('cell-editor-input')

  await clickProductCell(page, 1, 1)
  await page.keyboard.type('formula-focus-draft')
  await expect(cellEditor).toBeVisible()
  await expect(cellEditor).toBeFocused()

  await formulaInput.click()

  await expect(cellEditor).toBeHidden()
  await expect(formulaInput).toBeFocused()
  await expect(formulaInput).toHaveValue('formula-focus-draft')
  await expect(gridLocator).not.toBeFocused()
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

test('web app preserves editor multiline shortcuts across commit, formula bar, and reopen', async ({ page }) => {
  const documentId = createTestDocumentId('playwright-multiline-edit-shortcuts')
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
  await cellEditor.press(`${PRIMARY_MODIFIER}+Enter`)
  await page.keyboard.type('gamma')
  await expect(cellEditor).toHaveValue('alpha\nbeta\ngamma')

  await cellEditor.press('Enter')
  await expect(cellEditor).toBeHidden()
  await clickProductCell(page, 0, 0)
  await expect(formulaInput).toHaveValue('alpha\nbeta\ngamma')

  await page.getByTestId('sheet-grid').press('F2')
  await expect(cellEditor).toBeVisible()
  await expect(cellEditor).toHaveValue('alpha\nbeta\ngamma')
})

test('web app preserves multi-digit numeric type-to-replace input', async ({ page }) => {
  await page.goto('/workbook')
  await waitForWorkbookReady(page)

  const nameBox = page.getByTestId('name-box')
  const formulaInput = page.getByTestId('formula-input')
  const cellEditor = page.getByTestId('cell-editor-input')
  const grid = page.getByTestId('sheet-grid')

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
  await page.goto('/workbook')
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
  await page.goto('/workbook')
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
  await page.goto('/workbook')
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

test('@browser-ci web app offers formula autocomplete and inserts a function with Tab', async ({ page }) => {
  const documentId = createTestDocumentId('playwright-formula-autocomplete')
  await page.goto(`/?document=${encodeURIComponent(documentId)}`)
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
  await page.goto('/workbook')
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
  await page.goto('/workbook')
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
  await page.goto('/workbook')
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
  await page.goto('/workbook')
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

test('@browser-ci web app keeps fill-handle pointer-up exclusive before click-away selection', async ({ page }) => {
  await gotoWorkbookShell(page, `/?document=${encodeURIComponent(createTestDocumentId('fill-handle-exclusive-pointer-up'))}&persist=0`)
  await waitForWorkbookReady(page)

  const nameBox = page.getByTestId('name-box')
  const formulaInput = page.getByTestId('formula-input')

  await nameBox.fill('F6')
  await nameBox.press('Enter')
  await formulaInput.fill('7')
  await formulaInput.press('Enter')
  await nameBox.fill('F6')
  await nameBox.press('Enter')
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!F6')

  await dragProductFillHandle(page, 5, 5, 5, 7)
  await clickProductCell(page, 0, 0)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!A1')

  await nameBox.fill('F6')
  await nameBox.press('Enter')
  await expect(formulaInput).toHaveValue('7')
  await nameBox.fill('F7')
  await nameBox.press('Enter')
  await expect(formulaInput).toHaveValue('7')
  await nameBox.fill('F8')
  await nameBox.press('Enter')
  await expect(formulaInput).toHaveValue('7')
})

test('@browser-ci web app autoscrolls while dragging the fill handle past visible rows', async ({ page }) => {
  await page.setViewportSize({ width: 1180, height: 360 })
  await gotoWorkbookShell(page, `/?document=${encodeURIComponent(createTestDocumentId('fill-handle-edge-autoscroll'))}&persist=0`)
  await waitForWorkbookReady(page)

  const nameBox = page.getByTestId('name-box')
  const formulaInput = page.getByTestId('formula-input')
  const scrollViewport = page.getByTestId('grid-scroll-viewport')
  const grid = await page.getByTestId('sheet-grid').boundingBox()
  if (!grid) {
    throw new Error('sheet grid is not visible')
  }

  await nameBox.fill('B2')
  await nameBox.press('Enter')
  await formulaInput.fill('9')
  await formulaInput.press('Enter')

  const { sourceX, sourceY } = await getProductFillHandleDragPoints(page, 1, 1, 1, 1)
  await page.mouse.move(sourceX, sourceY)
  await page.mouse.down()
  await page.mouse.move(sourceX, grid.y + grid.height - 18, { steps: 8 })
  await expect
    .poll(() => scrollViewport.evaluate((node) => node.scrollTop), {
      message: 'fill-handle drag should autoscroll the grid at the lower edge',
    })
    .toBeGreaterThan(PRODUCT_ROW_HEIGHT * 12)
  await page.mouse.up()

  await expect(page.getByTestId('status-selection')).toContainText('!B2:B')
  await nameBox.fill('B20')
  await nameBox.press('Enter')
  await expect(formulaInput).toHaveValue('9')
})

test('@browser-sync web app enables undo and redo for a normal edit', async ({ page }) => {
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

test('@browser-sync web app preserves redo across a longer undo history', async ({ page }) => {
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

test('@browser-sync web app clears redo after a fresh edit branches history', async ({ page }) => {
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
  await expect(fillPreview).toBeVisible()

  await page.mouse.up()

  await expect(selectionStatus).toContainText('!F6:H6')

  await nameBox.fill('H6')
  await nameBox.press('Enter')
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!H6')
  await expect(formulaInput).toHaveValue('7')
  await expect(resolvedValue).toHaveText('7')
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
  const shortcutSearch = page.getByTestId('workbook-shortcut-search')
  await expect(shortcutSearch).toBeFocused()
  await shortcutSearch.fill('current region')
  await expect(page.getByTestId('workbook-shortcut-entry')).toContainText('Select current region')
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
  await installTypeGpuCellReadbackHarness(page)
  await page.goto(`/?document=${encodeURIComponent(createTestDocumentId('playwright-formula-clear-click-away'))}`)
  await waitForWorkbookReady(page)

  const formulaInput = page.getByTestId('formula-input')

  await clickProductCell(page, 3, 6)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!D7')
  await formulaInput.fill(staleText)
  await formulaInput.press('Enter')
  await expect(formulaInput).toHaveValue(staleText)
  await expectCellRenderedText(page, 3, 6, staleText, 'visible')

  await formulaInput.fill('')
  await expect(formulaInput).toHaveValue('')
  await clickProductCell(page, 4, 6)

  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!E7')
  await expect(formulaInput).toHaveValue('')
  await expectCellRenderedText(page, 3, 6, staleText, 'hidden')

  await clickProductCell(page, 3, 6)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!D7')
  await expect(formulaInput).toHaveValue('')
})

test('@browser-ci web app commits a first formula bar draft when focus leaves immediately', async ({ page }) => {
  const draftText = 'first-formula-blur-commit'
  await installTypeGpuCellReadbackHarness(page)
  await page.goto(`/?document=${encodeURIComponent(createTestDocumentId('playwright-formula-first-blur'))}`)
  await waitForWorkbookReady(page)

  const formulaInput = page.getByTestId('formula-input')

  await clickProductCell(page, 1, 1)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!B2')
  await formulaInput.click()
  await formulaInput.fill(draftText)
  await page.getByTestId('workbook-agent-input').click()

  await expect(formulaInput).toHaveValue(draftText)
  await expectCellRenderedText(page, 1, 1, draftText, 'visible')

  await clickProductCell(page, 2, 1)
  await clickProductCell(page, 1, 1)
  await expect(formulaInput).toHaveValue(draftText)
})

test('web app commits a cleared formula bar draft with Enter', async ({ page }) => {
  const staleText = 'formula-clear-enter'
  await installTypeGpuCellReadbackHarness(page)
  await page.goto(`/?document=${encodeURIComponent(createTestDocumentId('playwright-formula-clear-enter'))}`)
  await waitForWorkbookReady(page)

  const formulaInput = page.getByTestId('formula-input')

  await clickProductCell(page, 3, 6)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!D7')
  await formulaInput.fill(staleText)
  await formulaInput.press('Enter')
  await expect(formulaInput).toHaveValue(staleText)
  await expectCellRenderedText(page, 3, 6, staleText, 'visible')

  await formulaInput.fill('')
  await expect(formulaInput).toHaveValue('')
  await formulaInput.press('Enter')

  await expect(formulaInput).toHaveValue('')
  await expectCellRenderedText(page, 3, 6, staleText, 'hidden')

  await clickProductCell(page, 4, 6)
  await clickProductCell(page, 3, 6)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!D7')
  await expect(formulaInput).toHaveValue('')
  await expectCellRenderedText(page, 3, 6, staleText, 'hidden')
})

test('web app does not resurrect a keyboard-cleared cell after click-away', async ({ page }) => {
  const staleText = 'delete-clear-click-away'
  await installTypeGpuCellReadbackHarness(page)
  await page.goto(`/?document=${encodeURIComponent(createTestDocumentId('playwright-delete-clear-click-away'))}`)
  await waitForWorkbookReady(page)

  const formulaInput = page.getByTestId('formula-input')

  await clickProductCell(page, 3, 6)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!D7')
  await formulaInput.fill(staleText)
  await formulaInput.press('Enter')
  await expect(formulaInput).toHaveValue(staleText)
  await expectCellRenderedText(page, 3, 6, staleText, 'visible')

  await page.keyboard.press('Delete')
  await expect(formulaInput).toHaveValue('')
  await expectCellRenderedText(page, 3, 6, staleText, 'hidden')

  await clickProductCell(page, 4, 6)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!E7')
  await expect(formulaInput).toHaveValue('')
  await expectCellRenderedText(page, 3, 6, staleText, 'hidden')

  await clickProductCell(page, 3, 6)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!D7')
  await expect(formulaInput).toHaveValue('')
  await expectCellRenderedText(page, 3, 6, staleText, 'hidden')
})

test('web app clears selected in-cell editor text with primary-A and keeps continued typing anchored', async ({ page }) => {
  const staleText = 'select-all-delete-me'
  const replacementText = 'replacement'
  await installTypeGpuCellReadbackHarness(page)
  await page.keyboard.up(PRIMARY_MODIFIER)
  await page.goto(`/?document=${encodeURIComponent(createTestDocumentId('playwright-editor-primary-a-delete'))}`)
  await waitForWorkbookReady(page)

  const formulaInput = page.getByTestId('formula-input')
  const cellEditor = page.getByTestId('cell-editor-input')

  await clickProductCell(page, 3, 6)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!D7')
  await formulaInput.fill(staleText)
  await formulaInput.press('Enter')
  await expectCellRenderedText(page, 3, 6, staleText, 'visible')

  await clickProductCell(page, 3, 6)
  await page.keyboard.press('F2')
  await expect(cellEditor).toBeVisible()
  await expect(cellEditor).toHaveValue(staleText)
  await cellEditor.press(`${PRIMARY_MODIFIER}+A`)
  await expect
    .poll(async () => await cellEditor.evaluate((input) => (input instanceof HTMLTextAreaElement ? input.selectionStart : -1)))
    .toBe(0)
  await expect
    .poll(async () => await cellEditor.evaluate((input) => (input instanceof HTMLTextAreaElement ? input.selectionEnd : -1)))
    .toBe(staleText.length)

  await cellEditor.press('Backspace')
  await expect(cellEditor).toHaveValue('')
  await expect(formulaInput).toHaveValue('')
  await expectCellRenderedText(page, 3, 6, staleText, 'hidden')

  await page.keyboard.type(replacementText)
  await expect(cellEditor).toHaveValue(replacementText)
  await expect
    .poll(async () => await cellEditor.evaluate((input) => (input instanceof HTMLTextAreaElement ? input.selectionStart : -1)))
    .toBe(replacementText.length)

  await clickProductCell(page, 4, 6)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!E7')
  await expectCellRenderedText(page, 3, 6, replacementText, 'visible')
  await clickProductCell(page, 3, 6)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!D7')
  await expect(formulaInput).toHaveValue(replacementText)
})

test('@browser-ci web app keeps deleted content cleared through viewport churn and reload', async ({ page }) => {
  const staleText = 'delete-clear-viewport-reload'
  const documentId = createTestDocumentId('playwright-delete-clear-viewport-reload')
  await installTypeGpuCellReadbackHarness(page)
  await page.goto(`/?document=${encodeURIComponent(documentId)}&sheet=Sheet1&cell=D10`)
  await waitForWorkbookReady(page)

  const formulaInput = page.getByTestId('formula-input')

  await clickProductCell(page, 3, 9)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!D10')
  await formulaInput.fill(staleText)
  await formulaInput.press('Enter')
  await expect(formulaInput).toHaveValue(staleText)
  await expectCellRenderedText(page, 3, 9, staleText, 'visible')

  await page.keyboard.press('Delete')
  await expect(formulaInput).toHaveValue('')
  await expectCellRenderedText(page, 3, 9, staleText, 'hidden')

  await page.getByTestId('grid-scroll-viewport').evaluate((viewport) => {
    viewport.scrollTop = 900
    viewport.scrollLeft = 220
    viewport.dispatchEvent(new Event('scroll', { bubbles: true }))
  })
  await expectCellRenderedText(page, 3, 9, staleText, 'hidden')

  await page.getByTestId('grid-scroll-viewport').evaluate((viewport) => {
    viewport.scrollTop = 0
    viewport.scrollLeft = 0
    viewport.dispatchEvent(new Event('scroll', { bubbles: true }))
  })
  await expectCellRenderedText(page, 3, 9, staleText, 'hidden')

  await page.reload({ waitUntil: 'domcontentloaded' })
  await waitForWorkbookReady(page)
  await clickProductCell(page, 3, 9)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!D10')
  await expect(formulaInput).toHaveValue('')
  await expectCellRenderedText(page, 3, 9, staleText, 'hidden')
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

test('@browser-ci web app commits typed seeds before a same-frame click-away can retarget them', async ({ page }) => {
  await installTypeGpuCellReadbackHarness(page)
  await page.keyboard.up(PRIMARY_MODIFIER)
  await page.goto(`/?document=${encodeURIComponent(createTestDocumentId('playwright-pending-type-click-away'))}`)
  await waitForWorkbookReady(page)

  const formulaInput = page.getByTestId('formula-input')
  const renderer = page.getByTestId('grid-pane-renderer')

  await clickProductCell(page, 2, 11)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!C12')

  await page.evaluate(() => {
    const originalRequestAnimationFrame = window.requestAnimationFrame.bind(window)
    const originalCancelAnimationFrame = window.cancelAnimationFrame.bind(window)
    const callbacks = new Map<number, FrameRequestCallback>()
    let nextFrame = 1
    ;(
      window as typeof window & {
        __biligHeldAnimationFrames?: {
          flush(): void
          restore(): void
        }
      }
    ).__biligHeldAnimationFrames = {
      flush() {
        const pending = [...callbacks.values()]
        callbacks.clear()
        pending.forEach((callback) => callback(performance.now()))
      },
      restore() {
        window.requestAnimationFrame = originalRequestAnimationFrame
        window.cancelAnimationFrame = originalCancelAnimationFrame
      },
    }
    window.requestAnimationFrame = ((callback: FrameRequestCallback) => {
      const frame = nextFrame
      nextFrame += 1
      callbacks.set(frame, callback)
      return frame
    }) as typeof window.requestAnimationFrame
    window.cancelAnimationFrame = ((frame: number) => {
      callbacks.delete(frame)
    }) as typeof window.cancelAnimationFrame
  })

  await page.keyboard.press('g')
  await page.keyboard.press('o')
  await clickProductCell(page, 3, 11)
  await page.evaluate(() => {
    const held = (
      window as typeof window & {
        __biligHeldAnimationFrames?: {
          flush(): void
          restore(): void
        }
      }
    ).__biligHeldAnimationFrames
    held?.restore()
    held?.flush()
    delete (window as typeof window & { __biligHeldAnimationFrames?: unknown }).__biligHeldAnimationFrames
  })

  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!D12')
  await expect(page.getByTestId('cell-editor-input')).toHaveCount(0)
  await expect.poll(async () => Number((await renderer.getAttribute('data-v3-header-text-run-count')) ?? '0')).toBeGreaterThan(10)

  await clickProductCell(page, 2, 11)
  await expect(formulaInput).toHaveValue('go')
  await expectCellRenderedText(page, 2, 11, 'go', 'visible')
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

test('web app moves selected cell content only from the selection border', async ({ page }) => {
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

  await dragProductSelectionBorder(page, 1, 1, 1, 1, 3, 3)
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

test('web app returns grid focus after toolbar formatting commands and keeps shortcuts active', async ({ page }) => {
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
  await expect(boldButton).not.toBeFocused()
  await expect(page.getByLabel('Bold')).toHaveClass(/bg-\[var\(--wb-accent-soft\)\]/)

  await page.keyboard.press(`${PRIMARY_MODIFIER}+B`)
  await expect(page.getByLabel('Bold')).not.toHaveClass(/bg-\[var\(--wb-accent-soft\)\]/)

  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!C3')
  await page.keyboard.press('Delete')
  await expect(formulaInput).toHaveValue('')
})
