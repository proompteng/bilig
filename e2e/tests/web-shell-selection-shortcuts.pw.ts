import { expect, test } from '@playwright/test'
import * as fc from 'fast-check'
import { runProperty } from '../../packages/test-fuzz/src/index.ts'
import {
  PRIMARY_MODIFIER,
  clickGridRightEdge,
  clickProductCell,
  createTestDocumentId,
  dragProductBodySelection,
  gotoWorkbookShell,
  waitForWorkbookReady,
} from './web-shell-helpers.js'
import { runSelectionFuzzActions, type BrowserSelectionAction } from './web-shell-main-helpers.js'

test('web app routes row, column, and full-sheet selection shortcuts from toolbar focus', async ({ page }) => {
  const documentId = createTestDocumentId('playwright-toolbar-selection-shortcut-scope')
  await page.goto(`/?document=${encodeURIComponent(documentId)}`)
  await waitForWorkbookReady(page)

  const boldButton = page.getByLabel('Bold')
  await clickProductCell(page, 2, 2)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!C3')

  await boldButton.focus()
  await expect(boldButton).toBeFocused()
  await boldButton.evaluate((button) => {
    ;(window as Window & { __biligToolbarShortcutClickCount?: number }).__biligToolbarShortcutClickCount = 0
    button.addEventListener('click', () => {
      const hostWindow = window as Window & { __biligToolbarShortcutClickCount?: number }
      hostWindow.__biligToolbarShortcutClickCount = (hostWindow.__biligToolbarShortcutClickCount ?? 0) + 1
    })
  })

  await page.keyboard.press('Shift+Space')
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!3:3')
  await expect(boldButton).toBeFocused()

  await page.keyboard.press(`${PRIMARY_MODIFIER}+Space`)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!C:C')
  await expect(boldButton).toBeFocused()

  await page.keyboard.press(`${PRIMARY_MODIFIER}+Shift+Space`)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!All')
  await expect(boldButton).toBeFocused()
  await expect
    .poll(() =>
      page.evaluate(() => (window as Window & { __biligToolbarShortcutClickCount?: number }).__biligToolbarShortcutClickCount ?? 0),
    )
    .toBe(0)
})

test('web app supports row, column, and full-sheet selection shortcuts', async ({ page }) => {
  await page.goto('/workbook')
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

test('web app uses data-aware current-region and boundary navigation shortcuts', async ({ page }) => {
  const documentId = createTestDocumentId('playwright-data-aware-shortcuts')
  await page.goto(`/?document=${encodeURIComponent(documentId)}&persist=0&sheet=Sheet1&cell=B2`)
  await waitForWorkbookReady(page)

  const grid = page.getByTestId('sheet-grid')
  const formulaInput = page.getByTestId('formula-input')
  const cells = [
    [1, 1, 'B2'],
    [2, 1, 'C2'],
    [3, 1, 'D2'],
    [1, 2, 'B3'],
    [3, 2, 'D3'],
    [1, 3, 'B4'],
    [2, 3, 'C4'],
    [3, 3, 'D4'],
  ] as const

  await cells.reduce<Promise<void>>(async (previous, [col, row, value]) => {
    await previous
    await clickProductCell(page, col, row)
    await formulaInput.fill(value)
    await formulaInput.press('Enter')
    await expect(formulaInput).toHaveValue(value)
  }, Promise.resolve())

  await clickProductCell(page, 2, 2)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!C3')
  await grid.press(`${PRIMARY_MODIFIER}+Shift+Digit8`)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!B2:D4')

  await clickProductCell(page, 2, 2)
  await grid.press(`${PRIMARY_MODIFIER}+A`)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!B2:D4')
  await grid.press(`${PRIMARY_MODIFIER}+A`)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!All')

  await clickProductCell(page, 1, 1)
  await grid.press(`${PRIMARY_MODIFIER}+ArrowRight`)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!D2')

  await clickProductCell(page, 1, 1)
  await grid.press(`${PRIMARY_MODIFIER}+Shift+ArrowDown`)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!B2:B4')
})

test('web app fills the selected range from the active cell with the fill range shortcut', async ({ page }) => {
  const documentId = createTestDocumentId('playwright-fill-selected-range-shortcut')
  await page.goto(`/?document=${encodeURIComponent(documentId)}&persist=0&sheet=Sheet1&cell=B2`)
  await waitForWorkbookReady(page)

  const grid = page.getByTestId('sheet-grid')
  const formulaInput = page.getByTestId('formula-input')

  await clickProductCell(page, 1, 1)
  await formulaInput.fill('fill-selected-range')
  await formulaInput.press('Enter')
  await expect(formulaInput).toHaveValue('fill-selected-range')

  await dragProductBodySelection(page, 1, 1, 2, 2)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!B2:C3')

  await grid.press(`${PRIMARY_MODIFIER}+Enter`)

  await clickProductCell(page, 1, 1)
  await expect(formulaInput).toHaveValue('fill-selected-range')
  await clickProductCell(page, 2, 1)
  await expect(formulaInput).toHaveValue('fill-selected-range')
  await clickProductCell(page, 1, 2)
  await expect(formulaInput).toHaveValue('fill-selected-range')
  await clickProductCell(page, 2, 2)
  await expect(formulaInput).toHaveValue('fill-selected-range')
})

test('web app fills down and right with spreadsheet keyboard shortcuts', async ({ page }) => {
  const documentId = createTestDocumentId('playwright-fill-down-right-shortcuts')
  await page.goto(`/?document=${encodeURIComponent(documentId)}&persist=0&sheet=Sheet1&cell=B2`)
  await waitForWorkbookReady(page)

  const grid = page.getByTestId('sheet-grid')
  const formulaInput = page.getByTestId('formula-input')

  await clickProductCell(page, 1, 1)
  await formulaInput.fill('fill-down-source')
  await formulaInput.press('Enter')
  await dragProductBodySelection(page, 1, 1, 1, 3)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!B2:B4')

  await grid.press(`${PRIMARY_MODIFIER}+D`)

  await clickProductCell(page, 1, 1)
  await expect(formulaInput).toHaveValue('fill-down-source')
  await clickProductCell(page, 1, 2)
  await expect(formulaInput).toHaveValue('fill-down-source')
  await clickProductCell(page, 1, 3)
  await expect(formulaInput).toHaveValue('fill-down-source')

  await clickProductCell(page, 2, 1)
  await formulaInput.fill('fill-right-source')
  await formulaInput.press('Enter')
  await dragProductBodySelection(page, 2, 1, 4, 1)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!C2:E2')

  await grid.press(`${PRIMARY_MODIFIER}+R`)

  await clickProductCell(page, 2, 1)
  await expect(formulaInput).toHaveValue('fill-right-source')
  await clickProductCell(page, 3, 1)
  await expect(formulaInput).toHaveValue('fill-right-source')
  await clickProductCell(page, 4, 1)
  await expect(formulaInput).toHaveValue('fill-right-source')
})

test('web app expands the active range with repeated shift arrows', async ({ page }) => {
  await page.goto('/workbook')
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
  await page.goto('/workbook')
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
  await page.goto('/workbook')
  await waitForWorkbookReady(page)

  await clickProductCell(page, 1, 1)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!B2')

  await clickProductCell(page, 4, 5, { shift: true })
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!B2:E6')
})

test('web app cycles the active cell inside a selected range with Enter and Tab', async ({ page, context }) => {
  await context.grantPermissions(['clipboard-read', 'clipboard-write'])
  const documentId = createTestDocumentId('playwright-range-active-cell-cycle')
  await page.goto(`/?document=${encodeURIComponent(documentId)}&persist=0&sheet=Sheet1&cell=B2`)
  await waitForWorkbookReady(page)

  const grid = page.getByTestId('sheet-grid')
  const nameBox = page.getByTestId('name-box')
  const formulaInput = page.getByTestId('formula-input')

  await clickProductCell(page, 1, 1)
  await page.evaluate(() => navigator.clipboard.writeText('range-b2\trange-c2\nrange-b3\trange-c3'))
  await grid.press(`${PRIMARY_MODIFIER}+V`)
  await expect(formulaInput).toHaveValue('range-b2')

  await dragProductBodySelection(page, 1, 1, 2, 2)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!B2:C3')
  await expect(nameBox).toHaveValue('B2:C3')
  await expect(formulaInput).toHaveValue('range-b2')

  await grid.press('Tab')
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!B2:C3')
  await expect(nameBox).toHaveValue('B2:C3')
  await expect(formulaInput).toHaveValue('range-c2')

  await grid.press('Tab')
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!B2:C3')
  await expect(nameBox).toHaveValue('B2:C3')
  await expect(formulaInput).toHaveValue('range-b3')

  await grid.press('Shift+Tab')
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!B2:C3')
  await expect(nameBox).toHaveValue('B2:C3')
  await expect(formulaInput).toHaveValue('range-c2')

  await grid.press('Enter')
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!B2:C3')
  await expect(nameBox).toHaveValue('B2:C3')
  await expect(formulaInput).toHaveValue('range-c3')

  await grid.press('Enter')
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!B2:C3')
  await expect(nameBox).toHaveValue('B2:C3')
  await expect(formulaInput).toHaveValue('range-b2')
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

  await assertModifiedDeleteIgnored(`${PRIMARY_MODIFIER}+Delete`)
  await assertModifiedDeleteIgnored('Alt+Backspace')
  await assertModifiedDeleteIgnored('Shift+Delete')
})

test('web app scrolls to the active cell with primary Backspace without clearing it', async ({ page }) => {
  const documentId = createTestDocumentId('playwright-primary-backspace-scroll-active')
  await page.goto(`/?document=${encodeURIComponent(documentId)}`)
  await waitForWorkbookReady(page)

  const grid = page.getByTestId('sheet-grid')
  const formulaInput = page.getByTestId('formula-input')
  const scrollViewport = page.getByTestId('grid-scroll-viewport')

  await clickProductCell(page, 2, 2)
  await formulaInput.fill('scroll-active-keeps-content')
  await formulaInput.press('Enter')
  await clickProductCell(page, 2, 2)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!C3')
  await expect(formulaInput).toHaveValue('scroll-active-keeps-content')

  await scrollViewport.evaluate((element) => {
    element.scrollTop = 2_000
    element.scrollLeft = 600
    element.dispatchEvent(new Event('scroll', { bubbles: true }))
  })
  await expect.poll(async () => await scrollViewport.evaluate((element) => element.scrollTop)).toBeGreaterThan(1_000)

  await grid.press(`${PRIMARY_MODIFIER}+Backspace`)

  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!C3')
  await expect(formulaInput).toHaveValue('scroll-active-keeps-content')
  await expect.poll(async () => await scrollViewport.evaluate((element) => element.scrollTop)).toBeLessThan(200)
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
  await page.goto('/workbook')
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
  await page.goto('/workbook')
  await waitForWorkbookReady(page)

  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!A1')
  await clickGridRightEdge(page, 3)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!A1')
})

test('@fuzz-browser web app preserves valid selection geometry and focus under generated selection actions', async ({ page }) => {
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
