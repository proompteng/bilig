import { expect, test } from '@playwright/test'
import {
  PRIMARY_MODIFIER,
  clickProductCell,
  createTestDocumentId,
  dragProductBodySelection,
  waitForWorkbookReady,
} from './web-shell-helpers.js'

test.describe('@clipboard-global web app clipboard flows', () => {
  test.describe.configure({ mode: 'serial' })

  test('web app supports rectangular clipboard copy and external paste', async ({ page, context }) => {
    await context.grantPermissions(['clipboard-read', 'clipboard-write'])
    await page.goto('/workbook')
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
    await page.goto('/workbook')
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

  test('web app pastes copied formulas as resolved values with paste-values-only', async ({ page, context }) => {
    await context.grantPermissions(['clipboard-read', 'clipboard-write'])
    const documentId = createTestDocumentId('playwright-paste-values-only')
    await page.goto(`/?document=${encodeURIComponent(documentId)}&persist=0`)
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

    const expectCellValue = async (address: string, value: string, resolved = value) => {
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
    await expect.poll(() => page.evaluate(() => navigator.clipboard.readText())).toBe('3\t=B2*2\n4\t=B3*2')

    await clickProductCell(page, 3, 1)
    await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!D2')
    await grid.press(`${PRIMARY_MODIFIER}+Shift+V`)

    await expectCellValue('D2', '3')
    await expectCellValue('E2', '6')
    await expectCellValue('D3', '4')
    await expectCellValue('E3', '8')
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
