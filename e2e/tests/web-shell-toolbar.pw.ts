import { expect, test } from '@playwright/test'
import {
  expectToolbarColor,
  getBox,
  getToolbarButton,
  pickToolbarPresetColor,
  remoteSyncEnabled,
  waitForWorkbookReady,
} from './web-shell-helpers.js'

test('web app renders the minimal product shell without legacy demo chrome', async ({ page }) => {
  await page.goto('/')
  await waitForWorkbookReady(page)

  await expect(page.getByTestId('formula-bar')).toBeVisible()
  await expect(page.getByTestId('name-box')).toBeVisible()
  await expect(page.getByTestId('sheet-grid')).toBeVisible()
  await expect(page.getByRole('tab', { name: 'Sheet1' })).toBeVisible()
  await expect(page.getByRole('heading', { name: 'bilig-demo' })).toHaveCount(0)

  await expect(page.getByTestId('preset-strip')).toHaveCount(0)
  await expect(page.getByTestId('metrics-panel')).toHaveCount(0)
  await expect(page.getByTestId('replica-panel')).toHaveCount(0)

  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!A1')
  await expect(page.getByTestId('status-sync')).toHaveText(
    remoteSyncEnabled
      ? /^(Saved|Saving…|Sync pending|Sync issue)$/
      : /^(Saved|Saving…|Sync pending|Local saved|Local only|Read only|Offline|Sync issue)$/,
    { timeout: 15_000 },
  )
  await expect(page.locator('.formula-result-shell')).toHaveCount(0)
})

test('web app keeps toolbar controls aligned and consistently sized', async ({ page }) => {
  await page.goto('/')
  await waitForWorkbookReady(page)

  const toolbar = page.getByRole('toolbar', { name: 'Formatting toolbar' })
  await expect(toolbar).toBeVisible()

  const controls = [
    page.getByLabel('Number format'),
    page.getByLabel('Font size'),
    page.getByLabel('Bold'),
    page.getByLabel('Italic'),
    page.getByLabel('Underline'),
    page.getByLabel('Fill color'),
    page.getByLabel('Text color'),
    page.getByLabel('Align left'),
    page.getByLabel('Align center'),
    page.getByLabel('Align right'),
    page.getByLabel('Borders'),
    page.getByLabel('Wrap'),
    page.getByLabel('Clear style'),
  ]

  const metrics = await Promise.all(
    controls.map(async (locator) => {
      const label =
        (await locator.getAttribute('aria-label')) ?? (await locator.evaluate((element) => element.textContent?.trim() ?? '')) ?? 'unknown'
      const box = await getBox(locator)
      return {
        label,
        x: Math.round(box.x),
        y: Math.round(box.y),
        width: Math.round(box.width),
        height: Math.round(box.height),
      }
    }),
  )
  const boxes = metrics.map(({ x, y, width, height }) => ({ x, y, width, height }))
  const heights = boxes.map((box) => Math.round(box.height))
  const tops = boxes.map((box) => Math.round(box.y))
  const bottoms = boxes.map((box) => Math.round(box.y + box.height))

  const heightDelta = Math.max(...heights) - Math.min(...heights)
  const topDelta = Math.max(...tops) - Math.min(...tops)
  const bottomDelta = Math.max(...bottoms) - Math.min(...bottoms)
  if (heightDelta > 1 || topDelta > 1 || bottomDelta > 1) {
    throw new Error(`Toolbar geometry mismatch (height=${heightDelta}, top=${topDelta}, bottom=${bottomDelta}): ${JSON.stringify(metrics)}`)
  }

  const toolbarBox = await getBox(toolbar)
  expect(toolbarBox.height).toBeLessThanOrEqual(48)
})

test('web app keeps toolbar, formula bar, grid, and footer tightly stacked', async ({ page }) => {
  await page.goto('/')
  await waitForWorkbookReady(page)

  const toolbar = page.getByRole('toolbar', { name: 'Formatting toolbar' })
  const formulaBar = page.getByTestId('formula-bar')
  const grid = page.getByTestId('sheet-grid')
  const sheetTab = page.getByRole('tab', { name: 'Sheet1' })

  const [toolbarBox, formulaBarBox, gridBox, sheetTabBox] = await Promise.all([
    getBox(toolbar),
    getBox(formulaBar),
    getBox(grid),
    getBox(sheetTab),
  ])

  expect(Math.abs(formulaBarBox.y - (toolbarBox.y + toolbarBox.height))).toBeLessThanOrEqual(8)
  expect(Math.abs(gridBox.y - (formulaBarBox.y + formulaBarBox.height))).toBeLessThanOrEqual(8)
  expect(gridBox.height).toBeGreaterThan(300)
  expect(sheetTabBox.y).toBeGreaterThan(gridBox.y + gridBox.height - 40)
})

test('web app keeps formula bar controls aligned and consistently sized', async ({ page }) => {
  await page.goto('/')
  await waitForWorkbookReady(page)

  const nameBox = page.getByTestId('name-box')
  const formulaFrame = page.getByTestId('formula-input-frame')

  const [nameBoxBox, formulaFrameBox] = await Promise.all([getBox(nameBox), getBox(formulaFrame)])

  expect(Math.abs(nameBoxBox.height - formulaFrameBox.height)).toBeLessThanOrEqual(1)
  expect(Math.abs(nameBoxBox.y - formulaFrameBox.y)).toBeLessThanOrEqual(1)
  expect(Math.abs(nameBoxBox.y + nameBoxBox.height - (formulaFrameBox.y + formulaFrameBox.height))).toBeLessThanOrEqual(1)
})

test('web app keeps the formula input usable on phone-width screens', async ({ page }) => {
  await page.setViewportSize({ width: 390, height: 844 })
  await page.goto('/?sheet=Sheet1&cell=B10')
  await waitForWorkbookReady(page)

  const nameBox = page.getByTestId('name-box')
  const formulaInput = page.getByTestId('formula-input')
  const formulaFrame = page.getByTestId('formula-input-frame')

  const [nameBoxBox, formulaInputBox, formulaFrameBox] = await Promise.all([getBox(nameBox), getBox(formulaInput), getBox(formulaFrame)])
  const placeholderMetrics = await formulaInput.evaluate((element) => {
    if (!(element instanceof HTMLInputElement)) {
      throw new Error('Expected formula input')
    }
    const canvas = document.createElement('canvas')
    const context = canvas.getContext('2d')
    if (!context) {
      throw new Error('Expected canvas text measurement context')
    }
    const style = getComputedStyle(element)
    context.font = style.font
    const textWidth = context.measureText(element.placeholder).width
    const usableWidth = element.getBoundingClientRect().width - Number.parseFloat(style.paddingLeft) - Number.parseFloat(style.paddingRight)
    return { textWidth, usableWidth }
  })

  expect(nameBoxBox.width).toBeLessThanOrEqual(120)
  expect(formulaFrameBox.width).toBeGreaterThan(nameBoxBox.width)
  expect(formulaInputBox.width).toBeGreaterThan(200)
  expect(placeholderMetrics.usableWidth).toBeGreaterThan(placeholderMetrics.textWidth)
})

test('web app keeps the formula placeholder readable on tiny screens', async ({ page }) => {
  await page.setViewportSize({ width: 320, height: 700 })
  await page.goto('/?sheet=Sheet1&cell=B10')
  await waitForWorkbookReady(page)

  const formulaInput = page.getByTestId('formula-input')
  await expect(formulaInput).toHaveAttribute('placeholder', 'Value or =formula')

  const placeholderMetrics = await formulaInput.evaluate((element) => {
    if (!(element instanceof HTMLInputElement)) {
      throw new Error('Expected formula input')
    }
    const canvas = document.createElement('canvas')
    const context = canvas.getContext('2d')
    if (!context) {
      throw new Error('Expected canvas text measurement context')
    }
    const style = getComputedStyle(element)
    context.font = style.font
    const textWidth = context.measureText(element.placeholder).width
    const usableWidth = element.getBoundingClientRect().width - Number.parseFloat(style.paddingLeft) - Number.parseFloat(style.paddingRight)
    return { textWidth, usableWidth }
  })

  expect(placeholderMetrics.usableWidth).toBeGreaterThan(placeholderMetrics.textWidth)
})

test('web app keeps the workbook visible when the assistant rail becomes a phone overlay', async ({ page }) => {
  await page.setViewportSize({ width: 1280, height: 844 })
  await page.goto('/?sheet=Sheet1&cell=B10')
  await waitForWorkbookReady(page)

  await expect(page.getByTestId('workbook-side-panel')).toBeVisible()

  await page.setViewportSize({ width: 390, height: 844 })

  await expect(page.getByTestId('workbook-side-panel')).toHaveCount(0)
  await expect(page.getByTestId('workbook-side-panel-open')).toBeVisible()
  await expect(page.getByTestId('sheet-grid')).toBeVisible()

  const gridBox = await getBox(page.getByTestId('sheet-grid'))
  const horizontalOverflow = await page.evaluate(() => document.documentElement.scrollWidth - document.documentElement.clientWidth)
  expect(gridBox.width).toBeGreaterThan(320)
  expect(horizontalOverflow).toBeLessThanOrEqual(1)
})

test('web app keeps shell controls on one height and radius system', async ({ page }) => {
  await page.goto('/')
  await waitForWorkbookReady(page)

  const locators = [
    page.getByLabel('Number format'),
    page.getByTestId('name-box'),
    page.getByTestId('formula-input-frame'),
    page.getByRole('tab', { name: 'Sheet1' }),
  ]

  const metrics = await Promise.all(
    locators.map(async (locator) => ({
      height: Math.round((await getBox(locator)).height),
      radius: await locator.evaluate((element) => getComputedStyle(element).borderRadius),
    })),
  )

  const heights = metrics.map(({ height }) => height)
  expect(Math.max(...heights) - Math.min(...heights)).toBeLessThanOrEqual(1)
  expect(new Set(metrics.map(({ radius }) => radius)).size).toBe(1)
})

test('web app keeps the toolbar compact on narrow viewports', async ({ page }) => {
  await page.setViewportSize({ width: 620, height: 760 })
  await page.goto('/')
  await waitForWorkbookReady(page)

  const toolbar = page.getByRole('toolbar', { name: 'Formatting toolbar' })
  const firstControl = page.getByLabel('Number format')
  const lastControl = page.getByLabel('Clear style')
  await expect(toolbar).toBeVisible()
  const [toolbarBox, firstControlBox, lastControlBox] = await Promise.all([getBox(toolbar), getBox(firstControl), getBox(lastControl)])

  expect(toolbarBox.height).toBeLessThanOrEqual(48)
  expect(Math.abs(firstControlBox.y - lastControlBox.y)).toBeLessThanOrEqual(1)
  expect(lastControlBox.y + lastControlBox.height).toBeLessThanOrEqual(toolbarBox.y + toolbarBox.height + 1)
})

test('web app keeps tiny toolbar overflow controls from covering formatting actions', async ({ page }) => {
  await page.setViewportSize({ width: 320, height: 700 })
  await page.goto('/?sheet=Sheet1&cell=B10')
  await waitForWorkbookReady(page)

  const toolbar = page.getByRole('toolbar', { name: 'Formatting toolbar' })
  const formattingScroll = page.getByTestId('toolbar-formatting-scroll')
  const overflowCue = page.getByTestId('toolbar-overflow-cue')
  const numberFormat = page.getByLabel('Number format')

  await expect(toolbar).toBeVisible()
  await expect(overflowCue).toBeVisible()
  await expect(page.getByTestId('workbook-shortcut-button')).toBeHidden()
  await expect(page.getByTestId('workbook-import-toggle')).toBeHidden()

  const cuePosition = await overflowCue.evaluate((element) => getComputedStyle(element).position)
  const [toolbarBox, scrollBox, cueBox, numberFormatBox] = await Promise.all([
    getBox(toolbar),
    getBox(formattingScroll),
    getBox(overflowCue),
    getBox(numberFormat),
  ])
  const visibleNumberFormatWidth =
    Math.min(numberFormatBox.x + numberFormatBox.width, scrollBox.x + scrollBox.width) - Math.max(numberFormatBox.x, scrollBox.x)
  const horizontalOverflow = await page.evaluate(() => document.documentElement.scrollWidth - document.documentElement.clientWidth)

  expect(cuePosition).toBe('static')
  expect(cueBox.x).toBeGreaterThanOrEqual(scrollBox.x + scrollBox.width - 1)
  expect(cueBox.x + cueBox.width).toBeLessThanOrEqual(toolbarBox.x + toolbarBox.width + 1)
  expect(visibleNumberFormatWidth).toBeGreaterThanOrEqual(80)
  expect(horizontalOverflow).toBeLessThanOrEqual(1)

  await overflowCue.click()
  await expect
    .poll(async () =>
      formattingScroll.evaluate((element) => {
        if (!(element instanceof HTMLElement)) {
          throw new Error('Expected toolbar formatting scroll region')
        }
        return Math.round(element.scrollLeft)
      }),
    )
    .toBeGreaterThan(0)

  const backCue = page.getByTestId('toolbar-overflow-back-cue')
  await expect(backCue).toBeVisible()
  const [backCueBox, scrolledToolbarBox] = await Promise.all([getBox(backCue), getBox(toolbar)])
  expect(backCueBox.x).toBeGreaterThanOrEqual(scrolledToolbarBox.x - 1)
  expect(backCueBox.x + backCueBox.width).toBeLessThanOrEqual(scrolledToolbarBox.x + scrolledToolbarBox.width + 1)

  await backCue.click()
  await expect
    .poll(async () =>
      formattingScroll.evaluate((element) => {
        if (!(element instanceof HTMLElement)) {
          throw new Error('Expected toolbar formatting scroll region')
        }
        return Math.round(element.scrollLeft)
      }),
    )
    .toBe(0)
  await expect(backCue).toBeHidden()
})

test('web app shows preset color swatches first and only reveals the custom picker on demand', async ({ page }) => {
  await page.goto('/')
  await waitForWorkbookReady(page)

  await page.getByLabel('Fill color').click()
  await expect(page.getByRole('dialog', { name: 'Fill color palette' })).toBeVisible()
  await expect(page.getByLabel('Fill color white')).toBeVisible()
  await expect(page.getByLabel('Fill color light cornflower blue 3')).toBeVisible()
  await expect(page.getByLabel('Fill color dark cornflower blue 3')).toBeVisible()
  await expect(page.getByLabel('Fill color theme cornflower blue')).toBeVisible()
  await expect(page.getByLabel('Custom fill color', { exact: true })).toHaveCount(0)

  await page.getByLabel('Open custom fill color picker').click()
  await expect(page.getByLabel('Custom fill color', { exact: true })).toBeVisible()
})

test('web app renders the fill color palette as a visible popover below the toolbar', async ({ page }) => {
  await page.goto('/')
  await waitForWorkbookReady(page)

  await page.getByLabel('Fill color').click()

  const toolbar = page.getByRole('toolbar', { name: 'Formatting toolbar' })
  const palette = page.getByRole('dialog', { name: 'Fill color palette' })
  const swatch = page.getByLabel('Fill color light cornflower blue 3')
  const [toolbarBox, paletteBox, swatchBox] = await Promise.all([getBox(toolbar), getBox(palette), getBox(swatch)])

  expect(paletteBox.y).toBeGreaterThanOrEqual(toolbarBox.y + toolbarBox.height - 1)
  expect(paletteBox.height).toBeGreaterThan(120)
  expect(paletteBox.width).toBeGreaterThan(200)
  expect(swatchBox.y + swatchBox.height).toBeLessThanOrEqual(paletteBox.y + paletteBox.height)
  await expect(page.getByRole('button', { name: 'Show fill color swatches' })).toBeVisible()
})

test('web app applies preset swatch colors directly from the palette', async ({ page }) => {
  await page.goto('/')
  await waitForWorkbookReady(page)

  await pickToolbarPresetColor(page, 'Fill color', 'light cornflower blue 3')
  await expectToolbarColor(getToolbarButton(page, 'Fill color'), '#c9daf8')

  await pickToolbarPresetColor(page, 'Text color', 'dark blue 1')
  await expectToolbarColor(getToolbarButton(page, 'Text color'), '#3d85c6')
})
