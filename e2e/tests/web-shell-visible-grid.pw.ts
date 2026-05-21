import { expect, test, type Page } from '@playwright/test'
import {
  PRODUCT_HEADER_HEIGHT,
  PRODUCT_ROW_MARKER_WIDTH,
  PRIMARY_MODIFIER,
  clickProductCell,
  countBlueFillPixelsInCell,
  countGreenFillPixelsInCell,
  createTestDocumentId,
  dragProductBodySelection,
  dragProductHeaderSelection,
  getProductColumnLeft,
  getProductColumnWidth,
  getProductRowHeight,
  getProductRowTop,
  pickToolbarPresetColor,
  waitForWorkbookReady,
} from './web-shell-helpers.js'

const DEFAULT_WORKBOOK_CSS_FONT_SIZE = '13.333px'

test('@browser-ci web app paints a workbook skeleton before the app bundle mounts', async ({ page }) => {
  const documentId = createTestDocumentId('playwright-initial-workbook-shell')
  await page.setViewportSize({ width: 1166, height: 820 })
  await page.route('**/*.js', async (route) => {
    await route.fulfill({
      body: '',
      contentType: 'application/javascript',
      status: 200,
    })
  })
  await page.route('**/src/main.tsx**', async (route) => {
    await route.fulfill({
      body: '',
      contentType: 'application/javascript',
      status: 200,
    })
  })
  await page.goto(`/?document=${encodeURIComponent(documentId)}&sheet=Sheet1&cell=E6`, { waitUntil: 'domcontentloaded' })

  await expect(page.locator('#initial-workbook-shell')).toBeVisible()
  await expect(page.getByTestId('initial-workbook-grid')).toBeVisible()
  expect(
    await countInitialShellNonBlankPixels(page),
    'the first paint should show workbook chrome and gridlines instead of a blank page while JavaScript loads',
  ).toBeGreaterThan(1_000)
})

test('@browser-ci web app paints deep querystring-selected cell content in the visible grid', async ({ page }, testInfo) => {
  const documentId = createTestDocumentId('playwright-visible-deep-cell')
  await page.setViewportSize({ width: 1166, height: 820 })
  await page.goto(`/?document=${encodeURIComponent(documentId)}&sheet=Sheet1&cell=D53`)
  await waitForWorkbookReady(page)

  const formulaInput = page.getByTestId('formula-input')
  const nameBox = page.getByTestId('name-box')

  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!D53')
  await expect(nameBox).toHaveValue('D53')
  await formulaInput.fill('Month 1')
  await formulaInput.press('Enter')
  await expect(formulaInput).toHaveValue('Month 1')
  await expect
    .poll(readRendererSurfaceState(page), {
      message: 'TypeGPU should stay visible after its frame is presented without a fallback renderer masking the grid',
      timeout: 5_000,
    })
    .toMatchObject({
      fallbackMounted: false,
      textOverlayMounted: true,
      typeGpuMode: 'typegpu-v3',
      typeGpuOpacity: '1',
    })
  await expect
    .poll(readRenderRevisionState(page), {
      message: 'visible grid proof should expose projected, tile-scene, and visible TypeGPU render revisions',
      timeout: 5_000,
    })
    .toMatchObject({
      projectedRevisionPresent: true,
      typeGpuProjectedRevisionMatchesGrid: true,
      tileSceneRevisionPresent: true,
      visibleProjectedRevisionMatchesGrid: true,
      visibleRenderRevisionPresent: true,
    })

  const paintedPixels = await pollDarkInteriorPixelsInCell(page, 3, 52, (pixels) => pixels > 12)
  const screenshotProofAvailable = paintedPixels > 12
  if (!screenshotProofAvailable && shouldAllowHeadlessWebGpuScreenshotGap()) {
    testInfo.annotations.push({
      description:
        'Local headless Chromium did not composite the WebGPU canvas into page.screenshot; live in-app Browser proof covers this visual path.',
      type: 'visual-proof-limited',
    })
  } else {
    expect(paintedPixels, 'D53 should paint visible text pixels after the edit commits').toBeGreaterThan(12)
  }

  await page.keyboard.press('Delete')
  await expect(formulaInput).toHaveValue('')
  if (screenshotProofAvailable) {
    expect(
      await pollDarkInteriorPixelsInCell(page, 3, 52, (pixels) => pixels < 6),
      'D53 should stop painting stale text after Delete clears the selected cell',
    ).toBeLessThan(6)
  }
})

test('@browser-ci web app keeps table gridlines visible through TypeGPU without fallback masking', async ({ page }) => {
  const documentId = createTestDocumentId('playwright-visible-gridline-floor')
  await page.setViewportSize({ width: 1000, height: 760 })
  await page.goto(`/?document=${encodeURIComponent(documentId)}&persist=0&sheet=Sheet1&cell=A1`)
  await waitForWorkbookReady(page)

  await expect(page.getByTestId('grid-pane-renderer')).toHaveAttribute('data-renderer-mode', 'typegpu-v3')
  await expect.poll(async () => await page.getByTestId('grid-pane-renderer').getAttribute('data-v3-frame-proof-status')).toBe('presented')
  await expect(page.getByTestId('grid-pane-renderer-floor')).toHaveCount(0)
  await expect(page.getByTestId('grid-pane-renderer-fallback')).toHaveCount(0)

  await expect
    .poll(() => countVisibleGridLinePixels(page), {
      message: 'empty workbook view must still expose visible row and column gridlines, not blank white whitespace',
      timeout: 5_000,
    })
    .toBeGreaterThan(1_000)
})

test('@browser-ci web app keeps a user click selection after opening from a deep cell URL', async ({ page }) => {
  const documentId = createTestDocumentId('playwright-visible-click-deep-cell')
  await page.setViewportSize({ width: 1166, height: 820 })
  await page.goto(`/?document=${encodeURIComponent(documentId)}&sheet=Sheet1&cell=D53`)
  await waitForWorkbookReady(page)

  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!D53')
  await expect(page.getByTestId('name-box')).toHaveValue('D53')

  await clickVisibleGridBodySlot(page, 2, 3)

  await expect.poll(() => page.getByTestId('status-selection').textContent()).not.toBe('Sheet1!D53')
  await expect.poll(() => page.getByTestId('name-box').inputValue()).not.toBe('D53')
  await expect.poll(() => page.evaluate(() => new URL(window.location.href).searchParams.get('cell') ?? '')).not.toBe('D53')
})

test('@browser-ci web app keeps dense accounting-sheet text payloads complete in the TypeGPU layer', async ({
  page,
  context,
}, testInfo) => {
  const documentId = createTestDocumentId('playwright-dense-visible-payload')
  await context.grantPermissions(['clipboard-read', 'clipboard-write'])
  await page.setViewportSize({ width: 1166, height: 820 })
  await page.goto(`/?document=${encodeURIComponent(documentId)}&sheet=Sheet1&cell=A1`)
  await waitForWorkbookReady(page)

  const grid = page.getByTestId('sheet-grid')
  const formulaInput = page.getByTestId('formula-input')
  const nameBox = page.getByTestId('name-box')

  await page.evaluate((clipboardText) => navigator.clipboard.writeText(clipboardText), createDenseAccountingGridClipboardText())
  await clickVisibleGridBodySlot(page, 0, 0)
  await grid.press(`${PRIMARY_MODIFIER}+V`)

  await nameBox.fill('B34')
  await nameBox.press('Enter')
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!B34')
  await expect(nameBox).toHaveValue('B34')
  await expect(formulaInput).toHaveValue('Annual software subscription')
  await expect(page.getByTestId('sheet-grid')).toHaveCSS('font-family', /^Arial/)
  await expect(page.getByTestId('sheet-grid')).toHaveCSS('font-size', DEFAULT_WORKBOOK_CSS_FONT_SIZE)
  expect(
    await page.evaluate(() => [...document.fonts].some((fontFace) => fontFace.family === 'Bilig Sans' || fontFace.family === 'Bilig Mono')),
    'workbook shell should not register branded font faces that can swap under the grid',
  ).toBe(false)
  await expect
    .poll(readRendererSurfaceState(page), {
      message: 'dense accounting sheet should render through TypeGPU plus the native text layer, not the fallback canvas',
      timeout: 5_000,
    })
    .toMatchObject({
      fallbackMounted: false,
      textOverlayMounted: true,
      typeGpuMode: 'typegpu-v3',
      typeGpuOpacity: '1',
    })
  await expect
    .poll(readTypeGpuTextRunCount(page), {
      message: 'visible body tiles must carry dense text payloads; a blank successful frame is a visual regression',
      timeout: 5_000,
    })
    .toBeGreaterThan(40)
  await expect
    .poll(readNativeTextRunGeometryHealth(page), {
      message: 'native text runs must be clipped to the visible pane instead of creating huge offscreen DOM boxes',
      timeout: 5_000,
    })
    .toMatchObject({
      nativeRunCountMatchesDom: true,
      visibleOversizedRuns: 0,
    })

  const selectedCellTextPixels = await pollDarkInteriorPixelsInCell(page, 1, 33, (pixels) => pixels > 8)
  if (selectedCellTextPixels <= 8 && shouldAllowHeadlessWebGpuScreenshotGap()) {
    testInfo.annotations.push({
      description:
        'Local headless Chromium did not composite the WebGPU text into page.screenshot; TypeGPU text-run proof covers this visual path.',
      type: 'visual-proof-limited',
    })
  } else {
    expect(selectedCellTextPixels, 'B34 should visibly paint its text, not only expose formula-bar readback').toBeGreaterThan(8)
  }
})

test('@browser-ci web app keeps the live cell editor above the native grid text layer', async ({ page }) => {
  const documentId = createTestDocumentId('playwright-editor-native-text-z-order')
  await page.setViewportSize({ width: 1166, height: 820 })
  await page.goto(`/?document=${encodeURIComponent(documentId)}&sheet=Sheet1&cell=A1`)
  await waitForWorkbookReady(page)

  await clickVisibleGridBodySlot(page, 1, 1)
  await page.keyboard.type('editor-z-order')
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!B2')
  await expect(page.getByTestId('cell-editor-input')).toHaveValue('editor-z-order')
  await page.getByTestId('cell-editor-input').press('Enter')
  await expect(page.getByTestId('cell-editor-input')).toHaveCount(0)

  await clickVisibleGridBodySlot(page, 1, 1)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!B2')
  await page.getByTestId('sheet-grid').press('F2')
  await expect(page.getByTestId('cell-editor-input')).toHaveValue('editor-z-order')
  await expect(page.getByTestId('cell-editor-input')).not.toHaveCSS('opacity', '0')
  await expect
    .poll(readEditorLayerState(page, { col: 1, row: 1 }), {
      message:
        'the active editor must sit above native rendered text and suppress its own committed text run to prevent double-text artifacts',
      timeout: 5_000,
    })
    .toMatchObject({
      activeCellNativeTextSuppressed: true,
      editorAboveNativeText: true,
      editorMounted: true,
      nativeTextLayerMounted: true,
    })

  await clickVisibleGridBodySlot(page, 2, 1)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!C2')
  await expect(page.getByTestId('cell-editor-input')).toHaveCount(0)
  await expect
    .poll(readNativeTextLayerStabilityState(page, 'editor-z-order'), {
      message: 'click-away commit must not blank row headers or lose the committed cell text from the native text layer',
      timeout: 5_000,
    })
    .toMatchObject({
      containsText: true,
      declaredRunCountMatchesDom: true,
      rowHeaderRunCountHealthy: true,
    })
})

test('@browser-ci web app keeps rendered edits, clears, headers, and fills coherent across click-away and reload', async ({ page }) => {
  const documentId = createTestDocumentId('playwright-rendered-table-stakes')
  const editedText = 'rendered-table-stakes'
  await page.setViewportSize({ width: 1166, height: 820 })
  await page.goto(`/?document=${encodeURIComponent(documentId)}&sheet=Sheet1&cell=A1`)
  await waitForWorkbookReady(page)

  await clickProductCell(page, 1, 1)
  await page.keyboard.type(editedText)
  await expect(page.getByTestId('cell-editor-input')).toHaveValue(editedText)
  await expect
    .poll(readEditorLayerState(page, { col: 1, row: 1 }), {
      message: 'active editor text must replace, not duplicate, the committed native text for that cell',
      timeout: 5_000,
    })
    .toMatchObject({
      activeCellNativeTextSuppressed: true,
      editorAboveNativeText: true,
      editorMounted: true,
      nativeTextLayerMounted: true,
    })

  await clickProductCell(page, 2, 1)
  await expect(page.getByTestId('cell-editor-input')).toHaveCount(0)
  await expect
    .poll(readNativeTextLayerStabilityState(page, editedText), {
      message: 'click-away commit must keep one rendered text run and preserve row header numbers',
      timeout: 5_000,
    })
    .toMatchObject({
      containsText: true,
      declaredRunCountMatchesDom: true,
      duplicateTextRunCount: 1,
      rowHeaderRunCountHealthy: true,
    })

  await clickProductCell(page, 1, 1)
  await page.keyboard.press('Delete')
  await expect(page.getByTestId('formula-input')).toHaveValue('')
  await clickProductCell(page, 3, 2)
  await expect.poll(() => nativeTextRunsInclude(page, editedText)).toBe(false)

  await clickProductCell(page, 4, 5)
  await pickToolbarPresetColor(page, 'Fill color', 'green')
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!E6')
  await clickProductCell(page, 6, 7)
  await expect
    .poll(() => countGreenFillPixelsInCell(page, 4, 5), {
      message: 'toolbar fill color must repaint the actual visible grid, not only toolbar state',
      timeout: 5_000,
    })
    .toBeGreaterThan(120)

  await expect
    .poll(() => page.getByTestId('status-sync').textContent(), {
      message: 'workbook should finish saving before reload persistence proof',
      timeout: 15_000,
    })
    .toMatch(/^(Saved|Local saved|Local only)$/)
  await page.reload({ waitUntil: 'domcontentloaded' })
  await waitForWorkbookReady(page)
  await clickProductCell(page, 3, 2)
  await expect.poll(() => nativeTextRunsInclude(page, editedText)).toBe(false)
  await expect
    .poll(() => countGreenFillPixelsInCell(page, 4, 5), {
      message: 'toolbar fill color must survive reload and stay visibly painted',
      timeout: 5_000,
    })
    .toBeGreaterThan(120)
})

test('@browser-ci web app preserves visible fill while Delete clears only cell content', async ({ page }) => {
  const documentId = createTestDocumentId('playwright-delete-preserves-fill')
  const text = 'delete-keeps-fill'
  await page.setViewportSize({ width: 1166, height: 820 })
  await page.goto(`/?document=${encodeURIComponent(documentId)}&persist=0&sheet=Sheet1&cell=A1`)
  await waitForWorkbookReady(page)

  const grid = page.getByTestId('sheet-grid')
  const formulaInput = page.getByTestId('formula-input')

  await clickProductCell(page, 1, 1)
  await formulaInput.fill(text)
  await formulaInput.press('Enter')
  await expect(formulaInput).toHaveValue(text)
  await pickToolbarPresetColor(page, 'Fill color', 'green')
  await expect
    .poll(() => countGreenFillPixelsInCell(page, 1, 1), {
      message: 'test setup should visibly paint B2 green before Delete',
      timeout: 5_000,
    })
    .toBeGreaterThan(120)

  await clickProductCell(page, 1, 1)
  await grid.press('Delete')
  await expect(formulaInput).toHaveValue('')
  await expect.poll(() => nativeTextRunTextAt(page, 1, 1)).toBe('')

  const postDeleteFillSamples = await sampleGreenFillPixelsAcrossFrames(page, 1, 1, 4)
  expect(
    Math.min(...postDeleteFillSamples),
    'Delete must not flash the selected cell background to default while the clear mutation is optimistic or settling',
  ).toBeGreaterThan(120)
  await expect
    .poll(() => countGreenFillPixelsInCell(page, 1, 1), {
      message: 'Delete should clear content while leaving cell fill formatting visible after persistence settles',
      timeout: 5_000,
    })
    .toBeGreaterThan(120)

  await clickProductCell(page, 3, 3)
  await expect.poll(() => nativeTextRunsInclude(page, text)).toBe(false)
  await expect
    .poll(() => countGreenFillPixelsInCell(page, 1, 1), {
      message: 'deleted text must not come back after focus leaves the formatted cell',
      timeout: 5_000,
    })
    .toBeGreaterThan(120)

  const scrollViewport = page.getByTestId('grid-scroll-viewport')
  await scrollViewport.evaluate((viewport) => {
    viewport.scrollTop = 900
    viewport.scrollLeft = 220
    viewport.dispatchEvent(new Event('scroll', { bubbles: true }))
  })
  await expect.poll(() => nativeTextRunsInclude(page, text)).toBe(false)

  await scrollViewport.evaluate((viewport) => {
    viewport.scrollTop = 0
    viewport.scrollLeft = 0
    viewport.dispatchEvent(new Event('scroll', { bubbles: true }))
  })
  await clickProductCell(page, 1, 1)
  await expect(formulaInput).toHaveValue('')
  await expect.poll(() => nativeTextRunTextAt(page, 1, 1)).toBe('')
  await expect
    .poll(() => countGreenFillPixelsInCell(page, 1, 1), {
      message: 'formatted deleted cell must survive tile eviction and return without ghost content',
      timeout: 5_000,
    })
    .toBeGreaterThan(120)
})

test('@browser-ci web app repaints same-size TypeGPU fill color changes without stale tile colors', async ({ page }) => {
  const documentId = createTestDocumentId('playwright-fill-color-repaint')
  await page.setViewportSize({ width: 1166, height: 820 })
  await page.goto(`/?document=${encodeURIComponent(documentId)}&persist=0&sheet=Sheet1&cell=A1`)
  await waitForWorkbookReady(page)

  await clickProductCell(page, 1, 1)
  await pickToolbarPresetColor(page, 'Fill color', 'green')
  await expect
    .poll(() => countGreenFillPixelsInCell(page, 1, 1), {
      message: 'test setup should visibly paint B2 green before repainting the same rect buffer blue',
      timeout: 5_000,
    })
    .toBeGreaterThan(120)

  await pickToolbarPresetColor(page, 'Fill color', 'blue')
  const blueSamples = await sampleFillPixelsAcrossFrames(page, 1, 1, 'blue', 4)
  expect(
    Math.min(...blueSamples),
    'same-size fill color changes must upload new TypeGPU rect payloads instead of keeping stale green tiles',
  ).toBeGreaterThan(120)
  await expect.poll(() => countGreenFillPixelsInCell(page, 1, 1)).toBe(0)

  await page.keyboard.press('Delete')
  await expect(page.getByTestId('formula-input')).toHaveValue('')
  const postDeleteBlueSamples = await sampleFillPixelsAcrossFrames(page, 1, 1, 'blue', 4)
  expect(
    Math.min(...postDeleteBlueSamples),
    'Delete must clear content without flashing or removing the visible blue fill',
  ).toBeGreaterThan(120)
})

test('@browser-ci web app paints toolbar fill across a selected range without hiding the range color', async ({ page }) => {
  const documentId = createTestDocumentId('playwright-range-fill-visible')
  await page.setViewportSize({ width: 1166, height: 820 })
  await page.goto(`/?document=${encodeURIComponent(documentId)}&persist=0&sheet=Sheet1&cell=A1`)
  await waitForWorkbookReady(page)

  await dragProductBodySelection(page, 2, 4, 5, 10)
  await expect(page.getByTestId('name-box')).toHaveValue('C5:F11')

  await pickToolbarPresetColor(page, 'Fill color', 'green')
  await expect(page.getByTestId('name-box')).toHaveValue('C5:F11')
  await expect
    .poll(() => countGreenFillPixelsInCell(page, 2, 4), {
      message: 'selected range start cell should visibly repaint green while the range remains selected',
      timeout: 5_000,
    })
    .toBeGreaterThan(120)
  await expect.poll(() => countGreenFillPixelsInCell(page, 4, 7)).toBeGreaterThan(120)
  await expect.poll(() => countGreenFillPixelsInCell(page, 5, 10)).toBeGreaterThan(120)

  await pickToolbarPresetColor(page, 'Fill color', 'blue')
  await expect(page.getByTestId('name-box')).toHaveValue('C5:F11')
  await expect
    .poll(() => countBlueFillPixelsInCell(page, 2, 4), {
      message: 'same selected range should repaint blue instead of keeping stale green cells',
      timeout: 5_000,
    })
    .toBeGreaterThan(120)
  await expect.poll(() => countBlueFillPixelsInCell(page, 4, 7)).toBeGreaterThan(120)
  await expect.poll(() => countBlueFillPixelsInCell(page, 5, 10)).toBeGreaterThan(120)
  await expect.poll(() => countGreenFillPixelsInCell(page, 4, 7)).toBeLessThan(12)

  await clickProductCell(page, 7, 12)
  await expect.poll(() => countBlueFillPixelsInCell(page, 2, 4)).toBeGreaterThan(120)
  await expect.poll(() => countBlueFillPixelsInCell(page, 4, 7)).toBeGreaterThan(120)
  await expect.poll(() => countBlueFillPixelsInCell(page, 5, 10)).toBeGreaterThan(120)
})

test('@browser-ci web app repaints moved text cells when a background fill is applied', async ({ page, context }) => {
  const documentId = createTestDocumentId('playwright-moved-cell-fill-repaint')
  const movedText = 'moved-fill-proof'
  await context.grantPermissions(['clipboard-read', 'clipboard-write'])
  await page.setViewportSize({ width: 1166, height: 820 })
  await page.goto(`/?document=${encodeURIComponent(documentId)}&persist=0&sheet=Sheet1&cell=A1`)
  await waitForWorkbookReady(page)

  await clickProductCell(page, 1, 1)
  await page.keyboard.type(movedText)
  await expect(page.getByTestId('cell-editor-input')).toHaveValue(movedText)
  await page.getByTestId('cell-editor-input').press('Enter')
  await expect(page.getByTestId('cell-editor-input')).toHaveCount(0)

  await clickProductCell(page, 1, 1)
  await page.keyboard.press(`${PRIMARY_MODIFIER}+X`)
  await clickProductCell(page, 3, 3)
  await page.keyboard.press(`${PRIMARY_MODIFIER}+V`)

  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!D4')
  await expect(page.getByTestId('formula-input')).toHaveValue(movedText)
  await expect
    .poll(() => nativeTextRunTextAt(page, 1, 1), {
      message: 'cut source B2 must not leave ghost text in the native text layer',
      timeout: 5_000,
    })
    .toBe('')
  await expect
    .poll(() => nativeTextRunTextAt(page, 3, 3), {
      message: 'paste target D4 must own exactly the moved text before fill is applied',
      timeout: 5_000,
    })
    .toBe(movedText)

  await pickToolbarPresetColor(page, 'Fill color', 'green')
  await clickProductCell(page, 5, 5)
  await expect
    .poll(() => countGreenFillPixelsInCell(page, 3, 3), {
      message: 'background fill must repaint the actual moved-text cell, not only toolbar or model state',
      timeout: 5_000,
    })
    .toBeGreaterThan(120)
  await expect.poll(() => nativeTextRunTextAt(page, 3, 3)).toBe(movedText)
})

test('@browser-ci web app copies presentation and clears stale target fills in visible tiles', async ({ page, context }) => {
  const documentId = createTestDocumentId('playwright-copy-presentation-repaint')
  await context.grantPermissions(['clipboard-read', 'clipboard-write'])
  await page.setViewportSize({ width: 1166, height: 820 })
  await page.goto(`/?document=${encodeURIComponent(documentId)}&persist=0&sheet=Sheet1&cell=A1`)
  await waitForWorkbookReady(page)

  const grid = page.getByTestId('sheet-grid')
  const formulaInput = page.getByTestId('formula-input')

  await clickProductCell(page, 1, 1)
  await formulaInput.fill('styled-source')
  await formulaInput.press('Enter')
  await pickToolbarPresetColor(page, 'Fill color', 'green')
  await expect
    .poll(() => countGreenFillPixelsInCell(page, 1, 1), {
      message: 'styled source B2 should visibly paint the selected green fill',
      timeout: 5_000,
    })
    .toBeGreaterThan(120)

  await clickProductCell(page, 1, 2)
  await formulaInput.fill('plain-source')
  await formulaInput.press('Enter')

  await clickProductCell(page, 3, 1)
  await formulaInput.fill('stale-target-one')
  await formulaInput.press('Enter')
  await clickProductCell(page, 3, 2)
  await formulaInput.fill('stale-target-two')
  await formulaInput.press('Enter')
  await dragProductBodySelection(page, 3, 1, 3, 2)
  await pickToolbarPresetColor(page, 'Fill color', 'blue')
  await expect
    .poll(() => countBlueFillPixelsInCell(page, 3, 2), {
      message: 'test setup should visibly paint stale blue fill before the copy clears it',
      timeout: 5_000,
    })
    .toBeGreaterThan(120)

  await dragProductBodySelection(page, 1, 1, 1, 2)
  await grid.press(`${PRIMARY_MODIFIER}+C`)
  await clickProductCell(page, 3, 1)
  await grid.press(`${PRIMARY_MODIFIER}+V`)

  await expect.poll(() => nativeTextRunTextAt(page, 3, 1)).toBe('styled-source')
  await expect.poll(() => nativeTextRunTextAt(page, 3, 2)).toBe('plain-source')
  await expect
    .poll(() => countGreenFillPixelsInCell(page, 3, 1), {
      message: 'copying B2 over D2 must repaint the target tile with source presentation',
      timeout: 5_000,
    })
    .toBeGreaterThan(120)
  await expect
    .poll(() => countBlueFillPixelsInCell(page, 3, 2), {
      message: 'copying plain B3 over D3 must clear the stale target fill from the visible tile',
      timeout: 5_000,
    })
    .toBeLessThan(12)
})

test('@browser-ci web app moves background fill presentation without source or target ghosts', async ({ page, context }) => {
  const documentId = createTestDocumentId('playwright-move-fill-presentation')
  const movedText = 'move-fill-source'
  await context.grantPermissions(['clipboard-read', 'clipboard-write'])
  await page.setViewportSize({ width: 1166, height: 820 })
  await page.goto(`/?document=${encodeURIComponent(documentId)}&persist=0&sheet=Sheet1&cell=A1`)
  await waitForWorkbookReady(page)

  const grid = page.getByTestId('sheet-grid')
  const formulaInput = page.getByTestId('formula-input')

  await clickProductCell(page, 1, 1)
  await formulaInput.fill(movedText)
  await formulaInput.press('Enter')
  await pickToolbarPresetColor(page, 'Fill color', 'green')
  await expect
    .poll(() => countGreenFillPixelsInCell(page, 1, 1), {
      message: 'test setup should visibly paint green fill on the move source',
      timeout: 5_000,
    })
    .toBeGreaterThan(120)

  await clickProductCell(page, 3, 3)
  await formulaInput.fill('stale-target')
  await formulaInput.press('Enter')
  await pickToolbarPresetColor(page, 'Fill color', 'blue')
  await expect
    .poll(() => countBlueFillPixelsInCell(page, 3, 3), {
      message: 'test setup should visibly paint stale blue fill before the move overwrites it',
      timeout: 5_000,
    })
    .toBeGreaterThan(120)

  await clickProductCell(page, 1, 1)
  await grid.press(`${PRIMARY_MODIFIER}+X`)
  await clickProductCell(page, 3, 3)
  await grid.press(`${PRIMARY_MODIFIER}+V`)
  await clickProductCell(page, 5, 5)

  await expect
    .poll(() => nativeTextRunTextAt(page, 1, 1), {
      message: 'move must clear source text from the native text layer',
      timeout: 5_000,
    })
    .toBe('')
  await expect
    .poll(() => nativeTextRunTextAt(page, 3, 3), {
      message: 'move must render moved text at the target cell',
      timeout: 5_000,
    })
    .toBe(movedText)

  expect(
    Math.max(...(await sampleGreenFillPixelsAcrossFrames(page, 1, 1, 4))),
    'move must clear source fill instead of leaving a stale green tile',
  ).toBeLessThan(12)
  expect(
    Math.min(...(await sampleFillPixelsAcrossFrames(page, 3, 3, 'green', 4))),
    'move must paint the target with the moved green fill',
  ).toBeGreaterThan(120)
  await expect
    .poll(() => countBlueFillPixelsInCell(page, 3, 3), {
      message: 'move must clear stale target blue fill instead of blending or retaining old tile color',
      timeout: 5_000,
    })
    .toBeLessThan(12)
})

test('@browser-ci web app repaints shifted styled survivors after structural row delete', async ({ page }) => {
  const documentId = createTestDocumentId('playwright-structural-row-delete-visual-survivor')
  const survivorText = 'row-delete-survivor'
  await page.setViewportSize({ width: 1280, height: 1040 })
  await page.goto(`/?document=${encodeURIComponent(documentId)}&persist=0&sheet=Sheet1&cell=A1`)
  await waitForWorkbookReady(page)

  const formulaInput = page.getByTestId('formula-input')

  await clickProductCell(page, 1, 39)
  await formulaInput.fill(survivorText)
  await formulaInput.press('Enter')
  await pickToolbarPresetColor(page, 'Fill color', 'green')
  await expect
    .poll(() => nativeTextRunTextAt(page, 1, 39), {
      message: 'test setup should render the styled survivor before deleting an earlier row',
      timeout: 5_000,
    })
    .toBe(survivorText)
  await expect.poll(() => countGreenFillPixelsInCell(page, 1, 39)).toBeGreaterThan(120)

  await dragProductHeaderSelection(page, 'row', 1, 1)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!2:2')
  await page.keyboard.down(PRIMARY_MODIFIER)
  await page.keyboard.down('Alt')
  await page.keyboard.press('Minus')
  await page.keyboard.up('Alt')
  await page.keyboard.up(PRIMARY_MODIFIER)

  await expect
    .poll(() => nativeTextRunTextAt(page, 1, 38), {
      message: 'text below a deleted row must shift up in the native text layer',
      timeout: 5_000,
    })
    .toBe(survivorText)
  await expect
    .poll(() => nativeTextRunTextAt(page, 1, 39), {
      message: 'the old row tile must not keep ghost text after structural delete',
      timeout: 5_000,
    })
    .toBe('')
  await expect
    .poll(() => countGreenFillPixelsInCell(page, 1, 38), {
      message: 'the shifted survivor should keep its green fill',
      timeout: 5_000,
    })
    .toBeGreaterThan(120)
  await expect
    .poll(() => countGreenFillPixelsInCell(page, 1, 39), {
      message: 'the old row tile should not keep stale green fill',
      timeout: 5_000,
    })
    .toBeLessThan(12)
})

function createDenseAccountingGridClipboardText(): string {
  return Array.from({ length: 48 }, (_, rowIndex) => {
    const rowNumber = rowIndex + 1
    if (rowNumber === 1) {
      return ['Prepaid Expense Template', '', '', '', '', '', '', ''].join('\t')
    }
    if (rowNumber === 2) {
      return ['As of date', '2026-04-30', '', 'Purpose', 'Track prepaid assets and amortization', '', '', ''].join('\t')
    }
    if (rowNumber === 3) {
      return ['How to use', 'Enter/edit yellow input cells in the register', '', '', '', '', '', ''].join('\t')
    }
    if (rowNumber === 4) {
      return ['Summary', 'Total Prepaid', 'Expense Recognized', 'Remaining Prepaid', 'Monthly Expense', '', '', ''].join('\t')
    }
    if (rowNumber === 5) {
      return ['All register rows', '54600', '13600', '41000', '4000', '', '', ''].join('\t')
    }
    if (rowNumber === 7) {
      return ['Item ID', 'Description', 'Vendor', 'Prepaid Asset Account', 'Expense Account', 'Start Date', 'End Date', 'Total Cost'].join(
        '\t',
      )
    }
    if (rowNumber >= 8 && rowNumber <= 13) {
      const descriptions = [
        'General liability insurance',
        'Annual software subscription',
        'Quarterly office rent',
        'Support contract',
        'Equipment service',
        'Enter description',
      ]
      const vendors = ['Acme Insurance', 'SaaS Co', 'Landlord Co', 'Cloud Support', 'Service Co', 'Enter vendor']
      const index = rowNumber - 8
      return [
        `P-00${index + 1}`,
        descriptions[index] ?? '',
        vendors[index] ?? '',
        'Prepaid Asset',
        'Expense Account',
        '2026-01-01',
        '2026-12-31',
        '12000',
      ].join('\t')
    }
    if (rowNumber === 15) {
      return ['Amortization Schedule Examples', '', '', '', '', '', '', ''].join('\t')
    }
    if (rowNumber === 16) {
      return [
        'Item ID',
        'Description',
        'Period',
        'Period Label',
        'Opening Balance',
        'Monthly Expense',
        'Ending Balance',
        'Debit Account',
      ].join('\t')
    }
    if (rowNumber >= 17) {
      const period = rowNumber - 16
      const itemId = period <= 12 ? 'P-001' : period <= 24 ? 'P-002' : 'P-003'
      const description =
        itemId === 'P-002' ? 'Annual software subscription' : itemId === 'P-003' ? 'Quarterly office rent' : 'General liability insurance'
      const itemPeriod = ((period - 1) % 12) + 1
      const opening = itemId === 'P-002' ? Math.max(0, 26000 - itemPeriod * 2000) : Math.max(0, 13000 - itemPeriod * 1000)
      const monthly = itemId === 'P-002' ? 2000 : 1000
      return [
        itemId,
        description,
        String(itemPeriod),
        `Month ${itemPeriod}`,
        String(opening),
        String(monthly),
        String(Math.max(0, opening - monthly)),
        'Software Expense',
      ].join('\t')
    }
    return ['', '', '', '', '', '', '', ''].join('\t')
  }).join('\n')
}

async function pollDarkInteriorPixelsInCell(
  page: Page,
  columnIndex: number,
  rowIndex: number,
  predicate: (pixels: number) => boolean,
): Promise<number> {
  const deadline = Date.now() + 5_000
  const poll = async (): Promise<number> => {
    const lastPixels = await countDarkInteriorPixelsInCell(page, columnIndex, rowIndex)
    if (predicate(lastPixels)) {
      return lastPixels
    }
    if (Date.now() >= deadline) {
      return lastPixels
    }
    await page.waitForTimeout(100)
    return poll()
  }
  return await poll()
}

async function clickVisibleGridBodySlot(page: Page, columnIndex: number, visibleRowSlot: number): Promise<void> {
  const gridLocator = page.getByTestId('sheet-grid')
  await expect(gridLocator).toBeVisible()
  const grid = await gridLocator.boundingBox()
  if (!grid) {
    throw new Error('sheet grid is not visible')
  }

  const [columnLeft, columnWidth, rowHeight] = await Promise.all([
    getProductColumnLeft(page, columnIndex),
    getProductColumnWidth(page, columnIndex),
    getProductRowHeight(page, visibleRowSlot),
  ])
  await page.mouse.click(
    grid.x + columnLeft + Math.floor(columnWidth / 2),
    grid.y + PRODUCT_HEADER_HEIGHT + visibleRowSlot * rowHeight + Math.floor(rowHeight / 2),
  )
}

function shouldAllowHeadlessWebGpuScreenshotGap(): boolean {
  return process.platform === 'darwin' && process.env['CI'] !== '1' && process.env['CI'] !== 'true'
}

function readRenderRevisionState(page: Page): () => Promise<{
  readonly projectedRevision: string
  readonly projectedRevisionPresent: boolean
  readonly typeGpuProjectedRevision: string
  readonly typeGpuProjectedRevisionMatchesGrid: boolean
  readonly tileSceneRevisionPresent: boolean
  readonly visibleProjectedRevision: string
  readonly visibleProjectedRevisionMatchesGrid: boolean
  readonly visibleRenderRevisionPresent: boolean
}> {
  return async () =>
    await page.evaluate(() => {
      const grid = document.querySelector('[data-testid="sheet-grid"]')
      const typeGpu = document.querySelector('[data-testid="grid-pane-renderer"]')
      const projectedRevision = grid?.getAttribute('data-render-projected-revision') ?? ''
      const typeGpuProjectedRevision = typeGpu?.getAttribute('data-v3-projected-render-revision') ?? ''
      const tileSceneRevision = typeGpu?.getAttribute('data-v3-tile-scene-revision') ?? ''
      const visibleProjectedRevision = typeGpu?.getAttribute('data-v3-visible-projected-render-revision') ?? ''
      const visibleRenderRevision = typeGpu?.getAttribute('data-v3-visible-render-revision') ?? ''
      return {
        projectedRevision,
        projectedRevisionPresent: projectedRevision.length > 0,
        typeGpuProjectedRevision,
        typeGpuProjectedRevisionMatchesGrid: typeGpuProjectedRevision.length > 0 && typeGpuProjectedRevision === projectedRevision,
        tileSceneRevisionPresent: tileSceneRevision.length > 0,
        visibleProjectedRevision,
        visibleProjectedRevisionMatchesGrid: visibleProjectedRevision.length > 0 && visibleProjectedRevision === projectedRevision,
        visibleRenderRevisionPresent: visibleRenderRevision.length > 0,
      }
    })
}

function readRendererSurfaceState(page: Page): () => Promise<{
  readonly fallbackMounted: boolean
  readonly textOverlayMounted: boolean
  readonly typeGpuMode: string | null
  readonly typeGpuOpacity: string | null
}> {
  return async () =>
    await page.evaluate(() => {
      const typeGpu = document.querySelector('[data-testid="grid-pane-renderer"]')
      const fallback = document.querySelector('[data-testid="grid-pane-renderer-fallback"]')
      const textOverlay = document.querySelector('[data-testid="grid-native-text-layer"]')
      return {
        fallbackMounted: fallback instanceof HTMLCanvasElement,
        textOverlayMounted: textOverlay instanceof HTMLElement,
        typeGpuMode: typeGpu instanceof HTMLCanvasElement ? typeGpu.getAttribute('data-renderer-mode') : null,
        typeGpuOpacity: typeGpu instanceof HTMLElement ? getComputedStyle(typeGpu).opacity : null,
      }
    })
}

function readTypeGpuTextRunCount(page: Page): () => Promise<number> {
  return async () =>
    await page.evaluate(() => {
      const typeGpu = document.querySelector('[data-testid="grid-pane-renderer"]')
      if (!(typeGpu instanceof HTMLElement) || typeGpu.getAttribute('data-v3-frame-proof-status') !== 'presented') {
        return 0
      }
      return Number(typeGpu.getAttribute('data-v3-text-run-count') ?? '0')
    })
}

function readNativeTextRunGeometryHealth(page: Page): () => Promise<{
  readonly gridWidth: number
  readonly maxVisibleRunWidth: number
  readonly nativeRunCountMatchesDom: boolean
  readonly visibleOversizedRuns: number
}> {
  return async () =>
    await page.evaluate(() => {
      const grid = document.querySelector('[data-testid="sheet-grid"]')
      const gridRect = grid instanceof HTMLElement ? grid.getBoundingClientRect() : null
      if (!gridRect) {
        return {
          gridWidth: 0,
          maxVisibleRunWidth: 0,
          nativeRunCountMatchesDom: false,
          visibleOversizedRuns: 0,
        }
      }
      const nativeTextLayer = document.querySelector('[data-testid="grid-native-text-layer"]')
      const mountedTextRuns = document.querySelectorAll('[data-native-text-run]').length
      const declaredTextRuns =
        nativeTextLayer instanceof HTMLElement ? Number(nativeTextLayer.getAttribute('data-v3-native-text-run-count') ?? '-1') : -1
      let maxVisibleRunWidth = 0
      let visibleOversizedRuns = 0
      for (const outer of document.querySelectorAll('[data-native-text-run]')) {
        const outerRect = outer.getBoundingClientRect()
        const visible =
          outerRect.right > gridRect.left &&
          outerRect.left < gridRect.right &&
          outerRect.bottom > gridRect.top &&
          outerRect.top < gridRect.bottom
        if (!visible) {
          continue
        }
        const inner = outer.firstElementChild
        const innerRect = inner instanceof HTMLElement ? inner.getBoundingClientRect() : outerRect
        const runWidth = Math.max(outerRect.width, innerRect.width)
        maxVisibleRunWidth = Math.max(maxVisibleRunWidth, runWidth)
        if (runWidth > gridRect.width + 1) {
          visibleOversizedRuns += 1
        }
      }
      return {
        gridWidth: gridRect.width,
        maxVisibleRunWidth,
        nativeRunCountMatchesDom: declaredTextRuns === mountedTextRuns,
        visibleOversizedRuns,
      }
    })
}

function readEditorLayerState(
  page: Page,
  activeCell: { readonly col: number; readonly row: number },
): () => Promise<{
  readonly activeCellNativeTextSuppressed: boolean
  readonly editorAboveNativeText: boolean
  readonly editorMounted: boolean
  readonly nativeTextLayerMounted: boolean
}> {
  return async () =>
    await page.evaluate(({ col, row }) => {
      const editor = document.querySelector('[data-testid="cell-editor-overlay"]')
      const nativeTextLayer = document.querySelector('[data-testid="grid-native-text-layer"]')
      const editorZIndex = editor instanceof HTMLElement ? Number(getComputedStyle(editor).zIndex) : Number.NaN
      const nativeTextLayerZIndex = nativeTextLayer instanceof HTMLElement ? Number(getComputedStyle(nativeTextLayer).zIndex) : Number.NaN
      const activeCellNativeTextRun = document.querySelector(
        `[data-native-text-run-row="${String(row)}"][data-native-text-run-col="${String(col)}"]`,
      )
      return {
        activeCellNativeTextSuppressed: activeCellNativeTextRun === null,
        editorAboveNativeText:
          Number.isFinite(editorZIndex) && Number.isFinite(nativeTextLayerZIndex) && editorZIndex > nativeTextLayerZIndex,
        editorMounted: editor instanceof HTMLElement,
        nativeTextLayerMounted: nativeTextLayer instanceof HTMLElement,
      }
    }, activeCell)
}

function readNativeTextLayerStabilityState(
  page: Page,
  expectedText: string,
): () => Promise<{
  readonly containsText: boolean
  readonly declaredRunCountMatchesDom: boolean
  readonly duplicateTextRunCount: number
  readonly rowHeaderRunCount: number
  readonly rowHeaderRunCountHealthy: boolean
}> {
  return async () =>
    await page.evaluate(
      ({ rowMarkerWidth, text }) => {
        const nativeTextLayer = document.querySelector('[data-testid="grid-native-text-layer"]')
        const grid = document.querySelector('[data-testid="sheet-grid"]')
        const gridRect = grid instanceof HTMLElement ? grid.getBoundingClientRect() : null
        const textRuns = [...document.querySelectorAll('[data-native-text-run]')]
        const declaredRunCount =
          nativeTextLayer instanceof HTMLElement ? Number(nativeTextLayer.getAttribute('data-v3-native-text-run-count') ?? '-1') : -1
        const duplicateTextRunCount = textRuns.filter((run) => run.textContent === text).length
        const rowHeaderRunCount = textRuns.filter((run) => {
          if (!gridRect || !/^\d+$/.test(run.textContent ?? '')) {
            return false
          }
          const rect = run.getBoundingClientRect()
          return rect.left >= gridRect.left && rect.right <= gridRect.left + rowMarkerWidth + 1
        }).length
        return {
          containsText: duplicateTextRunCount > 0,
          declaredRunCountMatchesDom: declaredRunCount === textRuns.length,
          duplicateTextRunCount,
          rowHeaderRunCount,
          rowHeaderRunCountHealthy: rowHeaderRunCount >= 10,
        }
      },
      { rowMarkerWidth: PRODUCT_ROW_MARKER_WIDTH, text: expectedText },
    )
}

async function nativeTextRunsInclude(page: Page, text: string): Promise<boolean> {
  return await page.evaluate(
    (needle) => Array.from(document.querySelectorAll('[data-native-text-run]')).some((run) => run.textContent?.includes(needle) ?? false),
    text,
  )
}

async function nativeTextRunTextAt(page: Page, columnIndex: number, rowIndex: number): Promise<string> {
  return await page.evaluate(
    ({ col, row }) =>
      document.querySelector(`[data-native-text-run-row="${String(row)}"][data-native-text-run-col="${String(col)}"]`)?.textContent ?? '',
    { col: columnIndex, row: rowIndex },
  )
}

async function sampleGreenFillPixelsAcrossFrames(
  page: Page,
  columnIndex: number,
  rowIndex: number,
  remainingSamples: number,
  samples: readonly number[] = [],
): Promise<readonly number[]> {
  return await sampleFillPixelsAcrossFrames(page, columnIndex, rowIndex, 'green', remainingSamples, samples)
}

async function sampleFillPixelsAcrossFrames(
  page: Page,
  columnIndex: number,
  rowIndex: number,
  color: 'blue' | 'green',
  remainingSamples: number,
  samples: readonly number[] = [],
): Promise<readonly number[]> {
  if (remainingSamples <= 0) {
    return samples
  }
  const pixels =
    color === 'green'
      ? await countGreenFillPixelsInCell(page, columnIndex, rowIndex)
      : await countBlueFillPixelsInCell(page, columnIndex, rowIndex)
  await page.waitForTimeout(50)
  return await sampleFillPixelsAcrossFrames(page, columnIndex, rowIndex, color, remainingSamples - 1, [...samples, pixels])
}

async function countInitialShellNonBlankPixels(page: Page): Promise<number> {
  const gridLocator = page.getByTestId('initial-workbook-grid')
  await expect(gridLocator).toBeVisible()
  const grid = await gridLocator.boundingBox()
  if (!grid) {
    throw new Error('initial workbook grid is not visible')
  }

  const buffer = await page.screenshot({
    animations: 'disabled',
    caret: 'hide',
    clip: {
      height: Math.min(420, Math.round(grid.height)),
      width: Math.min(620, Math.round(grid.width)),
      x: Math.round(grid.x),
      y: Math.round(grid.y),
    },
  })

  return await page.evaluate(
    async ({ dataUrl }) => {
      const image = await new Promise<HTMLImageElement>((resolve, reject) => {
        const element = new Image()
        element.addEventListener('load', () => resolve(element), { once: true })
        element.addEventListener('error', () => reject(new Error('Failed to decode initial shell screenshot')), { once: true })
        element.src = dataUrl
      })
      const canvas = document.createElement('canvas')
      canvas.width = image.naturalWidth
      canvas.height = image.naturalHeight
      const context = canvas.getContext('2d')
      if (!context) {
        throw new Error('Missing 2d context for initial shell screenshot analysis')
      }
      context.drawImage(image, 0, 0)
      const pixels = context.getImageData(0, 0, canvas.width, canvas.height).data
      let nonBlankPixels = 0
      for (let index = 0; index < pixels.length; index += 4) {
        const alpha = pixels[index + 3] ?? 0
        const red = pixels[index] ?? 255
        const green = pixels[index + 1] ?? 255
        const blue = pixels[index + 2] ?? 255
        if (alpha > 220 && (red < 246 || green < 246 || blue < 246)) {
          nonBlankPixels += 1
        }
      }
      return nonBlankPixels
    },
    { dataUrl: `data:image/png;base64,${buffer.toString('base64')}` },
  )
}

async function countDarkInteriorPixelsInCell(page: Page, columnIndex: number, rowIndex: number): Promise<number> {
  const gridLocator = page.getByTestId('sheet-grid')
  await expect(gridLocator).toBeVisible()
  const grid = await gridLocator.boundingBox()
  if (!grid) {
    throw new Error('sheet grid is not visible')
  }

  const [columnLeft, columnWidth, rowTop, rowHeight, scroll] = await Promise.all([
    getProductColumnLeft(page, columnIndex),
    getProductColumnWidth(page, columnIndex),
    getProductRowTop(page, rowIndex),
    getProductRowHeight(page, rowIndex),
    page.getByTestId('grid-scroll-viewport').evaluate((node) => ({
      scrollLeft: node.scrollLeft,
      scrollTop: node.scrollTop,
    })),
  ])
  const buffer = await page.screenshot({
    animations: 'disabled',
    caret: 'hide',
    clip: {
      x: Math.round(grid.x + columnLeft - scroll.scrollLeft + 4),
      y: Math.round(grid.y + PRODUCT_HEADER_HEIGHT + rowTop - scroll.scrollTop + 4),
      width: Math.max(1, Math.round(columnWidth - 12)),
      height: Math.max(1, Math.round(rowHeight - 8)),
    },
  })

  return await page.evaluate(
    async ({ dataUrl }) => {
      const image = await new Promise<HTMLImageElement>((resolve, reject) => {
        const element = new Image()
        element.addEventListener('load', () => resolve(element), { once: true })
        element.addEventListener('error', () => reject(new Error('Failed to decode cell screenshot')), { once: true })
        element.src = dataUrl
      })
      const canvas = document.createElement('canvas')
      canvas.width = image.naturalWidth
      canvas.height = image.naturalHeight
      const context = canvas.getContext('2d')
      if (!context) {
        throw new Error('Missing 2d context for cell screenshot analysis')
      }
      context.drawImage(image, 0, 0)
      const pixels = context.getImageData(0, 0, canvas.width, canvas.height).data
      let darkPixels = 0
      for (let index = 0; index < pixels.length; index += 4) {
        const alpha = pixels[index + 3] ?? 0
        const red = pixels[index] ?? 255
        const green = pixels[index + 1] ?? 255
        const blue = pixels[index + 2] ?? 255
        if (alpha > 200 && red < 120 && green < 120 && blue < 120) {
          darkPixels += 1
        }
      }
      return darkPixels
    },
    { dataUrl: `data:image/png;base64,${buffer.toString('base64')}` },
  )
}

async function countVisibleGridLinePixels(page: Page): Promise<number> {
  const gridLocator = page.getByTestId('sheet-grid')
  await expect(gridLocator).toBeVisible()
  const grid = await gridLocator.boundingBox()
  if (!grid) {
    throw new Error('sheet grid is not visible')
  }

  const buffer = await page.screenshot({
    animations: 'disabled',
    caret: 'hide',
    clip: {
      x: Math.round(grid.x + PRODUCT_ROW_MARKER_WIDTH + 1),
      y: Math.round(grid.y + PRODUCT_HEADER_HEIGHT + 1),
      width: Math.max(1, Math.round(Math.min(grid.width - PRODUCT_ROW_MARKER_WIDTH - 2, 540))),
      height: Math.max(1, Math.round(Math.min(grid.height - PRODUCT_HEADER_HEIGHT - 2, 420))),
    },
  })

  return await page.evaluate(
    async ({ dataUrl }) => {
      const image = await new Promise<HTMLImageElement>((resolve, reject) => {
        const element = new Image()
        element.addEventListener('load', () => resolve(element), { once: true })
        element.addEventListener('error', () => reject(new Error('Failed to decode grid screenshot')), { once: true })
        element.src = dataUrl
      })
      const canvas = document.createElement('canvas')
      canvas.width = image.naturalWidth
      canvas.height = image.naturalHeight
      const context = canvas.getContext('2d')
      if (!context) {
        throw new Error('Missing 2d context for grid screenshot analysis')
      }
      context.drawImage(image, 0, 0)
      const pixels = context.getImageData(0, 0, canvas.width, canvas.height).data
      let gridlinePixels = 0
      for (let index = 0; index < pixels.length; index += 4) {
        const alpha = pixels[index + 3] ?? 0
        const red = pixels[index] ?? 255
        const green = pixels[index + 1] ?? 255
        const blue = pixels[index + 2] ?? 255
        if (alpha > 200 && red >= 205 && red <= 235 && green >= 198 && green <= 228 && blue >= 186 && blue <= 220) {
          gridlinePixels += 1
        }
      }
      return gridlinePixels
    },
    { dataUrl: `data:image/png;base64,${buffer.toString('base64')}` },
  )
}
