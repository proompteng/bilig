import { expect, test, type Page } from '@playwright/test'
import {
  PRODUCT_HEADER_HEIGHT,
  PRODUCT_ROW_MARKER_WIDTH,
  PRIMARY_MODIFIER,
  createTestDocumentId,
  getProductColumnLeft,
  getProductColumnWidth,
  getProductRowHeight,
  getProductRowTop,
  waitForWorkbookReady,
} from './web-shell-helpers.js'

const DEFAULT_WORKBOOK_CSS_FONT_SIZE = '13.333px'

test('web app paints deep querystring-selected cell content in the visible grid', async ({ page }, testInfo) => {
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
      message: 'TypeGPU should stay visible after its frame is presented; the Canvas2D fallback must not mask the grid',
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
      tileSceneRevisionPresent: true,
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

test('web app keeps a user click selection after opening from a deep cell URL', async ({ page }) => {
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

test('web app keeps dense accounting-sheet text payloads complete in the TypeGPU layer', async ({ page, context }, testInfo) => {
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
  await expect(page.getByTestId('sheet-grid')).toHaveCSS('font-family', /Arial|Helvetica/)
  await expect(page.getByTestId('sheet-grid')).toHaveCSS('font-size', DEFAULT_WORKBOOK_CSS_FONT_SIZE)
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

test('web app keeps the live cell editor above the native grid text layer', async ({ page }) => {
  const documentId = createTestDocumentId('playwright-editor-native-text-z-order')
  await page.setViewportSize({ width: 1166, height: 820 })
  await page.goto(`/?document=${encodeURIComponent(documentId)}&sheet=Sheet1&cell=A1`)
  await waitForWorkbookReady(page)

  await clickVisibleGridBodySlot(page, 1, 1)
  await page.keyboard.type('editor-z-order')

  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!B2')
  await expect(page.getByTestId('cell-editor-input')).toHaveValue('editor-z-order')
  await expect(page.getByTestId('cell-editor-input')).not.toHaveCSS('opacity', '0')
  await expect
    .poll(readEditorLayerState(page), {
      message: 'the active editor must sit above native rendered text to prevent double-text artifacts while typing or clicking away',
      timeout: 5_000,
    })
    .toMatchObject({
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
  readonly projectedRevisionPresent: boolean
  readonly tileSceneRevisionPresent: boolean
  readonly visibleRenderRevisionPresent: boolean
}> {
  return async () =>
    await page.evaluate(() => {
      const grid = document.querySelector('[data-testid="sheet-grid"]')
      const typeGpu = document.querySelector('[data-testid="grid-pane-renderer"]')
      const projectedRevision = grid?.getAttribute('data-render-projected-revision') ?? ''
      const tileSceneRevision = typeGpu?.getAttribute('data-v3-tile-scene-revision') ?? ''
      const visibleRenderRevision = typeGpu?.getAttribute('data-v3-visible-render-revision') ?? ''
      return {
        projectedRevisionPresent: projectedRevision.length > 0,
        tileSceneRevisionPresent: tileSceneRevision.length > 0,
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

function readEditorLayerState(page: Page): () => Promise<{
  readonly editorAboveNativeText: boolean
  readonly editorMounted: boolean
  readonly nativeTextLayerMounted: boolean
}> {
  return async () =>
    await page.evaluate(() => {
      const editor = document.querySelector('[data-testid="cell-editor-overlay"]')
      const nativeTextLayer = document.querySelector('[data-testid="grid-native-text-layer"]')
      const editorZIndex = editor instanceof HTMLElement ? Number(getComputedStyle(editor).zIndex) : Number.NaN
      const nativeTextLayerZIndex = nativeTextLayer instanceof HTMLElement ? Number(getComputedStyle(nativeTextLayer).zIndex) : Number.NaN
      return {
        editorAboveNativeText:
          Number.isFinite(editorZIndex) && Number.isFinite(nativeTextLayerZIndex) && editorZIndex > nativeTextLayerZIndex,
        editorMounted: editor instanceof HTMLElement,
        nativeTextLayerMounted: nativeTextLayer instanceof HTMLElement,
      }
    })
}

function readNativeTextLayerStabilityState(
  page: Page,
  expectedText: string,
): () => Promise<{
  readonly containsText: boolean
  readonly declaredRunCountMatchesDom: boolean
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
        const rowHeaderRunCount = textRuns.filter((run) => {
          if (!gridRect || !/^\d+$/.test(run.textContent ?? '')) {
            return false
          }
          const rect = run.getBoundingClientRect()
          return rect.left >= gridRect.left && rect.right <= gridRect.left + rowMarkerWidth + 1
        }).length
        return {
          containsText: textRuns.some((run) => run.textContent === text),
          declaredRunCountMatchesDom: declaredRunCount === textRuns.length,
          rowHeaderRunCount,
          rowHeaderRunCountHealthy: rowHeaderRunCount >= 10,
        }
      },
      { rowMarkerWidth: PRODUCT_ROW_MARKER_WIDTH, text: expectedText },
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
