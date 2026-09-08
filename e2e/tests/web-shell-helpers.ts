import { createHash } from 'node:crypto'
import { writeFile } from 'node:fs/promises'
import { expect, type Locator, type Page, type TestInfo } from '@playwright/test'
import {
  PRODUCT_HEADER_HEIGHT,
  PRODUCT_ROW_HEIGHT,
  PRODUCT_ROW_MARKER_WIDTH,
  clickProductCell,
  delay,
  getProductColumnLeft,
  getProductColumnWidth,
  getProductRowHeight,
  getProductRowTop,
} from './web-shell-product-grid-helpers.js'

export {
  PRODUCT_COLUMN_WIDTH,
  PRODUCT_HEADER_HEIGHT,
  PRODUCT_ROW_HEIGHT,
  PRODUCT_ROW_MARKER_WIDTH,
  PRIMARY_MODIFIER,
  clickProductCell,
  createTestDocumentId,
  getBox,
  getProductColumnLeft,
  getProductColumnWidth,
  getProductRowHeight,
  getProductRowTop,
  remoteSyncEnabled,
  waitForProductColumnWidthChange,
} from './web-shell-product-grid-helpers.js'

export {
  countBlueFillPixelsInCell,
  countBlueFillReadbackPixelsInCell,
  countDarkReadbackPixelsInCell,
  countGreenFillPixelsInCell,
  countGreenFillReadbackPixelsInCell,
  installTypeGpuCellReadbackHarness,
  sampleTypeGpuAveragePixelsInViewportRegion,
} from './web-shell-typegpu-readback-helpers.js'

interface ToolbarSyncAction {
  readonly label: string
  readonly apply: (page: Page) => Promise<void>
}

async function getProductColumnResizePoint(page: Page, columnIndex: number): Promise<{ readonly x: number; readonly y: number }> {
  const gridLocator = page.getByTestId('sheet-grid')
  await expect(gridLocator).toBeVisible()
  const grid = await gridLocator.boundingBox()
  if (!grid) {
    throw new Error('sheet grid is not visible')
  }

  const columnLeft = await getProductColumnLeft(page, columnIndex)
  const columnWidth = await getProductColumnWidth(page, columnIndex)
  return {
    x: grid.x + columnLeft + columnWidth - 1,
    y: grid.y + Math.floor(PRODUCT_HEADER_HEIGHT / 2),
  }
}

async function getProductRowResizePoint(page: Page, rowIndex: number): Promise<{ readonly x: number; readonly y: number }> {
  const gridLocator = page.getByTestId('sheet-grid')
  await expect(gridLocator).toBeVisible()
  const grid = await gridLocator.boundingBox()
  if (!grid) {
    throw new Error('sheet grid is not visible')
  }

  const rowTop = await getProductRowTop(page, rowIndex)
  const rowHeight = await getProductRowHeight(page, rowIndex)
  return {
    x: grid.x + Math.floor(PRODUCT_ROW_MARKER_WIDTH / 2),
    y: grid.y + PRODUCT_HEADER_HEIGHT + rowTop + rowHeight - 1,
  }
}

async function hoverProductColumnResizeHandle(page: Page, columnIndex: number): Promise<{ readonly x: number; readonly y: number }> {
  const point = await getProductColumnResizePoint(page, columnIndex)
  const gridLocator = page.getByTestId('sheet-grid')
  await page.mouse.move(point.x, point.y)
  await expect
    .poll(async () => gridLocator.evaluate((node) => window.getComputedStyle(node).cursor), {
      message: `column ${columnIndex} resize handle is hoverable`,
    })
    .toBe('col-resize')
  return point
}

export async function dragProductColumnResize(page: Page, columnIndex: number, deltaX: number) {
  const { x: edgeX, y: edgeY } = await hoverProductColumnResizeHandle(page, columnIndex)

  await page.mouse.down()
  await page.mouse.move(edgeX + deltaX, edgeY, { steps: 10 })
  await page.mouse.up()
}

export async function hoverProductRowResizeHandle(page: Page, rowIndex: number): Promise<{ readonly x: number; readonly y: number }> {
  const point = await getProductRowResizePoint(page, rowIndex)
  const gridLocator = page.getByTestId('sheet-grid')
  await page.mouse.move(point.x, point.y)
  await expect
    .poll(async () => gridLocator.evaluate((node) => window.getComputedStyle(node).cursor), {
      message: `row ${rowIndex} resize handle is hoverable`,
    })
    .toBe('row-resize')
  return point
}

export async function dragProductRowResize(page: Page, rowIndex: number, deltaY: number) {
  const { x: edgeX, y: edgeY } = await hoverProductRowResizeHandle(page, rowIndex)

  await page.mouse.down()
  await page.mouse.move(edgeX, edgeY + deltaY, { steps: 10 })
  await page.mouse.up()
}

export async function getProductFillHandleDragPoints(
  page: Page,
  sourceCol: number,
  sourceRow: number,
  targetCol: number,
  targetRow: number,
): Promise<{
  readonly sourceX: number
  readonly sourceY: number
  readonly targetX: number
  readonly targetY: number
}> {
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

export async function doubleClickProductColumnResizeHandle(page: Page, columnIndex: number) {
  const { x: edgeX, y: headerY } = await hoverProductColumnResizeHandle(page, columnIndex)
  await page.mouse.dblclick(edgeX, headerY)
}

export async function dragProductHeaderSelection(page: Page, axis: 'column' | 'row', startIndex: number, endIndex: number) {
  const gridLocator = page.getByTestId('sheet-grid')
  await expect(gridLocator).toBeVisible()
  const grid = await gridLocator.boundingBox()
  if (!grid) {
    throw new Error('sheet grid is not visible')
  }

  const startColumnLeft = axis === 'column' ? await getProductColumnLeft(page, startIndex) : 0
  const endColumnLeft = axis === 'column' ? await getProductColumnLeft(page, endIndex) : 0
  const startColumnWidth = axis === 'column' ? await getProductColumnWidth(page, startIndex) : 0
  const endColumnWidth = axis === 'column' ? await getProductColumnWidth(page, endIndex) : 0
  const startX =
    axis === 'column' ? grid.x + startColumnLeft + Math.floor(startColumnWidth / 2) : grid.x + Math.floor(PRODUCT_ROW_MARKER_WIDTH / 2)
  const startY =
    axis === 'column'
      ? grid.y + Math.floor(PRODUCT_HEADER_HEIGHT / 2)
      : grid.y + PRODUCT_HEADER_HEIGHT + startIndex * PRODUCT_ROW_HEIGHT + Math.floor(PRODUCT_ROW_HEIGHT / 2)
  const endX =
    axis === 'column' ? grid.x + endColumnLeft + Math.floor(endColumnWidth / 2) : grid.x + Math.floor(PRODUCT_ROW_MARKER_WIDTH / 2)
  const endY =
    axis === 'column'
      ? grid.y + Math.floor(PRODUCT_HEADER_HEIGHT / 2)
      : grid.y + PRODUCT_HEADER_HEIGHT + endIndex * PRODUCT_ROW_HEIGHT + Math.floor(PRODUCT_ROW_HEIGHT / 2)

  await page.mouse.move(startX, startY)
  await page.mouse.down()
  await page.mouse.move(endX, endY, { steps: 8 })
  await page.mouse.up()
}

export async function clickGridRightEdge(page: Page, rowIndex = 2) {
  const gridLocator = page.getByTestId('sheet-grid')
  await expect(gridLocator).toBeVisible()
  const grid = await gridLocator.boundingBox()
  if (!grid) {
    throw new Error('sheet grid is not visible')
  }

  const x = grid.x + grid.width - 3
  const y = grid.y + PRODUCT_HEADER_HEIGHT + rowIndex * PRODUCT_ROW_HEIGHT + Math.floor(PRODUCT_ROW_HEIGHT / 2)
  await page.mouse.click(x, y)
}

export function getToolbarButton(page: Page, label: string): Locator {
  return page.getByRole('button', { name: label, exact: true })
}

function getToolbarSelect(page: Page, label: string): Locator {
  return page.getByRole('combobox', { name: label, exact: true })
}

export async function expectToolbarColor(locator: Locator, value: string) {
  await expect(locator).toHaveAttribute('data-current-color', value.toLowerCase())
}

async function expectToolbarSelectValue(page: Page, label: string, value: string) {
  await expect(getToolbarSelect(page, label)).toHaveAttribute('data-current-value', value)
}

async function selectToolbarOption(page: Page, label: string, optionLabel: string, expectedValue = optionLabel) {
  await getToolbarSelect(page, label).click()
  await page.getByRole('option', { name: optionLabel, exact: true }).click()
  await expectToolbarSelectValue(page, label, expectedValue)
}

async function setColorInput(locator: Locator, value: string) {
  await locator.evaluate((element, nextValue) => {
    if (!(element instanceof HTMLInputElement)) {
      throw new Error('color input is not an HTMLInputElement')
    }
    const input = element
    const descriptor = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')
    descriptor?.set?.call(input, String(nextValue))
    input.dispatchEvent(new Event('input', { bubbles: true }))
    input.dispatchEvent(new Event('change', { bubbles: true }))
  }, value)
}

async function setToolbarCustomColor(page: Page, controlLabel: 'Fill color' | 'Text color', value: string) {
  await getToolbarButton(page, controlLabel).click()
  await page.getByLabel(`Open custom ${controlLabel.toLowerCase()} picker`).click()
  await setColorInput(
    page.getByLabel(controlLabel === 'Fill color' ? 'Custom fill color' : 'Custom text color', {
      exact: true,
    }),
    value,
  )
}

export async function pickToolbarPresetColor(page: Page, controlLabel: 'Fill color' | 'Text color', swatchLabel: string) {
  await getToolbarButton(page, controlLabel).click()
  await page.getByLabel(`${controlLabel} ${swatchLabel}`).click()
}

export async function pickToolbarBorderPreset(
  page: Page,
  presetLabel: 'All borders' | 'Outer borders' | 'Left border' | 'Top border' | 'Right border' | 'Bottom border' | 'Clear borders',
) {
  await getToolbarButton(page, 'Borders').click()
  await page.getByRole('button', { name: presetLabel }).click()
}

export const TOOLBAR_SYNC_ACTIONS: readonly ToolbarSyncAction[] = [
  {
    label: 'number-format-accounting',
    apply: async (activePage) => await selectToolbarOption(activePage, 'Number format', 'Accounting', 'accounting'),
  },
  {
    label: 'font-size-14',
    apply: async (activePage) => await selectToolbarOption(activePage, 'Font size', '14'),
  },
  { label: 'bold', apply: async (activePage) => await activePage.getByLabel('Bold').click() },
  {
    label: 'italic',
    apply: async (activePage) => await activePage.getByLabel('Italic').click(),
  },
  {
    label: 'underline',
    apply: async (activePage) => await activePage.getByLabel('Underline').click(),
  },
  {
    label: 'fill-color',
    apply: async (activePage) => await setToolbarCustomColor(activePage, 'Fill color', '#dbeafe'),
  },
  {
    label: 'text-color',
    apply: async (activePage) => await setToolbarCustomColor(activePage, 'Text color', '#7c2d12'),
  },
  {
    label: 'align-left',
    apply: async (activePage) => await activePage.getByLabel('Align left').click(),
  },
  {
    label: 'align-center',
    apply: async (activePage) => await activePage.getByLabel('Align center').click(),
  },
  {
    label: 'align-right',
    apply: async (activePage) => await activePage.getByLabel('Align right').click(),
  },
  {
    label: 'border-all',
    apply: async (activePage) => await pickToolbarBorderPreset(activePage, 'All borders'),
  },
  {
    label: 'border-outer',
    apply: async (activePage) => await pickToolbarBorderPreset(activePage, 'Outer borders'),
  },
  {
    label: 'border-left',
    apply: async (activePage) => await pickToolbarBorderPreset(activePage, 'Left border'),
  },
  {
    label: 'border-top',
    apply: async (activePage) => await pickToolbarBorderPreset(activePage, 'Top border'),
  },
  {
    label: 'border-right',
    apply: async (activePage) => await pickToolbarBorderPreset(activePage, 'Right border'),
  },
  {
    label: 'border-bottom',
    apply: async (activePage) => await pickToolbarBorderPreset(activePage, 'Bottom border'),
  },
  {
    label: 'border-clear',
    apply: async (activePage) => await pickToolbarBorderPreset(activePage, 'Clear borders'),
  },
  { label: 'wrap', apply: async (activePage) => await activePage.getByLabel('Wrap').click() },
  {
    label: 'clear-style',
    apply: async (activePage) => await activePage.getByLabel('Clear style').click(),
  },
  {
    label: 'number-format-general',
    apply: async (activePage) => await selectToolbarOption(activePage, 'Number format', 'General', 'general'),
  },
]

export async function selectToolbarActionRange(page: Page) {
  await clickProductCell(page, 1, 1)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!B2')
  await clickProductCell(page, 2, 2, { shift: true })
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!B2:C3')
}

export async function clickProductBodyOffset(page: Page, offsetX: number, rowIndex = 0) {
  const gridLocator = page.getByTestId('sheet-grid')
  await expect(gridLocator).toBeVisible()
  const grid = await gridLocator.boundingBox()
  if (!grid) {
    throw new Error('sheet grid is not visible')
  }

  await page.mouse.click(
    grid.x + PRODUCT_ROW_MARKER_WIDTH + offsetX,
    grid.y + PRODUCT_HEADER_HEIGHT + rowIndex * PRODUCT_ROW_HEIGHT + Math.floor(PRODUCT_ROW_HEIGHT / 2),
  )
}

export async function clickProductCellUpperHalf(page: Page, columnIndex: number, rowIndex: number) {
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
    grid.y + PRODUCT_HEADER_HEIGHT + rowIndex * PRODUCT_ROW_HEIGHT + 4,
  )
}

export async function dragProductBodySelection(page: Page, startColumn: number, startRow: number, endColumn: number, endRow: number) {
  const gridLocator = page.getByTestId('sheet-grid')
  await expect(gridLocator).toBeVisible()
  const grid = await gridLocator.boundingBox()
  if (!grid) {
    throw new Error('sheet grid is not visible')
  }

  const startLeft = await getProductColumnLeft(page, startColumn)
  const startWidth = await getProductColumnWidth(page, startColumn)
  const endLeft = await getProductColumnLeft(page, endColumn)
  const endWidth = await getProductColumnWidth(page, endColumn)

  const startX = grid.x + startLeft + Math.floor(startWidth / 2)
  const startY = grid.y + PRODUCT_HEADER_HEIGHT + startRow * PRODUCT_ROW_HEIGHT + Math.floor(PRODUCT_ROW_HEIGHT / 2)
  const endX = grid.x + endLeft + Math.floor(endWidth / 2)
  const endY = grid.y + PRODUCT_HEADER_HEIGHT + endRow * PRODUCT_ROW_HEIGHT + Math.floor(PRODUCT_ROW_HEIGHT / 2)
  await page.mouse.move(startX, startY)
  await page.mouse.down()
  await page.mouse.move(endX, endY, { steps: 8 })
  await page.mouse.up()
}

export async function seedToolbarActionRange(page: Page) {
  const nameBox = page.getByTestId('name-box')
  const formulaInput = page.getByTestId('formula-input')
  await nameBox.fill('B2')
  await nameBox.press('Enter')
  await formulaInput.fill('1234.5')
  await formulaInput.press('Enter')

  await nameBox.fill('C2')
  await nameBox.press('Enter')
  await formulaInput.fill('6789.125')
  await formulaInput.press('Enter')

  await nameBox.fill('B3')
  await nameBox.press('Enter')
  await formulaInput.fill('42.25')
  await formulaInput.press('Enter')

  await nameBox.fill('C3')
  await nameBox.press('Enter')
  await formulaInput.fill('-7.5')
  await formulaInput.press('Enter')
}

async function captureGridRangeScreenshot(page: Page, startColumn: number, startRow: number, endColumn: number, endRow: number) {
  await page.bringToFront()
  const gridLocator = page.getByTestId('sheet-grid')
  await expect(gridLocator).toBeVisible()
  const grid = await gridLocator.boundingBox()
  if (!grid) {
    throw new Error('sheet grid is not visible')
  }

  const minColumn = Math.min(startColumn, endColumn)
  const maxColumn = Math.max(startColumn, endColumn)
  const minRow = Math.min(startRow, endRow)
  const maxRow = Math.max(startRow, endRow)
  const startLeft = await getProductColumnLeft(page, minColumn)
  const endLeft = await getProductColumnLeft(page, maxColumn)
  const endWidth = await getProductColumnWidth(page, maxColumn)

  return await page.screenshot({
    animations: 'disabled',
    caret: 'hide',
    clip: {
      x: Math.round(grid.x + startLeft),
      y: Math.round(grid.y + PRODUCT_HEADER_HEIGHT + minRow * PRODUCT_ROW_HEIGHT),
      width: Math.round(endLeft + endWidth - startLeft),
      height: Math.round((maxRow - minRow + 1) * PRODUCT_ROW_HEIGHT),
    },
  })
}

async function compareScreenshotPixels(page: Page, left: Buffer, right: Buffer) {
  return await page.evaluate(
    async ({ leftDataUrl, rightDataUrl, channelTolerance }) => {
      const [leftImage, rightImage] = await Promise.all([
        new Promise<HTMLImageElement>((resolve, reject) => {
          const image = new Image()
          image.addEventListener('load', () => resolve(image), { once: true })
          image.addEventListener('error', () => reject(new Error('Failed to decode left screenshot data URL')), { once: true })
          image.src = leftDataUrl
        }),
        new Promise<HTMLImageElement>((resolve, reject) => {
          const image = new Image()
          image.addEventListener('load', () => resolve(image), { once: true })
          image.addEventListener('error', () => reject(new Error('Failed to decode right screenshot data URL')), { once: true })
          image.src = rightDataUrl
        }),
      ])
      if (leftImage.naturalWidth !== rightImage.naturalWidth || leftImage.naturalHeight !== rightImage.naturalHeight) {
        return {
          equal: false,
          diffPixels: Number.POSITIVE_INFINITY,
          width: leftImage.naturalWidth,
          height: leftImage.naturalHeight,
        }
      }

      const width = leftImage.naturalWidth
      const height = leftImage.naturalHeight
      const canvas = document.createElement('canvas')
      canvas.width = width
      canvas.height = height
      const context = canvas.getContext('2d')
      if (!context) {
        throw new Error('Missing 2d context for screenshot comparison')
      }

      context.clearRect(0, 0, width, height)
      context.drawImage(leftImage, 0, 0)
      const leftPixels = context.getImageData(0, 0, width, height).data
      context.clearRect(0, 0, width, height)
      context.drawImage(rightImage, 0, 0)
      const rightPixels = context.getImageData(0, 0, width, height).data

      let diffPixels = 0
      for (let index = 0; index < leftPixels.length; index += 4) {
        if (
          Math.abs(leftPixels[index] - rightPixels[index]) > channelTolerance ||
          Math.abs(leftPixels[index + 1] - rightPixels[index + 1]) > channelTolerance ||
          Math.abs(leftPixels[index + 2] - rightPixels[index + 2]) > channelTolerance ||
          Math.abs(leftPixels[index + 3] - rightPixels[index + 3]) > channelTolerance
        ) {
          diffPixels += 1
        }
      }

      return { equal: diffPixels === 0, diffPixels, width, height }
    },
    {
      leftDataUrl: `data:image/png;base64,${left.toString('base64')}`,
      rightDataUrl: `data:image/png;base64,${right.toString('base64')}`,
      channelTolerance: 2,
    },
  )
}

async function pollMatchingGridRangeScreenshots(
  primaryPage: Page,
  mirrorPage: Page,
  startedAt: number,
  timeoutMs: number,
  maxDiffPixels: number,
  startColumn: number,
  startRow: number,
  endColumn: number,
  endRow: number,
): Promise<{
  primaryBuffer: Buffer
  mirrorBuffer: Buffer
  diffPixels: number
  matched: boolean
}> {
  const [primaryBuffer, mirrorBuffer] = await Promise.all([
    captureGridRangeScreenshot(primaryPage, startColumn, startRow, endColumn, endRow),
    captureGridRangeScreenshot(mirrorPage, startColumn, startRow, endColumn, endRow),
  ])
  const comparison = await compareScreenshotPixels(primaryPage, primaryBuffer, mirrorBuffer)
  const matched = comparison.equal || comparison.diffPixels <= maxDiffPixels
  if (matched || Date.now() - startedAt > timeoutMs) {
    return {
      primaryBuffer,
      mirrorBuffer,
      diffPixels: comparison.diffPixels,
      matched,
    }
  }

  await delay(50)
  return await pollMatchingGridRangeScreenshots(
    primaryPage,
    mirrorPage,
    startedAt,
    timeoutMs,
    maxDiffPixels,
    startColumn,
    startRow,
    endColumn,
    endRow,
  )
}

export async function expectMatchingGridRangeScreenshots(
  primaryPage: Page,
  mirrorPage: Page,
  actionLabel: string,
  testInfo: TestInfo,
  startColumn: number,
  startRow: number,
  endColumn: number,
  endRow: number,
  timeoutMs = 1_500,
  maxDiffPixels = 8,
) {
  const startedAt = Date.now()
  const result = await pollMatchingGridRangeScreenshots(
    primaryPage,
    mirrorPage,
    startedAt,
    timeoutMs,
    maxDiffPixels,
    startColumn,
    startRow,
    endColumn,
    endRow,
  )
  if (result.matched) {
    return Date.now() - startedAt
  }

  const primaryHash = createHash('sha256').update(result.primaryBuffer).digest('hex')
  const mirrorHash = createHash('sha256').update(result.mirrorBuffer).digest('hex')
  await writeFile(testInfo.outputPath(`multiplayer-${actionLabel}-range-primary.png`), result.primaryBuffer)
  await writeFile(testInfo.outputPath(`multiplayer-${actionLabel}-range-mirror.png`), result.mirrorBuffer)

  throw new Error(
    `multiplayer grid range screenshots diverged for ${actionLabel} after ${timeoutMs}ms (primary=${primaryHash}, mirror=${mirrorHash}, diffPixels=${result.diffPixels}, maxDiffPixels=${maxDiffPixels})`,
  )
}

async function retryRuntimeConfigLoadIfVisible(page: Page): Promise<boolean> {
  const retryButton = page.getByRole('button', { name: 'Retry' })
  if (!(await retryButton.isVisible({ timeout: 250 }).catch(() => false))) {
    return false
  }
  const runtimeConfigError = page.getByText(/^Failed to load runtime config \(\d+\)$/)
  if (!(await runtimeConfigError.isVisible({ timeout: 250 }).catch(() => false))) {
    return false
  }
  await retryButton.click()
  return true
}

export async function waitForWorkbookReady(page: Page) {
  await retryRuntimeConfigLoadIfVisible(page)
  try {
    await waitForWorkbookReadyOnce(page)
    return
  } catch (error) {
    if (!(await retryRuntimeConfigLoadIfVisible(page))) {
      throw error
    }
  }
  await waitForWorkbookReadyOnce(page)
}

export async function waitForTypeGpuVisibleFrame(page: Page): Promise<void> {
  await expect
    .poll(
      async () =>
        await page.getByTestId('grid-pane-renderer').evaluate((node) => {
          if (!(node instanceof HTMLElement)) {
            return 'missing-renderer'
          }
          const backendStatus = node.getAttribute('data-v3-backend-status') ?? ''
          if (backendStatus !== 'ready') {
            return `backend:${backendStatus}`
          }
          const frameProofStatus = node.getAttribute('data-v3-frame-proof-status') ?? ''
          if (frameProofStatus !== 'presented') {
            return `frame:${frameProofStatus}`
          }
          if (node.getAttribute('data-v3-has-presented-visible-frame') !== 'true') {
            return 'visible-frame:pending'
          }
          const currentSelection = node.getAttribute('data-v3-current-selection-revision') ?? ''
          const presentedSelection = node.getAttribute('data-v3-presented-selection-revision') ?? ''
          if (currentSelection.length > 0 && currentSelection !== presentedSelection) {
            return 'selection:pending'
          }
          const currentViewport = node.getAttribute('data-v3-current-viewport-revision') ?? ''
          const presentedViewport = node.getAttribute('data-v3-presented-viewport-revision') ?? ''
          if (currentViewport.length > 0 && currentViewport !== presentedViewport) {
            return 'viewport:pending'
          }
          return 'ready'
        }),
      {
        message: 'TypeGPU visible frame should be presented before screenshot pixel assertions',
        timeout: 5_000,
      },
    )
    .toBe('ready')
}

async function waitForWorkbookReadyOnce(page: Page): Promise<void> {
  await expect(page.getByTestId('formula-bar')).toBeVisible({ timeout: 15_000 })
  await expect(page.getByTestId('sheet-grid')).toBeVisible({ timeout: 15_000 })
  await expect(page.getByTestId('status-sync')).toHaveText(
    /^(Saved|Saving…|Sync pending|Local saved|Local only|Offline|Sync issue|Read only)$/,
    {
      timeout: 15_000,
    },
  )
}

export async function gotoWorkbookShell(page: Page, path = '/workbook', timeoutMs = 15_000) {
  async function attempt(deadline: number, lastError: unknown): Promise<void> {
    if (Date.now() >= deadline) {
      throw lastError instanceof Error ? lastError : new Error(`Timed out navigating to ${path}`)
    }

    try {
      await page.goto(path)
      return
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error)
      if (!message.includes('ERR_CONNECTION_REFUSED')) {
        throw error
      }
      await delay(250)
      await attempt(deadline, error)
    }
  }

  await attempt(Date.now() + timeoutMs, null)
}

export async function openZeroWorkbookPage(page: Page, documentId: string) {
  await gotoWorkbookShell(page, `/?document=${encodeURIComponent(documentId)}`)
  await waitForWorkbookReady(page)
  await selectToolbarActionRange(page)
}

export {
  performDiagonalGridBrowse,
  performHorizontalGridBrowse,
  performVerticalGridBrowse,
  resetGridScroll,
  settleWorkbookScrollPerf,
  startWorkbookScrollPerf,
  stopWorkbookScrollPerf,
  waitForBenchmarkCorpus,
  warmStartWorkbookScrollPerf,
} from './web-shell-scroll-perf-helpers.js'

export async function runToolbarSyncActions(
  page: Page,
  mirrorPage: Page,
  actions: readonly ToolbarSyncAction[],
  testInfo: TestInfo,
  index = 0,
): Promise<void> {
  const action = actions[index]
  if (!action) {
    return
  }

  await action.apply(page)
  await selectToolbarActionRange(page)
  await selectToolbarActionRange(mirrorPage)
  const elapsed = await expectMatchingGridRangeScreenshots(page, mirrorPage, action.label, testInfo, 1, 1, 2, 2, 1_500)
  expect(elapsed).toBeLessThanOrEqual(1_500)
  await runToolbarSyncActions(page, mirrorPage, actions, testInfo, index + 1)
}
