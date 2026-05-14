import { expect, test, type Page } from '@playwright/test'
import {
  PRODUCT_HEADER_HEIGHT,
  createTestDocumentId,
  getProductColumnLeft,
  getProductColumnWidth,
  getProductRowHeight,
  getProductRowTop,
  waitForWorkbookReady,
} from './web-shell-helpers.js'

test('web app paints deep querystring-selected cell content in the visible grid', async ({ page }) => {
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
    .poll(() => countDarkInteriorPixelsInCell(page, 3, 52), {
      message: 'D53 should paint visible text pixels after the edit commits',
      timeout: 5_000,
    })
    .toBeGreaterThan(12)
})

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
      const textOverlay = document.querySelector('[data-testid="grid-pane-text-overlay"]')
      return {
        fallbackMounted: fallback instanceof HTMLCanvasElement,
        textOverlayMounted: textOverlay instanceof HTMLElement,
        typeGpuMode: typeGpu instanceof HTMLCanvasElement ? typeGpu.getAttribute('data-renderer-mode') : null,
        typeGpuOpacity: typeGpu instanceof HTMLElement ? getComputedStyle(typeGpu).opacity : null,
      }
    })
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
