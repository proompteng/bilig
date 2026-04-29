import { describe, expect, test, vi } from 'vitest'
import { drawTextRuns, type CanvasTextRunContext } from '../renderer-v3/WorkbookPaneCanvasFallbackV3.js'

function createCanvasContextMock(): {
  readonly context: CanvasTextRunContext
  readonly fillText: ReturnType<typeof vi.fn<(text: string, x: number, y: number, maxWidth?: number) => void>>
  readonly lineTo: ReturnType<typeof vi.fn<(x: number, y: number) => void>>
} {
  const fillText = vi.fn<(text: string, x: number, y: number, maxWidth?: number) => void>()
  const lineTo = vi.fn<(x: number, y: number) => void>()
  const context = {
    beginPath: vi.fn(),
    clip: vi.fn(),
    fillStyle: '',
    fillText,
    font: '',
    lineTo,
    lineWidth: 1,
    measureText: vi.fn(() => ({ width: 520 })),
    moveTo: vi.fn(),
    rect: vi.fn(),
    restore: vi.fn(),
    save: vi.fn(),
    stroke: vi.fn(),
    strokeStyle: '',
    textAlign: 'left' as CanvasTextAlign,
    textBaseline: 'middle' as CanvasTextBaseline,
  }
  return { context, fillText, lineTo }
}

describe('WorkbookPaneCanvasFallbackV3', () => {
  test('clips long text without using canvas maxWidth font compression', () => {
    const { context, fillText } = createCanvasContextMock()

    drawTextRuns(context, [
      {
        align: 'left',
        clipHeight: 22,
        clipWidth: 180,
        clipX: 10,
        clipY: 20,
        color: '#111827',
        font: '11px system-ui, sans-serif',
        fontSize: 11,
        height: 22,
        strike: false,
        text: 'Amortization schedule note that must clip instead of squeezing into one cell',
        underline: false,
        width: 180,
        x: 10,
        y: 20,
      },
    ])

    expect(fillText).toHaveBeenCalledWith('Amortization schedule note that must clip instead of squeezing into one cell', 16, 31)
    expect(fillText.mock.calls[0]).toHaveLength(3)
  })

  test('bounds fallback text decoration to the clip rectangle', () => {
    const { context, lineTo } = createCanvasContextMock()

    drawTextRuns(context, [
      {
        align: 'left',
        clipHeight: 22,
        clipWidth: 80,
        clipX: 10,
        clipY: 20,
        color: '#111827',
        font: '11px system-ui, sans-serif',
        fontSize: 11,
        height: 22,
        strike: false,
        text: 'Underlined text that is far wider than the visible cell',
        underline: true,
        width: 500,
        x: 10,
        y: 20,
      },
    ])

    expect(lineTo).toHaveBeenCalledWith(96, 38)
  })
})
