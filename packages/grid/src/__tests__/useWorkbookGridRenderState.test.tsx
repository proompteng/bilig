// @vitest-environment jsdom
import { act } from 'react'
import { createRoot } from 'react-dom/client'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { ValueTag, type CellSnapshot, type CellStyleRecord } from '@bilig/protocol'
import { useWorkbookGridRenderState } from '../useWorkbookGridRenderState.js'

function createEmptySnapshot(sheetName: string, address: string): CellSnapshot {
  return {
    sheetName,
    address,
    value: { tag: ValueTag.Empty },
    flags: 0,
    version: 0,
  }
}

const engine = {
  workbook: {
    getSheet: () => undefined,
  },
  getCell(sheetName: string, address: string): CellSnapshot {
    return createEmptySnapshot(sheetName, address)
  },
  getCellStyle(_styleId: string | undefined): CellStyleRecord | undefined {
    return undefined
  },
  subscribeCells(): () => void {
    return () => undefined
  },
}

describe('useWorkbookGridRenderState viewport residency', () => {
  const originalResizeObserver = globalThis.ResizeObserver
  const originalRequestAnimationFrame = window.requestAnimationFrame
  const originalCancelAnimationFrame = window.cancelAnimationFrame

  beforeEach(() => {
    class TestResizeObserver {
      constructor(private readonly listener: ResizeObserverCallback) {}

      observe() {
        Reflect.apply(this.listener, undefined, [[], undefined])
      }

      disconnect() {}

      unobserve() {}
    }

    Object.defineProperty(globalThis, 'ResizeObserver', {
      configurable: true,
      value: TestResizeObserver,
      writable: true,
    })
    window.requestAnimationFrame = ((callback: FrameRequestCallback) => {
      const handle = window.setTimeout(() => callback(performance.now()), 0)
      return handle
    }) as typeof window.requestAnimationFrame
    window.cancelAnimationFrame = ((handle: number) => {
      window.clearTimeout(handle)
    }) as typeof window.cancelAnimationFrame
  })

  afterEach(() => {
    Object.defineProperty(globalThis, 'ResizeObserver', {
      configurable: true,
      value: originalResizeObserver,
      writable: true,
    })
    window.requestAnimationFrame = originalRequestAnimationFrame
    window.cancelAnimationFrame = originalCancelAnimationFrame
    document.body.innerHTML = ''
  })

  it('keeps the viewport subscription stable while horizontal scroll stays inside one resident window', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const subscribeViewport = vi.fn(() => () => undefined)
    let hostElement: HTMLDivElement | null = null
    let scrollViewport: HTMLDivElement | null = null

    function Harness() {
      const renderState = useWorkbookGridRenderState({
        engine,
        sheetName: 'Sheet1',
        selectedAddr: 'A1',
        selectedCellSnapshot: createEmptySnapshot('Sheet1', 'A1'),
        editorValue: '',
        isEditingCell: false,
        subscribeViewport,
      })

      return (
        <div
          ref={(node) => {
            renderState.handleHostRef(node)
            hostElement = node
          }}
        >
          <div
            ref={(node) => {
              renderState.scrollViewportRef.current = node
              scrollViewport = node
            }}
          />
        </div>
      )
    }

    const rootHost = document.createElement('div')
    document.body.appendChild(rootHost)
    const root = createRoot(rootHost)

    await act(async () => {
      root.render(<Harness />)
    })

    Object.defineProperty(hostElement!, 'clientWidth', { configurable: true, value: 480 })
    Object.defineProperty(hostElement!, 'clientHeight', { configurable: true, value: 180 })
    Object.defineProperty(scrollViewport!, 'clientWidth', { configurable: true, value: 480 })
    Object.defineProperty(scrollViewport!, 'clientHeight', { configurable: true, value: 180 })

    await act(async () => {
      scrollViewport!.dispatchEvent(new Event('scroll'))
      await new Promise((resolve) => window.setTimeout(resolve, 0))
    })

    const initialSubscriptionCount = subscribeViewport.mock.calls.length
    expect(initialSubscriptionCount).toBeGreaterThan(0)
    expect(subscribeViewport).toHaveBeenLastCalledWith(
      'Sheet1',
      expect.objectContaining({
        rowStart: 0,
        rowEnd: 31,
        colStart: 0,
        colEnd: 127,
      }),
      expect.any(Function),
    )

    await act(async () => {
      scrollViewport!.scrollLeft = 64 * 104
      scrollViewport!.dispatchEvent(new Event('scroll'))
      await new Promise((resolve) => window.setTimeout(resolve, 0))
    })

    expect(subscribeViewport).toHaveBeenCalledTimes(initialSubscriptionCount)

    await act(async () => {
      scrollViewport!.scrollLeft = 128 * 104
      scrollViewport!.dispatchEvent(new Event('scroll'))
      await new Promise((resolve) => window.setTimeout(resolve, 0))
    })

    expect(subscribeViewport).toHaveBeenCalledTimes(initialSubscriptionCount + 1)
    expect(subscribeViewport).toHaveBeenLastCalledWith(
      'Sheet1',
      expect.objectContaining({
        rowStart: 0,
        rowEnd: 31,
        colStart: 128,
        colEnd: 255,
      }),
      expect.any(Function),
    )

    await act(async () => {
      root.unmount()
    })
  })
})
