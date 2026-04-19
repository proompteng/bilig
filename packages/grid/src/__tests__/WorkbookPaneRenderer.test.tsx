// @vitest-environment jsdom
import { act } from 'react'
import { createRoot } from 'react-dom/client'
import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import { WorkbookPaneRenderer } from '../renderer/WorkbookPaneRenderer.js'

describe('WorkbookPaneRenderer', () => {
  const originalResizeObserver = globalThis.ResizeObserver

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
  })

  afterEach(() => {
    Object.defineProperty(globalThis, 'ResizeObserver', {
      configurable: true,
      value: originalResizeObserver,
      writable: true,
    })
    document.body.innerHTML = ''
  })

  it('mounts a unified pane canvas host for resident panes and overlay scenes', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const host = document.createElement('div')
    Object.defineProperty(host, 'clientWidth', { configurable: true, value: 640 })
    Object.defineProperty(host, 'clientHeight', { configurable: true, value: 360 })
    document.body.appendChild(host)

    const root = createRoot(host)

    await act(async () => {
      root.render(
        <WorkbookPaneRenderer
          active
          host={host}
          overlay={{ gpuScene: { fillRects: [], borderRects: [] }, textScene: { items: [] } }}
          panes={[
            {
              generation: 1,
              paneId: 'body',
              viewport: { rowStart: 0, rowEnd: 20, colStart: 0, colEnd: 10 },
              surfaceSize: { width: 640, height: 360 },
              frame: { x: 46, y: 24, width: 594, height: 336 },
              contentOffset: { x: 0, y: 0 },
              gpuScene: { fillRects: [], borderRects: [] },
              textScene: { items: [] },
            },
          ]}
        />,
      )
    })

    expect(host.querySelector('[data-testid="workbook-pane-renderer"]')).toBeInstanceOf(HTMLCanvasElement)

    await act(async () => {
      root.unmount()
    })
  })
})
