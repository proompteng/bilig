// @vitest-environment jsdom
import { act } from 'react'
import { createRoot } from 'react-dom/client'
import { afterEach, beforeEach, describe, expect, test } from 'vitest'
import { WorkbookPaneRendererV2 } from '../renderer-v2/WorkbookPaneRendererV2.js'

describe('WorkbookPaneRendererV2', () => {
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

  test('mounts a distinct V2 canvas for the hard migration path', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    const host = document.createElement('div')
    Object.defineProperty(host, 'clientWidth', { configurable: true, value: 640 })
    Object.defineProperty(host, 'clientHeight', { configurable: true, value: 360 })
    const root = createRoot(host)
    const rendererHost = document.createElement('div')
    Object.defineProperty(rendererHost, 'clientWidth', { configurable: true, value: 640 })
    Object.defineProperty(rendererHost, 'clientHeight', { configurable: true, value: 360 })
    host.appendChild(rendererHost)

    await act(async () => {
      root.render(
        <WorkbookPaneRendererV2
          active
          host={rendererHost}
          geometry={null}
          panes={[
            {
              contentOffset: { x: 0, y: 0 },
              frame: { x: 46, y: 24, width: 594, height: 336 },
              generation: 1,
              gpuScene: { borderRects: [], fillRects: [] },
              paneId: 'body',
              scrollAxes: { x: true, y: true },
              surfaceSize: { width: 640, height: 360 },
              textScene: { items: [] },
            },
          ]}
        />,
      )
    })

    const canvas = host.querySelector('[data-testid="grid-pane-renderer"]')
    expect(canvas).toBeInstanceOf(HTMLCanvasElement)
    expect(canvas?.getAttribute('data-pane-renderer')).toBe('workbook-pane-renderer-v2')

    await act(async () => {
      root.unmount()
    })
  })
})
