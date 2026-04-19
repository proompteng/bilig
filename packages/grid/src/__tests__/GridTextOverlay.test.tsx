// @vitest-environment jsdom
import { act } from 'react'
import { createRoot } from 'react-dom/client'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { GridTextOverlay } from '../GridTextOverlay.js'

describe('GridTextOverlay', () => {
  const originalResizeObserver = globalThis.ResizeObserver
  let getContextSpy: { mockRestore(): void } | null = null

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
    getContextSpy = vi.spyOn(HTMLCanvasElement.prototype, 'getContext').mockImplementation(() => ({
      beginPath() {},
      clearRect() {},
      clip() {},
      fillText() {},
      lineTo() {},
      measureText(text: string) {
        return { width: text.length * 8 }
      },
      moveTo() {},
      rect() {},
      restore() {},
      save() {},
      scale() {},
      setTransform() {},
      stroke() {},
    }))
  })

  afterEach(() => {
    Object.defineProperty(globalThis, 'ResizeObserver', {
      configurable: true,
      value: originalResizeObserver,
      writable: true,
    })
    getContextSpy?.mockRestore()
    getContextSpy = null
  })

  it('renders one canvas surface instead of one DOM node per text item', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    const host = document.createElement('div')
    Object.defineProperty(host, 'clientWidth', { configurable: true, value: 480 })
    Object.defineProperty(host, 'clientHeight', { configurable: true, value: 240 })
    document.body.appendChild(host)
    const root = createRoot(host)
    const gridHost = document.createElement('div')
    Object.defineProperty(gridHost, 'clientWidth', { configurable: true, value: 480 })
    Object.defineProperty(gridHost, 'clientHeight', { configurable: true, value: 240 })
    host.appendChild(gridHost)

    await act(async () => {
      root.render(
        <GridTextOverlay
          active
          host={gridHost}
          scene={{
            items: [
              {
                x: 10,
                y: 10,
                width: 120,
                height: 22,
                clipInsetTop: 0,
                clipInsetRight: 0,
                clipInsetBottom: 0,
                clipInsetLeft: 0,
                text: 'Alpha',
                align: 'left',
                wrap: false,
                color: '#111111',
                font: '400 11px sans-serif',
                fontSize: 11,
                underline: false,
                strike: false,
              },
              {
                x: 10,
                y: 32,
                width: 120,
                height: 22,
                clipInsetTop: 0,
                clipInsetRight: 0,
                clipInsetBottom: 0,
                clipInsetLeft: 0,
                text: 'Beta',
                align: 'right',
                wrap: false,
                color: '#111111',
                font: '400 11px sans-serif',
                fontSize: 11,
                underline: true,
                strike: false,
              },
            ],
          }}
        />,
      )
    })

    expect(host.querySelectorAll('[data-testid="grid-text-overlay"]').length).toBe(1)
    expect(host.querySelector('[data-testid="grid-text-overlay"]')?.tagName).toBe('CANVAS')
    expect(host.querySelectorAll('span').length).toBe(0)
    expect(host.querySelectorAll('canvas').length).toBe(1)

    await act(async () => {
      root.unmount()
    })
    host.remove()
  })
})
