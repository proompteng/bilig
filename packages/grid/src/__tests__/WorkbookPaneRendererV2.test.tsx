// @vitest-environment jsdom
import { act } from 'react'
import { createRoot } from 'react-dom/client'
import { describe, expect, test } from 'vitest'
import { WorkbookPaneRendererV2 } from '../renderer-v2/WorkbookPaneRendererV2.js'

describe('WorkbookPaneRendererV2', () => {
  test('mounts a distinct V2 canvas for the hard migration path', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    const host = document.createElement('div')
    const root = createRoot(host)
    const rendererHost = document.createElement('div')
    host.appendChild(rendererHost)

    await act(async () => {
      root.render(<WorkbookPaneRendererV2 active host={rendererHost} geometry={null} />)
    })

    const canvas = host.querySelector('[data-testid="grid-pane-renderer-v2"]')
    expect(canvas).toBeInstanceOf(HTMLCanvasElement)
    expect(canvas?.getAttribute('data-pane-renderer')).toBe('workbook-pane-renderer-v2')

    await act(async () => {
      root.unmount()
    })
  })
})
