// @vitest-environment jsdom
import { act } from 'react'
import { createRoot } from 'react-dom/client'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { CellEditorOverlay } from '../CellEditorOverlay.js'

afterEach(() => {
  document.body.innerHTML = ''
})

describe('CellEditorOverlay', () => {
  it('renders a flat single-frame editor without rounded or shadow chrome', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(
        <CellEditorOverlay label="Sheet1!B2" onCancel={() => {}} onChange={() => {}} onCommit={() => {}} resolvedValue="" value="draft" />,
      )
    })

    const overlay = host.querySelector<HTMLElement>("[data-testid='cell-editor-overlay']")
    expect(overlay?.getAttribute('class')).not.toContain('rounded-')
    expect(overlay?.getAttribute('class')).not.toContain('shadow-')

    await act(async () => {
      root.unmount()
    })
  })

  it('renders a textarea editor and keeps Alt+Enter for multiline input instead of committing', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const onChange = vi.fn()
    const onCommit = vi.fn()
    const onCancel = vi.fn()
    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(
        <CellEditorOverlay label="Sheet1!B2" onCancel={onCancel} onChange={onChange} onCommit={onCommit} resolvedValue="" value="line 1" />,
      )
    })

    const textarea = host.querySelector<HTMLTextAreaElement>("[data-testid='cell-editor-input']")
    expect(textarea?.tagName).toBe('TEXTAREA')

    await act(async () => {
      textarea?.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', altKey: true, bubbles: true }))
    })

    expect(onCommit).not.toHaveBeenCalled()
    expect(onCancel).not.toHaveBeenCalled()

    await act(async () => {
      root.unmount()
    })
  })
})
