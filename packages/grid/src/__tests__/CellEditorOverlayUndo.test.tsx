// @vitest-environment jsdom
import { act } from 'react'
import { createRoot } from 'react-dom/client'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { CellEditorOverlay } from '../CellEditorOverlay.js'

afterEach(() => {
  document.body.innerHTML = ''
})

function makeTargetSelection() {
  return { sheetName: 'Sheet1', address: 'B2' }
}

function dispatchEditorKey(input: HTMLTextAreaElement, init: KeyboardEventInit) {
  input.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, cancelable: true, ...init }))
}

describe('CellEditorOverlay active draft history', () => {
  it('handles primary undo and redo inside the active in-cell editor without committing', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const onChange = vi.fn()
    const onCommit = vi.fn()
    const onCancel = vi.fn()
    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    try {
      await act(async () => {
        root.render(
          <CellEditorOverlay
            label="Sheet1!B2"
            targetSelection={makeTargetSelection()}
            onCancel={onCancel}
            onChange={onChange}
            onCommit={onCommit}
            resolvedValue=""
            selectionBehavior="caret-end"
            value=""
          />,
        )
      })

      const textarea = host.querySelector<HTMLTextAreaElement>("[data-testid='cell-editor-input']")
      expect(textarea).not.toBeNull()
      if (!textarea) {
        throw new Error('Expected mounted cell editor input')
      }

      await act(async () => {
        textarea.focus()
        textarea.setSelectionRange(0, 0)
        dispatchEditorKey(textarea, { key: 'a' })
        dispatchEditorKey(textarea, { key: 'b' })
        dispatchEditorKey(textarea, { key: 'c' })
      })
      expect(textarea.value).toBe('abc')
      expect(textarea.selectionStart).toBe(3)

      await act(async () => {
        dispatchEditorKey(textarea, { ctrlKey: true, key: 'z' })
      })
      expect(textarea.value).toBe('ab')
      expect(textarea.selectionStart).toBe(2)
      expect(onCommit).not.toHaveBeenCalled()
      expect(onCancel).not.toHaveBeenCalled()
      expect(onChange).toHaveBeenLastCalledWith('ab')

      await act(async () => {
        dispatchEditorKey(textarea, { ctrlKey: true, key: 'z' })
      })
      expect(textarea.value).toBe('a')
      expect(textarea.selectionStart).toBe(1)

      await act(async () => {
        dispatchEditorKey(textarea, { ctrlKey: true, key: 'z', shiftKey: true })
      })
      expect(textarea.value).toBe('ab')
      expect(textarea.selectionStart).toBe(2)

      await act(async () => {
        dispatchEditorKey(textarea, { ctrlKey: true, key: 'y' })
      })
      expect(textarea.value).toBe('abc')
      expect(textarea.selectionStart).toBe(3)
      expect(onChange).toHaveBeenLastCalledWith('abc')
    } finally {
      await act(async () => {
        root.unmount()
      })
    }
  })

  it('undoes a manual Delete clear in the active editor before click-away commit', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const onChange = vi.fn()
    const onCommit = vi.fn()
    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    try {
      await act(async () => {
        root.render(
          <CellEditorOverlay
            label="Sheet1!B2"
            targetSelection={makeTargetSelection()}
            onCancel={() => {}}
            onChange={onChange}
            onCommit={onCommit}
            resolvedValue="keep"
            value="keep"
          />,
        )
      })

      const textarea = host.querySelector<HTMLTextAreaElement>("[data-testid='cell-editor-input']")
      expect(textarea).not.toBeNull()
      if (!textarea) {
        throw new Error('Expected mounted cell editor input')
      }

      await act(async () => {
        textarea.focus()
        textarea.select()
        dispatchEditorKey(textarea, { key: 'Delete' })
      })
      expect(textarea.value).toBe('')

      await act(async () => {
        dispatchEditorKey(textarea, { metaKey: true, key: 'z' })
      })
      expect(textarea.value).toBe('keep')
      expect(textarea.selectionStart).toBe(0)
      expect(textarea.selectionEnd).toBe(4)
      expect(onCommit).not.toHaveBeenCalled()
      expect(onChange).toHaveBeenLastCalledWith('keep')
    } finally {
      await act(async () => {
        root.unmount()
      })
    }
  })

  it('does not insert normalized numpad text while a shortcut modifier is held', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const onChange = vi.fn()
    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    try {
      await act(async () => {
        root.render(
          <CellEditorOverlay
            label="Sheet1!B2"
            targetSelection={makeTargetSelection()}
            onCancel={() => {}}
            onChange={onChange}
            onCommit={() => {}}
            resolvedValue="keep"
            selectionBehavior="caret-end"
            value="keep"
          />,
        )
      })

      const textarea = host.querySelector<HTMLTextAreaElement>("[data-testid='cell-editor-input']")
      expect(textarea).not.toBeNull()
      if (!textarea) {
        throw new Error('Expected mounted cell editor input')
      }

      await act(async () => {
        textarea.focus()
        textarea.setSelectionRange(4, 4)
        dispatchEditorKey(textarea, { code: 'NumpadSubtract', ctrlKey: true, key: 'Subtract' })
      })

      expect(textarea.value).toBe('keep')
      expect(onChange).not.toHaveBeenCalled()
    } finally {
      await act(async () => {
        root.unmount()
      })
    }
  })
})
