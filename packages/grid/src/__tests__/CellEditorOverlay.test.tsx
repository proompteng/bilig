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

function installMockAnimationFrames() {
  const originalRequestAnimationFrame = window.requestAnimationFrame
  const originalCancelAnimationFrame = window.cancelAnimationFrame
  const frameCallbacks = new Map<number, FrameRequestCallback>()
  let nextFrame = 1
  window.requestAnimationFrame = ((callback: FrameRequestCallback) => {
    const id = nextFrame
    nextFrame += 1
    frameCallbacks.set(id, callback)
    return id
  }) as typeof window.requestAnimationFrame
  window.cancelAnimationFrame = ((id: number) => {
    frameCallbacks.delete(id)
  }) as typeof window.cancelAnimationFrame

  return {
    flushAnimationFrames() {
      const callbacks = [...frameCallbacks.values()]
      frameCallbacks.clear()
      callbacks.forEach((callback) => callback(performance.now()))
    },
    restore() {
      window.requestAnimationFrame = originalRequestAnimationFrame
      window.cancelAnimationFrame = originalCancelAnimationFrame
    },
  }
}

describe('CellEditorOverlay', () => {
  it('renders a flat single-frame editor without rounded or shadow chrome', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(
        <CellEditorOverlay
          label="Sheet1!B2"
          targetSelection={makeTargetSelection()}
          onCancel={() => {}}
          onChange={() => {}}
          onCommit={() => {}}
          resolvedValue=""
          value="draft"
        />,
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
        <CellEditorOverlay
          label="Sheet1!B2"
          targetSelection={makeTargetSelection()}
          onCancel={onCancel}
          onChange={onChange}
          onCommit={onCommit}
          resolvedValue=""
          value="line 1"
        />,
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

  it('keeps typed draft text local until the next frame', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const mockFrames = installMockAnimationFrames()
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
            resolvedValue=""
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
        textarea.dispatchEvent(new KeyboardEvent('keydown', { key: 'a', bubbles: true, cancelable: true }))
        textarea.dispatchEvent(new KeyboardEvent('keydown', { key: 'b', bubbles: true, cancelable: true }))
        textarea.dispatchEvent(new KeyboardEvent('keydown', { key: 'c', bubbles: true, cancelable: true }))
      })

      expect(textarea.value).toBe('abc')
      expect(onChange).not.toHaveBeenCalled()

      await act(async () => {
        mockFrames.flushAnimationFrames()
      })

      expect(onChange).toHaveBeenCalledTimes(1)
      expect(onChange).toHaveBeenLastCalledWith('abc')
    } finally {
      await act(async () => {
        root.unmount()
      })
      mockFrames.restore()
    }
  })

  it('flushes local draft text before committing', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const mockFrames = installMockAnimationFrames()
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
            resolvedValue=""
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
        textarea.dispatchEvent(new KeyboardEvent('keydown', { key: 'x', bubbles: true, cancelable: true }))
        textarea.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true }))
      })

      expect(onChange).toHaveBeenCalledTimes(1)
      expect(onChange).toHaveBeenLastCalledWith('x')
      expect(onCommit).toHaveBeenCalledTimes(1)
      expect(onCommit).toHaveBeenLastCalledWith([0, 1], 'x', makeTargetSelection())
    } finally {
      await act(async () => {
        root.unmount()
      })
      mockFrames.restore()
    }
  })

  it('keeps delete and backspace edits in the local draft before the next frame', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const mockFrames = installMockAnimationFrames()
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
            resolvedValue=""
            selectionBehavior="caret-end"
            value="abcdef"
          />,
        )
      })

      const textarea = host.querySelector<HTMLTextAreaElement>("[data-testid='cell-editor-input']")
      expect(textarea).not.toBeNull()
      if (!textarea) {
        throw new Error('Expected mounted cell editor input')
      }

      await act(async () => {
        textarea.setSelectionRange(6, 6)
        textarea.value = 'abcde'
        textarea.dispatchEvent(new InputEvent('input', { bubbles: true, data: null, inputType: 'deleteContentBackward' }))
        textarea.setSelectionRange(0, 0)
        textarea.value = 'bcde'
        textarea.dispatchEvent(new InputEvent('input', { bubbles: true, data: null, inputType: 'deleteContentForward' }))
      })

      expect(textarea.value).toBe('bcde')
      expect(onChange).not.toHaveBeenCalled()

      await act(async () => {
        textarea.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true }))
      })

      expect(onChange).toHaveBeenCalledTimes(1)
      expect(onChange).toHaveBeenLastCalledWith('bcde')
      expect(onCommit).toHaveBeenCalledTimes(1)
      expect(onCommit).toHaveBeenLastCalledWith([0, 1], 'bcde', makeTargetSelection())
    } finally {
      await act(async () => {
        root.unmount()
      })
      mockFrames.restore()
    }
  })

  it('lets movement commits win over a queued blur commit', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const mockFrames = installMockAnimationFrames()
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
            onChange={() => {}}
            onCommit={onCommit}
            resolvedValue=""
            value="line 1"
          />,
        )
      })
      await act(async () => {
        mockFrames.flushAnimationFrames()
      })

      const textarea = host.querySelector<HTMLTextAreaElement>("[data-testid='cell-editor-input']")

      await act(async () => {
        textarea?.dispatchEvent(new FocusEvent('focusout', { bubbles: true }))
      })
      expect(onCommit).not.toHaveBeenCalled()

      await act(async () => {
        textarea?.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }))
      })

      expect(onCommit).toHaveBeenCalledTimes(1)
      expect(onCommit).toHaveBeenLastCalledWith([0, 1], 'line 1', makeTargetSelection())
      expect(onCancel).not.toHaveBeenCalled()

      await act(async () => {
        mockFrames.flushAnimationFrames()
      })
      expect(onCommit).toHaveBeenCalledTimes(1)
    } finally {
      await act(async () => {
        root.unmount()
      })
      mockFrames.restore()
    }
  })
})
