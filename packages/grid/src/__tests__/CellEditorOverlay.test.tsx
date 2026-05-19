// @vitest-environment jsdom
import { act, useState } from 'react'
import { flushSync } from 'react-dom'
import { createRoot } from 'react-dom/client'
import type * as ReactDom from 'react-dom'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { CellEditorOverlay } from '../CellEditorOverlay.js'
import { WORKBOOK_DEFAULT_FONT_SIZE, workbookFontPointSizeToCssPx } from '../workbookTheme.js'

vi.mock('react-dom', async (importOriginal) => {
  const actual = await importOriginal<typeof ReactDom>()
  return {
    ...actual,
    flushSync: vi.fn(actual.flushSync),
  }
})

afterEach(() => {
  document.body.innerHTML = ''
  vi.mocked(flushSync).mockClear()
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

  it('uses compact workbook typography by default inside the cell editor', async () => {
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

    const textarea = host.querySelector<HTMLTextAreaElement>("[data-testid='cell-editor-input']")
    expect(textarea?.style.fontSize).toBe(`${workbookFontPointSizeToCssPx(WORKBOOK_DEFAULT_FONT_SIZE)}px`)
    expect(textarea?.style.fontFeatureSettings).toBe('normal')
    expect(textarea?.style.fontVariantNumeric).toBe('tabular-nums')
    expect(textarea?.style.letterSpacing).toBe('0px')
    expect(textarea?.style.fontOpticalSizing).toBe('auto')
    expect(textarea?.style.fontSynthesis).toBe('none')
    expect(textarea?.style.textRendering).toBe('optimizelegibility')
    expect(textarea?.style.WebkitFontSmoothing).toBe('antialiased')
    expect(textarea?.style.MozOsxFontSmoothing).toBe('grayscale')
    expect(textarea?.getAttribute('class')).toContain('py-[3px]')
    expect(textarea?.getAttribute('class')).toContain('leading-[1.2]')

    await act(async () => {
      root.unmount()
    })
  })

  it('reclaims focus from the grid focus target after editor mount', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const mockFrames = installMockAnimationFrames()
    const host = document.createElement('div')
    const gridFocusTarget = document.createElement('div')
    gridFocusTarget.tabIndex = 0
    gridFocusTarget.dataset['testid'] = 'sheet-grid-focus-target'
    document.body.append(host, gridFocusTarget)
    const root = createRoot(host)

    try {
      await act(async () => {
        root.render(
          <CellEditorOverlay
            label="Sheet1!B2"
            targetSelection={makeTargetSelection()}
            onCancel={() => {}}
            onChange={() => {}}
            onCommit={() => {}}
            resolvedValue=""
            selectionBehavior="caret-end"
            value="draft"
          />,
        )
      })

      const textarea = host.querySelector<HTMLTextAreaElement>("[data-testid='cell-editor-input']")
      expect(textarea).not.toBeNull()
      if (!textarea) {
        throw new Error('Expected mounted cell editor input')
      }

      gridFocusTarget.focus()
      expect(document.activeElement).toBe(gridFocusTarget)

      await act(async () => {
        mockFrames.flushAnimationFrames()
      })

      expect(document.activeElement).toBe(textarea)
      expect(textarea.selectionStart).toBe(5)
      expect(textarea.selectionEnd).toBe(5)
    } finally {
      await act(async () => {
        root.unmount()
      })
      mockFrames.restore()
    }
  })

  it('does not take focus back from a non-grid control during editor mount', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const mockFrames = installMockAnimationFrames()
    const host = document.createElement('div')
    const toolbarButton = document.createElement('button')
    toolbarButton.textContent = 'Toolbar'
    document.body.append(host, toolbarButton)
    const root = createRoot(host)

    try {
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

      toolbarButton.focus()
      expect(document.activeElement).toBe(toolbarButton)

      await act(async () => {
        mockFrames.flushAnimationFrames()
      })

      expect(document.activeElement).toBe(toolbarButton)
    } finally {
      await act(async () => {
        root.unmount()
      })
      mockFrames.restore()
    }
  })

  it.each([
    ['Alt+Enter', { altKey: true }],
    ['Ctrl+Enter', { ctrlKey: true }],
    ['Meta+Enter', { metaKey: true }],
  ] as const)('renders a textarea editor and keeps %s for multiline input instead of committing', async (_label, modifier) => {
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
    textarea?.setSelectionRange(textarea.value.length, textarea.value.length)

    await act(async () => {
      textarea?.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', ...modifier, bubbles: true }))
    })

    expect(onCommit).not.toHaveBeenCalled()
    expect(onCancel).not.toHaveBeenCalled()
    expect(onChange).toHaveBeenCalledWith('line 1\n')

    await act(async () => {
      root.unmount()
    })
  })

  it.each(['Enter', 'Tab', 'Escape'] as const)('leaves composing %s to the browser IME instead of committing navigation', async (key) => {
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
            value="compose"
          />,
        )
      })

      const textarea = host.querySelector<HTMLTextAreaElement>("[data-testid='cell-editor-input']")
      expect(textarea).not.toBeNull()
      if (!textarea) {
        throw new Error('Expected mounted cell editor input')
      }

      const event = new KeyboardEvent('keydown', { bubbles: true, cancelable: true, isComposing: true, key })
      let defaultAllowed = false
      await act(async () => {
        defaultAllowed = textarea.dispatchEvent(event)
      })

      expect(defaultAllowed).toBe(true)
      expect(event.defaultPrevented).toBe(false)
      expect(onCommit).not.toHaveBeenCalled()
      expect(onCancel).not.toHaveBeenCalled()
      expect(onChange).not.toHaveBeenCalled()
      expect(textarea.readOnly).toBe(false)
      expect(textarea.value).toBe('compose')
    } finally {
      await act(async () => {
        root.unmount()
      })
    }
  })

  it('keeps typed draft text local while mirroring parent state immediately', async () => {
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
      expect(textarea.selectionStart).toBe(3)
      expect(textarea.selectionEnd).toBe(3)
      expect(onChange).toHaveBeenCalledTimes(3)
      expect(onChange).toHaveBeenNthCalledWith(1, 'a')
      expect(onChange).toHaveBeenNthCalledWith(2, 'ab')
      expect(onChange).toHaveBeenLastCalledWith('abc')
    } finally {
      await act(async () => {
        root.unmount()
      })
    }
  })

  it('keeps rapid horizontal caret movement after manual printable insertion', async () => {
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
            resolvedValue=""
            selectionBehavior="caret-end"
            value="ab"
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
        textarea.setSelectionRange(2, 2)
        textarea.dispatchEvent(new KeyboardEvent('keydown', { key: 'c', bubbles: true, cancelable: true }))
        textarea.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowLeft', bubbles: true, cancelable: true }))
      })

      expect(textarea.value).toBe('abc')
      expect(textarea.selectionStart).toBe(2)
      expect(textarea.selectionEnd).toBe(2)
      expect(onChange).toHaveBeenCalledTimes(1)
      expect(onChange).toHaveBeenLastCalledWith('abc')
    } finally {
      await act(async () => {
        root.unmount()
      })
    }
  })

  it('extends selection from the latest typed caret after manual printable insertion', async () => {
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
            resolvedValue=""
            selectionBehavior="caret-end"
            value="ab"
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
        textarea.setSelectionRange(2, 2)
        textarea.dispatchEvent(new KeyboardEvent('keydown', { key: 'c', bubbles: true, cancelable: true }))
        textarea.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowLeft', shiftKey: true, bubbles: true, cancelable: true }))
      })

      expect(textarea.value).toBe('abc')
      expect(textarea.selectionStart).toBe(2)
      expect(textarea.selectionEnd).toBe(3)
      expect(textarea.selectionDirection).toBe('backward')
      expect(onChange).toHaveBeenCalledTimes(1)
      expect(onChange).toHaveBeenLastCalledWith('abc')
    } finally {
      await act(async () => {
        root.unmount()
      })
    }
  })

  it('moves Home and End from the pending keyboard caret during rapid manual editing', async () => {
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
            resolvedValue=""
            selectionBehavior="caret-end"
            value="ab"
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
        textarea.setSelectionRange(2, 2)
        textarea.dispatchEvent(new KeyboardEvent('keydown', { key: 'c', bubbles: true, cancelable: true }))
        textarea.dispatchEvent(new KeyboardEvent('keydown', { key: 'Home', bubbles: true, cancelable: true }))
        textarea.dispatchEvent(new KeyboardEvent('keydown', { key: 'z', bubbles: true, cancelable: true }))
        textarea.dispatchEvent(new KeyboardEvent('keydown', { key: 'End', bubbles: true, cancelable: true }))
        textarea.dispatchEvent(new KeyboardEvent('keydown', { key: '!', bubbles: true, cancelable: true }))
      })

      expect(textarea.value).toBe('zabc!')
      expect(textarea.selectionStart).toBe(5)
      expect(textarea.selectionEnd).toBe(5)
      expect(onChange).toHaveBeenLastCalledWith('zabc!')
    } finally {
      await act(async () => {
        root.unmount()
      })
    }
  })

  it('honors later manual caret placement after a pending keyboard restore finishes', async () => {
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
            selectionBehavior="caret-end"
            value="ab"
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
        textarea.setSelectionRange(2, 2)
        textarea.dispatchEvent(new KeyboardEvent('keydown', { key: 'c', bubbles: true, cancelable: true }))
        textarea.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowLeft', bubbles: true, cancelable: true }))
      })

      await act(async () => {
        mockFrames.flushAnimationFrames()
      })

      await act(async () => {
        textarea.setSelectionRange(0, 0)
        textarea.dispatchEvent(new KeyboardEvent('keydown', { key: 'x', bubbles: true, cancelable: true }))
      })

      expect(textarea.value).toBe('xabc')
      expect(textarea.selectionStart).toBe(1)
      expect(textarea.selectionEnd).toBe(1)
      expect(onChange).toHaveBeenCalledTimes(2)
      expect(onChange).toHaveBeenNthCalledWith(1, 'abc')
      expect(onChange).toHaveBeenLastCalledWith('xabc')
    } finally {
      await act(async () => {
        root.unmount()
      })
      mockFrames.restore()
    }
  })

  it('keeps printable key insertion off synchronous React flushes while syncing each draft', async () => {
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

      vi.mocked(flushSync).mockClear()

      await act(async () => {
        textarea.dispatchEvent(new KeyboardEvent('keydown', { key: 'a', bubbles: true, cancelable: true }))
        textarea.dispatchEvent(new KeyboardEvent('keydown', { key: 'b', bubbles: true, cancelable: true }))
        textarea.dispatchEvent(new KeyboardEvent('keydown', { key: 'c', bubbles: true, cancelable: true }))
      })

      expect(textarea.value).toBe('abc')
      expect(onChange).toHaveBeenCalledTimes(3)
      expect(onChange).toHaveBeenLastCalledWith('abc')
      expect(flushSync).not.toHaveBeenCalled()
    } finally {
      await act(async () => {
        root.unmount()
      })
    }
  })

  it('keeps local draft text and caret stable across parent renders after immediate parent sync', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    vi.useFakeTimers()
    const onChange = vi.fn()
    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    function RerenderHarness() {
      const [renderCount, setRenderCount] = useState(0)
      return (
        <div>
          <button
            data-testid="force-parent-render"
            type="button"
            onClick={() => {
              setRenderCount((current) => current + 1)
            }}
          >
            {renderCount}
          </button>
          <CellEditorOverlay
            label="Sheet1!B2"
            targetSelection={makeTargetSelection()}
            onCancel={() => {}}
            onChange={onChange}
            onCommit={() => {}}
            resolvedValue=""
            value=""
          />
        </div>
      )
    }

    try {
      await act(async () => {
        root.render(<RerenderHarness />)
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
      expect(textarea.selectionStart).toBe(3)
      expect(textarea.selectionEnd).toBe(3)
      expect(onChange).toHaveBeenCalledTimes(3)
      expect(onChange).toHaveBeenLastCalledWith('abc')

      await act(async () => {
        host.querySelector<HTMLButtonElement>("[data-testid='force-parent-render']")?.click()
      })

      expect(textarea.value).toBe('abc')
      expect(textarea.selectionStart).toBe(3)
      expect(textarea.selectionEnd).toBe(3)
      expect(onChange).toHaveBeenCalledTimes(3)

      await act(async () => {
        vi.advanceTimersByTime(250)
      })

      expect(onChange).toHaveBeenCalledTimes(3)
      expect(onChange).toHaveBeenLastCalledWith('abc')
    } finally {
      await act(async () => {
        root.unmount()
      })
      vi.useRealTimers()
    }
  })

  it('ignores stale same-cell parent values while the focused editor has a newer draft', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    vi.useFakeTimers()
    const onChange = vi.fn()
    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)
    let setParentValue: ((next: string) => void) | null = null

    function StaleParentHarness() {
      const [parentValue, setValue] = useState('')
      setParentValue = setValue
      return (
        <CellEditorOverlay
          label="Sheet1!B2"
          targetSelection={makeTargetSelection()}
          onCancel={() => {}}
          onChange={onChange}
          onCommit={() => {}}
          resolvedValue=""
          value={parentValue}
        />
      )
    }

    try {
      await act(async () => {
        root.render(<StaleParentHarness />)
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
      expect(textarea.selectionStart).toBe(3)

      await act(async () => {
        setParentValue?.('a')
      })

      expect(textarea.value).toBe('abc')
      expect(textarea.selectionStart).toBe(3)
      expect(onChange).toHaveBeenCalledTimes(3)
      expect(onChange).toHaveBeenLastCalledWith('abc')

      await act(async () => {
        vi.advanceTimersByTime(250)
      })

      expect(textarea.value).toBe('abc')
      expect(textarea.selectionStart).toBe(3)
      expect(onChange).toHaveBeenCalledTimes(3)
      expect(onChange).toHaveBeenLastCalledWith('abc')
    } finally {
      await act(async () => {
        root.unmount()
      })
      vi.useRealTimers()
    }
  })

  it('commits local draft text after immediate parent sync', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    vi.useFakeTimers()
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
      vi.useRealTimers()
    }
  })

  it('handles backspace and delete in the editor without relying on native textarea timing', async () => {
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
        textarea.focus()
        textarea.setSelectionRange(3, 3)
        textarea.dispatchEvent(new KeyboardEvent('keydown', { key: 'Backspace', bubbles: true, cancelable: true }))
      })
      expect(textarea.value).toBe('abdef')
      expect(textarea.selectionStart).toBe(2)
      expect(textarea.selectionEnd).toBe(2)
      expect(onChange).toHaveBeenLastCalledWith('abdef')

      await act(async () => {
        textarea.setSelectionRange(2, 4)
        textarea.dispatchEvent(new KeyboardEvent('keydown', { key: 'Delete', bubbles: true, cancelable: true }))
      })
      expect(textarea.value).toBe('abf')
      expect(textarea.selectionStart).toBe(2)
      expect(textarea.selectionEnd).toBe(2)
      expect(onChange).toHaveBeenLastCalledWith('abf')

      await act(async () => {
        textarea.dispatchEvent(new KeyboardEvent('keydown', { key: 'x', bubbles: true, cancelable: true }))
      })
      expect(textarea.value).toBe('abxf')
      expect(textarea.selectionStart).toBe(3)
      expect(textarea.selectionEnd).toBe(3)
      expect(onChange).toHaveBeenLastCalledWith('abxf')
    } finally {
      await act(async () => {
        root.unmount()
      })
    }
  })

  it('owns primary select-all before delete and continued typing in the editor', async () => {
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
            resolvedValue=""
            selectionBehavior="caret-end"
            value="delete-me"
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
        textarea.setSelectionRange(textarea.value.length, textarea.value.length)
        const selectAllEvent = new KeyboardEvent('keydown', { key: 'a', metaKey: true, bubbles: true, cancelable: true })
        textarea.dispatchEvent(selectAllEvent)
        expect(selectAllEvent.defaultPrevented).toBe(true)
      })
      expect(textarea.value).toBe('delete-me')
      expect(textarea.selectionStart).toBe(0)
      expect(textarea.selectionEnd).toBe('delete-me'.length)

      await act(async () => {
        textarea.dispatchEvent(new KeyboardEvent('keydown', { key: 'Backspace', bubbles: true, cancelable: true }))
      })
      expect(textarea.value).toBe('')
      expect(textarea.selectionStart).toBe(0)
      expect(textarea.selectionEnd).toBe(0)
      expect(onChange).toHaveBeenLastCalledWith('')

      await act(async () => {
        textarea.dispatchEvent(new KeyboardEvent('keydown', { key: 'n', bubbles: true, cancelable: true }))
        textarea.dispatchEvent(new KeyboardEvent('keydown', { key: 'e', bubbles: true, cancelable: true }))
        textarea.dispatchEvent(new KeyboardEvent('keydown', { key: 'w', bubbles: true, cancelable: true }))
      })
      expect(textarea.value).toBe('new')
      expect(textarea.selectionStart).toBe(3)
      expect(textarea.selectionEnd).toBe(3)
      expect(onChange).toHaveBeenLastCalledWith('new')
    } finally {
      await act(async () => {
        root.unmount()
      })
    }
  })

  it('keeps editable text visible while a commit is completing to avoid blank click-away flashes', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

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
            onChange={() => {}}
            onCommit={onCommit}
            resolvedValue=""
            value="visible draft"
          />,
        )
      })

      const overlay = host.querySelector<HTMLElement>("[data-testid='cell-editor-overlay']")
      const textarea = host.querySelector<HTMLTextAreaElement>("[data-testid='cell-editor-input']")
      expect(textarea).not.toBeNull()
      if (!textarea) {
        throw new Error('Expected mounted cell editor input')
      }

      await act(async () => {
        textarea.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true }))
      })

      expect(onCommit).toHaveBeenCalledTimes(1)
      expect(overlay?.dataset['completing']).toBe('true')
      expect(textarea.readOnly).toBe(true)
      expect(textarea.style.opacity).toBe('')
    } finally {
      await act(async () => {
        root.unmount()
      })
    }
  })

  it('does not swap the finishing editor draft when click-away selection props advance before unmount', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

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
            onChange={() => {}}
            onCommit={onCommit}
            resolvedValue=""
            value="visible draft"
          />,
        )
      })

      const initialTextarea = host.querySelector<HTMLTextAreaElement>("[data-testid='cell-editor-input']")
      expect(initialTextarea).not.toBeNull()
      if (!initialTextarea) {
        throw new Error('Expected mounted cell editor input')
      }

      await act(async () => {
        initialTextarea.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true }))
      })

      await act(async () => {
        root.render(
          <CellEditorOverlay
            label="Sheet1!C2"
            targetSelection={{ sheetName: 'Sheet1', address: 'C2' }}
            onCancel={() => {}}
            onChange={() => {}}
            onCommit={onCommit}
            resolvedValue=""
            value="next cell text"
          />,
        )
      })

      const completingTextarea = host.querySelector<HTMLTextAreaElement>("[data-testid='cell-editor-input']")
      expect(completingTextarea).not.toBeNull()
      expect(completingTextarea?.readOnly).toBe(true)
      expect(completingTextarea?.value).toBe('visible draft')
      expect(onCommit).toHaveBeenCalledTimes(1)
      expect(onCommit).toHaveBeenLastCalledWith([0, 1], 'visible draft', makeTargetSelection())
    } finally {
      await act(async () => {
        root.unmount()
      })
    }
  })

  it('commits delete and backspace DOM edits from the mounted editor', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    vi.useFakeTimers()
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

      expect(onChange).not.toHaveBeenCalled()
      expect(onCommit).toHaveBeenCalledTimes(1)
      expect(onCommit).toHaveBeenLastCalledWith([0, 1], 'bcde', makeTargetSelection())
    } finally {
      await act(async () => {
        root.unmount()
      })
      vi.useRealTimers()
    }
  })

  it('commits an immediate click-away blur even before the initial blur guard frame is armed', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const mockFrames = installMockAnimationFrames()
    const onCommit = vi.fn()
    const host = document.createElement('div')
    const toolbarButton = document.createElement('button')
    toolbarButton.textContent = 'Toolbar'
    document.body.append(host, toolbarButton)
    const root = createRoot(host)

    try {
      await act(async () => {
        root.render(
          <CellEditorOverlay
            label="Sheet1!B2"
            targetSelection={makeTargetSelection()}
            onCancel={() => {}}
            onChange={() => {}}
            onCommit={onCommit}
            resolvedValue=""
            selectionBehavior="caret-end"
            value="draft before click-away"
          />,
        )
      })

      const textarea = host.querySelector<HTMLTextAreaElement>("[data-testid='cell-editor-input']")
      expect(textarea).not.toBeNull()
      if (!textarea) {
        throw new Error('Expected mounted cell editor input')
      }
      expect(document.activeElement).toBe(textarea)

      await act(async () => {
        toolbarButton.focus()
      })

      expect(document.activeElement).toBe(toolbarButton)
      expect(onCommit).not.toHaveBeenCalled()

      await act(async () => {
        mockFrames.flushAnimationFrames()
      })
      expect(onCommit).not.toHaveBeenCalled()
      expect(textarea.readOnly).toBe(true)

      await act(async () => {
        mockFrames.flushAnimationFrames()
      })

      expect(onCommit).toHaveBeenCalledTimes(1)
      expect(onCommit).toHaveBeenLastCalledWith(undefined, 'draft before click-away', makeTargetSelection())
    } finally {
      await act(async () => {
        root.unmount()
      })
      mockFrames.restore()
    }
  })

  it('keeps Home and End caret movement inside the editor', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const mockFrames = installMockAnimationFrames()
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
            onChange={() => {}}
            onCommit={() => {}}
            resolvedValue=""
            selectionBehavior="caret-end"
            value="abcd"
          />,
        )
      })

      const textarea = host.querySelector<HTMLTextAreaElement>("[data-testid='cell-editor-input']")
      expect(textarea).not.toBeNull()
      if (!textarea) {
        throw new Error('Expected mounted cell editor input')
      }

      await act(async () => {
        textarea.dispatchEvent(new KeyboardEvent('keydown', { key: 'Home', bubbles: true, cancelable: true }))
      })
      expect(textarea.selectionStart).toBe(0)
      expect(textarea.selectionEnd).toBe(0)
      textarea.setSelectionRange(4, 4)
      mockFrames.flushAnimationFrames()
      expect(textarea.selectionStart).toBe(0)
      expect(textarea.selectionEnd).toBe(0)

      await act(async () => {
        textarea.dispatchEvent(new KeyboardEvent('keydown', { key: 'End', bubbles: true, cancelable: true }))
      })
      expect(textarea.selectionStart).toBe(4)
      expect(textarea.selectionEnd).toBe(4)
      textarea.setSelectionRange(0, 0)
      mockFrames.flushAnimationFrames()
      expect(textarea.selectionStart).toBe(4)
      expect(textarea.selectionEnd).toBe(4)
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

      const overlay = host.querySelector<HTMLElement>("[data-testid='cell-editor-overlay']")
      const textarea = host.querySelector<HTMLTextAreaElement>("[data-testid='cell-editor-input']")

      await act(async () => {
        textarea?.dispatchEvent(new FocusEvent('focusout', { bubbles: true }))
      })
      expect(onCommit).not.toHaveBeenCalled()
      expect(overlay?.dataset['completing']).toBe('true')
      expect(textarea?.readOnly).toBe(true)
      expect(textarea?.style.opacity).toBe('')

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
