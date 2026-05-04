// @vitest-environment jsdom
import { act } from 'react'
import { createRoot } from 'react-dom/client'
import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import { useAutoSizingTextarea } from '../use-autosizing-textarea.js'

function AutoSizingFixture(props: { readonly value: string; readonly minHeight: number; readonly maxHeight: number }) {
  const { textareaRef, viewportRef } = useAutoSizingTextarea({
    value: props.value,
    minHeight: props.minHeight,
    maxHeight: props.maxHeight,
  })

  return (
    <div>
      <div ref={viewportRef} data-testid="viewport" />
      <textarea ref={textareaRef} value={props.value} readOnly data-testid="textarea" />
    </div>
  )
}

describe('useAutoSizingTextarea', () => {
  let originalDescriptor: PropertyDescriptor | undefined

  beforeEach(() => {
    originalDescriptor = Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'scrollHeight')
    Object.defineProperty(HTMLTextAreaElement.prototype, 'scrollHeight', {
      configurable: true,
      get() {
        return this.value.length
      },
    })
  })

  afterEach(() => {
    if (originalDescriptor) {
      Object.defineProperty(HTMLTextAreaElement.prototype, 'scrollHeight', originalDescriptor)
    }
    document.body.innerHTML = ''
  })

  it('keeps textarea height at content size while capping viewport height', async () => {
    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    const minHeight = 112
    const maxHeight = 224
    const shortValue = 'short'
    const longValue = 'x'.repeat(300)

    await act(async () => {
      root.render(<AutoSizingFixture value={shortValue} minHeight={minHeight} maxHeight={maxHeight} />)
    })

    const shortTextarea = host.querySelector('[data-testid="textarea"]')
    const shortViewport = host.querySelector('[data-testid="viewport"]')

    expect(shortTextarea instanceof HTMLTextAreaElement).toBe(true)
    expect(shortViewport instanceof HTMLDivElement).toBe(true)
    expect(shortTextarea?.style.height).toBe(`${minHeight}px`)
    expect(shortViewport?.style.height).toBe(`${minHeight}px`)

    await act(async () => {
      root.render(<AutoSizingFixture value={longValue} minHeight={minHeight} maxHeight={maxHeight} />)
    })

    const longTextarea = host.querySelector('[data-testid="textarea"]')
    const longViewport = host.querySelector('[data-testid="viewport"]')

    expect(longTextarea instanceof HTMLTextAreaElement).toBe(true)
    expect(longViewport instanceof HTMLDivElement).toBe(true)
    expect(longTextarea?.style.height).toBe(`${longValue.length}px`)
    expect(longViewport?.style.height).toBe(`${maxHeight}px`)

    await act(async () => {
      root.unmount()
    })
  })
})
