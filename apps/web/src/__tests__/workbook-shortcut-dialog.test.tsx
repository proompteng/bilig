// @vitest-environment jsdom
import { act, useState } from 'react'
import { createRoot } from 'react-dom/client'
import { afterEach, describe, expect, it } from 'vitest'
import { WorkbookShortcutDialog } from '../WorkbookShortcutDialog.js'
import { getWorkbookShortcutLabel, getWorkbookShortcutParts } from '../shortcut-registry.js'
import { useWorkbookShortcutDialog } from '../use-workbook-shortcut-dialog.js'

function ShortcutDialogHarness() {
  const shortcuts = useWorkbookShortcutDialog()
  return (
    <>
      {shortcuts.shortcutHelpButton}
      {shortcuts.shortcutDialog}
    </>
  )
}

function ShortcutDialogFilterHarness() {
  const [query, setQuery] = useState('')
  return <WorkbookShortcutDialog open query={query} onOpenChange={() => {}} onQueryChange={setQuery} />
}

function setInputValue(input: HTMLInputElement, value: string) {
  const descriptor = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')
  descriptor?.set?.call(input, value)
  input.dispatchEvent(new Event('input', { bubbles: true }))
}

afterEach(() => {
  document.body.innerHTML = ''
})

describe('workbook shortcut dialog', () => {
  it('opens from the header button', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(<ShortcutDialogHarness />)
    })

    const openButton = document.querySelector("[data-testid='workbook-shortcut-button']")
    expect(openButton?.textContent?.trim()).toBe('')
    expect(openButton?.getAttribute('class')).toContain('max-[420px]:hidden')
    await act(async () => {
      openButton?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
    })

    expect(document.querySelector("[data-testid='workbook-shortcut-dialog']")).not.toBeNull()
    expect(document.querySelector("[data-testid='workbook-shortcut-search']")).not.toBeNull()
    expect(document.activeElement).toBe(document.querySelector("[data-testid='workbook-shortcut-search']"))

    await act(async () => {
      root.unmount()
    })
  })

  it('filters shortcut entries from the search box', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(<ShortcutDialogFilterHarness />)
    })

    const search = document.querySelector<HTMLInputElement>("[data-testid='workbook-shortcut-search']")
    expect(search).not.toBeNull()
    if (!search) {
      throw new Error('expected search input')
    }

    await act(async () => {
      setInputValue(search, 'bold')
    })

    const entries = [...document.querySelectorAll("[data-testid='workbook-shortcut-entry']")]
    expect(entries).toHaveLength(1)
    expect(entries[0]?.textContent).toContain('Bold')

    await act(async () => {
      root.unmount()
    })
  })

  it('uses a larger mauve dialog shell with flat shortcut rows', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(<ShortcutDialogFilterHarness />)
    })

    const dialog = document.querySelector<HTMLElement>("[data-testid='workbook-shortcut-dialog']")
    expect(dialog?.getAttribute('aria-modal')).toBe('true')
    expect(dialog?.getAttribute('class')).toContain('w-[min(56rem,calc(100vw-3rem))]')
    expect(dialog?.getAttribute('class')).toContain('max-h-[min(44rem,calc(100dvh-3rem))]')
    expect(dialog?.getAttribute('class')).toContain('bg-[var(--color-mauve-50)]')

    const scrollBody = document.querySelector<HTMLElement>("[data-testid='workbook-shortcut-list']")
    expect(scrollBody?.getAttribute('class')).toContain('min-h-0')
    expect(scrollBody?.getAttribute('class')).toContain('flex-1')
    expect(scrollBody?.getAttribute('class')).toContain('overflow-y-auto')

    const firstEntry = document.querySelector<HTMLElement>("[data-testid='workbook-shortcut-entry']")
    expect(firstEntry?.getAttribute('class')).toContain('border-b')
    expect(firstEntry?.getAttribute('class')).not.toContain('rounded-')
    expect(firstEntry?.getAttribute('class')).not.toContain('bg-[var(--wb-surface-muted)]')

    await act(async () => {
      root.unmount()
    })
  })

  it('opens from the global question-mark shortcut when focus is not in text entry', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(<ShortcutDialogHarness />)
    })

    await act(async () => {
      window.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key: '?' }))
    })

    expect(document.querySelector("[data-testid='workbook-shortcut-dialog']")).not.toBeNull()

    await act(async () => {
      root.unmount()
    })
  })

  it('does not open from question-mark while typing in an input', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const host = document.createElement('div')
    document.body.appendChild(host)
    const textInput = document.createElement('input')
    document.body.appendChild(textInput)
    const root = createRoot(host)

    await act(async () => {
      root.render(<ShortcutDialogHarness />)
    })

    textInput.focus()
    await act(async () => {
      textInput.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key: '?' }))
    })

    expect(document.querySelector("[data-testid='workbook-shortcut-dialog']")).toBeNull()

    await act(async () => {
      root.unmount()
    })
  })

  it('formats shortcut labels per platform', () => {
    expect(getWorkbookShortcutLabel('undo', 'MacIntel')).toBe('⌘Z')
    expect(getWorkbookShortcutLabel('undo', 'Win32')).toBe('Ctrl+Z')
  })

  it('splits shortcuts into readable keycap parts', () => {
    expect(getWorkbookShortcutParts('undo', 'MacIntel')).toEqual(['⌘', 'Z'])
    expect(getWorkbookShortcutParts('format-currency', 'MacIntel')).toEqual(['⇧', '⌘', '4'])
    expect(getWorkbookShortcutParts('undo', 'Win32')).toEqual(['Ctrl', 'Z'])
  })
})
