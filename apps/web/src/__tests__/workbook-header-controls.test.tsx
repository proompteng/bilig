// @vitest-environment jsdom
import { act } from 'react'
import { createRoot } from 'react-dom/client'
import { afterEach, describe, expect, it } from 'vitest'
import { WorkbookHeaderStatusChip } from '../workbook-header-controls.js'

afterEach(() => {
  document.body.innerHTML = ''
})

describe('WorkbookHeaderStatusChip', () => {
  it('renders the saved state as a readable compact indicator', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(<WorkbookHeaderStatusChip modeLabel="Live" syncLabel="Saved" tone="positive" />)
    })

    const status = host.querySelector<HTMLElement>("[data-testid='status-mode']")
    expect(status?.getAttribute('role')).toBe('status')
    expect(status?.getAttribute('class')).not.toContain('border')
    expect(status?.getAttribute('class')).not.toContain('bg-[')
    expect(status?.getAttribute('class')).not.toContain('rounded-')
    expect(status?.getAttribute('class')).toContain('max-[420px]:gap-0')
    expect(host.querySelector("[data-testid='status-label']")?.textContent).toBe('Saved')
    const sync = host.querySelector<HTMLElement>("[data-testid='status-sync']")
    expect(sync?.textContent).toBe('Saved')
    expect(sync?.hidden).toBe(true)
    expect(sync?.getAttribute('aria-hidden')).toBe('true')

    await act(async () => {
      root.unmount()
    })
  })

  it('keeps status text available while collapsing its visual width on tiny screens', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(<WorkbookHeaderStatusChip modeLabel="Live" syncLabel="Saving…" tone="progress" />)
    })

    const status = host.querySelector("[data-testid='status-mode']")
    const label = host.querySelector("[data-testid='status-label']")

    expect(status?.getAttribute('class')).toContain('max-[420px]:gap-0')
    expect(label?.textContent).toBe('Saving…')
    expect(label?.getAttribute('class')).toContain('max-[420px]:sr-only')
    const sync = host.querySelector<HTMLElement>("[data-testid='status-sync']")
    expect(sync?.textContent).toBe('Saving…')
    expect(sync?.hidden).toBe(true)

    await act(async () => {
      root.unmount()
    })
  })
})
