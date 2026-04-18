// @vitest-environment jsdom
import { act } from 'react'
import { createRoot } from 'react-dom/client'
import { afterEach, describe, expect, it } from 'vitest'
import { WorkbookHeaderStatusChip } from '../workbook-header-controls.js'

afterEach(() => {
  document.body.innerHTML = ''
})

describe('WorkbookHeaderStatusChip', () => {
  it('renders the saved state as a dot-only indicator without visible label text', async () => {
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
    expect(host.querySelector("[data-testid='status-label']")).toBeNull()
    expect(host.querySelector("[data-testid='status-sync']")?.textContent).toBe('Saved')

    await act(async () => {
      root.unmount()
    })
  })

  it('keeps visible text for non-saved states', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(<WorkbookHeaderStatusChip modeLabel="Live" syncLabel="Saving…" tone="progress" />)
    })

    expect(host.querySelector("[data-testid='status-label']")?.textContent).toBe('Saving…')
    expect(host.querySelector("[data-testid='status-sync']")?.textContent).toBe('Saving…')

    await act(async () => {
      root.unmount()
    })
  })
})
