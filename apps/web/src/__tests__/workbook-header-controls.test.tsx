// @vitest-environment jsdom
import { act } from 'react'
import { createRoot } from 'react-dom/client'
import { afterEach, describe, expect, it } from 'vitest'
import { WorkbookHeaderStatusChip } from '../workbook-header-controls.js'

afterEach(() => {
  document.body.innerHTML = ''
})

describe('WorkbookHeaderStatusChip', () => {
  it('renders readable saved state beside its indicator', async () => {
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
    expect(status?.getAttribute('class')).toContain('gap-2')
    expect(status?.getAttribute('aria-label')).toBe('Workbook status: Live, Saved')
    expect(status?.getAttribute('title')).toBe('Live • Saved')
    expect(status?.textContent).toBe('Saved')
    expect(host.querySelector("[data-testid='status-label']")).toBeNull()
    const sync = host.querySelector<HTMLElement>("[data-testid='status-sync']")
    expect(sync?.textContent).toBe('Saved')
    expect(sync?.hidden).toBe(false)
    expect(sync?.getAttribute('aria-hidden')).toBeNull()

    await act(async () => {
      root.unmount()
    })
  })

  it('shows saving progress as visible status text', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(<WorkbookHeaderStatusChip modeLabel="Live" syncLabel="Saving…" tone="progress" />)
    })

    const status = host.querySelector<HTMLElement>("[data-testid='status-mode']")

    expect(status?.textContent).toBe('Saving…')
    expect(status?.getAttribute('aria-label')).toBe('Workbook status: Live, Saving…')
    expect(status?.getAttribute('title')).toBe('Live • Saving…')
    expect(host.querySelector("[data-testid='status-label']")).toBeNull()
    const sync = host.querySelector<HTMLElement>("[data-testid='status-sync']")
    expect(sync?.textContent).toBe('Saving…')
    expect(sync?.hidden).toBe(false)

    await act(async () => {
      root.unmount()
    })
  })
})
