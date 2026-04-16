// @vitest-environment jsdom
import { act } from 'react'
import { createRoot } from 'react-dom/client'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { WorkbookAgentProgressRow } from '../workbook-agent-progress-row.js'

beforeEach(() => {
  ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
  vi.useFakeTimers()
})

afterEach(() => {
  vi.useRealTimers()
  vi.restoreAllMocks()
  document.body.innerHTML = ''
})

describe('WorkbookAgentProgressRow', () => {
  it('renders thinking text with a staged wave treatment', async () => {
    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(<WorkbookAgentProgressRow />)
    })

    expect(host.textContent).toContain('Thinking')
    expect(host.textContent).toContain('Reading selection')
    expect(host.textContent).toContain('Selection')
    expect(host.textContent).toContain('Context')
    expect(host.textContent).toContain('Draft')
    expect(host.querySelectorAll('.workbook-agent-thinking-letter')).toHaveLength(8)

    await act(async () => {
      vi.advanceTimersByTime(1500)
    })
    expect(host.textContent).toContain('Checking context')

    await act(async () => {
      vi.advanceTimersByTime(1500)
    })
    expect(host.textContent).toContain('Drafting reply')

    await act(async () => {
      root.unmount()
    })
  })
})
