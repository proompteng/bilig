// @vitest-environment jsdom
import { act } from 'react'
import { createRoot } from 'react-dom/client'
import { afterEach, expect, it } from 'vitest'
import { WorkbookSaveStatus } from '../workbook-save-status.js'
import type { WorkbookSaveState } from '../workbook-save-state.js'

afterEach(() => {
  document.body.innerHTML = ''
})

it.each<WorkbookSaveState>(['loading', 'saved', 'saving', 'local'])(
  'keeps %s feedback accessible without toolbar chrome',
  async (state) => {
    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)
    await act(async () => root.render(<WorkbookSaveStatus state={state} />))
    const status = host.querySelector('[role="status"]')
    expect(status?.className).toBe('sr-only')
    expect(status?.textContent).not.toBe('')
    expect(host.querySelector('[aria-hidden]')).toBeNull()
    await act(async () => root.unmount())
  },
)

it.each<WorkbookSaveState>(['error', 'offline', 'read-only'])('keeps %s visible when the user needs to know', async (state) => {
  const host = document.createElement('div')
  document.body.appendChild(host)
  const root = createRoot(host)
  await act(async () => root.render(<WorkbookSaveStatus state={state} />))
  const status = host.querySelector('[role="status"]')
  expect(status?.className).not.toContain('sr-only')
  expect(status?.textContent).not.toBe('')
  await act(async () => root.unmount())
})
