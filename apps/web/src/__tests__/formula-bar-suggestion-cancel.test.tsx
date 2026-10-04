// @vitest-environment jsdom
import { act, useState } from 'react'
import { createRoot } from 'react-dom/client'
import { expect, it, vi } from 'vitest'
import { FormulaBar } from '../../../../packages/grid/src/FormulaBar.js'

it('accepts SUM with Tab and cancels the draft with one Escape without committing', async () => {
  ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
  const onCommit = vi.fn()
  const onCancel = vi.fn()
  function Harness() {
    const [draft, setDraft] = useState('=SU')
    const [editing, setEditing] = useState(true)
    return (
      <FormulaBar
        address="D2"
        sheetName="Sheet1"
        value={draft}
        resolvedValue="12"
        isEditing={editing}
        onAddressCommit={() => true}
        onBeginEdit={() => setEditing(true)}
        onChange={setDraft}
        onCommit={onCommit}
        onCancel={() => {
          onCancel()
          setEditing(false)
          setDraft('=SUM(3,4,5)')
        }}
      />
    )
  }
  const host = document.createElement('div')
  document.body.appendChild(host)
  const root = createRoot(host)
  try {
    await act(async () => root.render(<Harness />))
    const input = host.querySelector<HTMLTextAreaElement>("[data-testid='formula-input']")!
    await act(async () => {
      input.focus()
      input.setSelectionRange(3, 3)
      input.dispatchEvent(new Event('select', { bubbles: true }))
    })
    expect(host.querySelector("[data-testid='formula-autocomplete']")?.textContent).toContain('SUM')
    await act(async () => {
      input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Tab', bubbles: true, cancelable: true }))
    })
    expect(input.value).toBe('=SUM()')
    expect(host.querySelector("[data-testid='formula-autocomplete']")).toBeNull()
    await act(async () => {
      input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true, cancelable: true }))
    })
    expect(onCancel).toHaveBeenCalledOnce()
    expect(onCommit).not.toHaveBeenCalled()
    expect(input.value).toBe('=SUM(3,4,5)')
  } finally {
    await act(async () => root.unmount())
    host.remove()
  }
})
