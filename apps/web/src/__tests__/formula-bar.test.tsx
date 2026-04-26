// @vitest-environment jsdom
import { act, useState } from 'react'
import { flushSync } from 'react-dom'
import { createRoot } from 'react-dom/client'
import { afterEach, describe, expect, it, vi } from 'vitest'
import type { WorkbookDefinedNameSnapshot } from '@bilig/protocol'
import { FormulaBar } from '../../../../packages/grid/src/FormulaBar.js'

function FormulaBarHarness(props: {
  initialValue: string
  initialEditing?: boolean
  definedNames?: readonly WorkbookDefinedNameSnapshot[]
  selectionLabel?: string
  onAddressCommitResult?: (next: string) => boolean
}) {
  const [value, setValue] = useState(props.initialValue)
  const [isEditing, setIsEditing] = useState(props.initialEditing ?? true)
  return (
    <FormulaBar
      address="B2"
      definedNames={props.definedNames}
      isEditing={isEditing}
      onAddressCommit={(next) => props.onAddressCommitResult?.(next) ?? true}
      onBeginEdit={(seed) => {
        if (seed !== undefined) {
          setValue(seed)
        }
        setIsEditing(true)
      }}
      onCancel={() => {
        setIsEditing(false)
      }}
      onChange={(next) => {
        setValue(next)
      }}
      onCommit={() => {
        setIsEditing(false)
      }}
      resolvedValue=""
      selectionLabel={props.selectionLabel}
      sheetName="Sheet1"
      value={value}
    />
  )
}

function dispatchInputValue(input: HTMLInputElement, value: string) {
  flushSync(() => {
    input.value = value
    input.dispatchEvent(new Event('input', { bubbles: true }))
  })
}

function dispatchTextControlValue(input: HTMLInputElement | HTMLTextAreaElement, value: string) {
  flushSync(() => {
    input.value = value
    input.dispatchEvent(new Event('input', { bubbles: true }))
  })
}

afterEach(() => {
  document.body.innerHTML = ''
})

describe('FormulaBar', () => {
  it('renders autocomplete suggestions and argument hints for formula edits', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(<FormulaBarHarness initialValue="=IF(A1," />)
    })

    const autocomplete = host.querySelector("[data-testid='formula-autocomplete']")
    const argHint = host.querySelector("[data-testid='formula-arg-hint']")
    expect(autocomplete?.textContent).toContain('IF')
    expect(argHint?.textContent).toContain('value_if_true')

    await act(async () => {
      root.unmount()
    })
  })

  it('shows a defined name in the name box when the current selection matches it', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(
        <FormulaBarHarness
          initialEditing={false}
          initialValue=""
          definedNames={[
            {
              name: 'TaxRate',
              value: {
                kind: 'cell-ref',
                sheetName: 'Sheet1',
                address: 'B2',
              },
            },
          ]}
        />,
      )
    })

    const nameBox = host.querySelector<HTMLInputElement>("[data-testid='name-box']")
    expect(nameBox?.value).toBe('TaxRate')

    await act(async () => {
      root.unmount()
    })
  })

  it('keeps selection and resolved status hidden instead of rendering a visible metadata strip', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(<FormulaBarHarness initialEditing={false} initialValue="" selectionLabel="B2:D5" />)
    })

    const visibleMeta = host.querySelector("[data-testid='formula-bar-meta']")
    const selectionStatus = host.querySelector("[data-testid='status-selection']")
    const resolvedValue = host.querySelector("[data-testid='formula-resolved-value']")
    expect(visibleMeta).toBeNull()
    expect(selectionStatus?.textContent).toBe('Sheet1!B2:D5')
    expect(resolvedValue?.textContent).toBe('∅')

    await act(async () => {
      root.unmount()
    })
  })

  it('shows the active range-ref defined name in the name box when the current selection matches it', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(
        <FormulaBarHarness
          initialEditing={false}
          initialValue=""
          selectionLabel="B2:D5"
          definedNames={[
            {
              name: 'QuarterlyData',
              value: {
                kind: 'range-ref',
                sheetName: 'Sheet1',
                startAddress: 'B2',
                endAddress: 'D5',
              },
            },
          ]}
        />,
      )
    })

    const nameBox = host.querySelector<HTMLInputElement>("[data-testid='name-box']")
    expect(nameBox?.value).toBe('QuarterlyData')

    await act(async () => {
      root.unmount()
    })
  })

  it('focuses the name box from the Go To shortcut', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(<FormulaBarHarness initialEditing={false} initialValue="" selectionLabel="B2:D5" />)
    })

    await act(async () => {
      window.dispatchEvent(new KeyboardEvent('keydown', { key: 'g', ctrlKey: true, bubbles: true }))
    })

    const nameBox = host.querySelector<HTMLInputElement>("[data-testid='name-box']")
    expect(document.activeElement).toBe(nameBox)

    await act(async () => {
      root.unmount()
    })
  })

  it('shows an inline validation message for invalid name-box input and clears it on escape', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(
        <FormulaBarHarness initialEditing={false} initialValue="" onAddressCommitResult={(next) => next === 'B2'} selectionLabel="B2" />,
      )
    })

    const nameBox = host.querySelector<HTMLInputElement>("[data-testid='name-box']")
    expect(nameBox).not.toBeNull()
    if (!nameBox) {
      throw new Error('Expected name box input')
    }

    dispatchInputValue(nameBox, 'NotARange')
    await act(async () => {
      nameBox.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }))
    })

    expect(nameBox.getAttribute('aria-invalid')).toBe('true')
    const errorMessage = host.querySelector("[data-testid='name-box-error']")
    expect(errorMessage?.textContent).toBe('Unknown range or name')

    await act(async () => {
      nameBox.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }))
    })

    expect(nameBox.getAttribute('aria-invalid')).toBeNull()
    expect(nameBox.value).toBe('B2')

    await act(async () => {
      root.unmount()
    })
  })

  it('keeps the formula bar as a compact single-line input without inline action buttons', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(<FormulaBarHarness initialEditing={true} initialValue="=SUM(B2:B4)" />)
    })

    const commitButton = host.querySelector<HTMLButtonElement>("[data-testid='formula-commit']")
    const cancelButton = host.querySelector<HTMLButtonElement>("[data-testid='formula-cancel']")
    const input = host.querySelector<HTMLInputElement>("[data-testid='formula-input']")
    expect(commitButton).toBeNull()
    expect(cancelButton).toBeNull()
    expect(input).not.toBeNull()

    if (!input) {
      throw new Error('Expected formula input')
    }

    expect(input.tagName).toBe('INPUT')
    dispatchTextControlValue(input, '=SUM(B2:B5)')
    expect(input.value).toBe('=SUM(B2:B5)')

    await act(async () => {
      root.unmount()
    })
  })

  it('does not commit twice when Enter is followed by blur before parent editing state catches up', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const host = document.createElement('div')
    document.body.appendChild(host)
    const outsideButton = document.createElement('button')
    document.body.appendChild(outsideButton)
    const root = createRoot(host)
    const onCommit = vi.fn()

    await act(async () => {
      root.render(
        <FormulaBar
          address="B2"
          isEditing={true}
          onAddressCommit={() => true}
          onBeginEdit={() => {}}
          onCancel={() => {}}
          onChange={() => {}}
          onCommit={onCommit}
          resolvedValue=""
          sheetName="Sheet1"
          value="draft"
        />,
      )
    })

    const input = host.querySelector<HTMLInputElement>("[data-testid='formula-input']")
    expect(input).not.toBeNull()
    if (!input) {
      throw new Error('Expected formula input')
    }

    await act(async () => {
      input.focus()
      input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }))
      outsideButton.focus()
    })

    expect(onCommit).toHaveBeenCalledTimes(1)

    await act(async () => {
      root.unmount()
    })
    outsideButton.remove()
  })

  it('keeps the name box and formula frame flat without raised shadow chrome', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(<FormulaBarHarness initialEditing={false} initialValue="" selectionLabel="D4" />)
    })

    const nameBox = host.querySelector<HTMLInputElement>("[data-testid='name-box']")
    const formulaFrame = host.querySelector<HTMLDivElement>("[data-testid='formula-input-frame']")
    expect(nameBox).not.toBeNull()
    expect(formulaFrame).not.toBeNull()

    if (!nameBox || !formulaFrame) {
      throw new Error('Expected formula bar controls')
    }

    expect(nameBox.className).not.toContain('shadow-[')
    expect(formulaFrame.className).not.toContain('shadow-[')

    await act(async () => {
      root.unmount()
    })
  })
})
