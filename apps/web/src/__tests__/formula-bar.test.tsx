// @vitest-environment jsdom
import { act, useState } from 'react'
import { flushSync } from 'react-dom'
import { createRoot } from 'react-dom/client'
import { afterEach, describe, expect, it, vi } from 'vitest'
import type { WorkbookDefinedNameSnapshot } from '@bilig/protocol'
import { FormulaBar } from '../../../../packages/grid/src/FormulaBar.js'

function FormulaBarHarness(props: {
  initialValue: string
  address?: string
  initialEditing?: boolean
  definedNames?: readonly WorkbookDefinedNameSnapshot[]
  selectionLabel?: string
  onAddressCommitResult?: (next: string) => boolean
  onAddressCommitSuccess?: () => void
  onFormulaCommitSuccess?: () => void
}) {
  const [value, setValue] = useState(props.initialValue)
  const [isEditing, setIsEditing] = useState(props.initialEditing ?? true)
  return (
    <FormulaBar
      address={props.address ?? 'B2'}
      definedNames={props.definedNames}
      isEditing={isEditing}
      onAddressCommit={(next) => props.onAddressCommitResult?.(next) ?? true}
      onAddressCommitSuccess={props.onAddressCommitSuccess}
      onFormulaCommitSuccess={props.onFormulaCommitSuccess}
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
    setNativeTextControlValue(input, value)
    input.dispatchEvent(new Event('input', { bubbles: true }))
  })
}

function dispatchTextControlValue(input: HTMLInputElement | HTMLTextAreaElement, value: string) {
  flushSync(() => {
    setNativeTextControlValue(input, value)
    input.dispatchEvent(new Event('input', { bubbles: true }))
  })
}

function setNativeTextControlValue(input: HTMLInputElement | HTMLTextAreaElement, value: string): void {
  const prototype = input instanceof HTMLTextAreaElement ? HTMLTextAreaElement.prototype : HTMLInputElement.prototype
  const descriptor = Object.getOwnPropertyDescriptor(prototype, 'value')
  if (typeof descriptor?.set !== 'function') {
    input.value = value
    return
  }
  // oxlint-disable-next-line typescript-eslint/unbound-method -- React input tests need the native setter with an explicit receiver.
  Reflect.apply(descriptor.set, input, [value])
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

  it('does not steal the Go To shortcut from an active text input', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(<FormulaBarHarness initialEditing initialValue="draft" selectionLabel="B2:D5" />)
    })

    const formulaInput = host.querySelector<HTMLTextAreaElement>("[data-testid='formula-input']")
    const nameBox = host.querySelector<HTMLInputElement>("[data-testid='name-box']")
    expect(formulaInput).not.toBeNull()
    expect(nameBox).not.toBeNull()
    formulaInput?.focus()

    await act(async () => {
      formulaInput?.dispatchEvent(new KeyboardEvent('keydown', { key: 'g', ctrlKey: true, bubbles: true, cancelable: true }))
    })

    expect(document.activeElement).toBe(formulaInput)
    expect(document.activeElement).not.toBe(nameBox)

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

  it('requests grid focus after a successful name-box commit only', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)
    const onAddressCommitSuccess = vi.fn()

    await act(async () => {
      root.render(
        <FormulaBarHarness
          initialEditing={false}
          initialValue=""
          onAddressCommitResult={(next) => next === 'C4'}
          onAddressCommitSuccess={onAddressCommitSuccess}
          selectionLabel="B2"
        />,
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

    expect(onAddressCommitSuccess).not.toHaveBeenCalled()

    nameBox.focus()
    dispatchInputValue(nameBox, 'C4')
    await act(async () => {
      nameBox.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }))
    })

    expect(onAddressCommitSuccess).toHaveBeenCalledTimes(1)
    expect(nameBox.getAttribute('aria-invalid')).toBeNull()
    expect(document.activeElement).not.toBe(nameBox)

    await act(async () => {
      root.unmount()
    })
  })

  it('returns keyboard ownership to the grid after committing the formula field with Enter', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)
    const onFormulaCommitSuccess = vi.fn()

    await act(async () => {
      root.render(<FormulaBarHarness initialEditing initialValue="draft" onFormulaCommitSuccess={onFormulaCommitSuccess} />)
    })

    const formulaInput = host.querySelector<HTMLTextAreaElement>("[data-testid='formula-input']")
    expect(formulaInput).not.toBeNull()
    if (!formulaInput) {
      throw new Error('Expected formula input')
    }

    formulaInput.focus()
    await act(async () => {
      formulaInput.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }))
    })

    expect(onFormulaCommitSuccess).toHaveBeenCalledTimes(1)
    expect(document.activeElement).not.toBe(formulaInput)

    await act(async () => {
      formulaInput.dispatchEvent(new FocusEvent('blur', { bubbles: true }))
    })

    expect(onFormulaCommitSuccess).toHaveBeenCalledTimes(1)

    await act(async () => {
      root.unmount()
    })
  })

  it('commits a first formula-bar draft on blur before the editing prop catches up', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)
    const onBeginEdit = vi.fn()
    const onChange = vi.fn()
    const onCommit = vi.fn()

    await act(async () => {
      root.render(
        <FormulaBar
          address="B2"
          isEditing={false}
          onAddressCommit={() => true}
          onBeginEdit={onBeginEdit}
          onCancel={() => {}}
          onChange={onChange}
          onCommit={onCommit}
          resolvedValue=""
          sheetName="Sheet1"
          value=""
        />,
      )
    })

    const formulaInput = host.querySelector<HTMLTextAreaElement>("[data-testid='formula-input']")
    expect(formulaInput).not.toBeNull()
    if (!formulaInput) {
      throw new Error('Expected formula input')
    }

    await act(async () => {
      formulaInput.focus()
      dispatchTextControlValue(formulaInput, 'fast blur draft')
    })
    await act(async () => {
      formulaInput.dispatchEvent(new FocusEvent('focusout', { bubbles: true }))
    })

    expect(onBeginEdit).toHaveBeenCalledTimes(1)
    expect(onBeginEdit).toHaveBeenCalledWith('')
    expect(onChange).toHaveBeenCalledWith('fast blur draft')
    expect(onCommit).toHaveBeenCalledTimes(1)
    expect(onCommit).toHaveBeenCalledWith('fast blur draft', { address: 'B2', sheetName: 'Sheet1' }, undefined)

    await act(async () => {
      root.unmount()
    })
  })

  it('commits formula-bar Tab navigation with a stable target and returns focus to the grid', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)
    const onCommit = vi.fn()
    const onFormulaCommitSuccess = vi.fn()

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
          onFormulaCommitSuccess={onFormulaCommitSuccess}
          resolvedValue=""
          sheetName="Sheet1"
          value="draft"
        />,
      )
    })

    const formulaInput = host.querySelector<HTMLTextAreaElement>("[data-testid='formula-input']")
    expect(formulaInput).not.toBeNull()
    if (!formulaInput) {
      throw new Error('Expected formula input')
    }

    formulaInput.focus()
    await act(async () => {
      formulaInput.dispatchEvent(new KeyboardEvent('keydown', { key: 'Tab', bubbles: true, cancelable: true }))
    })

    expect(onCommit).toHaveBeenCalledTimes(1)
    expect(onCommit).toHaveBeenCalledWith('draft', { address: 'B2', sheetName: 'Sheet1' }, [1, 0])
    expect(onFormulaCommitSuccess).toHaveBeenCalledTimes(1)
    expect(document.activeElement).not.toBe(formulaInput)

    await act(async () => {
      root.unmount()
    })
  })

  it('cancels formula-bar Escape and returns keyboard ownership to the grid', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)
    const onCancel = vi.fn()
    const onCommit = vi.fn()
    const onFormulaCommitSuccess = vi.fn()

    await act(async () => {
      root.render(
        <FormulaBar
          address="B2"
          isEditing={true}
          onAddressCommit={() => true}
          onBeginEdit={() => {}}
          onCancel={onCancel}
          onChange={() => {}}
          onCommit={onCommit}
          onFormulaCommitSuccess={onFormulaCommitSuccess}
          resolvedValue=""
          sheetName="Sheet1"
          value="draft"
        />,
      )
    })

    const formulaInput = host.querySelector<HTMLTextAreaElement>("[data-testid='formula-input']")
    expect(formulaInput).not.toBeNull()
    if (!formulaInput) {
      throw new Error('Expected formula input')
    }

    formulaInput.focus()
    await act(async () => {
      formulaInput.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true, cancelable: true }))
    })

    expect(onCancel).toHaveBeenCalledTimes(1)
    expect(onCommit).not.toHaveBeenCalled()
    expect(onFormulaCommitSuccess).toHaveBeenCalledTimes(1)
    expect(document.activeElement).not.toBe(formulaInput)

    await act(async () => {
      root.unmount()
    })
  })

  it('preserves dirty name-box input during late selection refreshes', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)
    const commits: string[] = []

    await act(async () => {
      root.render(
        <FormulaBarHarness
          address="C2"
          initialEditing={false}
          initialValue=""
          onAddressCommitResult={(next) => {
            commits.push(next)
            return true
          }}
          selectionLabel="C2"
        />,
      )
    })

    const nameBox = host.querySelector<HTMLInputElement>("[data-testid='name-box']")
    expect(nameBox).not.toBeNull()
    if (!nameBox) {
      throw new Error('Expected name box input')
    }

    nameBox.focus()
    dispatchInputValue(nameBox, 'D4')

    await act(async () => {
      root.render(
        <FormulaBarHarness
          address="B2"
          initialEditing={false}
          initialValue=""
          onAddressCommitResult={(next) => {
            commits.push(next)
            return true
          }}
          selectionLabel="B2"
        />,
      )
    })

    expect(nameBox.value).toBe('D4')

    await act(async () => {
      nameBox.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }))
    })

    expect(commits).toEqual(['D4'])

    await act(async () => {
      root.unmount()
    })
  })

  it('keeps the formula bar as a compact text field without inline action buttons', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(<FormulaBarHarness initialEditing={true} initialValue="=SUM(B2:B4)" />)
    })

    const commitButton = host.querySelector<HTMLButtonElement>("[data-testid='formula-commit']")
    const cancelButton = host.querySelector<HTMLButtonElement>("[data-testid='formula-cancel']")
    const input = host.querySelector<HTMLTextAreaElement>("[data-testid='formula-input']")
    expect(commitButton).toBeNull()
    expect(cancelButton).toBeNull()
    expect(input).not.toBeNull()

    if (!input) {
      throw new Error('Expected formula input')
    }

    expect(input.tagName).toBe('TEXTAREA')
    dispatchTextControlValue(input, '=SUM(B2:B5)')
    expect(input.value).toBe('=SUM(B2:B5)')

    await act(async () => {
      root.unmount()
    })
  })

  it('preserves multiline cell text in the formula field instead of flattening it', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(<FormulaBarHarness initialEditing={false} initialValue={'alpha\nbeta'} />)
    })

    const input = host.querySelector<HTMLTextAreaElement>("[data-testid='formula-input']")
    expect(input).not.toBeNull()

    if (!input) {
      throw new Error('Expected formula input')
    }

    expect(input.tagName).toBe('TEXTAREA')
    expect(input.value).toBe('alpha\nbeta')

    await act(async () => {
      root.unmount()
    })
  })

  it('inserts a line break in the formula field with Alt+Enter instead of committing', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)
    const onCommit = vi.fn()
    const onChange = vi.fn()

    await act(async () => {
      root.render(
        <FormulaBar
          address="B2"
          isEditing={true}
          onAddressCommit={() => true}
          onBeginEdit={() => {}}
          onCancel={() => {}}
          onChange={onChange}
          onCommit={onCommit}
          resolvedValue=""
          sheetName="Sheet1"
          value="alpha"
        />,
      )
    })

    const input = host.querySelector<HTMLTextAreaElement>("[data-testid='formula-input']")
    expect(input).not.toBeNull()

    if (!input) {
      throw new Error('Expected formula input')
    }

    input.focus()
    input.setSelectionRange(input.value.length, input.value.length)
    await act(async () => {
      input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', altKey: true, bubbles: true, cancelable: true }))
    })

    expect(onCommit).not.toHaveBeenCalled()
    expect(onChange).toHaveBeenCalledWith('alpha\n')
    expect(input.value).toBe('alpha')

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

    const input = host.querySelector<HTMLTextAreaElement>("[data-testid='formula-input']")
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
    expect(onCommit).toHaveBeenCalledWith('draft', { address: 'B2', sheetName: 'Sheet1' }, undefined)

    await act(async () => {
      root.unmount()
    })
    outsideButton.remove()
  })

  it('commits the live input value even when the controlled prop has not caught up', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const host = document.createElement('div')
    document.body.appendChild(host)
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
          value=""
        />,
      )
    })

    const input = host.querySelector<HTMLTextAreaElement>("[data-testid='formula-input']")
    expect(input).not.toBeNull()
    if (!input) {
      throw new Error('Expected formula input')
    }

    await act(async () => {
      input.focus()
      input.value = '=A1="HELLO"'
      input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }))
    })

    expect(onCommit).toHaveBeenCalledTimes(1)
    expect(onCommit).toHaveBeenCalledWith('=A1="HELLO"', { address: 'B2', sheetName: 'Sheet1' }, undefined)

    await act(async () => {
      root.unmount()
    })
  })

  it('keeps formula-bar blur commits pinned to the cell where editing started', async () => {
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

    const input = host.querySelector<HTMLTextAreaElement>("[data-testid='formula-input']")
    expect(input).not.toBeNull()
    if (!input) {
      throw new Error('Expected formula input')
    }

    input.focus()
    await act(async () => {
      root.render(
        <FormulaBar
          address="C3"
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

    const updatedInput = host.querySelector<HTMLTextAreaElement>("[data-testid='formula-input']")
    expect(updatedInput).not.toBeNull()
    if (!updatedInput) {
      throw new Error('Expected formula input')
    }

    await act(async () => {
      updatedInput.dispatchEvent(new FocusEvent('focusout', { bubbles: true, relatedTarget: outsideButton }))
    })

    expect(onCommit).toHaveBeenCalledTimes(1)
    expect(onCommit).toHaveBeenCalledWith('draft', { address: 'B2', sheetName: 'Sheet1' }, undefined)

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

  it('keeps the name box compact on phone-width formula bars', async () => {
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

    expect(nameBox.parentElement?.className).toContain('w-28')
    expect(nameBox.parentElement?.className).toContain('sm:w-[168px]')
    expect(formulaFrame.textContent).toContain('fx')
    expect(formulaFrame.querySelector('span')?.className).toContain('w-8')
    expect(formulaFrame.querySelector('span')?.className).toContain('sm:w-10')

    await act(async () => {
      root.unmount()
    })
  })
})
