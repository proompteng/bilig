// @vitest-environment jsdom
import { act, useState } from 'react'
import { createRoot } from 'react-dom/client'
import { afterEach, describe, expect, it } from 'vitest'
import type { WorkbookDefinedNameSnapshot } from '@bilig/protocol'
import { FormulaBar } from '../../../../packages/grid/src/FormulaBar.js'

function FormulaBarHarness(props: {
  initialValue: string
  initialEditing?: boolean
  definedNames?: readonly WorkbookDefinedNameSnapshot[]
  selectionLabel?: string
}) {
  const [value, setValue] = useState(props.initialValue)
  const [isEditing, setIsEditing] = useState(props.initialEditing ?? true)
  return (
    <FormulaBar
      address="B2"
      definedNames={props.definedNames}
      isEditing={isEditing}
      onAddressCommit={() => {}}
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

  it('publishes the canonical selection label for status consumers', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(<FormulaBarHarness initialEditing={false} initialValue="" selectionLabel="B2:D5" />)
    })

    const visibleSelectionLabel = host.querySelector("[data-testid='formula-selection-label']")
    const selectionStatus = host.querySelector("[data-testid='status-selection']")
    expect(visibleSelectionLabel?.textContent).toBe('B2:D5')
    expect(selectionStatus?.textContent).toBe('Sheet1!B2:D5')

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
})
