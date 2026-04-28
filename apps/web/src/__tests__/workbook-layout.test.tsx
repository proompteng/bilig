// @vitest-environment jsdom
import { createRoot } from 'react-dom/client'
import { act } from 'react'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { ValueTag } from '@bilig/protocol'
import { parseCellAddress } from '@bilig/formula'
import type { GridEngineLike } from '@bilig/grid'
import { WorkbookView } from '../../../../packages/grid/src/WorkbookView.js'

vi.mock('../../../../packages/grid/src/FormulaBar.js', () => ({
  FormulaBar: () => <div data-testid="formula-bar" />,
}))

vi.mock('../../../../packages/grid/src/WorkbookGridSurface.js', () => ({
  WorkbookGridSurface: () => <div data-testid="grid-surface" />,
}))

vi.mock('../../../../packages/grid/src/WorkbookSheetTabs.js', () => ({
  WorkbookSheetTabs: (props: { trailingContent?: React.ReactNode }) => (
    <div data-testid="sheet-tabs">
      <div data-testid="sheet-tabs-trailing">{props.trailingContent}</div>
    </div>
  ),
}))

afterEach(() => {
  document.body.innerHTML = ''
  Object.defineProperty(window, 'innerWidth', {
    configurable: true,
    value: 1024,
  })
})

function createEngineWithCells(
  cells: Record<string, { tag: typeof ValueTag.Number; value: number } | { tag: typeof ValueTag.String; value: string }>,
): GridEngineLike {
  return {
    getCell: (_sheetName, address) => {
      const current = cells[address]
      if (!current) {
        return {
          sheetName: 'Sheet1',
          address,
          value: { tag: ValueTag.Empty },
          flags: 0,
          version: 0,
        }
      }
      return current.tag === ValueTag.Number
        ? {
            sheetName: 'Sheet1',
            address,
            value: { tag: ValueTag.Number, value: current.value },
            flags: 0,
            version: 1,
          }
        : {
            sheetName: 'Sheet1',
            address,
            value: { tag: ValueTag.String, value: current.value, stringId: 1 },
            flags: 0,
            version: 1,
          }
    },
    getCellStyle: () => undefined,
    subscribeCells: () => () => {},
    workbook: {
      getSheet: () => ({
        grid: {
          forEachCellEntry(listener) {
            Object.keys(cells).forEach((address, index) => {
              const parsed = parseCellAddress(address, 'Sheet1')
              listener(index, parsed.row, parsed.col)
            })
          },
        },
      }),
    },
  }
}

describe('workbook layout', () => {
  it('renders the side panel beside the spreadsheet using the controlled panel width', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const engine: GridEngineLike = {
      getCell: () => ({
        sheetName: 'Sheet1',
        address: 'A1',
        value: { tag: ValueTag.Empty },
        flags: 0,
        version: 0,
      }),
      getCellStyle: () => undefined,
      subscribeCells: () => () => {},
      workbook: {
        getSheet: () => undefined,
      },
    }

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(
        <WorkbookView
          engine={engine}
          sheetNames={['Sheet1']}
          sheetName="Sheet1"
          selectedAddr="A1"
          selectedCellSnapshot={{
            sheetName: 'Sheet1',
            address: 'A1',
            value: { tag: ValueTag.Empty },
            flags: 0,
            version: 0,
          }}
          selectionSnapshot={{
            sheetName: 'Sheet1',
            address: 'A1',
            kind: 'cell',
            range: {
              startAddress: 'A1',
              endAddress: 'A1',
            },
          }}
          editorValue=""
          editorSelectionBehavior="select-all"
          resolvedValue=""
          isEditing={false}
          isEditingCell={false}
          onSelectSheet={() => {}}
          onSelectionChange={() => {}}
          onAddressCommit={() => true}
          onBeginEdit={() => {}}
          onBeginFormulaEdit={() => {}}
          onEditorChange={() => {}}
          onCommitEdit={() => {}}
          onCancelEdit={() => {}}
          onClearCell={() => {}}
          onFillRange={() => {}}
          onCopyRange={() => {}}
          onMoveRange={() => {}}
          onPaste={() => {}}
          onSidePanelWidthChange={() => {}}
          sidePanelId="workbook-side-panel-doc-1"
          sidePanel={<div data-testid="assistant-panel">Assistant panel</div>}
          sidePanelWidth={384}
        />,
      )
    })

    const sidePanel = host.querySelector("[data-testid='workbook-side-panel']")
    const gridSurface = host.querySelector("[data-testid='grid-surface']")
    const shell = host.querySelector("[data-testid='workbook-shell']")
    expect(sidePanel).not.toBeNull()
    expect(shell?.className).toContain('h-full')
    expect(shell?.className).not.toContain('h-screen')
    expect(sidePanel?.getAttribute('id')).toBe('workbook-side-panel-doc-1')
    expect(sidePanel?.textContent).toContain('Assistant panel')
    expect(sidePanel instanceof HTMLElement ? sidePanel.style.width : null).toBe('384px')
    const resizeHandle = host.querySelector("[data-testid='workbook-side-panel-resize-handle']")
    expect(resizeHandle).not.toBeNull()
    expect(sidePanel?.className).toContain('max-[900px]:absolute')
    expect(sidePanel?.className).toContain('max-[900px]:right-0')
    expect(resizeHandle?.className).toContain('w-4')
    expect(resizeHandle?.className).toContain('cursor-ew-resize')
    expect(resizeHandle?.className).toContain('-translate-x-2')
    expect(resizeHandle?.className).toContain('max-[900px]:hidden')
    expect(resizeHandle?.className).toContain('after:bg-[var(--wb-border)]')
    expect(gridSurface instanceof Node).toBe(true)
    expect(sidePanel instanceof Node).toBe(true)
    expect(gridSurface instanceof Node && sidePanel instanceof Node ? gridSurface.compareDocumentPosition(sidePanel) : 0).toBe(
      Node.DOCUMENT_POSITION_FOLLOWING,
    )

    await act(async () => {
      root.unmount()
    })
  })

  it('expands the assistant overlay to the full viewport width on phone-sized screens', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    Object.defineProperty(window, 'innerWidth', {
      configurable: true,
      value: 390,
    })

    const engine: GridEngineLike = {
      getCell: () => ({
        sheetName: 'Sheet1',
        address: 'A1',
        value: { tag: ValueTag.Empty },
        flags: 0,
        version: 0,
      }),
      getCellStyle: () => undefined,
      subscribeCells: () => () => {},
      workbook: {
        getSheet: () => undefined,
      },
    }

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(
        <WorkbookView
          engine={engine}
          sheetNames={['Sheet1']}
          sheetName="Sheet1"
          selectedAddr="A1"
          selectedCellSnapshot={{
            sheetName: 'Sheet1',
            address: 'A1',
            value: { tag: ValueTag.Empty },
            flags: 0,
            version: 0,
          }}
          selectionSnapshot={{
            sheetName: 'Sheet1',
            address: 'A1',
            kind: 'cell',
            range: {
              startAddress: 'A1',
              endAddress: 'A1',
            },
          }}
          editorValue=""
          editorSelectionBehavior="select-all"
          resolvedValue=""
          isEditing={false}
          isEditingCell={false}
          onSelectSheet={() => {}}
          onSelectionChange={() => {}}
          onAddressCommit={() => true}
          onBeginEdit={() => {}}
          onBeginFormulaEdit={() => {}}
          onEditorChange={() => {}}
          onCommitEdit={() => {}}
          onCancelEdit={() => {}}
          onClearCell={() => {}}
          onFillRange={() => {}}
          onCopyRange={() => {}}
          onMoveRange={() => {}}
          onPaste={() => {}}
          onSidePanelWidthChange={() => {}}
          sidePanel={<div data-testid="assistant-panel">Assistant panel</div>}
          sidePanelWidth={280}
        />,
      )
    })

    const sidePanel = host.querySelector("[data-testid='workbook-side-panel']")
    const resizeHandle = host.querySelector("[data-testid='workbook-side-panel-resize-handle']")
    expect(sidePanel instanceof HTMLElement ? sidePanel.style.width : null).toBe('390px')
    expect(sidePanel instanceof HTMLElement ? sidePanel.style.flexBasis : null).toBe('390px')
    expect(sidePanel?.className).toContain('max-[900px]:max-w-[100vw]')
    expect(sidePanel?.className).not.toContain('calc(100vw-56px)')
    expect(resizeHandle?.className).toContain('max-[900px]:hidden')

    await act(async () => {
      root.unmount()
    })
  })

  it('passes raw selection coordinates to the footer without the Selection label', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const engine: GridEngineLike = {
      getCell: () => ({
        sheetName: 'Sheet1',
        address: 'A1',
        value: { tag: ValueTag.Empty },
        flags: 0,
        version: 0,
      }),
      getCellStyle: () => undefined,
      subscribeCells: () => () => {},
      workbook: {
        getSheet: () => undefined,
      },
    }

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(
        <WorkbookView
          engine={engine}
          sheetNames={['Sheet1']}
          sheetName="Sheet1"
          selectedAddr="A3"
          selectedCellSnapshot={{
            sheetName: 'Sheet1',
            address: 'A3',
            value: { tag: ValueTag.Empty },
            flags: 0,
            version: 0,
          }}
          selectionSnapshot={{
            sheetName: 'Sheet1',
            address: 'A3',
            kind: 'range',
            range: {
              startAddress: 'A3',
              endAddress: 'C10',
            },
          }}
          editorValue=""
          editorSelectionBehavior="select-all"
          resolvedValue=""
          isEditing={false}
          isEditingCell={false}
          onSelectSheet={() => {}}
          onSelectionChange={() => {}}
          onAddressCommit={() => true}
          onBeginEdit={() => {}}
          onBeginFormulaEdit={() => {}}
          onEditorChange={() => {}}
          onCommitEdit={() => {}}
          onCancelEdit={() => {}}
          onClearCell={() => {}}
          onFillRange={() => {}}
          onCopyRange={() => {}}
          onMoveRange={() => {}}
          onPaste={() => {}}
        />,
      )
    })

    const trailing = host.querySelector("[data-testid='sheet-tabs-trailing']")
    expect(trailing?.textContent).toBe('A3:C10')
    expect(trailing?.textContent).not.toContain('Selection')

    await act(async () => {
      root.unmount()
    })
  })

  it('shows a sum summary chip in the footer for numeric range selections', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const engine = createEngineWithCells({
      A3: { tag: ValueTag.Number, value: 10 },
      B3: { tag: ValueTag.Number, value: 20 },
      C3: { tag: ValueTag.String, value: 'ignore' },
    })

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(
        <WorkbookView
          engine={engine}
          sheetNames={['Sheet1']}
          sheetName="Sheet1"
          selectedAddr="A3"
          selectedCellSnapshot={{
            sheetName: 'Sheet1',
            address: 'A3',
            value: { tag: ValueTag.Number, value: 10 },
            flags: 0,
            version: 1,
          }}
          selectionSnapshot={{
            sheetName: 'Sheet1',
            address: 'A3',
            kind: 'range',
            range: {
              startAddress: 'A3',
              endAddress: 'C3',
            },
          }}
          editorValue=""
          editorSelectionBehavior="select-all"
          resolvedValue=""
          isEditing={false}
          isEditingCell={false}
          onSelectSheet={() => {}}
          onSelectionChange={() => {}}
          onAddressCommit={() => true}
          onBeginEdit={() => {}}
          onBeginFormulaEdit={() => {}}
          onEditorChange={() => {}}
          onCommitEdit={() => {}}
          onCancelEdit={() => {}}
          onClearCell={() => {}}
          onFillRange={() => {}}
          onCopyRange={() => {}}
          onMoveRange={() => {}}
          onPaste={() => {}}
        />,
      )
    })

    const summary = host.querySelector("[data-testid='workbook-selection-summary']")
    expect(summary?.textContent).toContain('Sum: 30.00')

    await act(async () => {
      root.unmount()
    })
  })

  it('switches footer aggregate metrics from the chevron menu', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const engine = createEngineWithCells({
      A3: { tag: ValueTag.Number, value: 10 },
      B3: { tag: ValueTag.Number, value: 20 },
    })

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(
        <WorkbookView
          engine={engine}
          sheetNames={['Sheet1']}
          sheetName="Sheet1"
          selectedAddr="A3"
          selectedCellSnapshot={{
            sheetName: 'Sheet1',
            address: 'A3',
            value: { tag: ValueTag.Number, value: 10 },
            flags: 0,
            version: 1,
          }}
          selectionSnapshot={{
            sheetName: 'Sheet1',
            address: 'A3',
            kind: 'range',
            range: {
              startAddress: 'A3',
              endAddress: 'B3',
            },
          }}
          editorValue=""
          editorSelectionBehavior="select-all"
          resolvedValue=""
          isEditing={false}
          isEditingCell={false}
          onSelectSheet={() => {}}
          onSelectionChange={() => {}}
          onAddressCommit={() => true}
          onBeginEdit={() => {}}
          onBeginFormulaEdit={() => {}}
          onEditorChange={() => {}}
          onCommitEdit={() => {}}
          onCancelEdit={() => {}}
          onClearCell={() => {}}
          onFillRange={() => {}}
          onCopyRange={() => {}}
          onMoveRange={() => {}}
          onPaste={() => {}}
        />,
      )
    })

    const trigger = host.querySelector("[data-testid='workbook-selection-status-trigger']")
    expect(trigger?.textContent).toContain('Sum: 30.00')

    await act(async () => {
      trigger?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
    })

    const avgOption = document.querySelector("[data-testid='workbook-selection-status-option-avg']")
    const avgOptionText = document.querySelector("[data-testid='workbook-selection-status-option-avg-text']")
    const avgIndicatorSlot = document.querySelector("[data-testid='workbook-selection-status-option-avg-indicator-slot']")
    const sumIndicatorSlot = document.querySelector("[data-testid='workbook-selection-status-option-sum-indicator-slot']")
    const menu = document.querySelector("[data-testid='workbook-selection-status-menu']")
    expect(avgOption?.textContent).toContain('Avg: 15.00')
    expect(menu?.getAttribute('class')).toContain('w-max')
    expect(avgOption?.getAttribute('class')).toContain('whitespace-nowrap')
    expect(avgOptionText?.getAttribute('class')).toContain('whitespace-nowrap')
    expect(avgIndicatorSlot).not.toBeNull()
    expect(sumIndicatorSlot).not.toBeNull()

    await act(async () => {
      avgOption?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
    })

    expect(trigger?.textContent).toContain('Avg: 15.00')

    await act(async () => {
      root.unmount()
    })
  })
})
