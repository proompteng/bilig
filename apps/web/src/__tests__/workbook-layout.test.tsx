// @vitest-environment jsdom
import { createRoot } from 'react-dom/client'
import { act } from 'react'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { ValueTag } from '@bilig/protocol'
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
})

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
    expect(resizeHandle?.className).toContain('w-4')
    expect(resizeHandle?.className).toContain('cursor-ew-resize')
    expect(resizeHandle?.className).toContain('-translate-x-2')
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
})
