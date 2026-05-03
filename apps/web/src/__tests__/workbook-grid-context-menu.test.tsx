// @vitest-environment jsdom
import { act } from 'react'
import { createRoot } from 'react-dom/client'
import { afterEach, describe, expect, it, vi } from 'vitest'
import type { HeaderSelection, VisibleRegionState } from '../../../../packages/grid/src/gridPointer.js'
import type { GridSelection } from '../../../../packages/grid/src/gridTypes.js'
import { createColumnSliceSelection, createGridSelection, createRowSliceSelection } from '../../../../packages/grid/src/gridSelection.js'
import { WorkbookGridContextMenu } from '../../../../packages/grid/src/WorkbookGridContextMenu.js'
import { useWorkbookGridContextMenu } from '../../../../packages/grid/src/useWorkbookGridContextMenu.js'
import type { WorkbookGridContextMenuTarget } from '../../../../packages/grid/src/workbookGridContextMenuTarget.js'

afterEach(() => {
  document.body.innerHTML = ''
})

function ContextMenuHarness(props: {
  focusGrid?: (() => void) | undefined
  getGridSelection?: (() => GridSelection) | undefined
  headerSelection: HeaderSelection | null
  gridSelection?: GridSelection | undefined
  hiddenColumns?: Readonly<Record<number, true>> | undefined
  hiddenRows?: Readonly<Record<number, true>> | undefined
  isEditingCell?: boolean
  onCommitEdit?: (() => void) | undefined
  onDeleteColumn?: ((columnIndex: number, count: number) => void) | undefined
  onDeleteRow?: ((rowIndex: number, count: number) => void) | undefined
  onHideColumn?: ((columnIndex: number, hidden: boolean) => void) | undefined
  onHideRow?: ((rowIndex: number, hidden: boolean) => void) | undefined
  onInsertColumn?: ((columnIndex: number, count: number) => void) | undefined
  onInsertRow?: ((rowIndex: number, count: number) => void) | undefined
  onSetFreezePane?: ((rows: number, cols: number) => void) | undefined
  openTarget?: WorkbookGridContextMenuTarget | undefined
  onSelectionChange?: ((selection: GridSelection) => void) | undefined
  setGridSelection?: ((selection: GridSelection) => void) | undefined
  freezeRows?: number | undefined
  freezeCols?: number | undefined
}) {
  const visibleRegion: VisibleRegionState = {
    range: { x: 0, y: 0, width: 20, height: 20 },
    tx: 0,
    ty: 0,
    freezeRows: props.freezeRows ?? 0,
    freezeCols: props.freezeCols ?? 0,
  }
  const menu = useWorkbookGridContextMenu({
    focusGrid: props.focusGrid ?? (() => {}),
    getGridSelection: props.getGridSelection,
    getVisibleRegion() {
      return visibleRegion
    },
    isEditingCell: props.isEditingCell ?? false,
    onCommitEdit: props.onCommitEdit ?? (() => {}),
    onDeleteColumns: props.onDeleteColumn,
    onDeleteRows: props.onDeleteRow,
    onInsertColumns: props.onInsertColumn,
    onInsertRows: props.onInsertRow,
    onSelectionChange: props.onSelectionChange ?? (() => {}),
    onSetFreezePane: props.onSetFreezePane,
    hiddenColumnsByIndex: props.hiddenColumns,
    hiddenRowsByIndex: props.hiddenRows,
    onSetColumnHidden: props.onHideColumn,
    onSetRowHidden: props.onHideRow,
    resolveHeaderSelectionAtPointer() {
      return props.headerSelection
    },
    gridSelection: props.gridSelection ?? createGridSelection(4, 2),
    selectedCell: [4, 2],
    setGridSelection: props.setGridSelection ?? (() => {}),
  })

  return (
    <div data-testid="host" onContextMenuCapture={menu.handleHostContextMenuCapture}>
      <button
        data-testid="open-context-menu"
        type="button"
        onClick={() => {
          if (props.openTarget) {
            menu.openContextMenuForTarget(props.openTarget)
          }
        }}
      >
        Open
      </button>
      {menu.contextMenuState ? (
        <WorkbookGridContextMenu
          canUnfreezePanes={menu.canUnfreezePanes}
          menuRef={menu.menuRef}
          onClose={menu.closeContextMenu}
          onDeleteTarget={menu.deleteTarget}
          onFreezeTarget={menu.freezeTarget}
          onInsertAfterTarget={menu.insertAfterTarget}
          onInsertBeforeTarget={menu.insertBeforeTarget}
          onToggleTargetHidden={menu.toggleTargetHidden}
          onUnfreezePanes={menu.unfreezePanes}
          state={menu.contextMenuState}
        />
      ) : null}
    </div>
  )
}

describe('workbook grid context menu', () => {
  it('opens a row context menu and hides the targeted row', async () => {
    const onHideRow = vi.fn()
    const onSelectionChange = vi.fn()
    const setGridSelection = vi.fn()
    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(
        <ContextMenuHarness
          headerSelection={{ kind: 'row', index: 7 }}
          onHideRow={onHideRow}
          onSelectionChange={onSelectionChange}
          setGridSelection={setGridSelection}
        />,
      )
    })

    const target = host.querySelector("[data-testid='host']")
    await act(async () => {
      target?.dispatchEvent(
        new MouseEvent('contextmenu', {
          bubbles: true,
          button: 2,
          clientX: 120,
          clientY: 90,
        }),
      )
    })

    expect(host.querySelector("[data-testid='grid-context-menu']")).not.toBeNull()
    expect(onSelectionChange).toHaveBeenCalledTimes(1)

    const hideButton = host.querySelector("[data-testid='grid-context-action-hide-row']")
    await act(async () => {
      hideButton?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
    })

    expect(onHideRow).toHaveBeenCalledWith(7, true)
    expect(setGridSelection).toHaveBeenCalledTimes(1)
    expect(host.querySelector("[data-testid='grid-context-menu']")).toBeNull()

    await act(async () => {
      root.unmount()
    })
  })

  it('inserts a row above the targeted row', async () => {
    const onInsertRow = vi.fn()
    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(<ContextMenuHarness headerSelection={{ kind: 'row', index: 7 }} onHideRow={vi.fn()} onInsertRow={onInsertRow} />)
    })

    const target = host.querySelector("[data-testid='host']")
    await act(async () => {
      target?.dispatchEvent(
        new MouseEvent('contextmenu', {
          bubbles: true,
          button: 2,
          clientX: 120,
          clientY: 90,
        }),
      )
    })

    const insertButton = host.querySelector("[data-testid='grid-context-action-insert-before-row']")
    await act(async () => {
      insertButton?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
    })

    expect(onInsertRow).toHaveBeenCalledWith(7, 1)
    expect(host.querySelector("[data-testid='grid-context-menu']")).toBeNull()

    await act(async () => {
      root.unmount()
    })
  })

  it('freezes through the targeted row while preserving frozen columns', async () => {
    const onSetFreezePane = vi.fn()
    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(
        <ContextMenuHarness
          headerSelection={{ kind: 'row', index: 7 }}
          freezeCols={2}
          onHideRow={vi.fn()}
          onSetFreezePane={onSetFreezePane}
        />,
      )
    })

    const target = host.querySelector("[data-testid='host']")
    await act(async () => {
      target?.dispatchEvent(
        new MouseEvent('contextmenu', {
          bubbles: true,
          button: 2,
          clientX: 120,
          clientY: 90,
        }),
      )
    })

    const freezeButton = host.querySelector("[data-testid='grid-context-action-freeze-row']")
    expect(freezeButton).not.toBeNull()

    await act(async () => {
      freezeButton?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
    })

    expect(onSetFreezePane).toHaveBeenCalledWith(8, 2)
    expect(host.querySelector("[data-testid='grid-context-menu']")).toBeNull()

    await act(async () => {
      root.unmount()
    })
  })

  it('shows an unfreeze action when panes are frozen', async () => {
    const onSetFreezePane = vi.fn()
    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(
        <ContextMenuHarness
          headerSelection={{ kind: 'column', index: 3 }}
          freezeRows={1}
          freezeCols={2}
          onHideColumn={vi.fn()}
          onSetFreezePane={onSetFreezePane}
        />,
      )
    })

    const target = host.querySelector("[data-testid='host']")
    await act(async () => {
      target?.dispatchEvent(
        new MouseEvent('contextmenu', {
          bubbles: true,
          button: 2,
          clientX: 64,
          clientY: 24,
        }),
      )
    })

    const unfreezeButton = host.querySelector("[data-testid='grid-context-action-unfreeze-panes']")
    expect(unfreezeButton).not.toBeNull()

    await act(async () => {
      unfreezeButton?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
    })

    expect(onSetFreezePane).toHaveBeenCalledWith(0, 0)
    expect(host.querySelector("[data-testid='grid-context-menu']")).toBeNull()

    await act(async () => {
      root.unmount()
    })
  })

  it('shows unhide for a hidden row and restores the row', async () => {
    const onHideRow = vi.fn()
    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(<ContextMenuHarness headerSelection={{ kind: 'row', index: 7 }} hiddenRows={{ 7: true }} onHideRow={onHideRow} />)
    })

    const target = host.querySelector("[data-testid='host']")
    await act(async () => {
      target?.dispatchEvent(
        new MouseEvent('contextmenu', {
          bubbles: true,
          button: 2,
          clientX: 120,
          clientY: 90,
        }),
      )
    })

    const unhideButton = host.querySelector("[data-testid='grid-context-action-unhide-row']")
    expect(unhideButton).not.toBeNull()

    await act(async () => {
      unhideButton?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
    })

    expect(onHideRow).toHaveBeenCalledWith(7, false)
    expect(host.querySelector("[data-testid='grid-context-menu']")).toBeNull()

    await act(async () => {
      root.unmount()
    })
  })

  it('commits the active edit before opening a column context menu', async () => {
    const onCommitEdit = vi.fn()
    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(
        <ContextMenuHarness
          headerSelection={{ kind: 'column', index: 3 }}
          isEditingCell
          onCommitEdit={onCommitEdit}
          onHideColumn={vi.fn()}
        />,
      )
    })

    const target = host.querySelector("[data-testid='host']")
    await act(async () => {
      target?.dispatchEvent(
        new MouseEvent('contextmenu', {
          bubbles: true,
          button: 2,
          clientX: 64,
          clientY: 24,
        }),
      )
    })

    expect(onCommitEdit).toHaveBeenCalledTimes(1)
    expect(host.querySelector("[data-testid='grid-context-action-hide-column']")).not.toBeNull()

    await act(async () => {
      root.unmount()
    })
  })

  it('closes the context menu on escape', async () => {
    const focusGrid = vi.fn()
    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(<ContextMenuHarness focusGrid={focusGrid} headerSelection={{ kind: 'row', index: 1 }} onHideRow={vi.fn()} />)
    })

    const target = host.querySelector("[data-testid='host']")
    await act(async () => {
      target?.dispatchEvent(
        new MouseEvent('contextmenu', {
          bubbles: true,
          button: 2,
          clientX: 32,
          clientY: 60,
        }),
      )
    })

    expect(host.querySelector("[data-testid='grid-context-menu']")).not.toBeNull()
    expect(focusGrid).toHaveBeenCalledTimes(1)

    await act(async () => {
      window.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key: 'Escape' }))
    })

    expect(host.querySelector("[data-testid='grid-context-menu']")).toBeNull()
    expect(focusGrid).toHaveBeenCalledTimes(2)

    await act(async () => {
      root.unmount()
    })
  })

  it('moves focus into the menu on open and restores grid focus after outside close', async () => {
    const focusGrid = vi.fn()
    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(<ContextMenuHarness focusGrid={focusGrid} headerSelection={{ kind: 'column', index: 2 }} onHideColumn={vi.fn()} />)
    })

    const target = host.querySelector("[data-testid='host']")
    await act(async () => {
      target?.dispatchEvent(
        new MouseEvent('contextmenu', {
          bubbles: true,
          button: 2,
          clientX: 80,
          clientY: 18,
        }),
      )
    })

    const menuAction = host.querySelector("[data-testid='grid-context-action-insert-before-column']")
    expect(menuAction instanceof HTMLButtonElement).toBe(true)
    expect(document.activeElement).toBe(menuAction)
    expect(focusGrid).toHaveBeenCalledTimes(1)

    await act(async () => {
      window.dispatchEvent(new PointerEvent('pointerdown', { bubbles: true }))
    })

    expect(host.querySelector("[data-testid='grid-context-menu']")).toBeNull()
    expect(focusGrid).toHaveBeenCalledTimes(2)

    await act(async () => {
      root.unmount()
    })
  })

  it('opens a context menu programmatically for keyboard-triggered header actions', async () => {
    const onHideColumn = vi.fn()
    const onSelectionChange = vi.fn()
    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(
        <ContextMenuHarness
          headerSelection={null}
          onHideColumn={onHideColumn}
          onSelectionChange={onSelectionChange}
          openTarget={{ target: { kind: 'column', index: 2 }, x: 180, y: 24 }}
        />,
      )
    })

    const openButton = host.querySelector("[data-testid='open-context-menu']")
    await act(async () => {
      openButton?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
    })

    expect(host.querySelector("[data-testid='grid-context-menu']")).not.toBeNull()
    expect(onSelectionChange).toHaveBeenCalledTimes(1)

    const hideButton = host.querySelector("[data-testid='grid-context-action-hide-column']")
    await act(async () => {
      hideButton?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
    })

    expect(onHideColumn).toHaveBeenCalledWith(2, true)

    await act(async () => {
      root.unmount()
    })
  })

  it('deletes the targeted column', async () => {
    const onDeleteColumn = vi.fn()
    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(
        <ContextMenuHarness headerSelection={{ kind: 'column', index: 2 }} onDeleteColumn={onDeleteColumn} onHideColumn={vi.fn()} />,
      )
    })

    const target = host.querySelector("[data-testid='host']")
    await act(async () => {
      target?.dispatchEvent(
        new MouseEvent('contextmenu', {
          bubbles: true,
          button: 2,
          clientX: 64,
          clientY: 24,
        }),
      )
    })

    const deleteButton = host.querySelector("[data-testid='grid-context-action-delete-column']")
    await act(async () => {
      deleteButton?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
      await Promise.resolve()
    })

    expect(onDeleteColumn).toHaveBeenCalledWith(2, 1)

    await act(async () => {
      root.unmount()
    })
  })

  it('deletes an entire selected row range when deleting from inside it', async () => {
    const onDeleteRow = vi.fn()
    const onSelectionChange = vi.fn()
    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(
        <ContextMenuHarness
          headerSelection={{ kind: 'row', index: 4 }}
          gridSelection={createRowSliceSelection(0, 2, 5)}
          onDeleteRow={onDeleteRow}
          onHideRow={vi.fn()}
          onSelectionChange={onSelectionChange}
        />,
      )
    })

    const target = host.querySelector("[data-testid='host']")
    await act(async () => {
      target?.dispatchEvent(
        new MouseEvent('contextmenu', {
          bubbles: true,
          button: 2,
          clientX: 120,
          clientY: 90,
        }),
      )
    })

    const deleteButton = host.querySelector("[data-testid='grid-context-action-delete-row']")
    await act(async () => {
      deleteButton?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
      await Promise.resolve()
    })

    expect(onDeleteRow).toHaveBeenCalledWith(2, 4)
    expect(onSelectionChange).toHaveBeenCalledTimes(0)

    await act(async () => {
      root.unmount()
    })
  })

  it('uses the delete range from the moment the menu opened even if selection changes before delete', async () => {
    const onDeleteRow = vi.fn()
    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    const render = (selection: GridSelection) => (
      <ContextMenuHarness
        headerSelection={{ kind: 'row', index: 2 }}
        gridSelection={selection}
        onDeleteRow={onDeleteRow}
        onHideRow={vi.fn()}
      />
    )

    await act(async () => {
      root.render(render(createRowSliceSelection(0, 0, 6)))
    })

    const target = host.querySelector("[data-testid='host']")
    await act(async () => {
      target?.dispatchEvent(
        new MouseEvent('contextmenu', {
          bubbles: true,
          button: 2,
          clientX: 120,
          clientY: 90,
        }),
      )
    })

    await act(async () => {
      root.render(render(createRowSliceSelection(0, 2, 2)))
    })

    const deleteButton = host.querySelector("[data-testid='grid-context-action-delete-row']")
    await act(async () => {
      deleteButton?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
      await Promise.resolve()
    })

    expect(onDeleteRow).toHaveBeenCalledWith(0, 7)

    await act(async () => {
      root.unmount()
    })
  })

  it('prefers the live grid selection when opening a row context menu', async () => {
    const onDeleteRow = vi.fn()
    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)
    const staleSelection = createRowSliceSelection(0, 4, 4)
    const liveSelection = createRowSliceSelection(0, 2, 5)

    await act(async () => {
      root.render(
        <ContextMenuHarness
          headerSelection={{ kind: 'row', index: 4 }}
          getGridSelection={() => liveSelection}
          gridSelection={staleSelection}
          onDeleteRow={onDeleteRow}
          onHideRow={vi.fn()}
          openTarget={{
            target: { kind: 'row', index: 4 },
            x: 120,
            y: 90,
          }}
        />,
      )
    })

    const openButton = host.querySelector("[data-testid='open-context-menu']")
    await act(async () => {
      openButton?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
    })

    const deleteButton = host.querySelector("[data-testid='grid-context-action-delete-row']")
    await act(async () => {
      deleteButton?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
      await Promise.resolve()
    })

    expect(onDeleteRow).toHaveBeenCalledWith(2, 4)

    await act(async () => {
      root.unmount()
    })
  })

  it('deletes an entire selected column range when deleting from inside it', async () => {
    const onDeleteColumn = vi.fn()
    const onSelectionChange = vi.fn()
    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(
        <ContextMenuHarness
          headerSelection={{ kind: 'column', index: 4 }}
          gridSelection={createColumnSliceSelection(2, 5, 7)}
          onDeleteColumn={onDeleteColumn}
          onHideColumn={vi.fn()}
          onSelectionChange={onSelectionChange}
        />,
      )
    })

    const target = host.querySelector("[data-testid='host']")
    await act(async () => {
      target?.dispatchEvent(
        new MouseEvent('contextmenu', {
          bubbles: true,
          button: 2,
          clientX: 64,
          clientY: 24,
        }),
      )
    })

    const deleteButton = host.querySelector("[data-testid='grid-context-action-delete-column']")
    await act(async () => {
      deleteButton?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
      await Promise.resolve()
    })

    expect(onDeleteColumn).toHaveBeenCalledWith(2, 4)
    expect(onSelectionChange).toHaveBeenCalledTimes(0)

    await act(async () => {
      root.unmount()
    })
  })

  it('deletes only the matching disjoint column segment when selection contains multiple ranges', async () => {
    const onDeleteColumn = vi.fn()
    const onSelectionChange = vi.fn()
    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    const disjointColumnSelection = {
      current: {
        cell: [3, 7],
        range: { x: 1, y: 7, width: 1, height: 1 },
        rangeStack: [],
      },
      columns: {
        ranges: [
          [1, 4],
          [7, 10],
        ],
        first() {
          return 1
        },
        last() {
          return 9
        },
        hasIndex(index: number) {
          return (index >= 1 && index < 4) || (index >= 7 && index < 10)
        },
        length: 6,
      },
      rows: createGridSelection(0, 0).rows,
    }

    await act(async () => {
      root.render(
        <ContextMenuHarness
          headerSelection={{ kind: 'column', index: 7 }}
          gridSelection={disjointColumnSelection}
          onDeleteColumn={onDeleteColumn}
          onHideColumn={vi.fn()}
          onSelectionChange={onSelectionChange}
        />,
      )
    })

    const target = host.querySelector("[data-testid='host']")
    await act(async () => {
      target?.dispatchEvent(
        new MouseEvent('contextmenu', {
          bubbles: true,
          button: 2,
          clientX: 64,
          clientY: 24,
        }),
      )
    })

    const deleteButton = host.querySelector("[data-testid='grid-context-action-delete-column']")
    await act(async () => {
      deleteButton?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
      await Promise.resolve()
    })

    expect(onDeleteColumn).toHaveBeenCalledWith(7, 3)
    expect(onSelectionChange).toHaveBeenCalledTimes(0)

    await act(async () => {
      root.unmount()
    })
  })
})
