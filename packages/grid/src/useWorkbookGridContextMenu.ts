import { useCallback, useEffect, useMemo, useRef, useState, type MouseEvent as ReactMouseEvent } from 'react'
import { createColumnSliceSelection, createRowSliceSelection } from './gridSelection.js'
import type { HeaderSelection, VisibleRegionState } from './gridPointer.js'
import type { Item, GridSelection } from './gridTypes.js'
import type { WorkbookGridContextMenuState } from './WorkbookGridContextMenu.js'
import type { WorkbookGridContextMenuTarget } from './workbookGridContextMenuTarget.js'

type SingleAxisSelection = Pick<GridSelection['columns'], 'length' | 'hasIndex' | 'ranges'>

function resolveSingleAxisDeleteRange(selection: SingleAxisSelection, targetIndex: number): { start: number; count: number } | null {
  for (const [start, endExclusive] of selection.ranges) {
    if (targetIndex >= start && targetIndex < endExclusive) {
      return { start, count: endExclusive - start }
    }
  }
  return null
}

export function useWorkbookGridContextMenu(input: {
  focusGrid(this: void): void
  getGridSelection?(this: void): GridSelection
  getVisibleRegion(this: void): VisibleRegionState
  hiddenColumnsByIndex?: Readonly<Record<number, true>> | undefined
  hiddenRowsByIndex?: Readonly<Record<number, true>> | undefined
  isEditingCell: boolean
  onCommitEdit(this: void): void
  onDeleteColumns?: ((startCol: number, count: number) => void | Promise<void>) | undefined
  onDeleteRows?: ((startRow: number, count: number) => void | Promise<void>) | undefined
  onInsertColumns?: ((startCol: number, count: number) => void) | undefined
  onInsertRows?: ((startRow: number, count: number) => void) | undefined
  onSelectionChange(this: void, selection: GridSelection): void
  onSetFreezePane?: ((rows: number, cols: number) => void) | undefined
  onSetColumnHidden?: ((columnIndex: number, hidden: boolean) => void) | undefined
  onSetRowHidden?: ((rowIndex: number, hidden: boolean) => void) | undefined
  resolveHeaderSelectionAtPointer(this: void, clientX: number, clientY: number, region?: VisibleRegionState): HeaderSelection | null
  gridSelection: GridSelection
  selectedCell: Item
  setGridSelection(this: void, selection: GridSelection): void
}) {
  const {
    focusGrid,
    getGridSelection,
    getVisibleRegion,
    hiddenColumnsByIndex,
    hiddenRowsByIndex,
    isEditingCell,
    onCommitEdit,
    onDeleteColumns,
    onDeleteRows,
    onInsertColumns,
    onInsertRows,
    onSelectionChange,
    onSetFreezePane,
    onSetColumnHidden,
    onSetRowHidden,
    resolveHeaderSelectionAtPointer,
    gridSelection,
    selectedCell,
    setGridSelection,
  } = input
  const [contextMenuState, setContextMenuState] = useState<WorkbookGridContextMenuState | null>(null)
  const menuRef = useRef<HTMLDivElement | null>(null)

  const isTargetHidden = useCallback(
    (target: { kind: 'row' | 'column'; index: number }) =>
      target.kind === 'row' ? hiddenRowsByIndex?.[target.index] === true : hiddenColumnsByIndex?.[target.index] === true,
    [hiddenColumnsByIndex, hiddenRowsByIndex],
  )

  const closeContextMenu = useCallback(() => {
    setContextMenuState(null)
    focusGrid()
  }, [focusGrid])

  useEffect(() => {
    if (!contextMenuState) {
      return
    }
    const firstMenuItem = menuRef.current?.querySelector<HTMLElement>('[role="menuitem"]')
    firstMenuItem?.focus()
  }, [contextMenuState])

  useEffect(() => {
    if (!contextMenuState) {
      return
    }

    const handlePointerDown = (event: PointerEvent) => {
      if (menuRef.current && event.target instanceof Node && menuRef.current.contains(event.target)) {
        return
      }
      closeContextMenu()
    }

    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.key === 'Escape') {
        event.preventDefault()
        closeContextMenu()
      }
    }

    window.addEventListener('pointerdown', handlePointerDown, true)
    window.addEventListener('keydown', handleKeyDown, true)
    return () => {
      window.removeEventListener('pointerdown', handlePointerDown, true)
      window.removeEventListener('keydown', handleKeyDown, true)
    }
  }, [closeContextMenu, contextMenuState])

  const toggleTargetHidden = useCallback(() => {
    if (!contextMenuState) {
      return
    }
    if (contextMenuState.target.kind === 'row') {
      onSetRowHidden?.(contextMenuState.target.index, !contextMenuState.target.hidden)
    } else {
      onSetColumnHidden?.(contextMenuState.target.index, !contextMenuState.target.hidden)
    }
    closeContextMenu()
  }, [closeContextMenu, contextMenuState, onSetColumnHidden, onSetRowHidden])

  const insertBeforeTarget = useCallback(() => {
    if (!contextMenuState) {
      return
    }
    if (contextMenuState.target.kind === 'row') {
      onInsertRows?.(contextMenuState.target.index, 1)
    } else {
      onInsertColumns?.(contextMenuState.target.index, 1)
    }
    closeContextMenu()
  }, [closeContextMenu, contextMenuState, onInsertColumns, onInsertRows])

  const insertAfterTarget = useCallback(() => {
    if (!contextMenuState) {
      return
    }
    if (contextMenuState.target.kind === 'row') {
      onInsertRows?.(contextMenuState.target.index + 1, 1)
    } else {
      onInsertColumns?.(contextMenuState.target.index + 1, 1)
    }
    closeContextMenu()
  }, [closeContextMenu, contextMenuState, onInsertColumns, onInsertRows])

  const deleteTarget = useCallback(() => {
    if (!contextMenuState) {
      return
    }
    void (async () => {
      try {
        if (contextMenuState.target.kind === 'row') {
          await onDeleteRows?.(contextMenuState.deleteRange.start, contextMenuState.deleteRange.count)
        } else {
          await onDeleteColumns?.(contextMenuState.deleteRange.start, contextMenuState.deleteRange.count)
        }
      } finally {
        closeContextMenu()
      }
    })()
  }, [closeContextMenu, contextMenuState, onDeleteColumns, onDeleteRows])

  const freezeTarget = useCallback(() => {
    if (!contextMenuState || !onSetFreezePane) {
      return
    }
    const visibleRegion = getVisibleRegion()
    if (contextMenuState.target.kind === 'row') {
      onSetFreezePane(contextMenuState.target.index + 1, visibleRegion.freezeCols ?? 0)
    } else {
      onSetFreezePane(visibleRegion.freezeRows ?? 0, contextMenuState.target.index + 1)
    }
    closeContextMenu()
  }, [closeContextMenu, contextMenuState, getVisibleRegion, onSetFreezePane])

  const unfreezePanes = useCallback(() => {
    if (!onSetFreezePane) {
      return
    }
    onSetFreezePane(0, 0)
    closeContextMenu()
  }, [closeContextMenu, onSetFreezePane])

  const openContextMenuForTarget = useCallback(
    ({ target, x, y }: WorkbookGridContextMenuTarget): boolean => {
      const canOpen =
        (target.kind === 'row' && (onSetRowHidden || onInsertRows || onDeleteRows || onSetFreezePane)) ||
        (target.kind === 'column' && (onSetColumnHidden || onInsertColumns || onDeleteColumns || onSetFreezePane))
      if (!canOpen) {
        return false
      }

      if (isEditingCell) {
        onCommitEdit()
      }
      focusGrid()

      const activeGridSelection = getGridSelection?.() ?? gridSelection
      const activeSelectedCell = activeGridSelection.current?.cell ?? selectedCell
      const selectedAxis = target.kind === 'row' ? activeGridSelection.rows : activeGridSelection.columns
      const orthogonalAxis = target.kind === 'row' ? activeGridSelection.columns : activeGridSelection.rows
      const keepCurrentSelection = selectedAxis.length > 0 && orthogonalAxis.length === 0 && selectedAxis.hasIndex(target.index)

      if (!keepCurrentSelection) {
        if (target.kind === 'row') {
          const nextSelection = createRowSliceSelection(activeSelectedCell[0], target.index, target.index)
          setGridSelection(nextSelection)
          onSelectionChange(nextSelection)
        } else {
          const nextSelection = createColumnSliceSelection(target.index, target.index, activeSelectedCell[1])
          setGridSelection(nextSelection)
          onSelectionChange(nextSelection)
        }
      }
      const deleteRange = keepCurrentSelection ? resolveSingleAxisDeleteRange(selectedAxis, target.index) : null
      const deleteRangeStart = deleteRange?.start ?? target.index
      const deleteRangeCount = deleteRange?.count ?? 1

      setContextMenuState({
        x,
        y,
        target: {
          ...target,
          hidden: isTargetHidden(target),
        },
        deleteRange: {
          kind: target.kind,
          start: deleteRangeStart,
          count: deleteRangeCount,
        },
      })
      return true
    },
    [
      focusGrid,
      getGridSelection,
      isTargetHidden,
      isEditingCell,
      gridSelection,
      onCommitEdit,
      onDeleteColumns,
      onDeleteRows,
      onInsertColumns,
      onInsertRows,
      onSelectionChange,
      onSetFreezePane,
      onSetColumnHidden,
      onSetRowHidden,
      selectedCell,
      setGridSelection,
    ],
  )

  const handleHostContextMenuCapture = useCallback(
    (event: ReactMouseEvent<HTMLDivElement>) => {
      const headerSelection = resolveHeaderSelectionAtPointer(event.clientX, event.clientY, getVisibleRegion())
      if (!headerSelection) {
        closeContextMenu()
        return
      }
      if (
        !openContextMenuForTarget({
          target: headerSelection,
          x: event.clientX,
          y: event.clientY,
        })
      ) {
        return
      }

      event.preventDefault()
      event.stopPropagation()
    },
    [closeContextMenu, getVisibleRegion, openContextMenuForTarget, resolveHeaderSelectionAtPointer],
  )

  return useMemo(
    () => ({
      closeContextMenu,
      canUnfreezePanes: (getVisibleRegion().freezeRows ?? 0) > 0 || (getVisibleRegion().freezeCols ?? 0) > 0,
      contextMenuState,
      deleteTarget,
      freezeTarget,
      handleHostContextMenuCapture,
      insertAfterTarget,
      insertBeforeTarget,
      toggleTargetHidden,
      unfreezePanes,
      menuRef,
      openContextMenuForTarget,
    }),
    [
      closeContextMenu,
      contextMenuState,
      deleteTarget,
      freezeTarget,
      handleHostContextMenuCapture,
      insertAfterTarget,
      insertBeforeTarget,
      toggleTargetHidden,
      unfreezePanes,
      openContextMenuForTarget,
      getVisibleRegion,
    ],
  )
}
