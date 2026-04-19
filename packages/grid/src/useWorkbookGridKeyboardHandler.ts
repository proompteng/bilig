import { useCallback, useEffect, type Dispatch, type MutableRefObject, type SetStateAction } from 'react'
import { formatAddress } from '@bilig/formula'
import { ValueTag } from '@bilig/protocol'
import {
  getNormalizedGridKeyboardKey,
  handleGridKey as dispatchGridKey,
  shouldHandleGridWindowKey,
  type GridKeyboardEventLike,
} from './gridClipboardKeyboardController.js'
import type { InternalClipboardRange } from './gridInternalClipboard.js'
import type { GridSelection } from './gridTypes.js'

export function useWorkbookGridKeyboardHandler(input: {
  applyClipboardValues: (target: readonly [number, number], values: readonly (readonly string[])[]) => void
  beginSelectedEdit: (seed?: string, selectionBehavior?: 'select-all' | 'caret-end') => void
  captureInternalClipboardSelection: () => InternalClipboardRange | null
  editorValue: string
  engine: { getCell(sheetName: string, address: string): { value: { tag: ValueTag } } }
  gridSelection: GridSelection
  hostRef: MutableRefObject<HTMLDivElement | null>
  internalClipboardRef: MutableRefObject<InternalClipboardRange | null>
  isEditingCell: boolean
  onCancelEdit: () => void
  onClearCell: () => void
  onCommitEdit: (movement?: readonly [-1 | 0 | 1, -1 | 0 | 1]) => void
  onEditorChange: (next: string) => void
  onSelectionChange: (selection: GridSelection) => void
  pendingClipboardCopySequenceRef: MutableRefObject<number>
  pendingKeyboardPasteSequenceRef: MutableRefObject<number>
  pendingTypeSeedRef: MutableRefObject<string | null>
  selectedCell: { col: number; row: number }
  setGridSelection: Dispatch<SetStateAction<GridSelection>>
  sheetName: string
  suppressNextNativePasteRef: MutableRefObject<boolean>
  toggleSelectedBooleanCell: () => void
}) {
  const handleGridKey = useCallback(
    (event: GridKeyboardEventLike) => {
      dispatchGridKey({
        applyClipboardValues: input.applyClipboardValues,
        beginSelectedEdit: input.beginSelectedEdit,
        captureInternalClipboardSelection: input.captureInternalClipboardSelection,
        editorValue: input.editorValue,
        event,
        gridSelection: input.gridSelection,
        internalClipboardRef: input.internalClipboardRef,
        isSelectedCellBoolean: () =>
          input.engine.getCell(input.sheetName, formatAddress(input.selectedCell.row, input.selectedCell.col)).value.tag ===
          ValueTag.Boolean,
        isEditingCell: input.isEditingCell,
        onCancelEdit: input.onCancelEdit,
        onClearCell: input.onClearCell,
        onCommitEdit: input.onCommitEdit,
        onEditorChange: input.onEditorChange,
        onSelectionChange: input.onSelectionChange,
        pendingClipboardCopySequenceRef: input.pendingClipboardCopySequenceRef,
        pendingKeyboardPasteSequenceRef: input.pendingKeyboardPasteSequenceRef,
        pendingTypeSeedRef: input.pendingTypeSeedRef,
        selectedCell: input.selectedCell,
        setGridSelection: input.setGridSelection,
        suppressNextNativePasteRef: input.suppressNextNativePasteRef,
        toggleSelectedBooleanCell: input.toggleSelectedBooleanCell,
      })
    },
    [input],
  )

  useEffect(() => {
    const handleWindowKeyDown = (event: KeyboardEvent) => {
      const normalizedKey = getNormalizedGridKeyboardKey(event.key, event.code)
      const activeElement = document.activeElement
      if (
        !shouldHandleGridWindowKey(
          {
            altKey: event.altKey,
            ctrlKey: event.ctrlKey,
            key: normalizedKey,
            metaKey: event.metaKey,
            shiftKey: event.shiftKey,
          },
          activeElement,
          input.hostRef.current,
        )
      ) {
        return
      }

      handleGridKey({
        altKey: event.altKey,
        cancel: () => {
          event.stopPropagation()
        },
        ctrlKey: event.ctrlKey,
        key: normalizedKey,
        metaKey: event.metaKey,
        shiftKey: event.shiftKey,
        preventDefault: () => event.preventDefault(),
      })
      if (event.defaultPrevented) {
        ;(event as KeyboardEvent & { __biligGridHandled?: boolean }).__biligGridHandled = true
      }
    }

    window.addEventListener('keydown', handleWindowKeyDown, true)
    return () => {
      window.removeEventListener('keydown', handleWindowKeyDown, true)
    }
  }, [handleGridKey, input.hostRef])

  return {
    handleGridKey,
  }
}
