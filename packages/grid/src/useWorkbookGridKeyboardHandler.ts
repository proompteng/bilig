import { useCallback, useEffect, useLayoutEffect, useMemo, useRef, type Dispatch, type MutableRefObject, type SetStateAction } from 'react'
import { formatAddress } from '@bilig/formula'
import {
  getNormalizedGridKeyboardKey,
  handleGridKey as dispatchGridKey,
  isGridKeyboardEditableTarget,
  shouldHandleGridWindowKey,
  shouldHandleGridSurfaceKey,
  shouldSuppressWorkbookChromeClearKey,
  shouldSuppressWorkbookChromeSelectionKeyUp,
  type GridKeyboardEventLike,
} from './gridClipboardKeyboardController.js'
import type { GridEngineLike } from './grid-engine.js'
import { isToggleableBooleanCellSnapshot } from './gridInteractionCommands.js'
import type { InternalClipboardRange } from './gridInternalClipboard.js'
import { createGridNavigationResolver } from './gridNavigation.js'
import type { VisibleRegionState } from './gridPointer.js'
import type { GridSelection, GridSelectionSnapshot } from './gridTypes.js'

type BeginSelectedEdit = (seed?: string, selectionBehavior?: 'select-all' | 'caret-end') => void

interface DeferredBeginEditScheduler {
  beginImmediate(seed: string | undefined, selectionBehavior: 'select-all' | 'caret-end'): void
  cancel(): void
  consume(): { readonly seed: string; readonly selectionBehavior: 'select-all' | 'caret-end' } | null
  schedule(seed: string, selectionBehavior: 'select-all' | 'caret-end'): void
}

type PendingTypedEditResolution =
  | { readonly kind: 'begin'; readonly seed: (pendingSeed: string) => string }
  | { readonly kind: 'cancel' }
  | { readonly kind: 'commit'; readonly movement: readonly [-1 | 0 | 1, -1 | 0 | 1] }

function isPendingTypedEditTextContinuation(event: GridKeyboardEventLike): boolean {
  return event.key.length === 1 && !event.altKey && !event.ctrlKey && !event.metaKey && !(event.key === ' ' && event.shiftKey)
}

function resolvePendingTypedEdit(event: GridKeyboardEventLike): PendingTypedEditResolution | null {
  if (isPendingTypedEditTextContinuation(event) || !shouldHandleGridSurfaceKey(event)) {
    return null
  }

  const hasPlainGridModifierState = !event.altKey && !event.ctrlKey && !event.metaKey
  if (event.key === 'Enter' && hasPlainGridModifierState) {
    return { kind: 'commit', movement: [0, event.shiftKey ? -1 : 1] }
  }
  if (event.key === 'Tab' && hasPlainGridModifierState) {
    return { kind: 'commit', movement: [event.shiftKey ? -1 : 1, 0] }
  }
  if (event.key === 'ArrowUp' && hasPlainGridModifierState) {
    return { kind: 'commit', movement: [0, -1] }
  }
  if (event.key === 'ArrowDown' && hasPlainGridModifierState) {
    return { kind: 'commit', movement: [0, 1] }
  }
  if (event.key === 'ArrowLeft' && hasPlainGridModifierState) {
    return { kind: 'commit', movement: [-1, 0] }
  }
  if (event.key === 'ArrowRight' && hasPlainGridModifierState) {
    return { kind: 'commit', movement: [1, 0] }
  }
  if (event.key === 'Escape' && hasPlainGridModifierState) {
    return { kind: 'cancel' }
  }
  if (event.key === 'Backspace' && hasPlainGridModifierState) {
    return { kind: 'begin', seed: (pendingSeed) => pendingSeed.slice(0, -1) }
  }
  return { kind: 'begin', seed: (pendingSeed) => pendingSeed }
}

export function createDeferredBeginEditScheduler(input: {
  readonly beginSelectedEdit: BeginSelectedEdit
  readonly cancelAnimationFrame?: ((handle: number) => void) | undefined
  readonly requestAnimationFrame?: ((callback: FrameRequestCallback) => number) | undefined
}): DeferredBeginEditScheduler {
  let pendingFrame: number | null = null
  let pendingEdit: { readonly seed: string; readonly selectionBehavior: 'select-all' | 'caret-end' } | null = null

  const cancel = (): void => {
    if (pendingFrame !== null) {
      input.cancelAnimationFrame?.(pendingFrame)
    }
    pendingFrame = null
    pendingEdit = null
  }
  const flush = (): void => {
    const nextEdit = pendingEdit
    pendingFrame = null
    pendingEdit = null
    if (nextEdit) {
      input.beginSelectedEdit(nextEdit.seed, nextEdit.selectionBehavior)
    }
  }

  return {
    beginImmediate(seed, selectionBehavior) {
      cancel()
      input.beginSelectedEdit(seed, selectionBehavior)
    },
    cancel,
    consume() {
      const nextEdit = pendingEdit
      cancel()
      return nextEdit
    },
    schedule(seed, selectionBehavior) {
      pendingEdit = { seed, selectionBehavior }
      if (pendingFrame !== null) {
        return
      }
      if (!input.requestAnimationFrame) {
        flush()
        return
      }
      pendingFrame = input.requestAnimationFrame(flush)
    },
  }
}

export function useWorkbookGridKeyboardHandler(input: {
  applyClipboardValues: (
    target: readonly [number, number],
    values: readonly (readonly string[])[],
    options?: { readonly pasteValuesOnly?: boolean | undefined },
  ) => void
  beginSelectedEdit: (seed?: string, selectionBehavior?: 'select-all' | 'caret-end') => void
  captureInternalClipboardSelection: () => InternalClipboardRange | null
  editorValue: string
  engine: GridEngineLike
  getVisibleRegion?: (() => VisibleRegionState) | undefined
  gridSelection: GridSelection
  getGridSelection?: (() => GridSelection) | undefined
  hostRef: MutableRefObject<HTMLDivElement | null>
  internalClipboardRef: MutableRefObject<InternalClipboardRange | null>
  isEditingCell: boolean
  onCancelEdit: () => void
  onClearCell: (selection?: GridSelectionSnapshot) => void
  onCommitEdit: (movement?: readonly [-1 | 0 | 1, -1 | 0 | 1], valueOverride?: string) => void
  onDeleteColumns?: ((startCol: number, count: number) => void | Promise<void>) | undefined
  onDeleteRows?: ((startRow: number, count: number) => void | Promise<void>) | undefined
  onEditorChange: (next: string) => void
  onFillRange: (sourceStartAddr: string, sourceEndAddr: string, targetStartAddr: string, targetEndAddr: string) => void
  onSelectionChange: (selection: GridSelection) => void
  scrollActiveCellIntoView: () => void
  pendingClipboardCopySequenceRef: MutableRefObject<number>
  pendingKeyboardPasteSequenceRef: MutableRefObject<number>
  pendingTypeSeedRef: MutableRefObject<string | null>
  selectedCell: { col: number; row: number }
  setGridSelection: Dispatch<SetStateAction<GridSelection>>
  sheetName: string
  suppressNextNativePasteRef: MutableRefObject<boolean>
  toggleSelectedBooleanCell: () => void
}) {
  const beginSelectedEditRef = useRef(input.beginSelectedEdit)
  useLayoutEffect(() => {
    beginSelectedEditRef.current = input.beginSelectedEdit
  }, [input.beginSelectedEdit])

  const deferredBeginEditScheduler = useMemo(
    () =>
      createDeferredBeginEditScheduler({
        beginSelectedEdit: (seed, selectionBehavior) => {
          beginSelectedEditRef.current(seed, selectionBehavior)
        },
        cancelAnimationFrame: typeof window === 'undefined' ? undefined : (handle: number) => window.cancelAnimationFrame(handle),
        requestAnimationFrame:
          typeof window === 'undefined' ? undefined : (callback: FrameRequestCallback) => window.requestAnimationFrame(callback),
      }),
    [],
  )
  useEffect(() => () => deferredBeginEditScheduler.cancel(), [deferredBeginEditScheduler])

  const handleGridKey = useCallback(
    (event: GridKeyboardEventLike) => {
      const pendingTypedEdit = resolvePendingTypedEdit(event)
      if (!input.isEditingCell && pendingTypedEdit) {
        const pendingEdit = deferredBeginEditScheduler.consume()
        if (pendingEdit) {
          input.pendingTypeSeedRef.current = null
          event.preventDefault()
          event.cancel?.()
          if (pendingTypedEdit.kind === 'commit') {
            input.onCommitEdit(pendingTypedEdit.movement, pendingEdit.seed)
          } else if (pendingTypedEdit.kind === 'cancel') {
            input.onCancelEdit()
          } else {
            deferredBeginEditScheduler.beginImmediate(pendingTypedEdit.seed(pendingEdit.seed), pendingEdit.selectionBehavior)
          }
          return
        }
      }
      const gridSelection = input.getGridSelection?.() ?? input.gridSelection
      const selectedCell = gridSelection.current?.cell ?? [input.selectedCell.col, input.selectedCell.row]
      const visibleRegion = input.getVisibleRegion?.()
      dispatchGridKey({
        applyClipboardValues: input.applyClipboardValues,
        beginSelectedEdit: (seed, selectionBehavior = 'caret-end') => {
          if (seed === undefined) {
            deferredBeginEditScheduler.beginImmediate(seed, selectionBehavior)
            return
          }
          deferredBeginEditScheduler.schedule(seed, selectionBehavior)
        },
        captureInternalClipboardSelection: input.captureInternalClipboardSelection,
        editorValue: input.editorValue,
        event,
        gridSelection,
        internalClipboardRef: input.internalClipboardRef,
        isSelectedCellBoolean: () =>
          isToggleableBooleanCellSnapshot(input.engine.getCell(input.sheetName, formatAddress(selectedCell[1], selectedCell[0]))),
        isEditingCell: input.isEditingCell,
        onCancelEdit: input.onCancelEdit,
        onClearCell: input.onClearCell,
        onCommitEdit: input.onCommitEdit,
        onDeleteColumns: input.onDeleteColumns,
        onDeleteRows: input.onDeleteRows,
        onEditorChange: input.onEditorChange,
        onFillRange: input.onFillRange,
        onSelectionChange: input.onSelectionChange,
        navigation: createGridNavigationResolver({
          engine: input.engine,
          sheetName: input.sheetName,
        }),
        pageJumpRows: visibleRegion ? Math.max(1, visibleRegion.range.height - 1) : null,
        scrollActiveCellIntoView: input.scrollActiveCellIntoView,
        pendingClipboardCopySequenceRef: input.pendingClipboardCopySequenceRef,
        pendingKeyboardPasteSequenceRef: input.pendingKeyboardPasteSequenceRef,
        pendingTypeSeedRef: input.pendingTypeSeedRef,
        selectedCell: { col: selectedCell[0], row: selectedCell[1] },
        setGridSelection: input.setGridSelection,
        sheetName: input.sheetName,
        suppressNextNativePasteRef: input.suppressNextNativePasteRef,
        toggleSelectedBooleanCell: input.toggleSelectedBooleanCell,
      })
    },
    [deferredBeginEditScheduler, input],
  )

  useEffect(() => {
    const commitPendingTypedEditBeforePointerSelection = () => {
      if (input.isEditingCell) {
        return
      }
      const pendingEdit = deferredBeginEditScheduler.consume()
      if (!pendingEdit) {
        return
      }
      input.pendingTypeSeedRef.current = null
      input.onCommitEdit(undefined, pendingEdit.seed)
    }
    window.addEventListener('pointerdown', commitPendingTypedEditBeforePointerSelection, true)
    return () => {
      window.removeEventListener('pointerdown', commitPendingTypedEditBeforePointerSelection, true)
    }
  }, [deferredBeginEditScheduler, input])

  useEffect(() => {
    const handleWindowKeyDown = (event: KeyboardEvent) => {
      if (
        event.defaultPrevented ||
        (event as KeyboardEvent & { __biligGridHandled?: boolean }).__biligGridHandled === true ||
        isGridKeyboardEditableTarget(event.target)
      ) {
        return
      }
      const normalizedKey = getNormalizedGridKeyboardKey(event.key, event.code)
      const activeElement = document.activeElement
      if (
        shouldSuppressWorkbookChromeClearKey(
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
        event.preventDefault()
        event.stopPropagation()
        ;(event as KeyboardEvent & { __biligGridHandled?: boolean }).__biligGridHandled = true
        return
      }
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
    const handleWindowKeyUp = (event: KeyboardEvent) => {
      if (event.defaultPrevented || isGridKeyboardEditableTarget(event.target)) {
        return
      }
      const normalizedKey = getNormalizedGridKeyboardKey(event.key, event.code)
      if (
        !shouldSuppressWorkbookChromeSelectionKeyUp(
          {
            altKey: event.altKey,
            ctrlKey: event.ctrlKey,
            key: normalizedKey,
            metaKey: event.metaKey,
            shiftKey: event.shiftKey,
          },
          document.activeElement,
          input.hostRef.current,
        )
      ) {
        return
      }
      event.preventDefault()
      event.stopPropagation()
    }
    window.addEventListener('keyup', handleWindowKeyUp, true)
    return () => {
      window.removeEventListener('keydown', handleWindowKeyDown, true)
      window.removeEventListener('keyup', handleWindowKeyUp, true)
    }
  }, [handleGridKey, input.hostRef])

  return {
    handleGridKey,
  }
}
