import type { HeaderSelection } from '../gridPointer.js'
import { clearGridPendingPointerActivation, type GridInteractionStateRefs } from '../gridInteractionState.js'
import type { InternalClipboardRange } from '../gridInternalClipboard.js'
import { selectionToSnapshot, snapshotToSelection } from '../gridSelection.js'
import { resolveGridSelectionPendingSync } from '../gridSelectionPendingSync.js'
import type { GridSelection, GridSelectionSnapshot, Item, Rectangle } from '../gridTypes.js'

interface RuntimeRef<T> {
  current: T
}

type CleanupRef = RuntimeRef<(() => void) | null>

export interface InteriorRangeMoveCandidate {
  readonly pointerId: number
  readonly pointerCell: Item
  readonly startClientX: number
  readonly startClientY: number
}

export class GridInputController {
  readonly wasEditingOverlayRef = createRuntimeRef(false)
  readonly ignoreNextPointerSelectionRef = createRuntimeRef(false)
  readonly pendingPointerCellRef = createRuntimeRef<Item | null>(null)
  readonly dragAnchorCellRef = createRuntimeRef<Item | null>(null)
  readonly dragPointerCellRef = createRuntimeRef<Item | null>(null)
  readonly dragHeaderSelectionRef = createRuntimeRef<HeaderSelection | null>(null)
  readonly dragDidMoveRef = createRuntimeRef(false)
  readonly postDragSelectionExpiryRef = createRuntimeRef(0)
  readonly columnResizeActiveRef = createRuntimeRef(false)
  readonly lastBodyClickCellRef = createRuntimeRef<Item | null>(null)
  readonly internalClipboardRef = createRuntimeRef<InternalClipboardRange | null>(null)
  readonly pendingClipboardCopySequenceRef = createRuntimeRef(0)
  readonly pendingKeyboardPasteSequenceRef = createRuntimeRef(0)
  readonly suppressNextNativePasteRef = createRuntimeRef(false)
  readonly pendingTypeSeedRef = createRuntimeRef<string | null>(null)
  readonly pendingLocalSelectionSnapshotRef = createRuntimeRef<GridSelectionSnapshot | null>(null)
  readonly pendingLocalSelectionBaseSnapshotRef = createRuntimeRef<GridSelectionSnapshot | null>(null)
  readonly lastResizeHandleActivationRef = createRuntimeRef<{ columnIndex: number; at: number } | null>(null)
  readonly interiorRangeMoveCandidateRef = createRuntimeRef<InteriorRangeMoveCandidate | null>(null)
  readonly fillPreviewRangeRef = createRuntimeRef<Rectangle | null>(null)
  readonly fillHandleCleanupRef: CleanupRef = createRuntimeRef(null)
  readonly rangeMoveCleanupRef: CleanupRef = createRuntimeRef(null)
  readonly resizeCleanupRef: CleanupRef = createRuntimeRef(null)
  readonly activeSheetRef = createRuntimeRef<string | null>(null)
  readonly interactionState: GridInteractionStateRefs

  constructor() {
    this.interactionState = {
      columnResizeActiveRef: this.columnResizeActiveRef,
      dragAnchorCellRef: this.dragAnchorCellRef,
      dragDidMoveRef: this.dragDidMoveRef,
      dragHeaderSelectionRef: this.dragHeaderSelectionRef,
      dragPointerCellRef: this.dragPointerCellRef,
      ignoreNextPointerSelectionRef: this.ignoreNextPointerSelectionRef,
      pendingPointerCellRef: this.pendingPointerCellRef,
      postDragSelectionExpiryRef: this.postDragSelectionExpiryRef,
    }
  }

  syncActiveSheet(sheetName: string): boolean {
    const previousSheetName = this.activeSheetRef.current
    this.activeSheetRef.current = sheetName
    return previousSheetName !== null && previousSheetName !== sheetName
  }

  syncFillPreviewRange(fillPreviewRange: Rectangle | null): void {
    this.fillPreviewRangeRef.current = fillPreviewRange
  }

  syncExternalSelection(input: {
    readonly currentSelection: GridSelection
    readonly externalSnapshot: GridSelectionSnapshot
    readonly sheetName: string
  }): GridSelection | null {
    const sheetChanged = this.syncActiveSheet(input.sheetName)
    const currentSnapshot = selectionToSnapshot(input.currentSelection, input.externalSnapshot.sheetName, input.externalSnapshot.address)
    const sync = resolveGridSelectionPendingSync({
      currentSnapshot,
      externalSnapshot: input.externalSnapshot,
      pendingBaseSnapshot: this.pendingLocalSelectionBaseSnapshotRef.current,
      pendingLocalSnapshot: this.pendingLocalSelectionSnapshotRef.current,
      sheetChanged,
    })
    this.pendingLocalSelectionSnapshotRef.current = sync.pendingLocalSnapshot
    this.pendingLocalSelectionBaseSnapshotRef.current = sync.pendingBaseSnapshot
    if (sync.keepCurrentSelection) {
      return null
    }
    clearGridPendingPointerActivation(this.interactionState)
    return snapshotToSelection(input.externalSnapshot)
  }

  noteLocalSelectionChange(input: {
    readonly nextSelection: GridSelection
    readonly sheetName: string
    readonly baseSnapshot: GridSelectionSnapshot
  }): GridSelectionSnapshot {
    const nextSelectionSnapshot = selectionToSnapshot(input.nextSelection, input.sheetName, input.baseSnapshot.address)
    this.pendingLocalSelectionBaseSnapshotRef.current = input.baseSnapshot
    this.pendingLocalSelectionSnapshotRef.current = nextSelectionSnapshot
    return nextSelectionSnapshot
  }

  syncEditingState(input: {
    readonly isEditingCell: boolean
    readonly focusGrid: () => void
    readonly requestAnimationFrame?: ((callback: FrameRequestCallback) => number) | undefined
  }): void {
    if (this.wasEditingOverlayRef.current && !input.isEditingCell) {
      const requestAnimationFrame = input.requestAnimationFrame ?? globalThis.requestAnimationFrame?.bind(globalThis)
      requestAnimationFrame?.(() => {
        input.focusGrid()
      })
    }
    if (input.isEditingCell) {
      this.pendingTypeSeedRef.current = null
    }
    this.wasEditingOverlayRef.current = input.isEditingCell
  }

  syncMountedEditorValue(input: {
    readonly editorValue: string
    readonly onEditorChange: (value: string) => void
    readonly flushSync: (callback: () => void) => void
    readonly queryEditor?: (() => { readonly value: string } | null) | undefined
  }): string | null {
    const editor = input.queryEditor?.() ?? globalThis.document?.querySelector<HTMLTextAreaElement>('[data-testid="cell-editor-input"]')
    if (!editor) {
      return null
    }
    if (editor.value !== input.editorValue) {
      input.flushSync(() => {
        input.onEditorChange(editor.value)
      })
    }
    return editor.value
  }

  disconnect(): void {
    this.runCleanup(this.fillHandleCleanupRef)
    this.runCleanup(this.rangeMoveCleanupRef)
    this.runCleanup(this.resizeCleanupRef)
  }

  private runCleanup(ref: CleanupRef): void {
    const cleanup = ref.current
    ref.current = null
    cleanup?.()
  }
}

function createRuntimeRef<T>(current: T): RuntimeRef<T> {
  return { current }
}
