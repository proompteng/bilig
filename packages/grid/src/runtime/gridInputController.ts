import type { HeaderSelection } from '../gridPointer.js'
import type { GridInteractionStateRefs } from '../gridInteractionState.js'
import type { InternalClipboardRange } from '../gridInternalClipboard.js'
import type { GridSelectionSnapshot, Item, Rectangle } from '../gridTypes.js'

interface RuntimeRef<T> {
  current: T
}

type CleanupRef = RuntimeRef<(() => void) | null>

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
