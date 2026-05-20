import { describe, expect, it, vi } from 'vitest'
import { createGridSelection } from '../gridSelection.js'
import type { GridSelectionSnapshot } from '../gridTypes.js'
import { GridInputController } from '../runtime/gridInputController.js'

const SHEET1_A1: GridSelectionSnapshot = {
  address: 'A1',
  kind: 'cell',
  range: { endAddress: 'A1', startAddress: 'A1' },
  sheetName: 'Sheet1',
}

const SHEET1_B2: GridSelectionSnapshot = {
  address: 'B2',
  kind: 'cell',
  range: { endAddress: 'B2', startAddress: 'B2' },
  sheetName: 'Sheet1',
}

describe('GridInputController', () => {
  it('owns pointer interaction refs with a stable interaction-state object', () => {
    const controller = new GridInputController()
    const interactionState = controller.interactionState

    interactionState.dragAnchorCellRef.current = [2, 3]
    interactionState.dragDidMoveRef.current = true
    controller.dragPointerCellRef.current = [4, 5]

    expect(controller.interactionState).toBe(interactionState)
    expect(controller.dragAnchorCellRef.current).toEqual([2, 3])
    expect(controller.dragDidMoveRef.current).toBe(true)
    expect(controller.interactionState.dragPointerCellRef.current).toEqual([4, 5])
  })

  it('tracks sheet and fill-preview state outside the React hook', () => {
    const controller = new GridInputController()

    expect(controller.syncActiveSheet('Sheet1')).toBe(false)
    expect(controller.syncActiveSheet('Sheet1')).toBe(false)
    expect(controller.syncActiveSheet('Sheet2')).toBe(true)
    expect(controller.syncActiveSheet('Sheet2')).toBe(false)

    controller.syncFillPreviewRange({ height: 3, width: 2, x: 1, y: 4 })

    expect(controller.fillPreviewRangeRef.current).toEqual({ height: 3, width: 2, x: 1, y: 4 })
  })

  it('runs owned drag and resize cleanup once on disconnect', () => {
    const controller = new GridInputController()
    const fillCleanup = vi.fn()
    const moveCleanup = vi.fn()
    const resizeCleanup = vi.fn()

    controller.fillHandleCleanupRef.current = fillCleanup
    controller.rangeMoveCleanupRef.current = moveCleanup
    controller.resizeCleanupRef.current = resizeCleanup

    controller.disconnect()
    controller.disconnect()

    expect(fillCleanup).toHaveBeenCalledTimes(1)
    expect(moveCleanup).toHaveBeenCalledTimes(1)
    expect(resizeCleanup).toHaveBeenCalledTimes(1)
  })

  it('applies external selection sync while preserving pending local selection', () => {
    const controller = new GridInputController()

    expect(
      controller.syncExternalSelection({
        currentSelection: createGridSelection(0, 0),
        externalSnapshot: SHEET1_A1,
        sheetName: 'Sheet1',
      }),
    ).toBeNull()

    const localSnapshot = controller.noteLocalSelectionChange({
      baseSnapshot: SHEET1_A1,
      nextSelection: createGridSelection(1, 1),
      sheetName: 'Sheet1',
    })
    expect(localSnapshot).toEqual(SHEET1_B2)
    expect(controller.pendingLocalSelectionSnapshotRef.current).toEqual(SHEET1_B2)

    controller.interactionState.pendingPointerCellRef.current = [4, 5]
    expect(
      controller.syncExternalSelection({
        currentSelection: createGridSelection(1, 1),
        externalSnapshot: SHEET1_A1,
        sheetName: 'Sheet1',
      }),
    ).toBeNull()
    expect(controller.interactionState.pendingPointerCellRef.current).toEqual([4, 5])

    const authoritativeSelection = controller.syncExternalSelection({
      currentSelection: createGridSelection(1, 1),
      externalSnapshot: SHEET1_B2,
      sheetName: 'Sheet1',
    })
    expect(authoritativeSelection).toBeNull()
    expect(controller.pendingLocalSelectionSnapshotRef.current).toBeNull()

    const remoteSelection = controller.syncExternalSelection({
      currentSelection: createGridSelection(1, 1),
      externalSnapshot: SHEET1_A1,
      sheetName: 'Sheet1',
    })
    expect(remoteSelection).toEqual(createGridSelection(0, 0))
    expect(controller.interactionState.pendingPointerCellRef.current).toBeNull()
  })

  it('identifies the current render selection as pending only while the external snapshot is still at the base selection', () => {
    const controller = new GridInputController()
    const localSelection = createGridSelection(1, 1)

    controller.noteLocalSelectionChange({
      baseSnapshot: SHEET1_A1,
      nextSelection: localSelection,
      sheetName: 'Sheet1',
    })

    expect(
      controller.hasPendingLocalSelection({
        currentSelection: localSelection,
        externalSnapshot: SHEET1_A1,
        sheetName: 'Sheet1',
      }),
    ).toBe(true)
    expect(
      controller.hasPendingLocalSelection({
        currentSelection: createGridSelection(2, 2),
        externalSnapshot: SHEET1_A1,
        sheetName: 'Sheet1',
      }),
    ).toBe(false)
    expect(
      controller.hasPendingLocalSelection({
        currentSelection: localSelection,
        externalSnapshot: SHEET1_B2,
        sheetName: 'Sheet1',
      }),
    ).toBe(false)
  })

  it('syncs editor state and schedules focus after editor close', () => {
    const controller = new GridInputController()
    const focusGrid = vi.fn()
    const requestAnimationFrame = vi.fn((callback: FrameRequestCallback) => {
      callback(0)
      return 1
    })

    controller.pendingTypeSeedRef.current = 'x'
    controller.syncEditingState({
      focusGrid,
      isEditingCell: true,
      requestAnimationFrame,
    })
    expect(controller.pendingTypeSeedRef.current).toBeNull()
    expect(focusGrid).not.toHaveBeenCalled()

    controller.syncEditingState({
      focusGrid,
      isEditingCell: false,
      requestAnimationFrame,
    })
    expect(requestAnimationFrame).toHaveBeenCalledTimes(1)
    expect(focusGrid).toHaveBeenCalledTimes(1)
  })

  it('flushes mounted editor value when the DOM has newer text', () => {
    const controller = new GridInputController()
    const onEditorChange = vi.fn()
    const flushSync = vi.fn((callback: () => void) => {
      callback()
    })
    const editor = { value: 'typed' }
    const queryEditor = vi.fn(() => editor)

    expect(
      controller.syncMountedEditorValue({
        editorValue: 'old',
        flushSync,
        onEditorChange,
        queryEditor,
      }),
    ).toBe('typed')
    expect(queryEditor).toHaveBeenCalledTimes(1)
    expect(flushSync).toHaveBeenCalledTimes(1)
    expect(onEditorChange).toHaveBeenCalledWith('typed')

    expect(
      controller.syncMountedEditorValue({
        editorValue: 'typed',
        flushSync,
        onEditorChange,
        queryEditor,
      }),
    ).toBe('typed')
    expect(flushSync).toHaveBeenCalledTimes(1)
  })
})
