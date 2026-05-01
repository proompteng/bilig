import { describe, expect, it, vi } from 'vitest'
import { GridInputController } from '../runtime/gridInputController.js'

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
})
