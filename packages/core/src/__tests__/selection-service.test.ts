import { Effect } from 'effect'
import { describe, expect, it, vi } from 'vitest'
import type { SelectionState } from '@bilig/protocol'
import { createEngineSelectionService } from '../engine/services/selection-service.js'

function createSelectionState(initial: SelectionState) {
  let selection = initial
  return {
    selectionListeners: new Set<() => void>(),
    getSelection: () => selection,
    setSelection: (next: SelectionState) => {
      selection = next
    },
  }
}

describe('EngineSelectionService', () => {
  it('avoids duplicate notifications for equivalent selection writes', () => {
    const state = createSelectionState({
      sheetName: 'Sheet1',
      address: 'A1',
      anchorAddress: 'A1',
      range: { startAddress: 'A1', endAddress: 'A1' },
      editMode: 'idle',
    })
    const service = createEngineSelectionService(state)
    const listener = vi.fn()

    const unsubscribe = Effect.runSync(service.subscribe(listener))

    expect(Effect.runSync(service.setSelection('Sheet2', 'B3'))).toBe(true)
    expect(Effect.runSync(service.setSelection('Sheet2', 'B3'))).toBe(false)
    expect(Effect.runSync(service.getSelectionState())).toEqual({
      sheetName: 'Sheet2',
      address: 'B3',
      anchorAddress: 'B3',
      range: { startAddress: 'B3', endAddress: 'B3' },
      editMode: 'idle',
    })
    expect(listener).toHaveBeenCalledTimes(1)

    unsubscribe()
  })
})
