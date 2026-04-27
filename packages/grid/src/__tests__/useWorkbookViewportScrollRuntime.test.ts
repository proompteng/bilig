import { describe, expect, it } from 'vitest'
import type { VisibleRegionState } from '../gridPointer.js'
import { shouldCommitWorkbookVisibleRegion } from '../useWorkbookViewportScrollRuntime.js'

function region(input: {
  readonly x: number
  readonly y: number
  readonly width?: number | undefined
  readonly height?: number | undefined
  readonly freezeRows?: number | undefined
  readonly freezeCols?: number | undefined
  readonly tx?: number | undefined
  readonly ty?: number | undefined
}): VisibleRegionState {
  return {
    freezeCols: input.freezeCols ?? 0,
    freezeRows: input.freezeRows ?? 0,
    range: {
      height: input.height ?? 12,
      width: input.width ?? 12,
      x: input.x,
      y: input.y,
    },
    tx: input.tx ?? 0,
    ty: input.ty ?? 0,
  }
}

describe('shouldCommitWorkbookVisibleRegion', () => {
  it('keeps steady scroll inside the same resident window out of React state', () => {
    const current = region({ x: 0, y: 0 })
    const next = region({ x: 8, y: 8, tx: 22, ty: 11 })

    expect(
      shouldCommitWorkbookVisibleRegion({
        current,
        next,
        requiresLiveViewportState: false,
      }),
    ).toBe(false)
  })

  it('commits when the resident render window changes', () => {
    const current = region({ x: 0, y: 0 })
    const next = region({ x: 260, y: 100 })

    expect(
      shouldCommitWorkbookVisibleRegion({
        current,
        next,
        requiresLiveViewportState: false,
      }),
    ).toBe(true)
  })

  it('commits every visible window movement while a live overlay needs viewport state', () => {
    const current = region({ x: 0, y: 0 })
    const next = region({ x: 8, y: 8, tx: 22, ty: 11 })

    expect(
      shouldCommitWorkbookVisibleRegion({
        current,
        next,
        requiresLiveViewportState: true,
      }),
    ).toBe(true)
  })
})
