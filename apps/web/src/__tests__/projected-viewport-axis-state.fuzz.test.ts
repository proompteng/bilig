import { describe, expect, it } from 'vitest'
import * as fc from 'fast-check'
import { runProperty } from '@bilig/test-fuzz'
import { applyProjectedViewportAxisPatches, type ProjectedViewportAxisPatch } from '../projected-viewport-axis-patches.js'
import {
  rollbackProjectedViewportLocalAxisHidden,
  rollbackProjectedViewportLocalAxisSize,
  setProjectedViewportLocalAxisHidden,
  setProjectedViewportLocalAxisSize,
  type ProjectedViewportLocalAxisState,
} from '../projected-viewport-local-axis-state.js'

describe('projected viewport axis state fuzz', () => {
  it('should make acked axis patches converge with local optimistic state', async () => {
    await runProperty({
      suite: 'web/projected-viewport/axis-state-convergence',
      arbitrary: axisPatchArbitrary,
      predicate: async (patch) => {
        const emptyState = createEmptyAxisState()
        const local = setProjectedViewportLocalAxisHidden({
          state: setProjectedViewportLocalAxisSize({ state: emptyState, index: patch.index, size: patch.size }),
          index: patch.index,
          hidden: patch.hidden,
          size: patch.size,
        })
        const patched = applyProjectedViewportAxisPatches({
          patches: [patch],
          sizes: local.sizes,
          renderedSizes: local.renderedSizes,
          pendingSizes: local.pendingSizes,
          pendingHiddenAxes: local.pendingHiddenAxes,
          hiddenAxes: local.hiddenAxes,
        })

        expect(patched.pendingHiddenAxes[patch.index]).toBeUndefined()
        expect(Boolean(patched.hiddenAxes[patch.index])).toBe(patch.hidden)
        expect(patched.renderedSizes[patch.index]).toBe(patch.hidden ? 0 : patch.size)
      },
      parameters: { numRuns: 120 },
    })
  })

  it('should rollback generated local size and hidden edits to the previous sparse state', async () => {
    await runProperty({
      suite: 'web/projected-viewport/axis-state-rollback',
      arbitrary: fc.record({
        index: fc.integer({ min: 0, max: 50 }),
        nextSize: fc.integer({ min: 20, max: 300 }),
        previousSize: fc.option(fc.integer({ min: 20, max: 300 }), { nil: undefined }),
        previousHidden: fc.boolean(),
      }),
      predicate: async ({ index, nextSize, previousSize, previousHidden }) => {
        const emptyState = createEmptyAxisState()
        const sized = setProjectedViewportLocalAxisSize({ state: emptyState, index, size: nextSize })
        const sizeRolledBack = rollbackProjectedViewportLocalAxisSize({ state: sized, index, size: previousSize })
        const hidden = setProjectedViewportLocalAxisHidden({
          state: sizeRolledBack,
          index,
          hidden: !previousHidden,
          size: previousSize ?? nextSize,
        })
        const hiddenRolledBack = rollbackProjectedViewportLocalAxisHidden({
          state: hidden,
          index,
          previous: { hidden: previousHidden, size: previousSize },
        })

        expect(hiddenRolledBack.sizes[index]).toBe(previousSize)
        expect(Boolean(hiddenRolledBack.hiddenAxes[index])).toBe(previousHidden)
        expect(hiddenRolledBack.pendingSizes[index]).toBeUndefined()
        expect(hiddenRolledBack.pendingHiddenAxes[index]).toBeUndefined()
      },
      parameters: { numRuns: 120 },
    })
  })
})

// Helpers

const axisPatchArbitrary: fc.Arbitrary<ProjectedViewportAxisPatch> = fc.record({
  index: fc.integer({ min: 0, max: 50 }),
  size: fc.integer({ min: 20, max: 300 }),
  hidden: fc.boolean(),
})

function createEmptyAxisState(): ProjectedViewportLocalAxisState {
  return {
    sizes: {},
    renderedSizes: {},
    pendingSizes: {},
    pendingHiddenAxes: {},
    hiddenAxes: {},
  }
}
