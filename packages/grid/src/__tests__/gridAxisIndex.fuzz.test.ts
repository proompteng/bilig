import { describe, expect, it } from 'vitest'
import * as fc from 'fast-check'
import { runProperty } from '@bilig/test-fuzz'
import { createAxisIndex, type AxisEntryOverride } from '../gridAxisIndex.js'
import { createGridAxisWorldIndex, type GridAxisEntryOverride } from '../gridAxisWorldIndex.js'

describe('grid axis index fuzz', () => {
  it('should match a simple prefix-sum model for sparse axis sizes', async () => {
    await runProperty({
      suite: 'grid/axis-index/prefix-sum-reference',
      arbitrary: axisCaseArbitrary,
      predicate: async ({ axisLength, defaultSize, overrides, probe }) => {
        const axis = createAxisIndex({ axisLength, defaultSize, overrides })
        const sizes = referenceSizes(axisLength, defaultSize, overrides)
        const clampedIndex = Math.max(0, Math.min(axisLength - 1, Math.floor(probe)))
        const startExclusive = Math.max(0, Math.min(axisLength, Math.floor(probe)))
        const endExclusive = Math.max(0, Math.min(axisLength, Math.floor(probe + 3)))

        expect(axis.resolveSize(probe)).toBe(probe >= 0 && probe < axisLength ? sizes[clampedIndex] : 0)
        expect(axis.resolveOffset(probe)).toBe(referenceOffset(sizes, startExclusive))
        expect(axis.resolveSpan(probe, probe + 3)).toBe(
          Math.max(0, referenceOffset(sizes, endExclusive) - referenceOffset(sizes, startExclusive)),
        )
      },
      parameters: { numRuns: 120 },
    })
  })

  it('should keep world-index hit tests consistent with visible anchors', async () => {
    await runProperty({
      suite: 'grid/axis-world-index/hittest-anchor-consistency',
      arbitrary: axisCaseArbitrary,
      predicate: async ({ axisLength, defaultSize, overrides, scrollOffset }) => {
        const world = createGridAxisWorldIndex({
          axisLength,
          defaultSize,
          overrides: overrides.map((override) => ({
            hidden: override.hidden,
            index: override.index,
            size: override.size,
          })),
        })
        const anchor = world.anchorAt(scrollOffset)
        const hit = world.hitTest(scrollOffset)

        expect(anchor.index).toBeGreaterThanOrEqual(0)
        expect(anchor.index).toBeLessThan(axisLength)
        expect(anchor.offset).toBe(world.offsetOf(anchor.index))
        expect(anchor.size).toBe(world.sizeOf(anchor.index))
        if (hit !== null) {
          expect(world.isHidden(hit)).toBe(false)
          expect(scrollOffset).toBeGreaterThanOrEqual(world.offsetOf(hit))
          expect(scrollOffset).toBeLessThan(world.endOffsetOf(hit))
        }
      },
      parameters: { numRuns: 120 },
    })
  })
})

const axisOverrideArbitrary: fc.Arbitrary<AxisEntryOverride & GridAxisEntryOverride> = fc.record({
  hidden: fc.boolean(),
  index: fc.integer({ min: -5, max: 55 }),
  size: fc.integer({ min: 0, max: 90 }),
})

const axisCaseArbitrary = fc.record({
  axisLength: fc.integer({ min: 1, max: 50 }),
  defaultSize: fc.integer({ min: 1, max: 60 }),
  overrides: fc.array(axisOverrideArbitrary, { minLength: 0, maxLength: 20 }),
  probe: fc.integer({ min: -5, max: 55 }),
  scrollOffset: fc.integer({ min: -50, max: 2_500 }),
})

function referenceSizes(axisLength: number, defaultSize: number, overrides: readonly AxisEntryOverride[]): number[] {
  const sizes = Array.from({ length: axisLength }, () => defaultSize)
  for (const override of overrides) {
    const index = Math.floor(override.index)
    if (index < 0 || index >= axisLength) {
      continue
    }
    sizes[index] = override.hidden ? 0 : Math.max(0, override.size)
  }
  return sizes
}

function referenceOffset(sizes: readonly number[], endExclusive: number): number {
  let offset = 0
  for (let index = 0; index < endExclusive; index += 1) {
    offset += sizes[index] ?? 0
  }
  return offset
}
