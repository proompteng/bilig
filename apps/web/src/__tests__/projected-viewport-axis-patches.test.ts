import { describe, expect, it } from 'vitest'
import { applyProjectedViewportAxisPatches } from '../projected-viewport-axis-patches.js'

describe('projected viewport axis patches', () => {
  it('clears a matching pending optimistic size once the authoritative patch catches up', () => {
    const result = applyProjectedViewportAxisPatches({
      patches: [{ index: 2, size: 120, hidden: false }],
      sizes: { 2: 90 },
      renderedSizes: { 2: 140 },
      pendingSizes: { 2: 120 },
      pendingHiddenAxes: {},
      hiddenAxes: {},
    })

    expect(result.sizes[2]).toBe(120)
    expect(result.renderedSizes[2]).toBe(120)
    expect(result.pendingSizes[2]).toBeUndefined()
    expect(result.axisChanged).toBe(true)
  })

  it('preserves an optimistic rendered size while a visible authoritative patch still lags', () => {
    const result = applyProjectedViewportAxisPatches({
      patches: [{ index: 3, size: 80, hidden: false }],
      sizes: { 3: 110 },
      renderedSizes: { 3: 110 },
      pendingSizes: { 3: 110 },
      pendingHiddenAxes: {},
      hiddenAxes: {},
    })

    expect(result.sizes[3]).toBe(80)
    expect(result.renderedSizes[3]).toBe(110)
    expect(result.pendingSizes[3]).toBe(110)
    expect(result.axisChanged).toBe(false)
  })

  it('applies the authoritative size once a previously hidden axis becomes visible again', () => {
    const result = applyProjectedViewportAxisPatches({
      patches: [{ index: 4, size: 95, hidden: false }],
      sizes: { 4: 70 },
      renderedSizes: { 4: 0 },
      pendingSizes: { 4: 140 },
      pendingHiddenAxes: {},
      hiddenAxes: { 4: true },
    })

    expect(result.hiddenAxes[4]).toBeUndefined()
    expect(result.renderedSizes[4]).toBe(95)
    expect(result.pendingSizes[4]).toBe(140)
    expect(result.axisChanged).toBe(true)
  })

  it('zeros the rendered size while keeping the authoritative base size when hidden', () => {
    const result = applyProjectedViewportAxisPatches({
      patches: [{ index: 5, size: 60, hidden: true }],
      sizes: {},
      renderedSizes: {},
      pendingSizes: { 5: 75 },
      pendingHiddenAxes: {},
      hiddenAxes: {},
    })

    expect(result.sizes[5]).toBe(60)
    expect(result.renderedSizes[5]).toBe(0)
    expect(result.hiddenAxes[5]).toBe(true)
    expect(result.pendingSizes[5]).toBe(75)
    expect(result.axisChanged).toBe(true)
  })

  it('stays quiet for repeated identical visible patches', () => {
    const result = applyProjectedViewportAxisPatches({
      patches: [{ index: 6, size: 93, hidden: false }],
      sizes: { 6: 93 },
      renderedSizes: { 6: 93 },
      pendingSizes: {},
      pendingHiddenAxes: {},
      hiddenAxes: {},
    })

    expect(result.axisChanged).toBe(false)
  })

  it('stays quiet for repeated identical hidden patches', () => {
    const result = applyProjectedViewportAxisPatches({
      patches: [{ index: 7, size: 44, hidden: true }],
      sizes: { 7: 44 },
      renderedSizes: { 7: 0 },
      pendingSizes: {},
      pendingHiddenAxes: {},
      hiddenAxes: { 7: true },
    })

    expect(result.axisChanged).toBe(false)
  })

  it('preserves an optimistic hidden state while a visible authoritative patch still lags', () => {
    const result = applyProjectedViewportAxisPatches({
      patches: [{ index: 8, size: 22, hidden: false }],
      sizes: { 8: 22 },
      renderedSizes: { 8: 0 },
      pendingSizes: {},
      pendingHiddenAxes: { 8: true },
      hiddenAxes: { 8: true },
    })

    expect(result.sizes[8]).toBe(22)
    expect(result.renderedSizes[8]).toBe(0)
    expect(result.pendingHiddenAxes[8]).toBe(true)
    expect(result.hiddenAxes[8]).toBe(true)
    expect(result.axisChanged).toBe(false)
  })

  it('clears a matching pending optimistic hidden state once the authoritative patch catches up', () => {
    const result = applyProjectedViewportAxisPatches({
      patches: [{ index: 9, size: 22, hidden: true }],
      sizes: { 9: 22 },
      renderedSizes: { 9: 0 },
      pendingSizes: {},
      pendingHiddenAxes: { 9: true },
      hiddenAxes: { 9: true },
    })

    expect(result.renderedSizes[9]).toBe(0)
    expect(result.pendingHiddenAxes[9]).toBeUndefined()
    expect(result.hiddenAxes[9]).toBe(true)
    expect(result.axisChanged).toBe(false)
  })
})
