import { describe, expect, it } from 'vitest'
import {
  ackProjectedViewportLocalAxisSize,
  rollbackProjectedViewportLocalAxisHidden,
  rollbackProjectedViewportLocalAxisSize,
  setProjectedViewportLocalAxisHidden,
  setProjectedViewportLocalAxisSize,
  type ProjectedViewportLocalAxisState,
} from '../projected-viewport-local-axis-state.js'

const EMPTY_STATE: ProjectedViewportLocalAxisState = {
  sizes: {},
  renderedSizes: {},
  pendingSizes: {},
  hiddenAxes: {},
}

describe('projected viewport local axis state', () => {
  it('tracks optimistic size changes and matching acknowledgements', () => {
    const sized = setProjectedViewportLocalAxisSize({
      state: EMPTY_STATE,
      index: 2,
      size: 120,
    })

    expect(sized).toMatchObject({
      sizes: { 2: 120 },
      renderedSizes: { 2: 120 },
      pendingSizes: { 2: 120 },
      changed: true,
    })

    const acked = ackProjectedViewportLocalAxisSize({
      state: sized,
      index: 2,
      size: 120,
    })

    expect(acked).toMatchObject({
      sizes: { 2: 120 },
      renderedSizes: { 2: 120 },
      pendingSizes: {},
      changed: true,
    })
  })

  it('rolls back optimistic sizes while preserving hidden rendered sizes', () => {
    const hidden = setProjectedViewportLocalAxisHidden({
      state: EMPTY_STATE,
      index: 3,
      hidden: true,
      size: 88,
    })
    const sized = setProjectedViewportLocalAxisSize({
      state: hidden,
      index: 3,
      size: 120,
    })

    const rolledBack = rollbackProjectedViewportLocalAxisSize({
      state: sized,
      index: 3,
      size: 88,
    })

    expect(rolledBack).toMatchObject({
      sizes: { 3: 88 },
      renderedSizes: { 3: 0 },
      pendingSizes: {},
      hiddenAxes: { 3: true },
      changed: true,
    })
  })

  it('treats repeated hidden writes with the same effective state as no-ops', () => {
    const hidden = setProjectedViewportLocalAxisHidden({
      state: EMPTY_STATE,
      index: 4,
      hidden: true,
      size: 72,
    })

    const repeated = setProjectedViewportLocalAxisHidden({
      state: hidden,
      index: 4,
      hidden: true,
      size: 72,
    })

    expect(repeated.changed).toBe(false)
    expect(repeated).toMatchObject({
      sizes: { 4: 72 },
      renderedSizes: { 4: 0 },
      hiddenAxes: { 4: true },
    })
  })

  it('rolls back hidden state to a visible prior size', () => {
    const hidden = setProjectedViewportLocalAxisHidden({
      state: EMPTY_STATE,
      index: 1,
      hidden: true,
      size: 55,
    })

    const rolledBack = rollbackProjectedViewportLocalAxisHidden({
      state: hidden,
      index: 1,
      previous: { hidden: false, size: 44 },
    })

    expect(rolledBack).toMatchObject({
      sizes: { 1: 44 },
      renderedSizes: { 1: 44 },
      hiddenAxes: {},
      changed: true,
    })
  })
})
