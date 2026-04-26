import { createAxisIndex, type AxisEntryOverride, type AxisIndex } from '../gridAxisIndex.js'

export interface GridAxisRuntimeSnapshot {
  readonly axisLength: number
  readonly defaultSize: number
  readonly seq: number
}

export interface GridAxisRange {
  readonly start: number
  readonly endExclusive: number
}

export class GridAxisRuntime {
  private index: AxisIndex
  private seq: number
  private readonly axisLength: number
  private readonly defaultSize: number

  constructor(input: {
    readonly axisLength: number
    readonly defaultSize: number
    readonly seq?: number | undefined
    readonly overrides?: readonly AxisEntryOverride[] | undefined
  }) {
    this.axisLength = Math.max(1, Math.floor(input.axisLength))
    this.defaultSize = Math.max(0, input.defaultSize)
    this.seq = Math.max(0, Math.floor(input.seq ?? 0))
    this.index = createRuntimeAxisIndex(this.axisLength, this.defaultSize, input.overrides)
  }

  snapshot(): GridAxisRuntimeSnapshot {
    return {
      axisLength: this.axisLength,
      defaultSize: this.defaultSize,
      seq: this.seq,
    }
  }

  update(input: { readonly seq?: number | undefined; readonly overrides: readonly AxisEntryOverride[] }): void {
    this.seq = Math.max(this.seq + 1, Math.floor(input.seq ?? 0))
    this.index = createRuntimeAxisIndex(this.axisLength, this.defaultSize, input.overrides)
  }

  sizeAt(index: number): number {
    return this.index.resolveSize(index)
  }

  offsetOf(index: number): number {
    return this.index.resolveOffset(index)
  }

  span(start: number, endExclusive: number): number {
    return this.index.resolveSpan(start, endExclusive)
  }

  tileOrigin(tileStart: number): number {
    return this.offsetOf(tileStart)
  }

  visibleRangeForOffset(offset: number, extent: number, overscanPx = 0): GridAxisRange {
    const anchor = this.index.resolveAnchor(offset)
    const count = this.index.resolveVisibleCount(anchor.index, extent, overscanPx)
    return {
      start: anchor.index,
      endExclusive: Math.min(this.axisLength, anchor.index + count),
    }
  }
}

function createRuntimeAxisIndex(axisLength: number, defaultSize: number, overrides: readonly AxisEntryOverride[] | undefined): AxisIndex {
  return createAxisIndex({
    axisLength,
    defaultSize,
    ...(overrides ? { overrides } : {}),
  })
}
