export interface GridAxisEntryOverride {
  readonly index: number
  readonly size?: number
  readonly hidden?: boolean
}

export interface GridAxisAnchor {
  readonly index: number
  readonly offset: number
  readonly size: number
  readonly intraOffset: number
}

export interface GridAxisVisibleRange {
  readonly startIndex: number
  readonly endIndexExclusive: number
  readonly count: number
}

export interface GridAxisWorldIndex {
  readonly axisLength: number
  readonly defaultSize: number
  readonly version: number
  readonly totalSize: number
  sizeOf(index: number): number
  isHidden(index: number): boolean
  offsetOf(index: number): number
  endOffsetOf(index: number): number
  span(startInclusive: number, endExclusive: number): number
  anchorAt(worldOffset: number): GridAxisAnchor
  visibleCountFrom(startIndex: number, viewportSize: number): number
  visibleRangeForWorldRect(startOffset: number, size: number): GridAxisVisibleRange
  hitTest(worldOffset: number): number | null
}

interface NormalizedAxisOverride {
  readonly index: number
  readonly size: number
  readonly hidden: boolean
  readonly deltaPrefixBefore: number
}

export function createGridAxisWorldIndex(input: {
  readonly axisLength: number
  readonly defaultSize: number
  readonly version?: number | undefined
  readonly overrides?: readonly GridAxisEntryOverride[]
}): GridAxisWorldIndex {
  const axisLength = Math.max(1, Math.floor(input.axisLength))
  const defaultSize = Math.max(0, input.defaultSize)
  const overrides = normalizeOverrides(axisLength, defaultSize, input.overrides ?? [])
  const version = Math.max(0, Math.floor(input.version ?? hashAxisSnapshot(axisLength, defaultSize, overrides)))
  const totalSize = offsetOf(axisLength, axisLength, defaultSize, overrides)

  return {
    axisLength,
    defaultSize,
    version,
    totalSize,
    anchorAt(worldOffset) {
      if (totalSize <= 0) {
        return { index: 0, offset: 0, size: 0, intraOffset: 0 }
      }
      const target = Math.max(0, Math.min(worldOffset, totalSize))
      if (target >= totalSize) {
        const index = findLastVisibleIndex(axisLength, defaultSize, overrides)
        const offset = offsetOf(index, axisLength, defaultSize, overrides)
        const size = sizeOf(index, axisLength, defaultSize, overrides)
        return { index, offset, size, intraOffset: Math.max(0, size) }
      }
      let low = 0
      let high = axisLength - 1
      while (low < high) {
        const mid = Math.floor((low + high) / 2)
        if (offsetOf(mid + 1, axisLength, defaultSize, overrides) > target) {
          high = mid
        } else {
          low = mid + 1
        }
      }
      const index = skipHiddenForward(low, axisLength, defaultSize, overrides)
      const offset = offsetOf(index, axisLength, defaultSize, overrides)
      const size = sizeOf(index, axisLength, defaultSize, overrides)
      return { index, offset, size, intraOffset: Math.max(0, target - offset) }
    },
    endOffsetOf(index) {
      return offsetOf(Math.floor(index) + 1, axisLength, defaultSize, overrides)
    },
    hitTest(worldOffset) {
      if (worldOffset < 0 || worldOffset >= totalSize) {
        return null
      }
      const anchor = this.anchorAt(worldOffset)
      return anchor.size > 0 && !this.isHidden(anchor.index) ? anchor.index : null
    },
    isHidden(index) {
      return isHidden(index, axisLength, defaultSize, overrides)
    },
    offsetOf(index) {
      return offsetOf(index, axisLength, defaultSize, overrides)
    },
    sizeOf(index) {
      return sizeOf(index, axisLength, defaultSize, overrides)
    },
    span(startInclusive, endExclusive) {
      return Math.max(
        0,
        offsetOf(endExclusive, axisLength, defaultSize, overrides) - offsetOf(startInclusive, axisLength, defaultSize, overrides),
      )
    },
    visibleCountFrom(startIndex, viewportSize) {
      const range = this.visibleRangeForWorldRect(this.offsetOf(startIndex), viewportSize)
      return range.count
    },
    visibleRangeForWorldRect(startOffset, size) {
      const startAnchor = this.anchorAt(Math.max(0, startOffset))
      const endOffset = Math.max(startOffset, startOffset + Math.max(0, size))
      let endIndexExclusive = startAnchor.index + 1
      while (endIndexExclusive < axisLength && this.offsetOf(endIndexExclusive) < endOffset) {
        endIndexExclusive += 1
      }
      let count = 0
      for (let index = startAnchor.index; index < endIndexExclusive; index += 1) {
        if (!this.isHidden(index) && this.sizeOf(index) > 0) {
          count += 1
        }
      }
      return {
        count,
        endIndexExclusive,
        startIndex: startAnchor.index,
      }
    },
  }
}

export function createGridAxisWorldIndexFromRecords(input: {
  readonly axisLength: number
  readonly defaultSize: number
  readonly sizes?: Readonly<Record<number, number>> | undefined
  readonly hidden?: Readonly<Record<number, true>> | undefined
  readonly version?: number | undefined
}): GridAxisWorldIndex {
  const overrides = new Map<number, GridAxisEntryOverride>()
  for (const [rawIndex, size] of Object.entries(input.sizes ?? {})) {
    overrides.set(Number(rawIndex), { index: Number(rawIndex), size })
  }
  for (const rawIndex of Object.keys(input.hidden ?? {})) {
    const index = Number(rawIndex)
    overrides.set(index, { index, hidden: true, size: 0 })
  }
  return createGridAxisWorldIndex({
    axisLength: input.axisLength,
    defaultSize: input.defaultSize,
    version: input.version,
    overrides: [...overrides.values()],
  })
}

function normalizeOverrides(
  axisLength: number,
  defaultSize: number,
  overrides: readonly GridAxisEntryOverride[],
): readonly NormalizedAxisOverride[] {
  const byIndex = new Map<number, { size: number; hidden: boolean }>()
  for (const override of overrides) {
    const index = Math.floor(override.index)
    if (index < 0 || index >= axisLength) {
      continue
    }
    const hidden = override.hidden === true
    const size = hidden ? 0 : Math.max(0, override.size ?? defaultSize)
    if (!hidden && size === defaultSize) {
      byIndex.delete(index)
      continue
    }
    byIndex.set(index, { hidden, size })
  }

  let deltaPrefix = 0
  return [...byIndex.entries()]
    .toSorted((left, right) => left[0] - right[0])
    .map(([index, value]) => {
      const entry = {
        deltaPrefixBefore: deltaPrefix,
        hidden: value.hidden,
        index,
        size: value.size,
      }
      deltaPrefix += value.size - defaultSize
      return entry
    })
}

function hashAxisSnapshot(axisLength: number, defaultSize: number, overrides: readonly NormalizedAxisOverride[]): number {
  let hash = 2_166_136_261
  hash = mixHash(hash, axisLength)
  hash = mixHash(hash, Math.round(defaultSize * 1_000))
  for (const override of overrides) {
    hash = mixHash(hash, override.index)
    hash = mixHash(hash, Math.round(override.size * 1_000))
    hash = mixHash(hash, override.hidden ? 1 : 0)
  }
  return hash
}

function mixHash(hash: number, value: number): number {
  return Math.imul((hash ^ value) >>> 0, 16_777_619) >>> 0
}

function offsetOf(index: number, axisLength: number, defaultSize: number, overrides: readonly NormalizedAxisOverride[]): number {
  const clamped = clampEndExclusive(index, axisLength)
  const entryBeforeIndex = findLastEntryBefore(overrides, clamped)
  const deltaPrefix = entryBeforeIndex === null ? 0 : entryBeforeIndex.deltaPrefixBefore + entryBeforeIndex.size - defaultSize
  return clamped * defaultSize + deltaPrefix
}

function sizeOf(index: number, axisLength: number, defaultSize: number, overrides: readonly NormalizedAxisOverride[]): number {
  const clamped = Math.floor(index)
  if (clamped < 0 || clamped >= axisLength) {
    return 0
  }
  return findEntryAt(overrides, clamped)?.size ?? defaultSize
}

function isHidden(index: number, axisLength: number, defaultSize: number, overrides: readonly NormalizedAxisOverride[]): boolean {
  const clamped = Math.floor(index)
  if (clamped < 0 || clamped >= axisLength) {
    return true
  }
  const entry = findEntryAt(overrides, clamped)
  return entry?.hidden === true || (entry ? entry.size === 0 : defaultSize === 0)
}

function skipHiddenForward(index: number, axisLength: number, defaultSize: number, overrides: readonly NormalizedAxisOverride[]): number {
  let cursor = Math.max(0, Math.min(axisLength - 1, index))
  while (cursor < axisLength - 1 && isHidden(cursor, axisLength, defaultSize, overrides)) {
    cursor += 1
  }
  return cursor
}

function findLastVisibleIndex(axisLength: number, defaultSize: number, overrides: readonly NormalizedAxisOverride[]): number {
  for (let index = axisLength - 1; index >= 0; index -= 1) {
    if (!isHidden(index, axisLength, defaultSize, overrides)) {
      return index
    }
  }
  return 0
}

function findEntryAt(overrides: readonly NormalizedAxisOverride[], index: number): NormalizedAxisOverride | null {
  let low = 0
  let high = overrides.length
  while (low < high) {
    const mid = Math.floor((low + high) / 2)
    const entry = overrides[mid]
    if (!entry || entry.index >= index) {
      high = mid
    } else {
      low = mid + 1
    }
  }
  return overrides[low]?.index === index ? (overrides[low] ?? null) : null
}

function findLastEntryBefore(overrides: readonly NormalizedAxisOverride[], index: number): NormalizedAxisOverride | null {
  let low = 0
  let high = overrides.length
  while (low < high) {
    const mid = Math.floor((low + high) / 2)
    const entry = overrides[mid]
    if (!entry || entry.index >= index) {
      high = mid
    } else {
      low = mid + 1
    }
  }
  return low === 0 ? null : (overrides[low - 1] ?? null)
}

function clampEndExclusive(index: number, axisLength: number): number {
  return Math.max(0, Math.min(axisLength, Math.floor(index)))
}
