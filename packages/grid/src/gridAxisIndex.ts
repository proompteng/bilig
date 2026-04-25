export interface AxisEntryOverride {
  readonly index: number
  readonly size: number
  readonly hidden?: boolean
}

export interface AxisAnchor {
  readonly index: number
  readonly offset: number
}

export interface AxisIndex {
  readonly axisLength: number
  readonly defaultSize: number
  resolveOffset(index: number): number
  resolveAnchor(scrollOffset: number): AxisAnchor
  resolveSize(index: number): number
  resolveSpan(start: number, endExclusive: number): number
  resolveVisibleCount(start: number, viewportSize: number, overscanPx: number): number
}

interface NormalizedAxisEntry {
  readonly index: number
  readonly size: number
  readonly deltaPrefixBefore: number
}

export function createAxisIndex(input: {
  readonly axisLength: number
  readonly defaultSize: number
  readonly overrides?: readonly AxisEntryOverride[]
}): AxisIndex {
  const axisLength = Math.max(1, Math.floor(input.axisLength))
  const defaultSize = Math.max(0, input.defaultSize)
  const entries = normalizeOverrides(axisLength, defaultSize, input.overrides ?? [])
  if (entries.length === 0) {
    return createDefaultAxisIndex(axisLength, defaultSize)
  }
  return createSparseAxisIndex(axisLength, defaultSize, entries)
}

export function createAxisIndexFromRecord(input: {
  readonly axisLength: number
  readonly defaultSize: number
  readonly sizes: Readonly<Record<number, number>>
}): AxisIndex {
  return createAxisIndex({
    axisLength: input.axisLength,
    defaultSize: input.defaultSize,
    overrides: Object.entries(input.sizes).map(([index, size]) => ({
      index: Number(index),
      size,
    })),
  })
}

export function createAxisIndexFromSortedOverrides(input: {
  readonly axisLength: number
  readonly defaultSize: number
  readonly sortedOverrides: readonly (readonly [number, number])[]
}): AxisIndex {
  return createAxisIndex({
    axisLength: input.axisLength,
    defaultSize: input.defaultSize,
    overrides: input.sortedOverrides.map(([index, size]) => ({ index, size })),
  })
}

function createDefaultAxisIndex(axisLength: number, defaultSize: number): AxisIndex {
  const totalSize = axisLength * defaultSize
  return {
    axisLength,
    defaultSize,
    resolveAnchor(scrollOffset) {
      if (defaultSize <= 0 || scrollOffset >= totalSize) {
        return { index: axisLength - 1, offset: 0 }
      }
      const clamped = Math.max(0, scrollOffset)
      const index = Math.min(axisLength - 1, Math.floor(clamped / defaultSize))
      return {
        index,
        offset: clamped - index * defaultSize,
      }
    },
    resolveOffset(index) {
      return clampAxisIndex(index, axisLength) * defaultSize
    },
    resolveSize(index) {
      return index >= 0 && index < axisLength ? defaultSize : 0
    },
    resolveSpan(start, endExclusive) {
      const clampedStart = clampAxisIndex(start, axisLength)
      const clampedEnd = clampAxisEndExclusive(endExclusive, axisLength)
      return Math.max(0, clampedEnd - clampedStart) * defaultSize
    },
    resolveVisibleCount(start, viewportSize, overscanPx) {
      if (defaultSize <= 0) {
        return 1
      }
      const clampedStart = clampAxisIndex(start, axisLength)
      const targetSize = Math.max(0, viewportSize) + Math.max(0, overscanPx)
      const count = Math.ceil(targetSize / defaultSize)
      return Math.max(1, Math.min(axisLength - clampedStart, count))
    },
  }
}

function createSparseAxisIndex(axisLength: number, defaultSize: number, entries: readonly NormalizedAxisEntry[]): AxisIndex {
  const entryByIndex = new Map(entries.map((entry) => [entry.index, entry]))
  const totalSize = resolveSparseOffset(axisLength, axisLength, defaultSize, entries)

  return {
    axisLength,
    defaultSize,
    resolveAnchor(scrollOffset) {
      if (scrollOffset >= totalSize) {
        return { index: axisLength - 1, offset: 0 }
      }
      const target = Math.max(0, scrollOffset)
      let low = 0
      let high = axisLength - 1
      while (low < high) {
        const mid = Math.floor((low + high) / 2)
        if (resolveSparseOffset(mid + 1, axisLength, defaultSize, entries) > target) {
          high = mid
        } else {
          low = mid + 1
        }
      }
      return {
        index: low,
        offset: Math.max(0, target - resolveSparseOffset(low, axisLength, defaultSize, entries)),
      }
    },
    resolveOffset(index) {
      return resolveSparseOffset(index, axisLength, defaultSize, entries)
    },
    resolveSize(index) {
      const clamped = Math.floor(index)
      if (clamped < 0 || clamped >= axisLength) {
        return 0
      }
      return entryByIndex.get(clamped)?.size ?? defaultSize
    },
    resolveSpan(start, endExclusive) {
      return Math.max(
        0,
        resolveSparseOffset(endExclusive, axisLength, defaultSize, entries) - resolveSparseOffset(start, axisLength, defaultSize, entries),
      )
    },
    resolveVisibleCount(start, viewportSize, overscanPx) {
      const clampedStart = clampAxisIndex(start, axisLength)
      const startOffset = resolveSparseOffset(clampedStart, axisLength, defaultSize, entries)
      const targetOffset = startOffset + Math.max(0, viewportSize) + Math.max(0, overscanPx)
      let low = clampedStart + 1
      let high = axisLength
      while (low < high) {
        const mid = Math.floor((low + high) / 2)
        if (resolveSparseOffset(mid, axisLength, defaultSize, entries) >= targetOffset) {
          high = mid
        } else {
          low = mid + 1
        }
      }
      return Math.max(1, Math.min(axisLength - clampedStart, low - clampedStart))
    },
  }
}

function normalizeOverrides(
  axisLength: number,
  defaultSize: number,
  overrides: readonly AxisEntryOverride[],
): readonly NormalizedAxisEntry[] {
  const byIndex = new Map<number, number>()
  for (const override of overrides) {
    const index = Math.floor(override.index)
    if (index < 0 || index >= axisLength) {
      continue
    }
    const size = override.hidden ? 0 : Math.max(0, override.size)
    if (size === defaultSize) {
      byIndex.delete(index)
      continue
    }
    byIndex.set(index, size)
  }

  let deltaPrefix = 0
  return [...byIndex.entries()]
    .toSorted((left, right) => left[0] - right[0])
    .map(([index, size]) => {
      const entry = {
        deltaPrefixBefore: deltaPrefix,
        index,
        size,
      }
      deltaPrefix += size - defaultSize
      return entry
    })
}

function resolveSparseOffset(index: number, axisLength: number, defaultSize: number, entries: readonly NormalizedAxisEntry[]): number {
  const clamped = clampAxisEndExclusive(index, axisLength)
  const entryBeforeIndex = findLastEntryBefore(entries, clamped)
  const deltaPrefix = entryBeforeIndex === null ? 0 : entryBeforeIndex.deltaPrefixBefore + entryBeforeIndex.size - defaultSize
  return clamped * defaultSize + deltaPrefix
}

function findLastEntryBefore(entries: readonly NormalizedAxisEntry[], index: number): NormalizedAxisEntry | null {
  let low = 0
  let high = entries.length
  while (low < high) {
    const mid = Math.floor((low + high) / 2)
    const entry = entries[mid]
    if (!entry || entry.index >= index) {
      high = mid
    } else {
      low = mid + 1
    }
  }
  return entries[low - 1] ?? null
}

function clampAxisIndex(index: number, axisLength: number): number {
  return Math.max(0, Math.min(axisLength - 1, Math.floor(index)))
}

function clampAxisEndExclusive(index: number, axisLength: number): number {
  return Math.max(0, Math.min(axisLength, Math.floor(index)))
}
