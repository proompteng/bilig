import type { RegionId } from './region-node-store.js'

export interface IntervalRegionRef {
  readonly regionId: RegionId
  readonly rowStart: number
  readonly rowEnd: number
}

export interface IntervalTreeNode {
  readonly center: number
  readonly overlappingByStart: readonly IntervalRegionRef[]
  readonly overlappingByEnd: readonly IntervalRegionRef[]
  readonly left: IntervalTreeNode | undefined
  readonly right: IntervalTreeNode | undefined
}

export function buildIntervalTree(intervals: readonly IntervalRegionRef[]): IntervalTreeNode | undefined {
  if (intervals.length === 0) {
    return undefined
  }
  const centers = intervals.map((interval) => Math.floor((interval.rowStart + interval.rowEnd) / 2)).toSorted((a, b) => a - b)
  const center = centers[Math.floor(centers.length / 2)]!
  const left: IntervalRegionRef[] = []
  const right: IntervalRegionRef[] = []
  const overlapping: IntervalRegionRef[] = []
  intervals.forEach((interval) => {
    if (interval.rowEnd < center) {
      left.push(interval)
      return
    }
    if (interval.rowStart > center) {
      right.push(interval)
      return
    }
    overlapping.push(interval)
  })
  return {
    center,
    overlappingByStart: overlapping.toSorted((a, b) => a.rowStart - b.rowStart || a.rowEnd - b.rowEnd || a.regionId - b.regionId),
    overlappingByEnd: overlapping.toSorted((a, b) => b.rowEnd - a.rowEnd || b.rowStart - a.rowStart || a.regionId - b.regionId),
    left: buildIntervalTree(left),
    right: buildIntervalTree(right),
  }
}

export function collectIntervalsContainingRow(node: IntervalTreeNode | undefined, row: number, target: RegionId[]): void {
  if (!node) {
    return
  }
  if (row < node.center) {
    for (let index = 0; index < node.overlappingByStart.length; index += 1) {
      const interval = node.overlappingByStart[index]!
      if (interval.rowStart > row) {
        break
      }
      target.push(interval.regionId)
    }
    collectIntervalsContainingRow(node.left, row, target)
    return
  }
  if (row > node.center) {
    for (let index = 0; index < node.overlappingByEnd.length; index += 1) {
      const interval = node.overlappingByEnd[index]!
      if (interval.rowEnd < row) {
        break
      }
      target.push(interval.regionId)
    }
    collectIntervalsContainingRow(node.right, row, target)
    return
  }
  for (let index = 0; index < node.overlappingByStart.length; index += 1) {
    target.push(node.overlappingByStart[index]!.regionId)
  }
}
