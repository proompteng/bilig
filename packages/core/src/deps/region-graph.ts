import type { WorkbookStore } from '../workbook-store.js'
import { addEngineCounter, type EngineCounters } from '../perf/engine-counters.js'
import { createRegionNodeStore, type RegionId, type RegionNodeStore, type SingleColumnRegionNode } from './region-node-store.js'

interface IntervalRegionRef {
  readonly regionId: RegionId
  readonly rowStart: number
  readonly rowEnd: number
}

interface IntervalTreeNode {
  readonly center: number
  readonly overlappingByStart: readonly IntervalRegionRef[]
  readonly overlappingByEnd: readonly IntervalRegionRef[]
  readonly left: IntervalTreeNode | undefined
  readonly right: IntervalTreeNode | undefined
}

interface ColumnSubscriptionState {
  regionIds: Set<RegionId>
  tree: IntervalTreeNode | undefined
  dirty: boolean
}

export interface RegionGraph {
  readonly internSingleColumnRegion: (args: {
    readonly sheetName: string
    readonly rowStart: number
    readonly rowEnd: number
    readonly col: number
  }) => RegionId
  readonly getRegion: (regionId: RegionId) => SingleColumnRegionNode | undefined
  readonly replaceFormulaSubscriptions: (formulaCellIndex: number, regionIds: readonly RegionId[]) => void
  readonly clearFormulaSubscriptions: (formulaCellIndex: number) => void
  readonly getFormulaSubscriptions: (formulaCellIndex: number) => readonly RegionId[]
  readonly prepareQueryIndices: () => void
  readonly collectFormulaDependentsForCell: (sheetId: number, row: number, col: number) => Uint32Array
  readonly hasFormulaSubscriptionsForColumn: (sheetId: number, col: number) => boolean
  readonly reset: () => void
}

function columnKey(sheetId: number, col: number): string {
  return `${sheetId}\t${col}`
}

function buildIntervalTree(intervals: readonly IntervalRegionRef[]): IntervalTreeNode | undefined {
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

function collectIntervalsContainingRow(node: IntervalTreeNode | undefined, row: number, target: Set<RegionId>): void {
  if (!node) {
    return
  }
  if (row < node.center) {
    for (let index = 0; index < node.overlappingByStart.length; index += 1) {
      const interval = node.overlappingByStart[index]!
      if (interval.rowStart > row) {
        break
      }
      target.add(interval.regionId)
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
      target.add(interval.regionId)
    }
    collectIntervalsContainingRow(node.right, row, target)
    return
  }
  for (let index = 0; index < node.overlappingByStart.length; index += 1) {
    target.add(node.overlappingByStart[index]!.regionId)
  }
}

export function createRegionGraph(args: {
  readonly workbook: Pick<WorkbookStore, 'getSheet'>
  readonly counters?: EngineCounters
  readonly nodeStore?: RegionNodeStore
}): RegionGraph {
  const nodeStore = args.nodeStore ?? createRegionNodeStore()
  const regionSubscribers = new Map<RegionId, Set<number>>()
  const formulaRegions = new Map<number, Set<RegionId>>()
  const columnSubscriptions = new Map<string, ColumnSubscriptionState>()
  const dirtyColumns = new Set<string>()

  const getColumnSubscriptions = (sheetId: number, col: number): ColumnSubscriptionState => {
    const key = columnKey(sheetId, col)
    const existing = columnSubscriptions.get(key)
    if (existing) {
      return existing
    }
    const created: ColumnSubscriptionState = {
      regionIds: new Set(),
      tree: undefined,
      dirty: false,
    }
    columnSubscriptions.set(key, created)
    return created
  }

  const markColumnDirty = (regionId: RegionId): void => {
    const region = nodeStore.get(regionId)
    if (!region) {
      return
    }
    const key = columnKey(region.sheetId, region.col)
    getColumnSubscriptions(region.sheetId, region.col).dirty = true
    dirtyColumns.add(key)
  }

  const addRegionSubscription = (formulaCellIndex: number, regionId: RegionId): void => {
    let subscribers = regionSubscribers.get(regionId)
    const wasEmpty = subscribers === undefined || subscribers.size === 0
    if (!subscribers) {
      subscribers = new Set()
      regionSubscribers.set(regionId, subscribers)
    }
    subscribers.add(formulaCellIndex)
    if (wasEmpty) {
      const region = nodeStore.get(regionId)
      if (region) {
        getColumnSubscriptions(region.sheetId, region.col).regionIds.add(regionId)
        markColumnDirty(regionId)
      }
    }
  }

  const removeRegionSubscription = (formulaCellIndex: number, regionId: RegionId): void => {
    const subscribers = regionSubscribers.get(regionId)
    if (!subscribers) {
      return
    }
    subscribers.delete(formulaCellIndex)
    if (subscribers.size > 0) {
      return
    }
    regionSubscribers.delete(regionId)
    const region = nodeStore.get(regionId)
    if (!region) {
      return
    }
    const subscriptions = getColumnSubscriptions(region.sheetId, region.col)
    subscriptions.regionIds.delete(regionId)
    markColumnDirty(regionId)
  }

  const ensureIntervalTree = (sheetId: number, col: number): IntervalTreeNode | undefined => {
    const key = columnKey(sheetId, col)
    const subscriptions = getColumnSubscriptions(sheetId, col)
    if (!subscriptions.dirty) {
      return subscriptions.tree
    }
    const intervals: IntervalRegionRef[] = []
    subscriptions.regionIds.forEach((regionId) => {
      const region = nodeStore.get(regionId)
      if (!region) {
        return
      }
      intervals.push({
        regionId,
        rowStart: region.rowStart,
        rowEnd: region.rowEnd,
      })
    })
    if (args.counters) {
      addEngineCounter(args.counters, 'regionQueryIndexBuilds')
    }
    subscriptions.tree = buildIntervalTree(intervals)
    subscriptions.dirty = false
    dirtyColumns.delete(key)
    return subscriptions.tree
  }

  return {
    internSingleColumnRegion({ sheetName, rowStart, rowEnd, col }) {
      const sheet = args.workbook.getSheet(sheetName)
      if (!sheet) {
        throw new Error(`Unknown sheet for region graph: ${sheetName}`)
      }
      return nodeStore.internSingleColumnRegion({
        sheetId: sheet.id,
        sheetName,
        rowStart,
        rowEnd,
        col,
      })
    },
    getRegion(regionId) {
      return nodeStore.get(regionId)
    },
    replaceFormulaSubscriptions(formulaCellIndex, regionIds) {
      const previous = formulaRegions.get(formulaCellIndex)
      if (previous) {
        previous.forEach((regionId) => {
          removeRegionSubscription(formulaCellIndex, regionId)
        })
      }
      if (regionIds.length === 0) {
        formulaRegions.delete(formulaCellIndex)
        return
      }
      const next = new Set<RegionId>()
      regionIds.forEach((regionId) => {
        if (next.has(regionId)) {
          return
        }
        next.add(regionId)
        addRegionSubscription(formulaCellIndex, regionId)
      })
      formulaRegions.set(formulaCellIndex, next)
    },
    clearFormulaSubscriptions(formulaCellIndex) {
      const previous = formulaRegions.get(formulaCellIndex)
      if (!previous) {
        return
      }
      previous.forEach((regionId) => {
        removeRegionSubscription(formulaCellIndex, regionId)
      })
      formulaRegions.delete(formulaCellIndex)
    },
    getFormulaSubscriptions(formulaCellIndex) {
      return [...(formulaRegions.get(formulaCellIndex) ?? [])]
    },
    prepareQueryIndices() {
      dirtyColumns.forEach((key) => {
        const [sheetIdText, colText] = key.split('\t')
        if (sheetIdText === undefined || colText === undefined) {
          return
        }
        const sheetId = Number(sheetIdText)
        const col = Number(colText)
        if (!Number.isNaN(sheetId) && !Number.isNaN(col)) {
          ensureIntervalTree(sheetId, col)
        }
      })
    },
    collectFormulaDependentsForCell(sheetId, row, col) {
      const matchingRegions = new Set<RegionId>()
      collectIntervalsContainingRow(ensureIntervalTree(sheetId, col), row, matchingRegions)
      const dependents = new Set<number>()
      matchingRegions.forEach((regionId) => {
        regionSubscribers.get(regionId)?.forEach((formulaCellIndex) => {
          dependents.add(formulaCellIndex)
        })
      })
      return Uint32Array.from(dependents)
    },
    hasFormulaSubscriptionsForColumn(sheetId, col) {
      return getColumnSubscriptions(sheetId, col).regionIds.size > 0
    },
    reset() {
      regionSubscribers.clear()
      formulaRegions.clear()
      columnSubscriptions.clear()
      dirtyColumns.clear()
    },
  }
}
