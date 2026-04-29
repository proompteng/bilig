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
  dirtyPointQueries: number
  orderedByRowStart: boolean
  lastRowStart: number
  pointImpactIndex: Array<number | undefined> | undefined
  pointImpactCellCount: number
  pointImpactDirty: boolean
  pointImpactDisabled: boolean
}

type FormulaRegionSubscription = RegionId | Set<RegionId>
type RegionSubscriberSubscription = number | Set<number>

const DIRTY_LINEAR_POINT_QUERY_LIMIT = 2
const DIRTY_LINEAR_REGION_LIMIT = 2_048
const MULTIPLE_POINT_DEPENDENTS = -2
const POINT_IMPACT_MAX_REGION_LENGTH = 256
const POINT_IMPACT_MAX_CELLS = 65_536
const POINT_IMPACT_MAX_ROW_INDEX = 131_072

export interface RegionGraph {
  readonly internSingleColumnRegion: (args: {
    readonly sheetName: string
    readonly rowStart: number
    readonly rowEnd: number
    readonly col: number
  }) => RegionId
  readonly getRegion: (regionId: RegionId) => SingleColumnRegionNode | undefined
  readonly replaceFormulaSubscriptions: (formulaCellIndex: number, regionIds: readonly RegionId[]) => void
  readonly replaceSingleFormulaSubscription: (formulaCellIndex: number, previousRegionId: RegionId, nextRegionId: RegionId) => void
  readonly clearFormulaSubscriptions: (formulaCellIndex: number) => void
  readonly getFormulaSubscriptions: (formulaCellIndex: number) => readonly RegionId[]
  readonly prepareQueryIndices: () => void
  readonly collectFormulaDependentsForCell: (sheetId: number, row: number, col: number) => Uint32Array
  readonly collectSingleFormulaDependentForCell: (sheetId: number, row: number, col: number) => number
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

function collectIntervalsContainingRow(node: IntervalTreeNode | undefined, row: number, target: RegionId[]): void {
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

function collectRegionIdsContainingRow(
  regionIds: ReadonlySet<RegionId>,
  row: number,
  nodeStore: RegionNodeStore,
  target: RegionId[],
  orderedByRowStart: boolean,
): void {
  for (const regionId of regionIds) {
    const region = nodeStore.get(regionId)
    if (!region) {
      continue
    }
    if (orderedByRowStart && region.rowStart > row) {
      break
    }
    if (region.rowStart <= row && row <= region.rowEnd) {
      target.push(regionId)
    }
  }
}

function pushUniqueDependent(dependents: number[], formulaCellIndex: number): void {
  for (let index = 0; index < dependents.length; index += 1) {
    if (dependents[index] === formulaCellIndex) {
      return
    }
  }
  dependents.push(formulaCellIndex)
}

function mergeSingleDependent(current: number, next: number): number {
  if (current === MULTIPLE_POINT_DEPENDENTS || next === MULTIPLE_POINT_DEPENDENTS) {
    return MULTIPLE_POINT_DEPENDENTS
  }
  if (current === -1) {
    return next
  }
  return current === next ? current : MULTIPLE_POINT_DEPENDENTS
}

function mergePointImpactDependent(current: number | undefined, next: number): number {
  if (current === undefined || current === next) {
    return next
  }
  return MULTIPLE_POINT_DEPENDENTS
}

function forEachFormulaRegionSubscription(subscription: FormulaRegionSubscription, fn: (regionId: RegionId) => void): void {
  if (typeof subscription === 'number') {
    fn(subscription)
    return
  }
  subscription.forEach(fn)
}

function forEachRegionSubscriber(subscription: RegionSubscriberSubscription, fn: (formulaCellIndex: number) => void): void {
  if (typeof subscription === 'number') {
    fn(subscription)
    return
  }
  subscription.forEach(fn)
}

function regionSubscriberSize(subscription: RegionSubscriberSubscription): number {
  return typeof subscription === 'number' ? 1 : subscription.size
}

function regionSubscribersToUint32(subscription: RegionSubscriberSubscription | undefined): Uint32Array {
  if (subscription === undefined) {
    return new Uint32Array()
  }
  return typeof subscription === 'number' ? Uint32Array.of(subscription) : Uint32Array.from(subscription)
}

function formulaRegionSubscriptionToArray(subscription: FormulaRegionSubscription | undefined): readonly RegionId[] {
  if (subscription === undefined) {
    return []
  }
  return typeof subscription === 'number' ? [subscription] : [...subscription]
}

function collectSingleDependentForRegionIds(
  regionIds: ReadonlySet<RegionId>,
  row: number,
  nodeStore: RegionNodeStore,
  regionSubscribers: ReadonlyMap<RegionId, RegionSubscriberSubscription>,
  orderedByRowStart: boolean,
): number {
  let singleDependent = -1
  const mergeRegion = (regionId: RegionId): void => {
    const region = nodeStore.get(regionId)
    if (!region || region.rowStart > row || region.rowEnd < row) {
      return
    }
    const subscribers = regionSubscribers.get(regionId)
    if (subscribers === undefined || regionSubscriberSize(subscribers) === 0) {
      return
    }
    if (regionSubscriberSize(subscribers) !== 1) {
      singleDependent = -2
      return
    }
    forEachRegionSubscriber(subscribers, (formulaCellIndex) => {
      singleDependent = mergeSingleDependent(singleDependent, formulaCellIndex)
    })
  }
  for (const regionId of regionIds) {
    if (orderedByRowStart) {
      const region = nodeStore.get(regionId)
      if (region && region.rowStart > row) {
        break
      }
    }
    mergeRegion(regionId)
    if (singleDependent === -2) {
      return singleDependent
    }
  }
  return singleDependent
}

function disablePointImpactIndex(subscriptions: ColumnSubscriptionState): void {
  subscriptions.pointImpactIndex = undefined
  subscriptions.pointImpactCellCount = 0
  subscriptions.pointImpactDirty = false
  subscriptions.pointImpactDisabled = true
}

function markPointImpactIndexDirty(subscriptions: ColumnSubscriptionState): void {
  if (subscriptions.pointImpactIndex !== undefined) {
    subscriptions.pointImpactDirty = true
  }
}

function addSingleSubscriberPointImpacts(
  subscriptions: ColumnSubscriptionState,
  region: SingleColumnRegionNode,
  formulaCellIndex: number,
): void {
  if (subscriptions.pointImpactDisabled || subscriptions.pointImpactDirty) {
    return
  }
  const regionLength = region.rowEnd - region.rowStart + 1
  if (
    regionLength <= 0 ||
    region.rowEnd > POINT_IMPACT_MAX_ROW_INDEX ||
    regionLength > POINT_IMPACT_MAX_REGION_LENGTH ||
    subscriptions.pointImpactCellCount + regionLength > POINT_IMPACT_MAX_CELLS
  ) {
    disablePointImpactIndex(subscriptions)
    return
  }
  const index = (subscriptions.pointImpactIndex ??= [])
  for (let row = region.rowStart; row <= region.rowEnd; row += 1) {
    index[row] = mergePointImpactDependent(index[row], formulaCellIndex)
  }
  subscriptions.pointImpactCellCount += regionLength
}

function collectSingleDependentFromPointImpactIndex(subscriptions: ColumnSubscriptionState, row: number): number | undefined {
  if (subscriptions.pointImpactDirty || subscriptions.pointImpactDisabled || subscriptions.pointImpactIndex === undefined) {
    return undefined
  }
  return subscriptions.pointImpactIndex[row] ?? -1
}

function collectSingleDependentForRow(
  node: IntervalTreeNode | undefined,
  row: number,
  regionSubscribers: ReadonlyMap<RegionId, RegionSubscriberSubscription>,
): number {
  let singleDependent = -1
  const mergeRegion = (regionId: RegionId): void => {
    const subscribers = regionSubscribers.get(regionId)
    if (subscribers === undefined || regionSubscriberSize(subscribers) === 0) {
      return
    }
    if (regionSubscriberSize(subscribers) !== 1) {
      singleDependent = -2
      return
    }
    forEachRegionSubscriber(subscribers, (formulaCellIndex) => {
      singleDependent = mergeSingleDependent(singleDependent, formulaCellIndex)
    })
  }
  const visit = (current: IntervalTreeNode | undefined): void => {
    if (!current || singleDependent === -2) {
      return
    }
    if (row < current.center) {
      for (let index = 0; index < current.overlappingByStart.length; index += 1) {
        const interval = current.overlappingByStart[index]!
        if (interval.rowStart > row) {
          break
        }
        mergeRegion(interval.regionId)
        if (singleDependent === -2) {
          return
        }
      }
      visit(current.left)
      return
    }
    if (row > current.center) {
      for (let index = 0; index < current.overlappingByEnd.length; index += 1) {
        const interval = current.overlappingByEnd[index]!
        if (interval.rowEnd < row) {
          break
        }
        mergeRegion(interval.regionId)
        if (singleDependent === -2) {
          return
        }
      }
      visit(current.right)
      return
    }
    for (let index = 0; index < current.overlappingByStart.length; index += 1) {
      mergeRegion(current.overlappingByStart[index]!.regionId)
      if (singleDependent === -2) {
        return
      }
    }
  }
  visit(node)
  return singleDependent
}

export function createRegionGraph(args: {
  readonly workbook: Pick<WorkbookStore, 'getSheet'>
  readonly counters?: EngineCounters
  readonly nodeStore?: RegionNodeStore
}): RegionGraph {
  const nodeStore = args.nodeStore ?? createRegionNodeStore()
  const regionSubscribers = new Map<RegionId, RegionSubscriberSubscription>()
  const formulaRegions = new Map<number, FormulaRegionSubscription>()
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
      dirtyPointQueries: 0,
      orderedByRowStart: true,
      lastRowStart: Number.NEGATIVE_INFINITY,
      pointImpactIndex: undefined,
      pointImpactCellCount: 0,
      pointImpactDirty: false,
      pointImpactDisabled: false,
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
    const subscriptions = getColumnSubscriptions(region.sheetId, region.col)
    subscriptions.dirty = true
    subscriptions.dirtyPointQueries = 0
    dirtyColumns.add(key)
  }

  const addRegionSubscription = (formulaCellIndex: number, regionId: RegionId): void => {
    const subscribers = regionSubscribers.get(regionId)
    const wasEmpty = subscribers === undefined
    if (subscribers === undefined) {
      regionSubscribers.set(regionId, formulaCellIndex)
    } else if (typeof subscribers === 'number') {
      if (subscribers !== formulaCellIndex) {
        regionSubscribers.set(regionId, new Set([subscribers, formulaCellIndex]))
        const region = nodeStore.get(regionId)
        if (region) {
          markPointImpactIndexDirty(getColumnSubscriptions(region.sheetId, region.col))
        }
      }
    } else {
      const previousSize = subscribers.size
      subscribers.add(formulaCellIndex)
      if (subscribers.size !== previousSize) {
        const region = nodeStore.get(regionId)
        if (region) {
          markPointImpactIndexDirty(getColumnSubscriptions(region.sheetId, region.col))
        }
      }
    }
    if (wasEmpty) {
      const region = nodeStore.get(regionId)
      if (region) {
        const subscriptions = getColumnSubscriptions(region.sheetId, region.col)
        if (region.rowStart < subscriptions.lastRowStart) {
          subscriptions.orderedByRowStart = false
        }
        subscriptions.lastRowStart = Math.max(subscriptions.lastRowStart, region.rowStart)
        subscriptions.regionIds.add(regionId)
        addSingleSubscriberPointImpacts(subscriptions, region, formulaCellIndex)
        markColumnDirty(regionId)
      }
    }
  }

  const removeRegionSubscription = (formulaCellIndex: number, regionId: RegionId): void => {
    const subscribers = regionSubscribers.get(regionId)
    if (subscribers === undefined) {
      return
    }
    let removedSubscriber = false
    let shouldRemoveRegion = false
    if (typeof subscribers === 'number') {
      if (subscribers !== formulaCellIndex) {
        return
      }
      removedSubscriber = true
      shouldRemoveRegion = true
      regionSubscribers.delete(regionId)
    } else {
      const previousSize = subscribers.size
      subscribers.delete(formulaCellIndex)
      if (subscribers.size === previousSize) {
        return
      }
      removedSubscriber = true
      if (subscribers.size > 1) {
        shouldRemoveRegion = false
      } else if (subscribers.size === 1) {
        for (const remaining of subscribers) {
          regionSubscribers.set(regionId, remaining)
        }
        shouldRemoveRegion = false
      } else {
        shouldRemoveRegion = true
        regionSubscribers.delete(regionId)
      }
    }
    const region = nodeStore.get(regionId)
    if (!region) {
      return
    }
    const subscriptions = getColumnSubscriptions(region.sheetId, region.col)
    if (removedSubscriber) {
      markPointImpactIndexDirty(subscriptions)
    }
    if (!shouldRemoveRegion) {
      return
    }
    subscriptions.regionIds.delete(regionId)
    if (subscriptions.regionIds.size === 0) {
      subscriptions.orderedByRowStart = true
      subscriptions.lastRowStart = Number.NEGATIVE_INFINITY
      subscriptions.pointImpactIndex = undefined
      subscriptions.pointImpactCellCount = 0
      subscriptions.pointImpactDirty = false
      subscriptions.pointImpactDisabled = false
    }
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
      if (previous !== undefined) {
        forEachFormulaRegionSubscription(previous, (regionId) => {
          removeRegionSubscription(formulaCellIndex, regionId)
        })
      }
      if (regionIds.length === 0) {
        formulaRegions.delete(formulaCellIndex)
        return
      }
      if (regionIds.length === 1) {
        const regionId = regionIds[0]!
        addRegionSubscription(formulaCellIndex, regionId)
        formulaRegions.set(formulaCellIndex, regionId)
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
    replaceSingleFormulaSubscription(formulaCellIndex, previousRegionId, nextRegionId) {
      if (previousRegionId === nextRegionId) {
        return
      }
      const previous = formulaRegions.get(formulaCellIndex)
      if (previous === previousRegionId) {
        removeRegionSubscription(formulaCellIndex, previousRegionId)
        addRegionSubscription(formulaCellIndex, nextRegionId)
        formulaRegions.set(formulaCellIndex, nextRegionId)
        return
      }
      if (previous === undefined || typeof previous === 'number' || previous.size !== 1 || !previous.has(previousRegionId)) {
        if (previous !== undefined) {
          forEachFormulaRegionSubscription(previous, (regionId) => {
            removeRegionSubscription(formulaCellIndex, regionId)
          })
        }
        addRegionSubscription(formulaCellIndex, nextRegionId)
        formulaRegions.set(formulaCellIndex, nextRegionId)
        return
      }
      removeRegionSubscription(formulaCellIndex, previousRegionId)
      previous.clear()
      previous.add(nextRegionId)
      addRegionSubscription(formulaCellIndex, nextRegionId)
    },
    clearFormulaSubscriptions(formulaCellIndex) {
      const previous = formulaRegions.get(formulaCellIndex)
      if (previous === undefined) {
        return
      }
      forEachFormulaRegionSubscription(previous, (regionId) => {
        removeRegionSubscription(formulaCellIndex, regionId)
      })
      formulaRegions.delete(formulaCellIndex)
    },
    getFormulaSubscriptions(formulaCellIndex) {
      return formulaRegionSubscriptionToArray(formulaRegions.get(formulaCellIndex))
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
      const matchingRegions: RegionId[] = []
      const subscriptions = getColumnSubscriptions(sheetId, col)
      if (
        subscriptions.dirty &&
        subscriptions.dirtyPointQueries < DIRTY_LINEAR_POINT_QUERY_LIMIT &&
        subscriptions.regionIds.size <= DIRTY_LINEAR_REGION_LIMIT
      ) {
        subscriptions.dirtyPointQueries += 1
        collectRegionIdsContainingRow(subscriptions.regionIds, row, nodeStore, matchingRegions, subscriptions.orderedByRowStart)
      } else {
        collectIntervalsContainingRow(ensureIntervalTree(sheetId, col), row, matchingRegions)
      }
      if (matchingRegions.length === 0) {
        return new Uint32Array()
      }
      if (matchingRegions.length === 1) {
        const subscribers = regionSubscribers.get(matchingRegions[0]!)
        return regionSubscribersToUint32(subscribers)
      }
      const dependents: number[] = []
      for (let regionIndex = 0; regionIndex < matchingRegions.length; regionIndex += 1) {
        const subscribers = regionSubscribers.get(matchingRegions[regionIndex]!)
        if (subscribers === undefined) {
          continue
        }
        forEachRegionSubscriber(subscribers, (formulaCellIndex) => {
          pushUniqueDependent(dependents, formulaCellIndex)
        })
      }
      return Uint32Array.from(dependents)
    },
    collectSingleFormulaDependentForCell(sheetId, row, col) {
      const subscriptions = getColumnSubscriptions(sheetId, col)
      const indexedDependent = collectSingleDependentFromPointImpactIndex(subscriptions, row)
      if (indexedDependent !== undefined) {
        return indexedDependent
      }
      if (
        subscriptions.dirty &&
        subscriptions.dirtyPointQueries < DIRTY_LINEAR_POINT_QUERY_LIMIT &&
        subscriptions.regionIds.size <= DIRTY_LINEAR_REGION_LIMIT
      ) {
        subscriptions.dirtyPointQueries += 1
        return collectSingleDependentForRegionIds(
          subscriptions.regionIds,
          row,
          nodeStore,
          regionSubscribers,
          subscriptions.orderedByRowStart,
        )
      }
      return collectSingleDependentForRow(ensureIntervalTree(sheetId, col), row, regionSubscribers)
    },
    hasFormulaSubscriptionsForColumn(sheetId, col) {
      return (columnSubscriptions.get(columnKey(sheetId, col))?.regionIds.size ?? 0) > 0
    },
    reset() {
      regionSubscribers.clear()
      formulaRegions.clear()
      columnSubscriptions.clear()
      dirtyColumns.clear()
    },
  }
}
