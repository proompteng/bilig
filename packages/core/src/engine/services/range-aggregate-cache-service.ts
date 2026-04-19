import type { AggregateStateEntry, AggregateStateStore } from '../../deps/aggregate-state-store.js'
import type { RegionGraph } from '../../deps/region-graph.js'

export interface RangeAggregateCacheService {
  readonly getOrBuildPrefix: (
    request: { sheetName: string; rowStart: number; rowEnd: number; col: number },
    aggregateKind?: 'sum' | 'average' | 'count' | 'min' | 'max',
  ) => AggregateStateEntry
}

export function createRangeAggregateCacheService(args: {
  readonly regionGraph: Pick<RegionGraph, 'internSingleColumnRegion'>
  readonly aggregateStateStore: AggregateStateStore
}): RangeAggregateCacheService {
  return {
    getOrBuildPrefix(request, aggregateKind) {
      const regionId = args.regionGraph.internSingleColumnRegion(request)
      return args.aggregateStateStore.getOrBuildPrefixForRegion(regionId, aggregateKind)
    },
  }
}
