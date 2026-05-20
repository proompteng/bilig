import type { AggregateColumnWindowSummary, AggregateStateEntry, AggregateStateStore } from '../../deps/aggregate-state-store.js'
import type { RegionGraph } from '../../deps/region-graph.js'

export interface RangeAggregateCacheService {
  readonly getOrBuildPrefix: (
    request: { sheetName: string; rowStart: number; rowEnd: number; col: number },
    aggregateKind?: 'sum' | 'average' | 'count' | 'min' | 'max',
  ) => AggregateStateEntry
  readonly getOrBuildColumnPrefix: (
    request: { sheetName: string; rowStart: number; rowEnd: number; col: number },
    aggregateKind?: 'sum' | 'average' | 'count' | 'min' | 'max',
  ) => AggregateStateEntry
  readonly summarizeColumnWindow: (request: {
    sheetName: string
    rowStart: number
    rowEnd: number
    col: number
  }) => AggregateColumnWindowSummary | undefined
  readonly hasReusableColumnPrefix: (
    request: { sheetName: string; rowStart: number; rowEnd: number; col: number },
    aggregateKind?: 'sum' | 'average' | 'count' | 'min' | 'max',
  ) => boolean
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
    getOrBuildColumnPrefix(request, aggregateKind) {
      const regionId = args.regionGraph.internSingleColumnRegion(request)
      return args.aggregateStateStore.getOrBuildPrefixForRegion(regionId, aggregateKind)
    },
    summarizeColumnWindow(request) {
      return args.aggregateStateStore.summarizeColumnWindow(request)
    },
    hasReusableColumnPrefix(request, aggregateKind) {
      return args.aggregateStateStore.hasReusableColumnPrefix(request, aggregateKind)
    },
  }
}
