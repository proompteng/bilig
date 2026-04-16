import type { CellValue } from '@bilig/protocol'
import type { PreparedApproximateVectorLookup, PreparedExactVectorLookup } from '../runtime-state.js'
import type { ExactColumnIndexService, ExactVectorMatchRequest, ExactVectorMatchResult } from './exact-column-index-service.js'
import type {
  ApproximateVectorMatchRequest,
  ApproximateVectorMatchResult,
  SortedColumnSearchService,
} from './sorted-column-search-service.js'

export type { ExactVectorMatchRequest, ExactVectorMatchResult } from './exact-column-index-service.js'
export type { ApproximateVectorMatchRequest, ApproximateVectorMatchResult } from './sorted-column-search-service.js'

export interface EngineLookupService {
  readonly findExactVectorMatch: (request: ExactVectorMatchRequest) => ExactVectorMatchResult
  readonly findApproximateVectorMatch: (request: ApproximateVectorMatchRequest) => ApproximateVectorMatchResult
  readonly primeExactColumnIndex: (request: { sheetName: string; rowStart: number; rowEnd: number; col: number }) => void
  readonly primeApproximateColumnIndex: (request: { sheetName: string; rowStart: number; rowEnd: number; col: number }) => void
  readonly prepareExactVectorLookup: (request: {
    sheetName: string
    rowStart: number
    rowEnd: number
    col: number
  }) => PreparedExactVectorLookup
  readonly findPreparedExactVectorMatch: (request: {
    lookupValue: CellValue
    prepared: PreparedExactVectorLookup
    searchMode: 1 | -1
  }) => ExactVectorMatchResult
  readonly prepareApproximateVectorLookup: (request: {
    sheetName: string
    rowStart: number
    rowEnd: number
    col: number
  }) => PreparedApproximateVectorLookup
  readonly findPreparedApproximateVectorMatch: (request: {
    lookupValue: CellValue
    prepared: PreparedApproximateVectorLookup
    matchMode: 1 | -1
  }) => ApproximateVectorMatchResult
}

export function createEngineLookupService(args: {
  readonly exact: ExactColumnIndexService
  readonly sorted: SortedColumnSearchService
}): EngineLookupService {
  return {
    findExactVectorMatch(request) {
      return args.exact.findVectorMatch(request)
    },
    findApproximateVectorMatch(request) {
      return args.sorted.findVectorMatch(request)
    },
    primeExactColumnIndex(request) {
      args.exact.primeColumnIndex(request)
    },
    primeApproximateColumnIndex(request) {
      args.sorted.primeColumnIndex(request)
    },
    prepareExactVectorLookup(request) {
      return args.exact.prepareVectorLookup(request)
    },
    findPreparedExactVectorMatch(request) {
      return args.exact.findPreparedVectorMatch(request)
    },
    prepareApproximateVectorLookup(request) {
      return args.sorted.prepareVectorLookup(request)
    },
    findPreparedApproximateVectorMatch(request) {
      return args.sorted.findPreparedVectorMatch(request)
    },
  }
}
