import { ValueTag, type CellValue } from '@bilig/protocol'
import type { RangeBuiltinArgument } from './lookup-core-helpers.js'

export type LookupReferenceMatchMode = -1 | 0 | 1 | 2
export type LookupReferenceSearchMode = -2 | -1 | 1 | 2

export interface LookupReferenceSearchDeps {
  compareScalars: (left: CellValue, right: CellValue) => number | undefined
  getRangeValue: (range: RangeBuiltinArgument, row: number, col: number) => CellValue
}

export function vectorLength(range: RangeBuiltinArgument): number | undefined {
  if (range.rows !== 1 && range.cols !== 1) {
    return undefined
  }
  return range.rows === 1 ? range.cols : range.rows
}

export function getVectorValue(range: RangeBuiltinArgument, index: number, deps: LookupReferenceSearchDeps): CellValue {
  return range.rows === 1 ? deps.getRangeValue(range, 0, index) : deps.getRangeValue(range, index, 0)
}

export function hasLookupWildcardSyntax(value: CellValue): boolean {
  return value.tag === ValueTag.String && /[*?~]/.test(value.value)
}

export function exactMatch(
  lookupValue: CellValue,
  range: RangeBuiltinArgument,
  deps: LookupReferenceSearchDeps,
  options: { wildcard?: boolean; searchMode?: 1 | -1 } = {},
): number {
  const length = vectorLength(range)
  if (length === undefined) {
    return -1
  }
  const first = options.searchMode === -1 ? length - 1 : 0
  const last = options.searchMode === -1 ? -1 : length
  const step = options.searchMode === -1 ? -1 : 1
  for (let index = first; index !== last; index += step) {
    const candidate = getVectorValue(range, index, deps)
    if (options.wildcard ? wildcardMatches(lookupValue, candidate) : deps.compareScalars(candidate, lookupValue) === 0) {
      return index + 1
    }
  }
  return -1
}

export function approximateMatchAscending(lookupValue: CellValue, range: RangeBuiltinArgument, deps: LookupReferenceSearchDeps): number {
  const length = vectorLength(range)
  if (length === undefined) {
    return -1
  }
  let best = -1
  for (let index = 0; index < length; index += 1) {
    const comparison = deps.compareScalars(getVectorValue(range, index, deps), lookupValue)
    if (comparison === undefined) {
      return -1
    }
    if (comparison <= 0) {
      best = index + 1
    } else {
      break
    }
  }
  return best
}

export function approximateLookupAscending(
  lookupValue: CellValue,
  range: RangeBuiltinArgument,
  deps: LookupReferenceSearchDeps & {
    isError: (value: CellValue) => boolean
  },
): number {
  const length = vectorLength(range)
  if (length === undefined) {
    return -1
  }
  let best = -1
  for (let index = 0; index < length; index += 1) {
    const value = getVectorValue(range, index, deps)
    if (deps.isError(value)) {
      continue
    }
    const comparison = deps.compareScalars(value, lookupValue)
    if (comparison === undefined) {
      return -1
    }
    if (comparison <= 0) {
      best = index + 1
    } else {
      break
    }
  }
  return best
}

export function approximateMatchDescending(lookupValue: CellValue, range: RangeBuiltinArgument, deps: LookupReferenceSearchDeps): number {
  const length = vectorLength(range)
  if (length === undefined) {
    return -1
  }
  let best = -1
  for (let index = 0; index < length; index += 1) {
    const comparison = deps.compareScalars(getVectorValue(range, index, deps), lookupValue)
    if (comparison === undefined) {
      return -1
    }
    if (comparison >= 0) {
      best = index + 1
      continue
    }
    break
  }
  return best
}

export function findReferenceMatchIndex(
  lookupValue: CellValue,
  lookupRange: RangeBuiltinArgument,
  matchMode: LookupReferenceMatchMode,
  searchMode: LookupReferenceSearchMode,
  deps: LookupReferenceSearchDeps,
): number {
  const length = vectorLength(lookupRange)
  if (length === undefined) {
    return -1
  }

  if (searchMode === 2 || searchMode === -2) {
    if (matchMode === 2) {
      return -1
    }
    return binaryMatchIndex(lookupValue, lookupRange, matchMode, searchMode === 2, deps)
  }

  if (matchMode === 0 || matchMode === 2) {
    const position = exactMatch(lookupValue, lookupRange, deps, { wildcard: matchMode === 2, searchMode })
    return position < 0 ? -1 : position - 1
  }

  const exactPosition = exactMatch(lookupValue, lookupRange, deps, { searchMode })
  if (exactPosition >= 0) {
    return exactPosition - 1
  }
  return linearNearestMatchIndex(lookupValue, lookupRange, matchMode, searchMode, deps)
}

function linearNearestMatchIndex(
  lookupValue: CellValue,
  lookupRange: RangeBuiltinArgument,
  matchMode: -1 | 1,
  searchMode: -1 | 1,
  deps: LookupReferenceSearchDeps,
): number {
  const length = vectorLength(lookupRange)
  if (length === undefined) {
    return -1
  }

  const first = searchMode === -1 ? length - 1 : 0
  const last = searchMode === -1 ? -1 : length
  const step = searchMode === -1 ? -1 : 1
  let bestIndex = -1
  let bestValue: CellValue | undefined
  for (let index = first; index !== last; index += step) {
    const candidate = getVectorValue(lookupRange, index, deps)
    const comparison = deps.compareScalars(candidate, lookupValue)
    if (comparison === undefined) {
      continue
    }
    const qualifies = matchMode === -1 ? comparison < 0 : comparison > 0
    if (!qualifies) {
      continue
    }
    if (bestValue === undefined) {
      bestIndex = index
      bestValue = candidate
      continue
    }
    const bestComparison = deps.compareScalars(candidate, bestValue)
    if (bestComparison === undefined) {
      continue
    }
    if ((matchMode === -1 && bestComparison > 0) || (matchMode === 1 && bestComparison < 0)) {
      bestIndex = index
      bestValue = candidate
    }
  }
  return bestIndex
}

function binaryMatchIndex(
  lookupValue: CellValue,
  lookupRange: RangeBuiltinArgument,
  matchMode: Exclude<LookupReferenceMatchMode, 2>,
  isAscending: boolean,
  deps: LookupReferenceSearchDeps,
): number {
  let low = 0
  let high = (vectorLength(lookupRange) ?? 0) - 1
  let candidate = -1

  while (low <= high) {
    const mid = Math.floor((low + high) / 2)
    const comparison = deps.compareScalars(getVectorValue(lookupRange, mid, deps), lookupValue)
    if (comparison === undefined) {
      return -1
    }
    if (comparison === 0) {
      return mid
    }

    if (isAscending) {
      if (comparison < 0) {
        if (matchMode === -1) {
          candidate = mid
        }
        low = mid + 1
      } else {
        if (matchMode === 1) {
          candidate = mid
        }
        high = mid - 1
      }
      continue
    }

    if (comparison > 0) {
      if (matchMode === 1) {
        candidate = mid
      }
      low = mid + 1
    } else {
      if (matchMode === -1) {
        candidate = mid
      }
      high = mid - 1
    }
  }

  return matchMode === 0 ? -1 : candidate
}

function wildcardMatches(patternValue: CellValue, candidateValue: CellValue | undefined): boolean {
  if (patternValue.tag !== ValueTag.String || candidateValue === undefined) {
    return false
  }
  const candidate = wildcardText(candidateValue)
  return candidate === undefined ? false : wildcardPatternToRegExp(patternValue.value).test(candidate)
}

function wildcardText(value: CellValue): string | undefined {
  if (value.tag === ValueTag.String) {
    return value.value
  }
  if (value.tag === ValueTag.Empty) {
    return ''
  }
  return undefined
}

function wildcardPatternToRegExp(pattern: string): RegExp {
  let source = '^'
  for (let index = 0; index < pattern.length; index += 1) {
    const char = pattern[index]
    if (char === undefined) {
      continue
    }
    if (char === '~') {
      const escaped = pattern[index + 1]
      if (escaped !== undefined) {
        source += escapeRegexFragment(escaped)
        index += 1
        continue
      }
      source += escapeRegexFragment(char)
      continue
    }
    if (char === '*') {
      source += '.*'
      continue
    }
    if (char === '?') {
      source += '.'
      continue
    }
    source += escapeRegexFragment(char)
  }
  return new RegExp(`${source}$`, 'i')
}

function escapeRegexFragment(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
}
