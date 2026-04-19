import type { RegionId } from './region-node-store.js'

export interface DepPatternEntry {
  readonly rows: Uint32Array
  readonly length: number
  readonly versionStamp: string
}

export interface DepPatternStore {
  readonly getCriteriaPattern: (request: {
    readonly regionIds: readonly RegionId[]
    readonly criteriaKeys: readonly string[]
    readonly versionStamp: string
  }) => DepPatternEntry | undefined
  readonly setCriteriaPattern: (request: {
    readonly regionIds: readonly RegionId[]
    readonly criteriaKeys: readonly string[]
    readonly versionStamp: string
    readonly rows: Uint32Array
    readonly length: number
  }) => DepPatternEntry
}

function patternKey(regionIds: readonly RegionId[], criteriaKeys: readonly string[]): string {
  return `${regionIds.join(',')}\u0001${criteriaKeys.join('\u0001')}`
}

export function createDepPatternStore(): DepPatternStore {
  const criteriaPatterns = new Map<string, DepPatternEntry>()

  return {
    getCriteriaPattern({ regionIds, criteriaKeys, versionStamp }) {
      const existing = criteriaPatterns.get(patternKey(regionIds, criteriaKeys))
      return existing?.versionStamp === versionStamp ? existing : undefined
    },
    setCriteriaPattern({ regionIds, criteriaKeys, versionStamp, rows, length }) {
      const entry: DepPatternEntry = {
        rows,
        length,
        versionStamp,
      }
      criteriaPatterns.set(patternKey(regionIds, criteriaKeys), entry)
      return entry
    },
  }
}
