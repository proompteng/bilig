import { ValueTag, type CellValue } from '@bilig/protocol'
import type { EngineRuntimeState, RuntimeDirectCriteriaOperand, RuntimeFormula } from '../runtime-state.js'
import type { CriterionRangePair } from './criterion-range-cache-service.js'
import { readRuntimeDirectCriteriaOperandValue } from './direct-criteria-operands.js'
import { directCriteriaRangeVersionKey } from './formula-evaluation-direct-criteria-cache.js'
import { cellValueCriteriaString, directCriteriaCacheValueKey } from './formula-evaluation-helpers.js'

const DIRECT_CRITERIA_SHARE_COUNT_CACHE_LIMIT = 4096

export interface ResolvedDirectCriteriaPairs {
  readonly pairs: CriterionRangePair[]
  readonly error: CellValue | undefined
}

export function createDirectCriteriaSharingContext(args: {
  readonly state: Pick<EngineRuntimeState, 'workbook' | 'formulas'>
  readonly readCellValueByIndex: (cellIndex: number | undefined) => CellValue
}): {
  readonly directCriteriaShareCount: (criteriaVersionKey: string) => number
  readonly directCriteriaVersionKeyForPairs: (pairs: readonly CriterionRangePair[]) => string
  readonly resolveDirectCriteriaPairs: (formula: RuntimeFormula) => ResolvedDirectCriteriaPairs | undefined
} {
  const directCriteriaShareCountCache = new Map<string, number>()
  const readDirectCriteriaOperandValue = (operand: RuntimeDirectCriteriaOperand): CellValue => {
    return readRuntimeDirectCriteriaOperandValue({
      operand,
      readCellValueByIndex: args.readCellValueByIndex,
      stringifyCriteriaValue: cellValueCriteriaString,
    })
  }
  const directCriteriaVersionKeyForPairs = (pairs: readonly CriterionRangePair[]): string =>
    pairs.map((pair) => `${directCriteriaRangeVersionKey(args.state, pair.range)}:${directCriteriaCacheValueKey(pair.criteria)}`).join('|')
  const resolveDirectCriteriaPairs = (formula: RuntimeFormula): ResolvedDirectCriteriaPairs | undefined => {
    const directCriteria = formula.directCriteria
    if (!directCriteria) {
      return undefined
    }
    const pairs = directCriteria.criteriaPairs.map((pair) => ({
      range: pair.range,
      criteria: readDirectCriteriaOperandValue(pair.criterion),
    }))
    return { pairs, error: pairs.find((pair) => pair.criteria.tag === ValueTag.Error)?.criteria }
  }
  const directCriteriaShareCount = (criteriaVersionKey: string): number => {
    const cacheKey = `${args.state.formulas.version}\u0000${criteriaVersionKey}`
    const cached = directCriteriaShareCountCache.get(cacheKey)
    if (cached !== undefined) {
      return cached
    }
    let count = 0
    for (const candidate of args.state.formulas.values()) {
      const candidateCriteria = resolveDirectCriteriaPairs(candidate)
      if (candidateCriteria === undefined || candidateCriteria.error !== undefined) {
        continue
      }
      if (directCriteriaVersionKeyForPairs(candidateCriteria.pairs) === criteriaVersionKey) {
        count += 1
        if (count > 1) {
          break
        }
      }
    }
    if (directCriteriaShareCountCache.size >= DIRECT_CRITERIA_SHARE_COUNT_CACHE_LIMIT) {
      directCriteriaShareCountCache.clear()
    }
    directCriteriaShareCountCache.set(cacheKey, count)
    return count
  }
  return { directCriteriaShareCount, directCriteriaVersionKeyForPairs, resolveDirectCriteriaPairs }
}
