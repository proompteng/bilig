import type { CellValue } from '@bilig/protocol'
import type { EngineRuntimeState } from '../runtime-state.js'

const DIRECT_CRITERIA_AGGREGATE_CACHE_LIMIT = 16_384

export function directCriteriaRangeVersionKey(
  state: Pick<EngineRuntimeState, 'workbook'>,
  range: { sheetName: string; rowStart: number; rowEnd: number; col: number },
): string {
  const sheet = state.workbook.getSheet(range.sheetName)
  return `${range.sheetName}:${range.rowStart}:${range.rowEnd}:${range.col}:${sheet?.columnVersions[range.col] ?? 0}:${
    sheet?.structureVersion ?? 0
  }`
}

export function rememberDirectCriteriaResult(cache: Map<string, CellValue>, key: string, value: CellValue): CellValue {
  if (cache.size >= DIRECT_CRITERIA_AGGREGATE_CACHE_LIMIT) {
    const firstKey = cache.keys().next().value
    if (firstKey !== undefined) {
      cache.delete(firstKey)
    }
  }
  cache.set(key, value)
  return value
}
