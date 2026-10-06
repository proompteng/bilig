export interface SelectionAggregateSummary {
  readonly nonEmptyCount: number
  readonly numericCount: number
  readonly sum: number
  readonly min: number | null
  readonly max: number | null
}

export function isSelectionAggregateSummary(value: unknown): value is SelectionAggregateSummary {
  if (typeof value !== 'object' || value === null) return false
  if (!('nonEmptyCount' in value) || !('numericCount' in value) || !('sum' in value) || !('min' in value) || !('max' in value)) return false
  return (
    typeof value.nonEmptyCount === 'number' &&
    Number.isSafeInteger(value.nonEmptyCount) &&
    value.nonEmptyCount >= 0 &&
    typeof value.numericCount === 'number' &&
    Number.isSafeInteger(value.numericCount) &&
    value.numericCount >= 0 &&
    value.numericCount <= value.nonEmptyCount &&
    typeof value.sum === 'number' &&
    !Number.isNaN(value.sum) &&
    (value.numericCount === 0
      ? value.min === null && value.max === null
      : typeof value.min === 'number' &&
        Number.isFinite(value.min) &&
        typeof value.max === 'number' &&
        Number.isFinite(value.max) &&
        value.min <= value.max)
  )
}
