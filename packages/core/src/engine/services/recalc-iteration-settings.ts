import type { WorkbookCalculationSettingsSnapshot } from '@bilig/protocol'

export const DEFAULT_ITERATION_COUNT = 100
export const DEFAULT_ITERATION_DELTA = 0.001

export interface RecalcIterationSettings {
  readonly enabled: boolean
  readonly count: number
  readonly delta: number
}

function parsePositiveFiniteDecimal(value: string | null | undefined): number | null {
  if (typeof value !== 'string' || value.trim() !== value || value.length === 0) {
    return null
  }
  const parsed = Number(value)
  return Number.isFinite(parsed) && parsed > 0 ? parsed : null
}

export function resolveRecalcIterationSettings(settings: WorkbookCalculationSettingsSnapshot): RecalcIterationSettings {
  const count =
    typeof settings.iterateCount === 'number' && Number.isSafeInteger(settings.iterateCount) && settings.iterateCount > 0
      ? settings.iterateCount
      : DEFAULT_ITERATION_COUNT
  const parsedDelta = parsePositiveFiniteDecimal(settings.iterateDelta)
  return {
    enabled: settings.iterate === true,
    count,
    delta: parsedDelta === null ? DEFAULT_ITERATION_DELTA : Math.abs(parsedDelta),
  }
}
