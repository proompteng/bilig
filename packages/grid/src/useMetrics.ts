import { useSyncExternalStore } from 'react'
import { selectors, type SpreadsheetEngine } from '@bilig/core'

export function useMetrics(engine: SpreadsheetEngine) {
  useSyncExternalStore(
    engine.subscribe.bind(engine),
    () => selectors.selectMetrics(engine).batchId,
    () => selectors.selectMetrics(engine).batchId,
  )

  return { ...selectors.selectMetrics(engine) }
}
