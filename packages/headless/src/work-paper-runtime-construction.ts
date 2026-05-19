import { SpreadsheetEngine } from '@bilig/core/headless-runtime'
import { WorkPaperEvaluationTimeoutError } from './work-paper-errors.js'
import { normalizeConfiguredWorkPaperCalculationSettings } from './work-paper-config.js'
import type { WorkPaperClipboardPayload } from './work-paper-clipboard.js'
import type { WorkPaperHistoryRecord } from './work-paper-history.js'
import type { SerializedWorkPaperNamedExpression, WorkPaperConfig, WorkPaperSheets } from './work-paper-types.js'

export interface WorkPaperTransactionSnapshot {
  readonly clipboard: WorkPaperClipboardPayload | null
  readonly config: WorkPaperConfig
  readonly namedExpressions: readonly SerializedWorkPaperNamedExpression[]
  readonly redoStack: readonly WorkPaperHistoryRecord[]
  readonly sheets: WorkPaperSheets
  readonly undoStack: readonly WorkPaperHistoryRecord[]
}

export function workPaperEvaluationTimeoutErrorFrom(error: unknown): WorkPaperEvaluationTimeoutError | undefined {
  let current: unknown = error
  while (typeof current === 'object' && current !== null) {
    if (current instanceof WorkPaperEvaluationTimeoutError) {
      return current
    }
    const name = current instanceof Error ? current.name : undefined
    if (name === 'WorkPaperEvaluationTimeoutError' || name === 'EngineEvaluationTimeoutError') {
      const timeoutMs = Reflect.get(current, 'timeoutMs')
      return new WorkPaperEvaluationTimeoutError(typeof timeoutMs === 'number' ? timeoutMs : 0)
    }
    current = Reflect.get(current, 'cause')
  }
  return undefined
}

export function createWorkPaperEngine(config: WorkPaperConfig): SpreadsheetEngine {
  const engine = new SpreadsheetEngine({
    workbookName: 'Workbook',
    trackReplicaVersions: false,
    ...(config.useColumnIndex !== undefined ? { useColumnIndex: config.useColumnIndex } : {}),
    ...(config.evaluationTimeoutMs !== undefined ? { evaluationTimeoutMs: config.evaluationTimeoutMs } : {}),
  })
  const calculationSettings = normalizeConfiguredWorkPaperCalculationSettings(config.calculationSettings)
  if (calculationSettings !== undefined) {
    engine.setCalculationSettings(calculationSettings)
  }
  return engine
}
