import { SpreadsheetEngine } from '@bilig/core/headless-runtime'
import { WorkPaperEvaluationTimeoutError } from './work-paper-errors.js'
import { normalizeConfiguredWorkPaperCalculationSettings } from './work-paper-config.js'
import type { WorkPaperClipboardPayload } from './work-paper-clipboard.js'
import type { WorkPaperHistoryRecord } from './work-paper-history.js'
import type { SerializedWorkPaperNamedExpression, WorkPaperConfig, WorkPaperSheets } from './work-paper-types.js'

const MAX_POOLED_WORKPAPER_ENGINES = 8
const MAX_POOLED_WORKPAPER_ENGINE_CELL_CAPACITY = 16_384
const pooledWorkPaperEngines: SpreadsheetEngine[] = []

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

export function createWorkPaperEngine(config: WorkPaperConfig, options: { readonly fresh?: boolean } = {}): SpreadsheetEngine {
  const engine =
    options.fresh === true
      ? new SpreadsheetEngine({ workbookName: 'Workbook', trackReplicaVersions: false })
      : (pooledWorkPaperEngines.pop() ?? new SpreadsheetEngine({ workbookName: 'Workbook', trackReplicaVersions: false }))
  engine.setUseColumnIndexEnabled(config.useColumnIndex ?? true)
  engine.setEvaluationTimeoutMs(config.evaluationTimeoutMs)
  engine.resetPerformanceCounters()
  const calculationSettings = normalizeConfiguredWorkPaperCalculationSettings(config.calculationSettings)
  if (calculationSettings !== undefined) {
    engine.setCalculationSettings(calculationSettings)
  }
  return engine
}

export function releaseWorkPaperEngine(engine: SpreadsheetEngine): void {
  if (
    engine.workbook.cellStore.capacity > MAX_POOLED_WORKPAPER_ENGINE_CELL_CAPACITY ||
    pooledWorkPaperEngines.length >= MAX_POOLED_WORKPAPER_ENGINES
  ) {
    engine.workbook.releaseReusableBuffers()
    engine.resetForReuse({ workbookName: 'Workbook', trackReplicaVersions: false })
    return
  }
  engine.resetForReuse({ workbookName: 'Workbook', trackReplicaVersions: false })
  pooledWorkPaperEngines.push(engine)
}
