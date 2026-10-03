import { type CellValue, ValueTag, type WorkbookDefinedNameValueSnapshot } from '@bilig/protocol'
import { WORKPAPER_PUBLIC_ERROR_NAMES } from './work-paper-config.js'
import { WorkPaperOperationError } from './work-paper-errors.js'
import {
  cloneNamedExpressionValue,
  createInternalNamedExpressionRecord,
  createSerializedWorkPaperNamedExpression,
  createWorkPaperNamedExpressionChange,
  type InternalNamedExpression,
} from './work-paper-named-expression-helpers.js'
import { makeNamedExpressionKey, matrixValuesEqual, tryEvaluateSimpleNamedExpression, valuesEqual } from './work-paper-runtime-helpers.js'
import type { TrackedEngineEvent } from './tracked-engine-event-refs.js'
import type { RawCellContent, SerializedWorkPaperNamedExpression, WorkPaperCellChange, WorkPaperChange } from './work-paper-types.js'

type NamedExpressionValue = CellValue | CellValue[][]

export interface WorkPaperNamedExpressionFastPathRuntime {
  readonly canUseNamedExpressionChangeFastPath: () => boolean
  readonly assertNotDisposed: () => void
  readonly materializePendingLazyTrackedChanges: () => void
  readonly downgradeTrackedBatchFastPath: () => void
  readonly getCachedNamedExpressionValue: (key: string) => NamedExpressionValue | undefined
  readonly setCachedNamedExpressionValue: (key: string, value: NamedExpressionValue) => void
  readonly evaluateNamedExpression: (expression: InternalNamedExpression) => NamedExpressionValue
  readonly hasAnyListeners: () => boolean
  readonly toDefinedNameSnapshot: (expression: RawCellContent, scope: number | undefined) => WorkbookDefinedNameValueSnapshot
  readonly upsertNumericDefinedNameFast: (
    name: string,
    value: WorkbookDefinedNameValueSnapshot,
    numericValue: number,
  ) => readonly number[] | null
  readonly setNamedExpressionRecord: (key: string, record: InternalNamedExpression) => void
  readonly readSingleTrackedCellChange: (cellIndex: number) => WorkPaperCellChange | undefined
  readonly orderCellChanges: (changes: WorkPaperCellChange[], explicitChangedCount: number | undefined) => WorkPaperChange[]
  readonly drainTrackedEngineEvents: () => readonly TrackedEngineEvent[]
  readonly withRetainedTrackedEngineEventIndices: (callback: () => void) => void
  readonly upsertNamedExpressionInternal: (expression: SerializedWorkPaperNamedExpression) => void
  readonly namedExpressionRecord: (name: string, scope: number | undefined) => InternalNamedExpression
  readonly computeTrackedChangesWithoutVisibilityCache: (
    events: readonly TrackedEngineEvent[],
    options: { readonly preferLazyPublicChanges?: boolean },
  ) => WorkPaperChange[]
  readonly messageOf: (error: unknown, fallback: string) => string
}

export interface WorkPaperNamedExpressionFastPathRequest {
  readonly existing: InternalNamedExpression
  readonly expressionName: string
  readonly expression: RawCellContent
  readonly scope?: number
  readonly options?: Record<string, string | number | boolean>
}

export function tryCaptureWorkPaperNamedExpressionChangeWithoutSnapshots(
  runtime: WorkPaperNamedExpressionFastPathRuntime,
  request: WorkPaperNamedExpressionFastPathRequest,
): WorkPaperChange[] | null {
  if (!runtime.canUseNamedExpressionChangeFastPath()) {
    return null
  }
  runtime.assertNotDisposed()
  runtime.materializePendingLazyTrackedChanges()
  runtime.downgradeTrackedBatchFastPath()
  if (!runtime.canUseNamedExpressionChangeFastPath()) {
    return null
  }

  const key = makeNamedExpressionKey(request.existing.publicName, request.existing.scope)
  const cachedBeforeValue = runtime.getCachedNamedExpressionValue(key)
  const beforeValue =
    cachedBeforeValue ?? tryEvaluateSimpleNamedExpression(request.existing.expression) ?? runtime.evaluateNamedExpression(request.existing)
  const afterScalarValue = tryEvaluateSimpleNamedExpression(request.expression)
  if (afterScalarValue?.tag === ValueTag.Number && !runtime.hasAnyListeners()) {
    const directChanges = tryCaptureNumericNamedExpressionChange(runtime, request, key, beforeValue, afterScalarValue)
    if (directChanges) {
      return directChanges
    }
  }
  runtime.drainTrackedEngineEvents()

  try {
    runtime.withRetainedTrackedEngineEventIndices(() => {
      runtime.upsertNamedExpressionInternal({
        ...createSerializedWorkPaperNamedExpression({
          name: request.expressionName,
          expression: request.expression,
          scope: request.scope,
          options: request.options,
        }),
      })
    })
  } catch (error) {
    throw publicOrOperationError(runtime, error)
  }

  const updated = runtime.namedExpressionRecord(request.expressionName, request.scope)
  const afterValue = cloneNamedExpressionValue(runtime.evaluateNamedExpression(updated))
  runtime.setCachedNamedExpressionValue(key, cloneNamedExpressionValue(afterValue))
  const cellChanges = runtime.computeTrackedChangesWithoutVisibilityCache(runtime.drainTrackedEngineEvents(), {
    preferLazyPublicChanges: !runtime.hasAnyListeners(),
  })
  if (matrixValuesEqual(beforeValue, afterValue)) {
    return cellChanges
  }
  return [
    ...cellChanges,
    createWorkPaperNamedExpressionChange({
      name: updated.publicName,
      scope: updated.scope,
      newValue: cloneNamedExpressionValue(afterValue),
    }),
  ]
}

function tryCaptureNumericNamedExpressionChange(
  runtime: WorkPaperNamedExpressionFastPathRuntime,
  request: WorkPaperNamedExpressionFastPathRequest,
  cacheKey: string,
  beforeValue: NamedExpressionValue,
  afterValue: CellValue & { readonly tag: ValueTag.Number },
): WorkPaperChange[] | null {
  const record = createInternalNamedExpressionRecord(
    createSerializedWorkPaperNamedExpression({
      name: request.expressionName,
      expression: request.expression,
      scope: request.scope,
      options: request.options,
    }),
  )
  const definedNameSnapshot = runtime.toDefinedNameSnapshot(request.expression, request.scope)
  try {
    const changedCellIndices = runtime.upsertNumericDefinedNameFast(record.internalName, definedNameSnapshot, afterValue.value)
    if (changedCellIndices) {
      runtime.setNamedExpressionRecord(makeNamedExpressionKey(record.publicName, record.scope), record)
      runtime.setCachedNamedExpressionValue(cacheKey, cloneNamedExpressionValue(afterValue))
      const cellChanges = changedCellIndices
        .map((cellIndex) => runtime.readSingleTrackedCellChange(cellIndex))
        .filter((change): change is WorkPaperCellChange => change !== undefined)
      const orderedCellChanges = cellChanges.length > 1 ? runtime.orderCellChanges(cellChanges, cellChanges.length) : cellChanges
      return !Array.isArray(beforeValue) && valuesEqual(beforeValue, afterValue)
        ? orderedCellChanges
        : [
            ...orderedCellChanges,
            createWorkPaperNamedExpressionChange({
              name: record.publicName,
              scope: record.scope,
              newValue: cloneNamedExpressionValue(afterValue),
            }),
          ]
    }
  } catch (error) {
    throw publicOrOperationError(runtime, error)
  }
  return null
}

function publicOrOperationError(runtime: WorkPaperNamedExpressionFastPathRuntime, error: unknown): Error {
  if (error instanceof Error && WORKPAPER_PUBLIC_ERROR_NAMES.has(error.name)) {
    return error
  }
  return new WorkPaperOperationError(runtime.messageOf(error, 'Mutation failed'))
}
