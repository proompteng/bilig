import type { SpreadsheetEngine, SheetRecord } from '@bilig/core/headless-runtime'
import { ValueTag, type WorkbookDefinedNameValueSnapshot } from '@bilig/protocol'
import { orderWorkPaperCellChanges } from './change-order.js'
import { WorkPaperOperationError } from './work-paper-errors.js'
import {
  cloneNamedExpressionValue,
  createInternalNamedExpressionRecord,
  createSerializedWorkPaperNamedExpression,
  createWorkPaperNamedExpressionChange,
  type InternalNamedExpression,
  type WorkPaperNamedExpressionValueSnapshot,
} from './work-paper-named-expression-helpers.js'
import { makeNamedExpressionKey, tryEvaluateSimpleNamedExpression, valuesEqual } from './work-paper-runtime-helpers.js'
import type { RawCellContent, WorkPaperCellChange, WorkPaperChange } from './work-paper-types.js'

export function tryChangeSimpleNumericNamedExpressionFastPath(args: {
  readonly assertNotDisposed: () => void
  readonly canUseNamedExpressionChangeFastPath: () => boolean
  readonly downgradeTrackedBatchFastPath: () => void
  readonly engine: SpreadsheetEngine
  readonly listSheetRecords: () => readonly SheetRecord[]
  readonly materializePendingLazyChanges: () => void
  readonly messageOf: (error: unknown, fallback: string) => string
  readonly namedExpressionValueCache: WorkPaperNamedExpressionValueSnapshot | null
  readonly namedExpressions: Map<string, InternalNamedExpression>
  readonly publicErrorNames: ReadonlySet<string>
  readonly readSingleTrackedCellChange: (cellIndex: number) => WorkPaperCellChange | undefined
  readonly toDefinedNameSnapshot: (expression: RawCellContent, scope?: number) => WorkbookDefinedNameValueSnapshot
  readonly validateNamedExpression: (expressionName: string, expression: RawCellContent, scope?: number) => void
  readonly expressionName: string
  readonly expression: RawCellContent
  readonly scope: number | undefined
}): WorkPaperChange[] | null {
  if (!args.canUseNamedExpressionChangeFastPath()) {
    return null
  }
  args.validateNamedExpression(args.expressionName, args.expression, args.scope)
  const existing = args.namedExpressions.get(makeNamedExpressionKey(args.expressionName, args.scope))
  if (!existing) {
    return null
  }
  const beforeValue = tryEvaluateSimpleNamedExpression(existing.expression)
  const afterValue = tryEvaluateSimpleNamedExpression(args.expression)
  if (beforeValue === undefined || afterValue?.tag !== ValueTag.Number) {
    return null
  }

  args.assertNotDisposed()
  args.materializePendingLazyChanges()
  args.downgradeTrackedBatchFastPath()
  if (!args.canUseNamedExpressionChangeFastPath()) {
    return null
  }

  const record = createInternalNamedExpressionRecord(
    createSerializedWorkPaperNamedExpression({
      name: args.expressionName,
      expression: args.expression,
      scope: args.scope,
      options: undefined,
    }),
  )
  const key = makeNamedExpressionKey(record.publicName, record.scope)
  try {
    const changedCellIndices = args.engine.upsertNumericDefinedNameFast(
      record.internalName,
      args.toDefinedNameSnapshot(record.expression, record.scope),
      afterValue.value,
    )
    if (changedCellIndices === null) {
      return null
    }
    args.namedExpressions.set(key, record)
    args.namedExpressionValueCache?.set(key, cloneNamedExpressionValue(afterValue))
    const cellChanges = collectNamedExpressionCellChanges(changedCellIndices, args.readSingleTrackedCellChange)
    const orderedCellChanges =
      cellChanges.length > 1 ? orderWorkPaperCellChanges(cellChanges, args.listSheetRecords(), cellChanges.length) : cellChanges
    return valuesEqual(beforeValue, afterValue)
      ? orderedCellChanges
      : [
          ...orderedCellChanges,
          createWorkPaperNamedExpressionChange({
            name: record.publicName,
            scope: record.scope,
            newValue: cloneNamedExpressionValue(afterValue),
          }),
        ]
  } catch (error) {
    if (error instanceof Error && args.publicErrorNames.has(error.name)) {
      throw error
    }
    throw new WorkPaperOperationError(args.messageOf(error, 'Mutation failed'))
  }
}

function collectNamedExpressionCellChanges(
  changedCellIndices: readonly number[],
  readSingleTrackedCellChange: (cellIndex: number) => WorkPaperCellChange | undefined,
): WorkPaperCellChange[] {
  const cellChanges: WorkPaperCellChange[] = []
  for (let index = 0; index < changedCellIndices.length; index += 1) {
    const change = readSingleTrackedCellChange(changedCellIndices[index]!)
    if (change) {
      cellChanges.push(change)
    }
  }
  return cellChanges
}
