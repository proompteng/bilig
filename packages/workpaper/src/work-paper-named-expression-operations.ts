import type { CellValue } from '@bilig/protocol'
import { isFormulaContent } from './work-paper-runtime-helpers.js'
import {
  createSerializedWorkPaperNamedExpression,
  listWorkPaperNamedExpressions,
  publicWorkPaperNamedExpressionFromInternal,
  serializeWorkPaperNamedExpressions,
  type InternalNamedExpression,
} from './work-paper-named-expression-helpers.js'
import { WorkPaperNamedExpressionDoesNotExistError, WorkPaperNamedExpressionNameIsAlreadyTakenError } from './work-paper-errors.js'
import type {
  RawCellContent,
  SerializedWorkPaperNamedExpression,
  WorkPaperChange,
  WorkPaperDetailedEventMap,
  WorkPaperNamedExpression,
} from './work-paper-types.js'

type NamedExpressionEvent =
  | {
      readonly eventName: 'namedExpressionAdded'
      readonly payload: WorkPaperDetailedEventMap['namedExpressionAdded']
    }
  | {
      readonly eventName: 'namedExpressionRemoved'
      readonly payload: WorkPaperDetailedEventMap['namedExpressionRemoved']
    }

export interface WorkPaperNamedExpressionOperationsRuntime {
  readonly assertReadable: () => void
  readonly getNamedExpression: (name: string, scope?: number) => InternalNamedExpression | undefined
  readonly getNamedExpressionValues: () => Iterable<InternalNamedExpression>
  readonly evaluateNamedExpression: (expression: InternalNamedExpression) => CellValue | CellValue[][]
  readonly isItPossibleToAddNamedExpression: (expressionName: string, expression: RawCellContent, scope?: number) => boolean
  readonly isItPossibleToRemoveNamedExpression: (expressionName: string, scope?: number) => boolean
  readonly validateNamedExpression: (expressionName: string, expression: RawCellContent, scope?: number) => void
  readonly tryCaptureNamedExpressionChangeWithoutSnapshots: (
    existing: InternalNamedExpression,
    expressionName: string,
    expression: RawCellContent,
    scope?: number,
    options?: Record<string, string | number | boolean>,
  ) => WorkPaperChange[] | null
  readonly captureChanges: (event: NamedExpressionEvent | undefined, mutate: () => void) => WorkPaperChange[]
  readonly upsertNamedExpressionInternal: (
    expression: SerializedWorkPaperNamedExpression,
    options: { readonly duringInitialization: boolean; readonly skipValidation?: boolean },
  ) => void
  readonly deleteNamedExpressionRecord: (name: string, scope?: number) => void
  readonly deleteDefinedName: (internalName: string) => void
}

export interface WorkPaperNamedExpressionOperations {
  readonly getNamedExpressionValue: (name: string, scope?: number) => CellValue | CellValue[][] | undefined
  readonly getNamedExpressionFormula: (name: string, scope?: number) => string | undefined
  readonly getNamedExpression: (name: string, scope?: number) => WorkPaperNamedExpression | undefined
  readonly addNamedExpression: (
    expressionName: string,
    expression: RawCellContent,
    scope?: number,
    options?: Record<string, string | number | boolean>,
  ) => WorkPaperChange[]
  readonly changeNamedExpression: (
    expressionName: string,
    expression: RawCellContent,
    scope?: number,
    options?: Record<string, string | number | boolean>,
  ) => WorkPaperChange[]
  readonly removeNamedExpression: (expressionName: string, scope?: number) => WorkPaperChange[]
  readonly listNamedExpressions: (scope?: number) => string[]
  readonly getAllNamedExpressionsSerialized: () => SerializedWorkPaperNamedExpression[]
}

export function createWorkPaperNamedExpressionOperations(
  runtime: WorkPaperNamedExpressionOperationsRuntime,
): WorkPaperNamedExpressionOperations {
  return {
    getNamedExpressionValue(name, scope) {
      runtime.assertReadable()
      const expression = runtime.getNamedExpression(name, scope)
      return expression ? runtime.evaluateNamedExpression(expression) : undefined
    },

    getNamedExpressionFormula(name, scope) {
      const expression = runtime.getNamedExpression(name, scope)
      if (!expression) {
        return undefined
      }
      return isFormulaContent(expression.expression) ? expression.expression : undefined
    },

    getNamedExpression(name, scope) {
      const expression = runtime.getNamedExpression(name, scope)
      if (!expression) {
        return undefined
      }
      return publicWorkPaperNamedExpressionFromInternal(expression)
    },

    addNamedExpression(expressionName, expression, scope, options) {
      if (!runtime.isItPossibleToAddNamedExpression(expressionName, expression, scope)) {
        throw new WorkPaperNamedExpressionNameIsAlreadyTakenError(expressionName)
      }
      return runtime.captureChanges(
        {
          eventName: 'namedExpressionAdded',
          payload: {
            name: expressionName.trim(),
            changes: [],
            ...(scope !== undefined ? { scope } : {}),
          },
        },
        () => {
          runtime.upsertNamedExpressionInternal(
            createSerializedWorkPaperNamedExpression({ name: expressionName, expression, scope, options }),
            { duringInitialization: false },
          )
        },
      )
    },

    changeNamedExpression(expressionName, expression, scope, options) {
      runtime.validateNamedExpression(expressionName, expression, scope)
      const existing = runtime.getNamedExpression(expressionName, scope)
      if (!existing) {
        throw new WorkPaperNamedExpressionDoesNotExistError(expressionName)
      }
      const fastPathChanges = runtime.tryCaptureNamedExpressionChangeWithoutSnapshots(existing, expressionName, expression, scope, options)
      if (fastPathChanges) {
        return fastPathChanges
      }
      return runtime.captureChanges(undefined, () => {
        runtime.upsertNamedExpressionInternal(
          createSerializedWorkPaperNamedExpression({ name: expressionName, expression, scope, options }),
          { duringInitialization: false },
        )
      })
    },

    removeNamedExpression(expressionName, scope) {
      if (!runtime.isItPossibleToRemoveNamedExpression(expressionName, scope)) {
        throw new WorkPaperNamedExpressionDoesNotExistError(expressionName)
      }
      const existing = runtime.getNamedExpression(expressionName, scope)
      if (!existing) {
        throw new WorkPaperNamedExpressionDoesNotExistError(expressionName)
      }
      return runtime.captureChanges(
        {
          eventName: 'namedExpressionRemoved',
          payload: {
            name: existing.publicName,
            changes: [],
            ...(existing.scope !== undefined ? { scope: existing.scope } : {}),
          },
        },
        () => {
          runtime.deleteNamedExpressionRecord(expressionName, scope)
          runtime.deleteDefinedName(existing.internalName)
        },
      )
    },

    listNamedExpressions(scope) {
      return listWorkPaperNamedExpressions(runtime.getNamedExpressionValues(), scope)
    },

    getAllNamedExpressionsSerialized() {
      return serializeWorkPaperNamedExpressions(runtime.getNamedExpressionValues())
    },
  }
}
