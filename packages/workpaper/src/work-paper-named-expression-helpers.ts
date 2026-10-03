import { isCellReferenceText, parseFormula } from '@bilig/formula'
import { ValueTag, type CellSnapshot, type CellValue, type WorkbookDefinedNameValueSnapshot } from '@bilig/protocol'
import {
  WorkPaperNamedExpressionNameIsInvalidError,
  WorkPaperNoRelativeAddressesAllowedError,
  WorkPaperUnableToParseError,
} from './work-paper-errors.js'
import {
  cloneCellValue,
  formulaHasRelativeReferences,
  isFormulaContent,
  makeInternalScopedName,
  makeNamedExpressionKey,
  matrixValuesEqual,
  normalizeName,
  scalarValueFromLiteral,
  stripLeadingEquals,
  tryEvaluateSimpleNamedExpression,
  tryEvaluateSimpleScalarFormulaBody,
  tryReadSimpleScalarFormulaBody,
} from './work-paper-runtime-helpers.js'
import type { RawCellContent, SerializedWorkPaperNamedExpression, WorkPaperChange, WorkPaperNamedExpression } from './work-paper-types.js'
import { compareSheetNames } from './work-paper-sheet-inspection.js'

export interface InternalNamedExpression {
  publicName: string
  normalizedName: string
  internalName: string
  scope?: number
  expression: RawCellContent
  options?: Record<string, string | number | boolean>
}

export type WorkPaperNamedExpressionValue = CellValue | CellValue[][]
export type WorkPaperNamedExpressionValueSnapshot = Map<string, WorkPaperNamedExpressionValue>

export function cloneNamedExpressionValue(value: CellValue | CellValue[][]): CellValue | CellValue[][] {
  if (!Array.isArray(value)) {
    return cloneCellValue(value)
  }
  return value.map((row) => row.map((cell) => cloneCellValue(cell)))
}

export function captureWorkPaperNamedExpressionValueSnapshot(
  expressions: Iterable<InternalNamedExpression>,
  evaluateExpression: (expression: InternalNamedExpression) => WorkPaperNamedExpressionValue,
): WorkPaperNamedExpressionValueSnapshot {
  const snapshot: WorkPaperNamedExpressionValueSnapshot = new Map()
  ;[...expressions].forEach((expression) => {
    snapshot.set(makeNamedExpressionKey(expression.publicName, expression.scope), cloneNamedExpressionValue(evaluateExpression(expression)))
  })
  return snapshot
}

export function compareWorkPaperNamedExpressionChanges(left: WorkPaperChange, right: WorkPaperChange): number {
  if (left.kind !== 'named-expression' || right.kind !== 'named-expression') {
    return 0
  }
  return (left.scope ?? -1) - (right.scope ?? -1) || left.name.localeCompare(right.name)
}

export function createWorkPaperNamedExpressionChange(args: {
  readonly name: string
  readonly scope: number | undefined
  readonly newValue: CellValue | CellValue[][]
}): WorkPaperChange {
  const change: WorkPaperChange = {
    kind: 'named-expression',
    name: args.name,
    newValue: args.newValue,
  }
  if (args.scope !== undefined) {
    change.scope = args.scope
  }
  return change
}

export function createSerializedWorkPaperNamedExpression(args: {
  readonly name: string
  readonly expression: RawCellContent
  readonly scope: number | undefined
  readonly options: Record<string, string | number | boolean> | undefined
}): SerializedWorkPaperNamedExpression {
  const expression: SerializedWorkPaperNamedExpression = {
    name: args.name,
    expression: args.expression,
  }
  if (args.scope !== undefined) {
    expression.scope = args.scope
  }
  if (args.options !== undefined) {
    expression.options = structuredClone(args.options)
  }
  return expression
}

export function publicWorkPaperNamedExpressionFromInternal(expression: InternalNamedExpression): WorkPaperNamedExpression {
  return createSerializedWorkPaperNamedExpression({
    name: expression.publicName,
    expression: expression.expression,
    scope: expression.scope,
    options: expression.options,
  })
}

export function computeWorkPaperNamedExpressionChanges(args: {
  readonly beforeNames: WorkPaperNamedExpressionValueSnapshot
  readonly afterNames: WorkPaperNamedExpressionValueSnapshot
  readonly expressionsByKey: ReadonlyMap<string, InternalNamedExpression>
}): WorkPaperChange[] {
  const namedExpressionChanges: WorkPaperChange[] = []
  args.afterNames.forEach((afterValue, key) => {
    const beforeValue = args.beforeNames.get(key)
    if (matrixValuesEqual(beforeValue, afterValue)) {
      return
    }
    const expression = args.expressionsByKey.get(key)
    if (!expression) {
      return
    }
    namedExpressionChanges.push(
      createWorkPaperNamedExpressionChange({
        name: expression.publicName,
        scope: expression.scope,
        newValue: cloneNamedExpressionValue(afterValue),
      }),
    )
  })
  return namedExpressionChanges.toSorted(compareWorkPaperNamedExpressionChanges)
}

export function validateWorkPaperNamedExpression(args: {
  readonly expressionName: string
  readonly expression: RawCellContent
  readonly scope?: number
  readonly requireScope: (scope: number) => void
  readonly messageOf: (error: unknown, fallback: string) => string
}): void {
  const trimmed = args.expressionName.trim()
  if (!/^[A-Za-z_][A-Za-z0-9_.]*$/.test(trimmed) || isCellReferenceText(trimmed)) {
    throw new WorkPaperNamedExpressionNameIsInvalidError(args.expressionName)
  }
  if (args.scope !== undefined) {
    args.requireScope(args.scope)
  }
  if (isFormulaContent(args.expression)) {
    const simpleBody = tryReadSimpleScalarFormulaBody(args.expression)
    if (simpleBody !== undefined && tryEvaluateSimpleScalarFormulaBody(simpleBody) !== undefined) {
      return
    }
    try {
      const parsed = parseFormula(stripLeadingEquals(args.expression))
      if (formulaHasRelativeReferences(parsed)) {
        throw new WorkPaperNoRelativeAddressesAllowedError()
      }
    } catch (error) {
      if (error instanceof WorkPaperNoRelativeAddressesAllowedError) {
        throw error
      }
      throw new WorkPaperUnableToParseError({
        expressionName: args.expressionName,
        reason: args.messageOf(error, `Invalid named expression formula for '${args.expressionName}'`),
      })
    }
  }
}

export function createInternalNamedExpressionRecord(expression: SerializedWorkPaperNamedExpression): InternalNamedExpression {
  const trimmed = expression.name.trim()
  const record: InternalNamedExpression = {
    publicName: trimmed,
    normalizedName: normalizeName(trimmed),
    internalName: expression.scope === undefined ? trimmed : makeInternalScopedName(expression.scope, trimmed),
    expression: expression.expression,
  }
  if (expression.scope !== undefined) {
    record.scope = expression.scope
  }
  if (expression.options !== undefined) {
    record.options = structuredClone(expression.options)
  }
  return record
}

export function listWorkPaperNamedExpressions(expressions: Iterable<InternalNamedExpression>, scope?: number): string[] {
  return [...expressions]
    .filter((expression) => expression.scope === scope)
    .map((expression) => expression.publicName)
    .toSorted(compareSheetNames)
}

export function serializeWorkPaperNamedExpressions(expressions: Iterable<InternalNamedExpression>): SerializedWorkPaperNamedExpression[] {
  return [...expressions]
    .map((expression) => publicWorkPaperNamedExpressionFromInternal(expression))
    .toSorted((left, right) => (left.scope ?? -1) - (right.scope ?? -1) || left.name.localeCompare(right.name))
}

export function workPaperNamedExpressionToDefinedNameSnapshot(args: {
  readonly expression: RawCellContent
  readonly scope?: number
  readonly defaultScopeId: number
  readonly rewriteFormulaForStorage: (formula: string, ownerSheetId: number) => string
}): WorkbookDefinedNameValueSnapshot {
  const simpleSnapshot = trySimpleWorkPaperNamedExpressionDefinedNameSnapshot(args.expression)
  if (simpleSnapshot !== undefined) {
    return simpleSnapshot
  }
  if (typeof args.expression === 'string' && args.expression.trim().startsWith('=')) {
    return {
      kind: 'formula',
      formula: `=${args.rewriteFormulaForStorage(stripLeadingEquals(args.expression), args.scope ?? args.defaultScopeId)}`,
    }
  }
  return args.expression
}

export function trySimpleWorkPaperNamedExpressionDefinedNameSnapshot(
  expression: RawCellContent,
): WorkbookDefinedNameValueSnapshot | undefined {
  if (expression === null || typeof expression === 'number' || typeof expression === 'boolean') {
    return expression
  }
  if (typeof expression !== 'string') {
    return undefined
  }
  if (!expression.trim().startsWith('=')) {
    return expression
  }
  const simpleBody = tryReadSimpleScalarFormulaBody(expression)
  if (simpleBody !== undefined && tryEvaluateSimpleScalarFormulaBody(simpleBody) !== undefined) {
    return {
      kind: 'formula',
      formula: `=${simpleBody}`,
    }
  }
  return undefined
}

export function evaluateWorkPaperNamedExpression(
  expression: InternalNamedExpression,
  calculateFormula: (formula: string, scope?: number) => CellValue | CellValue[][],
): CellValue | CellValue[][] {
  const raw = expression.expression
  const simpleValue = tryEvaluateSimpleNamedExpression(raw)
  if (simpleValue !== undefined) {
    return simpleValue
  }
  return typeof raw === 'string' ? calculateFormula(raw, expression.scope) : scalarValueFromLiteral(raw)
}

export function workPaperCellSnapshotToRawContent(args: {
  readonly cell: CellSnapshot
  readonly ownerSheetId: number
  readonly restorePublicFormula: (formula: string, ownerSheetId: number) => string
}): RawCellContent {
  if (args.cell.formula) {
    return `=${args.restorePublicFormula(args.cell.formula, args.ownerSheetId)}`
  }
  if (args.cell.input !== undefined) {
    return args.cell.input
  }
  switch (args.cell.value.tag) {
    case ValueTag.Empty:
    case ValueTag.Error:
      return null
    case ValueTag.Number:
      return args.cell.value.value
    case ValueTag.Boolean:
      return args.cell.value.value
    case ValueTag.String:
      return args.cell.value.value
  }
}
