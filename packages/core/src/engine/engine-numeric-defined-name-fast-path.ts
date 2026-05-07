import type { CellValue, WorkbookDefinedNameValueSnapshot } from '@bilig/protocol'
import { ErrorCode, ValueTag } from '@bilig/protocol'
import type { FormulaNode } from '@bilig/formula'
import type { EngineOp } from '@bilig/workbook-domain'
import { CellFlags } from '../cell-store.js'
import type { FormulaTable } from '../formula-table.js'
import { cloneDefinedNameValue } from '../workbook-metadata-records.js'
import { normalizeDefinedName, type WorkbookStore } from '../workbook-store.js'
import { batchOpOrder, createBatch, type OpOrder, type ReplicaState } from '../replica-state.js'
import type { RuntimeFormula, TransactionLogEntry } from './runtime-state.js'

function namedNumberOperand(node: FormulaNode, normalizedName: string, namedValue: number): number | undefined {
  if (node.kind === 'NumberLiteral') {
    return node.value
  }
  if (node.kind === 'NameRef' && normalizeDefinedName(node.name) === normalizedName) {
    return namedValue
  }
  return undefined
}

function evaluateNumericDefinedNameFormula(node: FormulaNode, normalizedName: string, namedValue: number): CellValue | undefined {
  if (node.kind === 'NameRef' && normalizeDefinedName(node.name) === normalizedName) {
    return { tag: ValueTag.Number, value: namedValue }
  }
  if (node.kind !== 'BinaryExpr') {
    return undefined
  }
  const left = namedNumberOperand(node.left, normalizedName, namedValue)
  const right = namedNumberOperand(node.right, normalizedName, namedValue)
  if (left === undefined || right === undefined) {
    return undefined
  }
  if (node.operator === '+') {
    return { tag: ValueTag.Number, value: left + right }
  }
  if (node.operator === '-') {
    return { tag: ValueTag.Number, value: left - right }
  }
  if (node.operator === '*') {
    return { tag: ValueTag.Number, value: left * right }
  }
  if (node.operator === '/') {
    return right === 0 ? { tag: ValueTag.Error, code: ErrorCode.Div0 } : { tag: ValueTag.Number, value: left / right }
  }
  return undefined
}

export function upsertNumericDefinedNameFast(args: {
  readonly workbook: WorkbookStore
  readonly formulas: FormulaTable<RuntimeFormula>
  readonly replicaState: ReplicaState
  readonly entityVersions: Map<string, OpOrder>
  readonly undoStack: TransactionLogEntry[]
  readonly redoStack: TransactionLogEntry[]
  readonly collectDependentFormulaCells: (normalizedName: string) => readonly number[]
  readonly name: string
  readonly value: WorkbookDefinedNameValueSnapshot
  readonly numericValue: number
}): readonly number[] | null {
  const trimmedName = args.name.trim()
  const normalizedName = normalizeDefinedName(trimmedName)
  const dependentCellIndices = args.collectDependentFormulaCells(normalizedName)
  const evaluated: Array<{ cellIndex: number; value: CellValue }> = []
  for (let index = 0; index < dependentCellIndices.length; index += 1) {
    const cellIndex = dependentCellIndices[index]!
    const formula = args.formulas.get(cellIndex)
    if (
      formula === undefined ||
      formula.compiled.symbolicNames.length !== 1 ||
      normalizeDefinedName(formula.compiled.symbolicNames[0]!) !== normalizedName ||
      formula.compiled.symbolicRanges.length !== 0 ||
      formula.compiled.symbolicTables.length !== 0 ||
      formula.compiled.symbolicSpills.length !== 0
    ) {
      return null
    }
    const nextValue = evaluateNumericDefinedNameFormula(formula.compiled.ast, normalizedName, args.numericValue)
    if (nextValue === undefined) {
      return null
    }
    evaluated.push({ cellIndex, value: nextValue })
  }

  const existing = args.workbook.getDefinedName(trimmedName)
  const op: EngineOp = { kind: 'upsertDefinedName', name: trimmedName, value: args.value }
  const batch = createBatch(args.replicaState, [op])
  const order = batchOpOrder(batch, 0)
  args.workbook.setDefinedName(trimmedName, args.value)
  args.entityVersions.set(`defined-name:${normalizedName}`, order)

  const changedCellIndices: number[] = []
  const cellStore = args.workbook.cellStore
  for (let index = 0; index < evaluated.length; index += 1) {
    const { cellIndex, value: nextValue } = evaluated[index]!
    const beforeTag = (cellStore.tags[cellIndex] as ValueTag | undefined) ?? ValueTag.Empty
    const beforeNumber = cellStore.numbers[cellIndex] ?? 0
    const beforeError = cellStore.errors[cellIndex] ?? ErrorCode.None
    const changed =
      beforeTag !== nextValue.tag ||
      (nextValue.tag === ValueTag.Number && !Object.is(beforeNumber, nextValue.value)) ||
      (nextValue.tag === ValueTag.Error && (beforeError as ErrorCode) !== nextValue.code)
    cellStore.flags[cellIndex] = (cellStore.flags[cellIndex] ?? 0) & ~(CellFlags.SpillChild | CellFlags.PivotOutput)
    cellStore.setValue(cellIndex, nextValue, 0)
    if (changed) {
      args.workbook.notifyCellValueWritten(cellIndex)
      changedCellIndices.push(cellIndex)
    }
  }

  const inverseOp: EngineOp =
    existing === undefined
      ? { kind: 'deleteDefinedName', name: trimmedName }
      : { kind: 'upsertDefinedName', name: existing.name, value: cloneDefinedNameValue(existing.value) }
  args.undoStack.push({
    forward: { kind: 'single-op', op: { kind: 'upsertDefinedName', name: trimmedName, value: cloneDefinedNameValue(args.value) } },
    inverse: { kind: 'single-op', op: inverseOp },
  })
  args.redoStack.length = 0
  return changedCellIndices
}
