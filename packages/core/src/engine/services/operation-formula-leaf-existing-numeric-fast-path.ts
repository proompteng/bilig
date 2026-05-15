import { getBuiltin, parseNumericText, type FormulaNode } from '@bilig/formula'
import { ErrorCode, type CellValue, FormulaMode, type RecalcMetrics, ValueTag } from '@bilig/protocol'
import { CellFlags } from '../../cell-store.js'
import type { EngineExistingNumericCellMutationResult } from '../../cell-mutations-at.js'
import { makeCellEntity } from '../../entity-ids.js'
import { addEngineCounter } from '../../perf/engine-counters.js'
import type { SheetRecord } from '../../workbook-store.js'
import type { RuntimeFormula } from '../runtime-state.js'
import { cellValuesEqual } from './formula-evaluation-helpers.js'
import { makeCompactExistingNumericMutationResult } from './operation-change-helpers.js'
import { emitOperationTrackedCellsBatch } from './operation-tracked-event-helpers.js'
import type { CreateEngineOperationServiceArgs } from './operation-service-types.js'

export interface OperationTrustedFormulaLeafExistingNumericMutationRequest {
  readonly existingIndex: number
  readonly formulaCellIndex: number
  readonly sheet: SheetRecord
  readonly col: number
  readonly value: number
  readonly hasTrackedEventListeners: boolean
}

export interface OperationFormulaLeafExistingNumericFastPathArgs {
  readonly state: Pick<
    CreateEngineOperationServiceArgs['state'],
    'workbook' | 'strings' | 'wasm' | 'formulas' | 'counters' | 'events' | 'setLastMetrics'
  >
  readonly getSingleEntityDependent: (entityId: number) => number
  readonly writeTrustedExistingNumericLiteralToCell: (existingIndex: number, sheet: SheetRecord, col: number, value: number) => void
  readonly evaluateFormulaCell: (formulaCellIndex: number) => readonly number[]
  readonly deferSingleCellKernelSync: (cellIndex: number) => void
  readonly makeSingleLiteralSkipMetrics: () => RecalcMetrics
}

export function tryApplyTrustedFormulaLeafExistingNumericMutation(
  args: OperationFormulaLeafExistingNumericFastPathArgs,
  request: OperationTrustedFormulaLeafExistingNumericMutationRequest,
): EngineExistingNumericCellMutationResult | null {
  if (request.formulaCellIndex < 0 || args.getSingleEntityDependent(makeCellEntity(request.formulaCellIndex)) !== -1) {
    return null
  }
  const formula = args.state.formulas.get(request.formulaCellIndex)
  if (
    !formula ||
    formula.directLookup !== undefined ||
    formula.directAggregate !== undefined ||
    formula.directCriteria !== undefined ||
    formula.directScalar !== undefined ||
    formula.compiled.volatile ||
    formula.compiled.producesSpill ||
    ((args.state.workbook.cellStore.flags[request.formulaCellIndex] ?? 0) & CellFlags.InCycle) !== 0
  ) {
    return null
  }

  const cellStore = args.state.workbook.cellStore
  const beforeFormulaValue = readFormulaCellValue(args, request.formulaCellIndex)
  args.writeTrustedExistingNumericLiteralToCell(request.existingIndex, request.sheet, request.col, request.value)
  const inlineResult = tryEvaluateLeafFormulaInline(args, formula)
  let evaluatedByWasm = false
  if (inlineResult !== undefined) {
    writeFormulaLeafValue(args, request.formulaCellIndex, inlineResult)
  } else {
    evaluatedByWasm = tryEvaluateLeafFormulaWithWasm(args, request.existingIndex, request.formulaCellIndex, formula)
  }
  if (inlineResult === undefined && !evaluatedByWasm) {
    args.evaluateFormulaCell(request.formulaCellIndex)
  }
  const afterFormulaValue = readFormulaCellValue(args, request.formulaCellIndex)
  const formulaChanged = !cellValuesEqual(beforeFormulaValue, afterFormulaValue)
  addEngineCounter(args.state.counters, 'directFormulaKernelSyncOnlyRecalcSkips')
  args.deferSingleCellKernelSync(request.existingIndex)
  const lastMetrics = {
    ...args.makeSingleLiteralSkipMetrics(),
    wasmFormulaCount: evaluatedByWasm ? 1 : 0,
    jsFormulaCount: evaluatedByWasm ? 0 : 1,
  }
  args.state.setLastMetrics(lastMetrics)

  const result = formulaChanged
    ? makeCompactExistingNumericMutationResult(
        request.existingIndex,
        request.formulaCellIndex,
        1,
        afterFormulaValue.tag === ValueTag.Number ? afterFormulaValue.value : undefined,
        {
          row: cellStore.rows[request.formulaCellIndex] ?? 0,
          col: cellStore.cols[request.formulaCellIndex] ?? 0,
        },
      )
    : makeCompactExistingNumericMutationResult(request.existingIndex, undefined, 1)
  if (request.hasTrackedEventListeners) {
    const changedCellIndices = formulaChanged
      ? Uint32Array.of(request.existingIndex, request.formulaCellIndex)
      : Uint32Array.of(request.existingIndex)
    emitOperationTrackedCellsBatch({
      events: args.state.events,
      changedCellIndices,
      metrics: lastMetrics,
    })
  }
  return result
}

function tryEvaluateLeafFormulaInline(
  args: Pick<OperationFormulaLeafExistingNumericFastPathArgs, 'state'>,
  formula: RuntimeFormula,
): CellValue | undefined {
  return evaluateLeafNode(args, formula, formula.compiled.optimizedAst)
}

function evaluateLeafNode(
  args: Pick<OperationFormulaLeafExistingNumericFastPathArgs, 'state'>,
  formula: RuntimeFormula,
  node: FormulaNode,
): CellValue | undefined {
  switch (node.kind) {
    case 'NumberLiteral':
      return { tag: ValueTag.Number, value: node.value }
    case 'BooleanLiteral':
      return { tag: ValueTag.Boolean, value: node.value }
    case 'StringLiteral':
      return { tag: ValueTag.String, value: node.value, stringId: 0 }
    case 'ErrorLiteral':
      return { tag: ValueTag.Error, code: node.code as ErrorCode }
    case 'CellRef':
      return readLeafCellRef(args, formula, node)
    case 'UnaryExpr': {
      const value = evaluateLeafNode(args, formula, node.argument)
      const numeric = coerceLeafNumber(value)
      if (numeric === undefined) {
        return undefined
      }
      if (numeric.kind === 'error') {
        return { tag: ValueTag.Error, code: numeric.code }
      }
      return { tag: ValueTag.Number, value: node.operator === '-' ? -numeric.value : numeric.value }
    }
    case 'BinaryExpr':
      return evaluateLeafBinary(args, formula, node.operator, node.left, node.right)
    case 'CallExpr': {
      const builtin = getBuiltin(node.callee)
      if (!builtin) {
        return undefined
      }
      const values: CellValue[] = []
      for (let index = 0; index < node.args.length; index += 1) {
        const value = evaluateLeafNode(args, formula, node.args[index]!)
        if (value === undefined) {
          return undefined
        }
        values.push(value)
      }
      return asScalarCellValue(builtin(...values))
    }
    case 'ArrayConstant':
    case 'ColumnRef':
    case 'InvokeExpr':
    case 'NameRef':
    case 'OmittedArgument':
    case 'RangeRef':
    case 'RowRef':
    case 'SpillRef':
    case 'StructuredRef': {
      return undefined
    }
  }
}

function evaluateLeafBinary(
  args: Pick<OperationFormulaLeafExistingNumericFastPathArgs, 'state'>,
  formula: RuntimeFormula,
  operator: Extract<FormulaNode, { kind: 'BinaryExpr' }>['operator'],
  leftNode: FormulaNode,
  rightNode: FormulaNode,
): CellValue | undefined {
  const left = evaluateLeafNode(args, formula, leftNode)
  const right = evaluateLeafNode(args, formula, rightNode)
  if (left === undefined || right === undefined) {
    return undefined
  }
  if (left.tag === ValueTag.Error) {
    return left
  }
  if (right.tag === ValueTag.Error) {
    return right
  }
  if (operator === '&') {
    return { tag: ValueTag.String, value: `${leafText(left)}${leafText(right)}`, stringId: 0 }
  }
  if (operator === '=' || operator === '<>' || operator === '>' || operator === '>=' || operator === '<' || operator === '<=') {
    return evaluateLeafComparison(operator, left, right)
  }
  const leftNumber = coerceLeafNumber(left)
  const rightNumber = coerceLeafNumber(right)
  if (leftNumber === undefined || rightNumber === undefined) {
    return undefined
  }
  if (leftNumber.kind === 'error') {
    return { tag: ValueTag.Error, code: leftNumber.code }
  }
  if (rightNumber.kind === 'error') {
    return { tag: ValueTag.Error, code: rightNumber.code }
  }
  switch (operator) {
    case '+':
      return { tag: ValueTag.Number, value: leftNumber.value + rightNumber.value }
    case '-':
      return { tag: ValueTag.Number, value: leftNumber.value - rightNumber.value }
    case '*':
      return { tag: ValueTag.Number, value: leftNumber.value * rightNumber.value }
    case '/':
      return rightNumber.value === 0
        ? { tag: ValueTag.Error, code: ErrorCode.Div0 }
        : { tag: ValueTag.Number, value: leftNumber.value / rightNumber.value }
    case '^':
      return { tag: ValueTag.Number, value: leftNumber.value ** rightNumber.value }
    case ':':
      return undefined
  }
}

function asScalarCellValue(
  value: CellValue | { readonly values: readonly CellValue[]; readonly rows: number; readonly cols: number },
): CellValue | undefined {
  return 'tag' in value ? value : undefined
}

function evaluateLeafComparison(operator: '=' | '<>' | '>' | '>=' | '<' | '<=', left: CellValue, right: CellValue): CellValue {
  const leftComparable = comparableLeafValue(left)
  const rightComparable = comparableLeafValue(right)
  let result: boolean
  switch (operator) {
    case '=':
      result = leftComparable === rightComparable
      break
    case '<>':
      result = leftComparable !== rightComparable
      break
    case '>':
      result = leftComparable > rightComparable
      break
    case '>=':
      result = leftComparable >= rightComparable
      break
    case '<':
      result = leftComparable < rightComparable
      break
    case '<=':
      result = leftComparable <= rightComparable
      break
  }
  return { tag: ValueTag.Boolean, value: result }
}

function readLeafCellRef(
  args: Pick<OperationFormulaLeafExistingNumericFastPathArgs, 'state'>,
  formula: RuntimeFormula,
  node: Extract<FormulaNode, { kind: 'CellRef' }>,
): CellValue | undefined {
  const ownerSheetId = args.state.workbook.cellStore.sheetIds[formula.cellIndex]
  const ownerSheetName = ownerSheetId === undefined ? undefined : args.state.workbook.getSheetNameById(ownerSheetId)
  const sheetName = node.sheetName ?? ownerSheetName
  if (!sheetName) {
    return undefined
  }
  const cellIndex = args.state.workbook.getCellIndex(sheetName, node.ref)
  return cellIndex === undefined ? { tag: ValueTag.Empty } : readFormulaCellValue(args, cellIndex)
}

function coerceLeafNumber(
  value: CellValue | undefined,
): { readonly kind: 'number'; readonly value: number } | { readonly kind: 'error'; readonly code: ErrorCode } | undefined {
  if (value === undefined) {
    return undefined
  }
  switch (value.tag) {
    case ValueTag.Number:
      return { kind: 'number', value: value.value }
    case ValueTag.Boolean:
      return { kind: 'number', value: value.value ? 1 : 0 }
    case ValueTag.Empty:
      return { kind: 'number', value: 0 }
    case ValueTag.Error:
      return { kind: 'error', code: value.code }
    case ValueTag.String: {
      const trimmed = value.value.trim()
      if (trimmed.length === 0) {
        return { kind: 'number', value: 0 }
      }
      const parsed = parseNumericText(trimmed)
      return parsed === undefined ? { kind: 'error', code: ErrorCode.Value } : { kind: 'number', value: parsed }
    }
  }
}

function comparableLeafValue(value: CellValue): number | string {
  switch (value.tag) {
    case ValueTag.Number:
      return value.value
    case ValueTag.Boolean:
      return value.value ? 1 : 0
    case ValueTag.Empty:
      return 0
    case ValueTag.String: {
      const parsed = parseNumericText(value.value.trim())
      return parsed ?? value.value
    }
    case ValueTag.Error:
      return String(value.code)
  }
}

function leafText(value: CellValue): string {
  switch (value.tag) {
    case ValueTag.Empty:
      return ''
    case ValueTag.Number:
      return String(Object.is(value.value, -0) ? 0 : value.value)
    case ValueTag.Boolean:
      return value.value ? 'TRUE' : 'FALSE'
    case ValueTag.String:
      return value.value
    case ValueTag.Error:
      return String(value.code)
  }
}

function writeFormulaLeafValue(
  args: Pick<OperationFormulaLeafExistingNumericFastPathArgs, 'state'>,
  formulaCellIndex: number,
  value: CellValue,
): void {
  const cellStore = args.state.workbook.cellStore
  const before = readFormulaCellValue(args, formulaCellIndex)
  cellStore.flags[formulaCellIndex] = (cellStore.flags[formulaCellIndex] ?? 0) & ~(CellFlags.SpillChild | CellFlags.PivotOutput)
  cellStore.setValue(formulaCellIndex, value, value.tag === ValueTag.String ? args.state.strings.intern(value.value) : 0)
  if (!cellValuesEqual(before, value)) {
    args.state.workbook.notifyCellValueWritten(formulaCellIndex)
  }
}

function tryEvaluateLeafFormulaWithWasm(
  args: Pick<OperationFormulaLeafExistingNumericFastPathArgs, 'state'>,
  existingIndex: number,
  formulaCellIndex: number,
  formula: RuntimeFormula,
): boolean {
  if (!shouldUseWasmLeafFormulaFastPath(formula) || !args.state.wasm.ready) {
    return false
  }
  args.state.wasm.syncFromStore(args.state.workbook.cellStore, Uint32Array.of(existingIndex))
  const formulaIndices = Uint32Array.of(formulaCellIndex)
  args.state.wasm.evalBatch(formulaIndices)
  args.state.wasm.syncToStore(args.state.workbook.cellStore, formulaIndices, args.state.strings, (changedCellIndex: number) =>
    args.state.workbook.notifyCellValueWritten(changedCellIndex),
  )
  return true
}

function shouldUseWasmLeafFormulaFastPath(formula: RuntimeFormula): boolean {
  if (formula.compiled.mode !== FormulaMode.WasmFastPath) {
    return false
  }
  const source = formula.source.toUpperCase()
  if (!/[A-Z][A-Z0-9.]*\s*\(/.test(source)) {
    return false
  }
  return !/\b(?:CONCAT|CONCATENATE|LEFT|LEN|MID|RIGHT|TEXT)\s*\(/.test(source)
}

function readFormulaCellValue(args: Pick<OperationFormulaLeafExistingNumericFastPathArgs, 'state'>, formulaCellIndex: number): CellValue {
  return args.state.workbook.cellStore.getValue(formulaCellIndex, (stringId) => (stringId === 0 ? '' : args.state.strings.get(stringId)))
}
