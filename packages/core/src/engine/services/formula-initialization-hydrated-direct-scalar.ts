import { addEngineCounter } from '../../perf/engine-counters.js'
import type {
  EngineFormulaInitializationServiceArgs,
  HydratedPreparedFormulaInitializationRef,
} from './formula-initialization-service-types.js'
import type { FreshDirectAggregateFormulaBindingMember } from './formula-binding-service-types.js'
import { unwrapDirectScalarBinaryNode } from './formula-binding-direct-scalar.js'

export function tryBindHydratedFreshDirectFormula(
  serviceArgs: EngineFormulaInitializationServiceArgs,
  hadExistingFormulas: boolean,
  cellIndex: number,
  ownerSheetName: string,
  ref: HydratedPreparedFormulaInitializationRef,
): boolean {
  return (
    tryBindHydratedFreshDirectScalarFormula(serviceArgs, hadExistingFormulas, cellIndex, ownerSheetName, ref) ||
    tryBindHydratedFreshDirectAggregateFormula(serviceArgs, hadExistingFormulas, cellIndex, ownerSheetName, ref)
  )
}

function tryBindHydratedFreshDirectScalarFormula(
  serviceArgs: EngineFormulaInitializationServiceArgs,
  hadExistingFormulas: boolean,
  cellIndex: number,
  ownerSheetName: string,
  ref: HydratedPreparedFormulaInitializationRef,
): boolean {
  if (
    hadExistingFormulas ||
    serviceArgs.bindFreshDirectScalarFormulaRun === undefined ||
    ref.templateId === undefined ||
    ref.preserveCachedValueOnFullRecalc === true ||
    !canUseFreshDirectScalarFormulaBinding(ref.compiled)
  ) {
    return false
  }
  serviceArgs.bindFreshDirectScalarFormulaRun({
    sheetId: ref.sheetId,
    ownerSheetName,
    cellIndex,
    member: {
      row: ref.row,
      col: ref.col,
      source: ref.source,
      compiled: ref.compiled,
      templateId: ref.templateId,
    },
  })
  addEngineCounter(serviceArgs.state.counters, 'runtimeHydratedDirectScalarFastBindings')
  return true
}

function tryBindHydratedFreshDirectAggregateFormula(
  serviceArgs: EngineFormulaInitializationServiceArgs,
  hadExistingFormulas: boolean,
  cellIndex: number,
  ownerSheetName: string,
  ref: HydratedPreparedFormulaInitializationRef,
): boolean {
  if (
    hadExistingFormulas ||
    serviceArgs.bindFreshDirectAggregateFormulaRun === undefined ||
    ref.templateId === undefined ||
    ref.preserveCachedValueOnFullRecalc === true
  ) {
    return false
  }
  const member = buildFreshDirectAggregateMember(ownerSheetName, ref)
  if (member === undefined) {
    return false
  }
  serviceArgs.bindFreshDirectAggregateFormulaRun({
    sheetId: ref.sheetId,
    ownerSheetName,
    cellIndex,
    member,
  })
  addEngineCounter(serviceArgs.state.counters, 'runtimeHydratedDirectAggregateFastBindings')
  return true
}

function buildFreshDirectAggregateMember(
  ownerSheetName: string,
  ref: HydratedPreparedFormulaInitializationRef,
): FreshDirectAggregateFormulaBindingMember | undefined {
  const compiled = ref.compiled
  if (
    compiled.volatile ||
    compiled.producesSpill ||
    compiled.symbolicRanges.length !== 1 ||
    compiled.symbolicNames.length !== 0 ||
    compiled.symbolicTables.length !== 0 ||
    compiled.symbolicSpills.length !== 0
  ) {
    return undefined
  }
  const aggregate = compiled.directAggregateCandidate
  const range = aggregate === undefined ? undefined : compiled.parsedSymbolicRanges?.[aggregate.symbolicRangeIndex]
  if (
    aggregate === undefined ||
    range === undefined ||
    range.refKind !== 'cells' ||
    (range.sheetName ?? ownerSheetName) !== ownerSheetName ||
    range.startRow > range.endRow ||
    range.startCol > range.endCol ||
    (range.startRow <= ref.row && ref.row <= range.endRow && range.startCol <= ref.col && ref.col <= range.endCol)
  ) {
    return undefined
  }
  return {
    row: ref.row,
    col: ref.col,
    source: ref.source,
    compiled,
    templateId: ref.templateId!,
    aggregateKind: aggregate.aggregateKind,
    aggregateRowStart: range.startRow,
    aggregateRowEnd: range.endRow,
    aggregateColStart: range.startCol,
    aggregateColEnd: range.endCol,
    resultOffset: normalizeDirectAggregateResultOffset(aggregate.resultOffset),
  }
}

function normalizeDirectAggregateResultOffset(offset: number | undefined): number | undefined {
  return offset === undefined || Object.is(offset, 0) ? undefined : offset
}

function canUseFreshDirectScalarFormulaBinding(compiled: HydratedPreparedFormulaInitializationRef['compiled']): boolean {
  if (
    compiled.volatile ||
    compiled.producesSpill ||
    compiled.symbolicRanges.length !== 0 ||
    compiled.symbolicNames.length !== 0 ||
    compiled.symbolicTables.length !== 0 ||
    compiled.symbolicSpills.length !== 0
  ) {
    return false
  }
  const node = unwrapDirectScalarBinaryNode(compiled.optimizedAst).node
  if (node.kind === 'BinaryExpr' && (node.operator === '+' || node.operator === '-' || node.operator === '*' || node.operator === '/')) {
    return isFreshDirectScalarOperand(node.left) && isFreshDirectScalarOperand(node.right)
  }
  return (
    node.kind === 'CallExpr' &&
    node.callee.trim().toUpperCase() === 'ABS' &&
    node.args.length === 1 &&
    isFreshDirectScalarOperand(node.args[0]!)
  )
}

function isFreshDirectScalarOperand(node: HydratedPreparedFormulaInitializationRef['compiled']['optimizedAst']): boolean {
  return node.kind === 'NumberLiteral' || node.kind === 'CellRef'
}
