import { addEngineCounter } from '../../perf/engine-counters.js'
import type {
  EngineFormulaInitializationServiceArgs,
  HydratedPreparedFormulaInitializationRef,
} from './formula-initialization-service-types.js'
import { unwrapDirectScalarBinaryNode } from './formula-binding-direct-scalar.js'

export function tryBindHydratedFreshDirectScalarFormula(
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
    cellIndices: [cellIndex],
    members: [
      {
        row: ref.row,
        col: ref.col,
        source: ref.source,
        compiled: ref.compiled,
        templateId: ref.templateId,
      },
    ],
  })
  addEngineCounter(serviceArgs.state.counters, 'runtimeHydratedDirectScalarFastBindings')
  return true
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
