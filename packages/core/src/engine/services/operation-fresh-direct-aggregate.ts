import { ErrorCode, ValueTag } from '@bilig/protocol'
import type { EngineCellMutationAt } from '../../cell-mutations-at.js'
import type { DirectScalarCurrentOperand } from './direct-formula-index-collection.js'
import type { RuntimeDirectAggregateDescriptor, RuntimeFormula } from '../runtime-state.js'
import type { CreateEngineOperationServiceArgs } from './operation-service-types.js'

const DIRECT_AGGREGATE_TOPO_SKIP_SCAN_LIMIT = 4096

export interface FreshDirectAggregateFormulaAnalysis {
  readonly canSkipTopoRepair: boolean
  readonly currentResult?: DirectScalarCurrentOperand
}

export function analyzeFreshDirectAggregateFormula(
  args: CreateEngineOperationServiceArgs,
  input: {
    readonly priorHadFormula: boolean
    readonly formulaCellIndex: number
    readonly formula: RuntimeFormula | undefined
  },
): FreshDirectAggregateFormulaAnalysis {
  const formula = input.formula
  const directAggregate = formula?.directAggregate
  if (
    input.priorHadFormula ||
    formula === undefined ||
    directAggregate === undefined ||
    formula.dependencyIndices.length !== 0 ||
    formula.rangeDependencies.length !== 0 ||
    formula.graphRangeDependencies.length !== 0 ||
    directAggregate.length > DIRECT_AGGREGATE_TOPO_SKIP_SCAN_LIMIT
  ) {
    return { canSkipTopoRepair: false }
  }
  const aggregateSheet = args.state.workbook.getSheet(directAggregate.sheetName)
  if (!aggregateSheet) {
    return { canSkipTopoRepair: false }
  }
  const canEvaluateCurrentResult =
    !formula.compiled.producesSpill &&
    formula.directCriteria === undefined &&
    formula.directLookup === undefined &&
    formula.directScalar === undefined
  if (!canEvaluateCurrentResult) {
    return {
      canSkipTopoRepair: scanFreshDirectAggregateDependenciesForTopoSkip(args, input.formulaCellIndex, aggregateSheet, directAggregate),
    }
  }
  const analysis = analyzeFreshDirectAggregateDependencies(args, input.formulaCellIndex, aggregateSheet, directAggregate)
  if (analysis === undefined) {
    return { canSkipTopoRepair: false }
  }
  return {
    canSkipTopoRepair: true,
    currentResult:
      analysis.kind === 'number' && directAggregate.resultOffset !== undefined
        ? { kind: 'number', value: analysis.value + directAggregate.resultOffset }
        : analysis,
  }
}

export function canSkipTopoRepairForFreshDirectAggregate(
  args: CreateEngineOperationServiceArgs,
  input: {
    readonly priorHadFormula: boolean
    readonly formulaCellIndex: number
    readonly formula: RuntimeFormula | undefined
  },
): boolean {
  return analyzeFreshDirectAggregateFormula(args, input).canSkipTopoRepair
}

export function tryEvaluateFreshDirectAggregateCurrentResult(
  args: CreateEngineOperationServiceArgs,
  formula: RuntimeFormula | undefined,
): DirectScalarCurrentOperand | undefined {
  if (
    formula === undefined ||
    formula.compiled.producesSpill ||
    formula.directAggregate === undefined ||
    formula.directCriteria !== undefined ||
    formula.directLookup !== undefined ||
    formula.directScalar !== undefined ||
    formula.dependencyIndices.length !== 0 ||
    formula.rangeDependencies.length !== 0 ||
    formula.graphRangeDependencies.length !== 0
  ) {
    return undefined
  }
  const directAggregate = formula.directAggregate
  const aggregateSheet = args.state.workbook.getSheet(directAggregate.sheetName)
  if (!aggregateSheet) {
    return undefined
  }
  const result = evaluateDirectAggregateFromCellStore(args, aggregateSheet, directAggregate)
  if (result.kind === 'number' && directAggregate.resultOffset !== undefined) {
    return { kind: 'number', value: result.value + directAggregate.resultOffset }
  }
  return result
}

export function bindFreshTemplateFormula(
  args: CreateEngineOperationServiceArgs,
  cellIndex: number,
  sheetName: string,
  mutation: Extract<EngineCellMutationAt, { kind: 'setCellFormula' }>,
): boolean {
  if (args.bindPreparedFormula === undefined || args.compileTemplateFormula === undefined) {
    return args.bindFormula(cellIndex, sheetName, mutation.formula)
  }
  const template = args.compileTemplateFormula(mutation.formula, mutation.row, mutation.col)
  if (
    template.compiled.symbolicNames.length !== 0 ||
    template.compiled.symbolicTables.length !== 0 ||
    template.compiled.symbolicSpills.length !== 0
  ) {
    return args.bindFormula(cellIndex, sheetName, mutation.formula)
  }
  return args.bindPreparedFormula(cellIndex, sheetName, mutation.formula, template.compiled, template.templateId, {
    assumeFreshFormula: true,
  })
}

function evaluateDirectAggregateFromCellStore(
  args: CreateEngineOperationServiceArgs,
  aggregateSheet: NonNullable<ReturnType<CreateEngineOperationServiceArgs['state']['workbook']['getSheet']>>,
  directAggregate: RuntimeDirectAggregateDescriptor,
): DirectScalarCurrentOperand {
  const cellStore = args.state.workbook.cellStore
  let sum = 0
  let count = 0
  let averageCount = 0
  let minimum = Number.POSITIVE_INFINITY
  let maximum = Number.NEGATIVE_INFINITY
  for (let col = directAggregate.col; col <= directAggregate.colEnd; col += 1) {
    for (let row = directAggregate.rowStart; row <= directAggregate.rowEnd; row += 1) {
      const memberCellIndex =
        aggregateSheet.structureVersion === 1 ? aggregateSheet.grid.getPhysical(row, col) : aggregateSheet.grid.get(row, col)
      if (memberCellIndex === -1) {
        continue
      }
      const tag = (cellStore.tags[memberCellIndex] as ValueTag | undefined) ?? ValueTag.Empty
      switch (tag) {
        case ValueTag.Number: {
          const value = cellStore.numbers[memberCellIndex] ?? 0
          sum += value
          count += 1
          averageCount += 1
          minimum = Math.min(minimum, value)
          maximum = Math.max(maximum, value)
          break
        }
        case ValueTag.Boolean: {
          const value = (cellStore.numbers[memberCellIndex] ?? 0) !== 0 ? 1 : 0
          sum += value
          count += 1
          averageCount += 1
          minimum = Math.min(minimum, value)
          maximum = Math.max(maximum, value)
          break
        }
        case ValueTag.Error:
          if (directAggregate.aggregateKind === 'sum' || directAggregate.aggregateKind === 'average') {
            return { kind: 'error', code: (cellStore.errors[memberCellIndex] as ErrorCode | undefined) ?? ErrorCode.None }
          }
          break
        case ValueTag.Empty:
        case ValueTag.String:
          break
      }
    }
  }
  switch (directAggregate.aggregateKind) {
    case 'sum':
      return { kind: 'number', value: sum }
    case 'count':
      return { kind: 'number', value: count }
    case 'average':
      return averageCount === 0 ? { kind: 'error', code: ErrorCode.Div0 } : { kind: 'number', value: sum / averageCount }
    case 'min':
      return { kind: 'number', value: minimum === Number.POSITIVE_INFINITY ? 0 : minimum }
    case 'max':
      return { kind: 'number', value: maximum === Number.NEGATIVE_INFINITY ? 0 : maximum }
  }
}

function scanFreshDirectAggregateDependenciesForTopoSkip(
  args: CreateEngineOperationServiceArgs,
  formulaCellIndex: number,
  aggregateSheet: NonNullable<ReturnType<CreateEngineOperationServiceArgs['state']['workbook']['getSheet']>>,
  directAggregate: RuntimeDirectAggregateDescriptor,
): boolean {
  for (let row = directAggregate.rowStart; row <= directAggregate.rowEnd; row += 1) {
    for (let col = directAggregate.col; col <= directAggregate.colEnd; col += 1) {
      const dependencyCellIndex = aggregateSheet.grid.getPhysical(row, col)
      if (dependencyCellIndex !== -1 && dependencyCellIndex !== formulaCellIndex && args.state.formulas.has(dependencyCellIndex)) {
        return false
      }
    }
  }
  return true
}

function analyzeFreshDirectAggregateDependencies(
  args: CreateEngineOperationServiceArgs,
  formulaCellIndex: number,
  aggregateSheet: NonNullable<ReturnType<CreateEngineOperationServiceArgs['state']['workbook']['getSheet']>>,
  directAggregate: RuntimeDirectAggregateDescriptor,
): DirectScalarCurrentOperand | undefined {
  const cellStore = args.state.workbook.cellStore
  let sum = 0
  let count = 0
  let averageCount = 0
  let minimum = Number.POSITIVE_INFINITY
  let maximum = Number.NEGATIVE_INFINITY
  for (let col = directAggregate.col; col <= directAggregate.colEnd; col += 1) {
    for (let row = directAggregate.rowStart; row <= directAggregate.rowEnd; row += 1) {
      const dependencyCellIndex = aggregateSheet.grid.getPhysical(row, col)
      if (dependencyCellIndex !== -1 && dependencyCellIndex !== formulaCellIndex && args.state.formulas.has(dependencyCellIndex)) {
        return undefined
      }
      const memberCellIndex = aggregateSheet.structureVersion === 1 ? dependencyCellIndex : aggregateSheet.grid.get(row, col)
      if (memberCellIndex === -1) {
        continue
      }
      const tag = (cellStore.tags[memberCellIndex] as ValueTag | undefined) ?? ValueTag.Empty
      switch (tag) {
        case ValueTag.Number: {
          const value = cellStore.numbers[memberCellIndex] ?? 0
          sum += value
          count += 1
          averageCount += 1
          minimum = Math.min(minimum, value)
          maximum = Math.max(maximum, value)
          break
        }
        case ValueTag.Boolean: {
          const value = (cellStore.numbers[memberCellIndex] ?? 0) !== 0 ? 1 : 0
          sum += value
          count += 1
          averageCount += 1
          minimum = Math.min(minimum, value)
          maximum = Math.max(maximum, value)
          break
        }
        case ValueTag.Error:
          if (directAggregate.aggregateKind === 'sum' || directAggregate.aggregateKind === 'average') {
            return { kind: 'error', code: (cellStore.errors[memberCellIndex] as ErrorCode | undefined) ?? ErrorCode.None }
          }
          break
        case ValueTag.Empty:
        case ValueTag.String:
          break
      }
    }
  }
  switch (directAggregate.aggregateKind) {
    case 'sum':
      return { kind: 'number', value: sum }
    case 'count':
      return { kind: 'number', value: count }
    case 'average':
      return averageCount === 0 ? { kind: 'error', code: ErrorCode.Div0 } : { kind: 'number', value: sum / averageCount }
    case 'min':
      return { kind: 'number', value: minimum === Number.POSITIVE_INFINITY ? 0 : minimum }
    case 'max':
      return { kind: 'number', value: maximum === Number.NEGATIVE_INFINITY ? 0 : maximum }
  }
}
