import { ErrorCode } from '@bilig/protocol'
import { getExternalLookupFunction, type ExternalLookupFunctionArgument } from '../external-function-adapter.js'
import { createLookupArrayShapeBuiltins } from './lookup-array-shape-builtins.js'
import {
  arrayResult,
  areRangeArgs,
  collectNumericSeries,
  compareScalars,
  compileCriteriaMatcher,
  errorValue,
  findFirstNonRange,
  firstLookupError,
  flattenNumbers,
  getRangeValue,
  isError,
  isRangeArg,
  matchesCompiledCriteria,
  matchesCriteria,
  numberResult,
  numericAggregateCandidate,
  pickRangeRow,
  requireCellRange,
  requireCellVector,
  toBoolean,
  toCellRange,
  toInteger,
  toNumber,
  toNumericMatrix,
  toStringValue,
  type CompiledCriteriaMatcher,
  type CriteriaOperator,
  type LookupBuiltin,
  type LookupBuiltinArgument,
  type LookupBuiltinResolverOptions,
  type RangeBuiltinArgument,
} from './lookup-core-helpers.js'
import { createLookupCriteriaBuiltins } from './lookup-criteria-builtins.js'
import { createLookupDatabaseBuiltins } from './lookup-database-builtins.js'
import { createLookupFinancialBuiltins } from './lookup-financial-builtins.js'
import { createLookupHypothesisBuiltins } from './lookup-hypothesis-builtins.js'
import { createLookupMatrixBuiltins } from './lookup-matrix-builtins.js'
import { createLookupOrderStatisticsBuiltins } from './lookup-order-statistics-builtins.js'
import { createLookupReferenceBuiltins } from './lookup-reference-builtins.js'
import { createLookupRegressionBuiltins } from './lookup-regression-builtins.js'
import { createLookupSortFilterBuiltins } from './lookup-sort-filter-builtins.js'

export { compileCriteriaMatcher, matchesCompiledCriteria }
export { hasLookupWildcardSyntax } from './lookup-reference-search.js'
export { exactLookupNumberKey, normalizeExactLookupNumber, sameExactLookupNumber } from './lookup-core-helpers.js'
export type {
  CompiledCriteriaMatcher,
  CriteriaOperator,
  LookupBuiltin,
  LookupBuiltinArgument,
  LookupBuiltinResolverOptions,
  RangeBuiltinArgument,
}

const externalLookupBuiltinNames = ['FILTERXML', 'STOCKHISTORY'] as const

function isExternalLookupFunctionArgument(arg: LookupBuiltinArgument): arg is ExternalLookupFunctionArgument {
  return arg !== undefined
}

function createExternalLookupBuiltin(name: string): LookupBuiltin {
  return (...args) => {
    const existingError = firstLookupError(args)
    if (existingError) {
      return existingError
    }
    if (!args.every(isExternalLookupFunctionArgument)) {
      return errorValue(ErrorCode.Value)
    }
    const external = getExternalLookupFunction(name)
    return external ? external(...args) : errorValue(ErrorCode.Blocked)
  }
}

const externalLookupBuiltins = Object.fromEntries(
  externalLookupBuiltinNames.map((name) => [name, createExternalLookupBuiltin(name)]),
) as Record<string, LookupBuiltin>

const lookupRegressionBuiltins = createLookupRegressionBuiltins({
  errorValue,
  numberResult,
  isRangeArg,
  toNumber,
  toBoolean,
  flattenNumbers,
})

const lookupOrderStatisticsBuiltins = createLookupOrderStatisticsBuiltins({
  errorValue,
  numberResult,
  arrayResult,
  requireCellRange,
  isError,
  isRangeArg,
  toNumber,
  toInteger,
  flattenNumbers,
})

const lookupFinancialBuiltins = createLookupFinancialBuiltins({
  errorValue,
  numberResult,
  isRangeArg,
  toNumber,
  collectNumericSeries,
})

const lookupHypothesisBuiltins = createLookupHypothesisBuiltins({
  errorValue,
  isRangeArg,
  toNumber,
  toNumericMatrix,
})

const lookupMatrixBuiltins = createLookupMatrixBuiltins({
  errorValue,
  numberResult,
  arrayResult,
  isRangeArg,
  requireCellRange,
  findFirstNonRange,
  areRangeArgs,
  toNumber,
  toNumericMatrix,
  flattenNumbers,
})

const lookupSortFilterBuiltins = createLookupSortFilterBuiltins({
  errorValue,
  arrayResult,
  isError,
  isRangeArg,
  toBoolean,
  toInteger,
  requireCellRange,
  toCellRange,
  compareScalars,
  getRangeValue,
  pickRangeRow,
})

const lookupDatabaseBuiltins = createLookupDatabaseBuiltins({
  errorValue,
  numberResult,
  isError,
  isRangeArg,
  toNumber,
  toStringValue,
  requireCellRange,
  getRangeValue,
  matchesCriteria,
})

const lookupCriteriaBuiltins = createLookupCriteriaBuiltins({
  errorValue,
  numberResult,
  isError,
  isRangeArg,
  toNumber,
  requireCellRange,
  matchesCriteria,
  numericAggregateCandidate,
})

function createLookupBuiltinMap(options: LookupBuiltinResolverOptions = {}): Record<string, LookupBuiltin> {
  const lookupReferenceBuiltins = createLookupReferenceBuiltins({
    errorValue,
    numberResult,
    arrayResult,
    isError,
    isRangeArg,
    toBoolean,
    toInteger,
    requireCellVector,
    toCellRange,
    compareScalars,
    getRangeValue,
    ...(options.resolveIndexedExactMatch ? { resolveIndexedExactMatch: options.resolveIndexedExactMatch } : {}),
  })

  return {
    ...lookupArrayShapeBuiltins,
    ...lookupReferenceBuiltins,
    ...lookupCriteriaBuiltins,
    ...lookupDatabaseBuiltins,
    ...lookupFinancialBuiltins,
    ...lookupHypothesisBuiltins,
    ...lookupRegressionBuiltins,
    ...lookupOrderStatisticsBuiltins,
    ...lookupMatrixBuiltins,
    ...lookupSortFilterBuiltins,
    ...externalLookupBuiltins,
  }
}

const lookupArrayShapeBuiltins = createLookupArrayShapeBuiltins({
  errorValue,
  arrayResult,
  isError,
  isRangeArg,
  toBoolean,
  toInteger,
  requireCellRange,
  toCellRange,
  getRangeValue,
  findFirstNonRange,
  areRangeArgs,
  pickRangeRow,
})

export const lookupBuiltins: Record<string, LookupBuiltin> = createLookupBuiltinMap()

export function createLookupBuiltinResolver(options: LookupBuiltinResolverOptions = {}): (name: string) => LookupBuiltin | undefined {
  const builtins = options.resolveIndexedExactMatch === undefined ? lookupBuiltins : createLookupBuiltinMap(options)
  return (name: string) => {
    const upper = name.toUpperCase()
    if (upper === 'USE.THE.COUNTIF') {
      return builtins['COUNTIF']
    }
    const external = getExternalLookupFunction(name)
    return builtins[upper] ?? (external ? (...args) => createExternalLookupBuiltin(name)(...args) : undefined)
  }
}

export function getLookupBuiltin(name: string): LookupBuiltin | undefined {
  return createLookupBuiltinResolver()(name)
}
