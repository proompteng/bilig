import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import { hasLookupWildcardSyntax } from '@bilig/formula'
import type { RuntimeDirectLookupDescriptor } from '../runtime-state.js'
import type { DirectScalarCurrentOperand } from './direct-formula-index-collection.js'
import type { ExactColumnIndexService } from './exact-column-index-service.js'
import type { SortedColumnSearchService } from './sorted-column-search-service.js'
import {
  approximateRepeatedUniformLookupCurrentResult,
  approximateUniformLookupCurrentResult,
  approximateUniformLookupNumericResult,
  directLookupVersionMatches,
  exactUniformLookupCurrentResult,
  exactUniformLookupNumericResult,
  type DirectLookupSheetVersionState,
  type UniformNumericDirectLookup,
} from './direct-lookup-helpers.js'

interface OperationDirectLookupFormulaAccess {
  get(cellIndex: number): { readonly directLookup: RuntimeDirectLookupDescriptor | undefined } | undefined
}

interface OperationDirectLookupWorkbookAccess {
  readonly cellStore: {
    readonly tags: ArrayLike<ValueTag | undefined>
    readonly numbers: ArrayLike<number | undefined>
  }
  getSheetById(sheetId: number): (DirectLookupSheetVersionState & { readonly id?: number }) | undefined
}

export interface OperationDirectLookupCurrentService {
  readonly tryDirectUniformLookupCurrentResult: (formulaCellIndex: number) => DirectScalarCurrentOperand | undefined
  readonly tryDirectUniformLookupCurrentResultFromNumeric: (
    formulaCellIndex: number,
    exactLookupValue: number | undefined,
    approximateLookupValue: number | undefined,
    lookupSheetHint?: DirectLookupSheetVersionState & { readonly id?: number },
  ) => DirectScalarCurrentOperand | undefined
  readonly tryDirectUniformLookupNumericResultFromDescriptor: (
    directLookup: RuntimeDirectLookupDescriptor | undefined,
    exactLookupValue: number | undefined,
    approximateLookupValue: number | undefined,
    lookupSheetHint?: DirectLookupSheetVersionState & { readonly id?: number },
  ) => number | undefined
  readonly canEvaluateDirectUniformLookupCurrentResultFromNumeric: (
    formulaCellIndex: number,
    exactLookupValue: number | undefined,
    approximateLookupValue: number | undefined,
  ) => boolean
  readonly tryDirectApproximateLookupCurrentResultFromNumeric: (
    directLookup: Extract<RuntimeDirectLookupDescriptor, { kind: 'approximate' }>,
    lookupValue: number,
  ) => DirectScalarCurrentOperand | undefined
  readonly tryDirectExactLookupCurrentResult: (
    directLookup: Extract<RuntimeDirectLookupDescriptor, { kind: 'exact' }>,
    lookupValue: CellValue,
  ) => DirectScalarCurrentOperand | undefined
}

export function createOperationDirectLookupCurrentService(args: {
  readonly state: {
    readonly workbook: OperationDirectLookupWorkbookAccess
    readonly formulas: OperationDirectLookupFormulaAccess
  }
  readonly exactLookup: Pick<ExactColumnIndexService, 'findPreparedVectorMatch'>
  readonly sortedLookup: Pick<SortedColumnSearchService, 'findPreparedVectorMatch'>
}): OperationDirectLookupCurrentService {
  const lookupSheetForUniformLookup = (
    directLookup: UniformNumericDirectLookup,
    lookupSheetHint?: DirectLookupSheetVersionState & { readonly id?: number },
  ): DirectLookupSheetVersionState | undefined =>
    lookupSheetHint?.id === directLookup.sheetId ? lookupSheetHint : args.state.workbook.getSheetById(directLookup.sheetId)

  const tryDirectUniformLookupCurrentResult = (formulaCellIndex: number): DirectScalarCurrentOperand | undefined => {
    const formula = args.state.formulas.get(formulaCellIndex)
    const directLookup = formula?.directLookup
    if (directLookup === undefined) {
      return undefined
    }
    const cellStore = args.state.workbook.cellStore
    if (directLookup.kind === 'exact-uniform-numeric') {
      const lookupSheet = args.state.workbook.getSheetById(directLookup.sheetId)
      if (!directLookupVersionMatches(lookupSheet, directLookup)) {
        return undefined
      }
      const tag = cellStore.tags[directLookup.operandCellIndex]
      if (tag === ValueTag.Error) {
        return undefined
      }
      if (tag !== ValueTag.Number) {
        return { kind: 'error', code: ErrorCode.NA }
      }
      const lookupValue = Object.is(cellStore.numbers[directLookup.operandCellIndex] ?? 0, -0)
        ? 0
        : (cellStore.numbers[directLookup.operandCellIndex] ?? 0)
      return exactUniformLookupCurrentResult(directLookup, lookupValue)
    }
    if (directLookup.kind !== 'approximate-uniform-numeric') {
      return undefined
    }
    const lookupSheet = args.state.workbook.getSheetById(directLookup.sheetId)
    if (!directLookupVersionMatches(lookupSheet, directLookup)) {
      return undefined
    }
    const tag = cellStore.tags[directLookup.operandCellIndex]
    let lookupValue = 0
    switch (tag) {
      case undefined:
      case ValueTag.Empty:
        lookupValue = 0
        break
      case ValueTag.Number:
        lookupValue = Object.is(cellStore.numbers[directLookup.operandCellIndex] ?? 0, -0)
          ? 0
          : (cellStore.numbers[directLookup.operandCellIndex] ?? 0)
        break
      case ValueTag.Boolean:
        lookupValue = (cellStore.numbers[directLookup.operandCellIndex] ?? 0) !== 0 ? 1 : 0
        break
      case ValueTag.Error:
      case ValueTag.String:
        return undefined
    }
    return approximateUniformLookupCurrentResult(directLookup, lookupValue)
  }

  const tryDirectUniformLookupCurrentResultFromNumeric = (
    formulaCellIndex: number,
    exactLookupValue: number | undefined,
    approximateLookupValue: number | undefined,
    lookupSheetHint?: DirectLookupSheetVersionState & { readonly id?: number },
  ): DirectScalarCurrentOperand | undefined => {
    const formula = args.state.formulas.get(formulaCellIndex)
    const directLookup = formula?.directLookup
    if (directLookup === undefined) {
      return undefined
    }
    if (directLookup.kind === 'exact-uniform-numeric') {
      if (exactLookupValue === undefined) {
        return undefined
      }
      const lookupSheet = lookupSheetForUniformLookup(directLookup, lookupSheetHint)
      if (!directLookupVersionMatches(lookupSheet, directLookup)) {
        return undefined
      }
      return exactUniformLookupCurrentResult(directLookup, exactLookupValue)
    }
    if (directLookup.kind !== 'approximate-uniform-numeric' || approximateLookupValue === undefined) {
      return undefined
    }
    const lookupSheet = lookupSheetForUniformLookup(directLookup, lookupSheetHint)
    if (!directLookupVersionMatches(lookupSheet, directLookup)) {
      return undefined
    }
    return approximateUniformLookupCurrentResult(directLookup, approximateLookupValue)
  }

  const tryDirectUniformLookupNumericResultFromDescriptor = (
    directLookup: RuntimeDirectLookupDescriptor | undefined,
    exactLookupValue: number | undefined,
    approximateLookupValue: number | undefined,
    lookupSheetHint?: DirectLookupSheetVersionState & { readonly id?: number },
  ): number | undefined => {
    if (directLookup?.kind === 'exact-uniform-numeric') {
      if (exactLookupValue === undefined) {
        return undefined
      }
      const lookupSheet = lookupSheetForUniformLookup(directLookup, lookupSheetHint)
      return directLookupVersionMatches(lookupSheet, directLookup)
        ? exactUniformLookupNumericResult(directLookup, exactLookupValue)
        : undefined
    }
    if (directLookup?.kind === 'approximate-uniform-numeric') {
      if (approximateLookupValue === undefined) {
        return undefined
      }
      const lookupSheet = lookupSheetForUniformLookup(directLookup, lookupSheetHint)
      return directLookupVersionMatches(lookupSheet, directLookup)
        ? approximateUniformLookupNumericResult(directLookup, approximateLookupValue)
        : undefined
    }
    return undefined
  }

  const canEvaluateDirectUniformLookupCurrentResultFromNumeric = (
    formulaCellIndex: number,
    exactLookupValue: number | undefined,
    approximateLookupValue: number | undefined,
  ): boolean => {
    const directLookup = args.state.formulas.get(formulaCellIndex)?.directLookup
    if (directLookup?.kind === 'exact-uniform-numeric') {
      const lookupSheet = args.state.workbook.getSheetById(directLookup.sheetId)
      if (!directLookupVersionMatches(lookupSheet, directLookup)) {
        return false
      }
      return exactLookupValue !== undefined
    }
    if (directLookup?.kind !== 'approximate-uniform-numeric') {
      return false
    }
    const lookupSheet = args.state.workbook.getSheetById(directLookup.sheetId)
    if (!directLookupVersionMatches(lookupSheet, directLookup)) {
      return false
    }
    return approximateLookupValue !== undefined
  }

  const tryDirectApproximateLookupCurrentResultFromNumeric = (
    directLookup: Extract<RuntimeDirectLookupDescriptor, { kind: 'approximate' }>,
    lookupValue: number,
  ): DirectScalarCurrentOperand | undefined => {
    const prepared = directLookup.prepared
    const values = prepared.numericValues
    if (
      values !== undefined &&
      prepared.comparableKind === 'numeric' &&
      (directLookup.matchMode === 1 ? prepared.sortedAscending : prepared.sortedDescending)
    ) {
      let position: number | undefined
      let handledUniform = false
      if (prepared.uniformStart !== undefined && prepared.uniformStep !== undefined) {
        const lastValue = prepared.uniformStart + prepared.uniformStep * (values.length - 1)
        if (directLookup.matchMode === 1 && prepared.uniformStep > 0) {
          handledUniform = true
          if (lookupValue < prepared.uniformStart) {
            position = undefined
          } else if (lookupValue >= lastValue) {
            position = values.length
          } else {
            position = Math.min(values.length, Math.max(1, Math.floor((lookupValue - prepared.uniformStart) / prepared.uniformStep) + 1))
          }
        } else if (directLookup.matchMode === -1 && prepared.uniformStep < 0) {
          handledUniform = true
          if (lookupValue > prepared.uniformStart) {
            position = undefined
          } else if (lookupValue <= lastValue) {
            position = values.length
          } else {
            position = Math.min(values.length, Math.max(1, Math.floor((prepared.uniformStart - lookupValue) / -prepared.uniformStep) + 1))
          }
        }
      }
      if (!handledUniform) {
        const repeatedUniformResult = approximateRepeatedUniformLookupCurrentResult(prepared, directLookup.matchMode, lookupValue)
        if (repeatedUniformResult?.kind === 'number') {
          position = repeatedUniformResult.value
        } else if (repeatedUniformResult?.kind === 'error') {
          position = undefined
        } else {
          let low = 0
          let high = values.length - 1
          let best = -1
          while (low <= high) {
            const mid = (low + high) >> 1
            const midValue = values[mid]!
            if (midValue === lookupValue) {
              best = mid
              low = mid + 1
            } else if (directLookup.matchMode === 1 ? midValue < lookupValue : midValue > lookupValue) {
              best = mid
              low = mid + 1
            } else {
              high = mid - 1
            }
          }
          position = best === -1 ? undefined : best + 1
        }
      }
      return position === undefined ? { kind: 'error', code: ErrorCode.NA } : { kind: 'number', value: position }
    }
    if (prepared.comparableKind === 'numeric' && (directLookup.matchMode === 1 ? prepared.sortedAscending : prepared.sortedDescending)) {
      const repeatedUniformResult = approximateRepeatedUniformLookupCurrentResult(prepared, directLookup.matchMode, lookupValue)
      if (repeatedUniformResult !== undefined) {
        return repeatedUniformResult
      }
    }
    const result = args.sortedLookup.findPreparedVectorMatch({
      lookupValue: { tag: ValueTag.Number, value: lookupValue },
      prepared: directLookup.prepared,
      matchMode: directLookup.matchMode,
    })
    if (!result.handled) {
      return undefined
    }
    return result.position === undefined ? { kind: 'error', code: ErrorCode.NA } : { kind: 'number', value: result.position }
  }

  const tryDirectExactLookupCurrentResult = (
    directLookup: Extract<RuntimeDirectLookupDescriptor, { kind: 'exact' }>,
    lookupValue: CellValue,
  ): DirectScalarCurrentOperand | undefined => {
    if (hasLookupWildcardSyntax(lookupValue)) {
      return undefined
    }
    const result = args.exactLookup.findPreparedVectorMatch({
      lookupValue,
      prepared: directLookup.prepared,
      searchMode: directLookup.searchMode,
    })
    if (!result.handled) {
      return undefined
    }
    return result.position === undefined ? { kind: 'error', code: ErrorCode.NA } : { kind: 'number', value: result.position }
  }

  return {
    tryDirectUniformLookupCurrentResult,
    tryDirectUniformLookupCurrentResultFromNumeric,
    tryDirectUniformLookupNumericResultFromDescriptor,
    canEvaluateDirectUniformLookupCurrentResultFromNumeric,
    tryDirectApproximateLookupCurrentResultFromNumeric,
    tryDirectExactLookupCurrentResult,
  }
}
