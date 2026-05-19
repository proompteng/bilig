import { makeCellKey } from '../../workbook-store.js'
import type { EngineCellMutationRef } from '../../cell-mutations-at.js'
import type { PreparedCellAddress } from '../runtime-state.js'

interface OperationPreparedCellWorkbook {
  readonly cellKeyToIndex: {
    get(key: number): number | undefined
  }
  readonly getCellIndex: (sheetName: string, address: string) => number | undefined
  readonly getSheet: (sheetName: string) => { readonly id: number } | undefined
  readonly getSheetById: (sheetId: number) => { readonly id: number } | undefined
  readonly getOrCreateSheet: (sheetName: string) => { readonly id: number } | undefined
  readonly ensureCellAt: (sheetId: number, row: number, col: number) => { readonly cellIndex: number }
}

interface OperationMutationCellResolverWorkbook {
  readonly cellStore: {
    readonly sheetIds: ArrayLike<number | undefined>
    readonly rows: ArrayLike<number | undefined>
    readonly cols: ArrayLike<number | undefined>
  }
  readonly cellKeyToIndex: {
    get(key: number): number | undefined
  }
  readonly getSheetById: (sheetId: number) =>
    | {
        readonly structureVersion: number
        readonly logical?: {
          getCellVisiblePosition(cellIndex: number): { readonly row: number; readonly col: number } | undefined
          cellIdentityMatchesVisiblePosition?: (cellIndex: number, row: number, col: number) => boolean
        }
      }
    | undefined
}

interface OperationSheetNameResolverWorkbook {
  readonly getSheetById: (sheetId: number) => { readonly name: string } | undefined
}

export interface OperationPreparedCellTracker {
  readonly invalidateSheetName: (sheetName: string) => void
  readonly getExistingCellIndex: (sheetName: string, address: string, preparedCellAddress: PreparedCellAddress | null) => number | undefined
  readonly ensureCellTracked: (sheetName: string, address: string, preparedCellAddress: PreparedCellAddress | null) => number
}

export interface OperationSheetNameResolver {
  readonly resolve: (sheetId: number) => string
}

export function createOperationPreparedCellTracker(args: {
  readonly workbook: OperationPreparedCellWorkbook
  readonly ensureCellTracked: (sheetName: string, address: string) => number
}): OperationPreparedCellTracker {
  const preparedSheetIdByName = new Map<string, number>()

  const resolveSheetId = (sheetName: string, create: boolean): number | undefined => {
    const cachedSheetId = preparedSheetIdByName.get(sheetName)
    if (cachedSheetId !== undefined) {
      if (args.workbook.getSheetById(cachedSheetId)) {
        return cachedSheetId
      }
      preparedSheetIdByName.delete(sheetName)
    }
    const sheet = create ? args.workbook.getOrCreateSheet(sheetName) : args.workbook.getSheet(sheetName)
    if (!sheet) {
      return undefined
    }
    preparedSheetIdByName.set(sheetName, sheet.id)
    return sheet.id
  }

  return {
    invalidateSheetName(sheetName) {
      preparedSheetIdByName.delete(sheetName)
    },
    getExistingCellIndex(sheetName, address, preparedCellAddress) {
      if (!preparedCellAddress) {
        return args.workbook.getCellIndex(sheetName, address)
      }
      const sheetId = resolveSheetId(sheetName, false)
      if (sheetId === undefined) {
        return undefined
      }
      return args.workbook.cellKeyToIndex.get(makeCellKey(sheetId, preparedCellAddress.row, preparedCellAddress.col))
    },
    ensureCellTracked(sheetName, address, preparedCellAddress) {
      if (!preparedCellAddress) {
        return args.ensureCellTracked(sheetName, address)
      }
      const sheetId = resolveSheetId(sheetName, true)
      if (sheetId === undefined) {
        throw new Error(`Unknown sheet: ${sheetName}`)
      }
      return args.workbook.ensureCellAt(sheetId, preparedCellAddress.row, preparedCellAddress.col).cellIndex
    },
  }
}

export function createOperationSheetNameResolver(workbook: OperationSheetNameResolverWorkbook): OperationSheetNameResolver {
  const sheetNameById = new Map<number, string>()

  return {
    resolve(sheetId) {
      const cached = sheetNameById.get(sheetId)
      if (cached !== undefined) {
        return cached
      }
      const sheet = workbook.getSheetById(sheetId)
      if (!sheet) {
        throw new Error(`Unknown sheet id: ${sheetId}`)
      }
      sheetNameById.set(sheetId, sheet.name)
      return sheet.name
    },
  }
}

export function resolveOperationExistingMutationCellIndex(
  workbook: OperationMutationCellResolverWorkbook,
  ref: EngineCellMutationRef,
): number | undefined {
  const candidate = ref.cellIndex
  const { sheetId, mutation } = ref
  if (candidate !== undefined && workbook.cellStore.sheetIds[candidate] === sheetId) {
    const sheet = workbook.getSheetById(sheetId)
    if (sheet?.structureVersion === 1) {
      if (
        workbook.cellStore.rows[candidate] === mutation.row &&
        workbook.cellStore.cols[candidate] === mutation.col &&
        sheet.logical?.cellIdentityMatchesVisiblePosition?.(candidate, mutation.row, mutation.col) === true
      ) {
        return candidate
      }
    } else {
      const position = sheet?.logical?.getCellVisiblePosition(candidate)
      if (position?.row === mutation.row && position.col === mutation.col) {
        return candidate
      }
    }
  }
  return workbook.cellKeyToIndex.get(makeCellKey(sheetId, mutation.row, mutation.col))
}
