import type { EngineCellMutationRef, EngineExistingNumericCellMutationResult } from '@bilig/core/headless-runtime'
import { ValueTag, type CellValue, type LiteralInput } from '@bilig/protocol'
import type { WorkPaperDetailedEvent } from './work-paper-emitter.js'
import type { WorkPaperCellAddress, WorkPaperCellChange, WorkPaperChange } from './work-paper-types.js'
import type { TrackedEngineEvent } from './tracked-engine-event-refs.js'
import { scalarValueFromLiteral } from './work-paper-runtime-helpers.js'

export const TRUSTED_TRACKED_PHYSICAL_SHEET_ID_PROPERTY = '__biligTrackedPhysicalSheetId'
export const TRUSTED_TRACKED_PHYSICAL_SORTED_SPLIT_PROPERTY = '__biligTrackedPhysicalSortedSliceSplit'
export const TINY_TRACKED_CHANGE_LIMIT = 4

export interface WorkPaperTrackedRuntimeCellStore {
  readonly tags: ArrayLike<ValueTag | undefined>
  readonly numbers: ArrayLike<number | undefined>
  readonly errors: ArrayLike<number | undefined>
  readonly getValue: (cellIndex: number, readString: (stringId: number) => string) => CellValue
}

export interface WorkPaperTrackedStringPool {
  readonly get: (id: number) => string
}

export interface WorkPaperTrackedSheetRecord {
  readonly name: string
  readonly structureVersion: number
}

export interface WorkPaperTrackedWorkbookAccess {
  readonly cellStore: WorkPaperTrackedAddressedCellStore
  readonly getCellPosition: (cellIndex: number) => { readonly row: number; readonly col: number } | undefined
  readonly getSheetById: (sheetId: number) => WorkPaperTrackedSheetRecord | undefined
  readonly getSheetNameById: (sheetId: number) => string | undefined
}

export interface WorkPaperTrackedAddressedCellStore extends WorkPaperTrackedRuntimeCellStore {
  readonly sheetIds: ArrayLike<number | undefined>
  readonly rows: ArrayLike<number | undefined>
  readonly cols: ArrayLike<number | undefined>
}

export interface DirectSingleLiteralTrackedChangeExpectation {
  readonly address: WorkPaperCellAddress
  readonly cellIndex?: number
  readonly isPhysicalSheet: boolean
  readonly sheetName: string
  readonly value: LiteralInput
}

export type QueuedEvent = Extract<
  WorkPaperDetailedEvent,
  {
    eventName: 'sheetAdded' | 'sheetRemoved' | 'sheetRenamed' | 'namedExpressionAdded' | 'namedExpressionRemoved'
  }
>

export function readTrustedPhysicalTrackedChangeMetadata(
  changedCellIndices: Uint32Array,
): { readonly trustedPhysicalSheetId: number; readonly trustedSortedSliceSplit?: number } | undefined {
  const trustedPhysicalSheetId = Reflect.get(changedCellIndices, TRUSTED_TRACKED_PHYSICAL_SHEET_ID_PROPERTY)
  if (typeof trustedPhysicalSheetId !== 'number' || !Number.isInteger(trustedPhysicalSheetId) || trustedPhysicalSheetId < 0) {
    return undefined
  }
  const trustedSortedSliceSplit = Reflect.get(changedCellIndices, TRUSTED_TRACKED_PHYSICAL_SORTED_SPLIT_PROPERTY)
  return typeof trustedSortedSliceSplit === 'number' && Number.isInteger(trustedSortedSliceSplit) && trustedSortedSliceSplit > 0
    ? { trustedPhysicalSheetId, trustedSortedSliceSplit }
    : { trustedPhysicalSheetId }
}

export function trackedEventHasNoValueChanges(event: TrackedEngineEvent): boolean {
  return (
    event.invalidation !== 'full' &&
    event.changedCellIndices.length === 0 &&
    !(event.patches?.some((patch) => patch.kind === 'cell') ?? false)
  )
}

export function countPotentialNewTrackedCellMutations(refs: readonly EngineCellMutationRef[]): number {
  let count = 0
  for (let index = 0; index < refs.length; index += 1) {
    const ref = refs[index]
    if (ref && ref.cellIndex === undefined && ref.mutation.kind !== 'clearCell') {
      count += 1
    }
  }
  return count
}

export function canSkipDimensionUpdateAfterLiteralMutation(
  refs: readonly EngineCellMutationRef[],
  potentialNewCells: number | undefined,
): boolean {
  if (potentialNewCells !== 0 || refs.length === 0) {
    return false
  }
  for (let index = 0; index < refs.length; index += 1) {
    const ref = refs[index]
    if (ref?.cellIndex === undefined || ref.mutation.kind !== 'setCellValue' || ref.mutation.value === null) {
      return false
    }
  }
  return true
}

export function readTrackedRuntimeCellValue(
  cellStore: WorkPaperTrackedRuntimeCellStore,
  cellIndex: number,
  strings: WorkPaperTrackedStringPool,
): CellValue {
  const tag = cellStore.tags[cellIndex] ?? ValueTag.Empty
  switch (tag) {
    case ValueTag.Number:
      return { tag: ValueTag.Number, value: cellStore.numbers[cellIndex] ?? 0 }
    case ValueTag.Boolean:
      return { tag: ValueTag.Boolean, value: (cellStore.numbers[cellIndex] ?? 0) !== 0 }
    case ValueTag.String:
      return cellStore.getValue(cellIndex, (stringId) => strings.get(stringId))
    case ValueTag.Error:
      return { tag: ValueTag.Error, code: cellStore.errors[cellIndex]! }
    case ValueTag.Empty:
    default:
      return { tag: ValueTag.Empty }
  }
}

export function readCompactSecondChangedValue(result: EngineExistingNumericCellMutationResult): CellValue | undefined {
  return (result as EngineExistingNumericCellMutationResult & { readonly secondChangedValue?: CellValue }).secondChangedValue
}

export function existingNumericMutationChangedCellCount(result: EngineExistingNumericCellMutationResult): number {
  return result.changedCellIndices?.length ?? result.changedCellCount ?? 0
}

export function existingNumericMutationChangedCellAt(result: EngineExistingNumericCellMutationResult, index: number): number | undefined {
  if (result.changedCellIndices) {
    return result.changedCellIndices[index]
  }
  if (index === 0) {
    return result.firstChangedCellIndex
  }
  if (index === 1) {
    return result.secondChangedCellIndex
  }
  return undefined
}

export function materializeExistingNumericMutationChangedCellIndices(result: EngineExistingNumericCellMutationResult): Uint32Array {
  if (result.changedCellIndices) {
    return result.changedCellIndices
  }
  const count = existingNumericMutationChangedCellCount(result)
  const changed = new Uint32Array(count)
  for (let index = 0; index < count; index += 1) {
    changed[index] = existingNumericMutationChangedCellAt(result, index) ?? 0
  }
  return changed
}

export function trackedEventFromExistingNumericMutationResult(result: EngineExistingNumericCellMutationResult): TrackedEngineEvent {
  let sortedDisjoint = true
  let previous = -1
  let firstChangedCellIndex: number | undefined
  let lastChangedCellIndex: number | undefined
  const changedCellCount = existingNumericMutationChangedCellCount(result)
  for (let index = 0; index < changedCellCount; index += 1) {
    const cellIndex = existingNumericMutationChangedCellAt(result, index) ?? -1
    if (index === 0) {
      firstChangedCellIndex = cellIndex
    }
    if (!Number.isInteger(cellIndex) || cellIndex < 0 || cellIndex <= previous) {
      sortedDisjoint = false
    }
    previous = cellIndex
    lastChangedCellIndex = cellIndex
  }
  return {
    invalidation: 'cells',
    changedCellIndices: materializeExistingNumericMutationChangedCellIndices(result),
    changedInputCount: 1,
    explicitChangedCount: result.explicitChangedCount,
    changedCellIndicesSortedDisjoint: sortedDisjoint,
    ...(firstChangedCellIndex === undefined ? {} : { firstChangedCellIndex }),
    ...(lastChangedCellIndex === undefined ? {} : { lastChangedCellIndex }),
    hasInvalidatedRanges: false,
    hasInvalidatedRows: false,
    hasInvalidatedColumns: false,
  }
}

export function readTrackedCellChange(input: {
  readonly cellIndex: number
  readonly workbook: WorkPaperTrackedWorkbookAccess
  readonly strings: WorkPaperTrackedStringPool
  readonly trackedA1: (row: number, col: number) => string
}): WorkPaperCellChange | undefined {
  const { cellIndex, strings, trackedA1, workbook } = input
  const cellStore = workbook.cellStore
  const sheetId = cellStore.sheetIds[cellIndex]
  if (sheetId === undefined) {
    return undefined
  }
  const sheet = workbook.getSheetById(sheetId)
  const sheetName = sheet?.name ?? workbook.getSheetNameById(sheetId)
  if (sheetName === undefined) {
    return undefined
  }
  let row: number
  let col: number
  if (!sheet || sheet.structureVersion === 1) {
    const physicalRow = cellStore.rows[cellIndex]
    const physicalCol = cellStore.cols[cellIndex]
    if (physicalRow === undefined || physicalCol === undefined) {
      return undefined
    }
    row = physicalRow
    col = physicalCol
  } else {
    const position = workbook.getCellPosition(cellIndex)
    if (!position) {
      return undefined
    }
    row = position.row
    col = position.col
  }
  return {
    kind: 'cell',
    address: { sheet: sheetId, row, col },
    sheetName,
    a1: trackedA1(row, col),
    newValue: readTrackedRuntimeCellValue(cellStore, cellIndex, strings),
  }
}

export function readTinySortedPhysicalTrackedEventChanges(input: {
  readonly event: TrackedEngineEvent
  readonly workbook: WorkPaperTrackedWorkbookAccess
  readonly strings: WorkPaperTrackedStringPool
  readonly trackedA1: (row: number, col: number) => string
}): WorkPaperCellChange[] | null {
  const { event, strings, trackedA1, workbook } = input
  if (!event.changedCellIndicesSortedDisjoint) {
    return null
  }
  const cellStore = workbook.cellStore
  const firstCellIndex = event.changedCellIndices[0]
  if (firstCellIndex === undefined) {
    return []
  }
  const sheetId = cellStore.sheetIds[firstCellIndex]
  if (sheetId === undefined) {
    return []
  }
  const sheet = workbook.getSheetById(sheetId)
  if (sheet && sheet.structureVersion !== 1) {
    return null
  }
  const sheetName = sheet?.name ?? workbook.getSheetNameById(sheetId)
  if (sheetName === undefined) {
    return null
  }
  const changes: WorkPaperCellChange[] = []
  let previousRow = -1
  let previousCol = -1
  for (let index = 0; index < event.changedCellIndices.length; index += 1) {
    const cellIndex = event.changedCellIndices[index]!
    if (cellStore.sheetIds[cellIndex] !== sheetId) {
      return null
    }
    const row = cellStore.rows[cellIndex]
    const col = cellStore.cols[cellIndex]
    if (row === undefined || col === undefined) {
      return null
    }
    if (row < previousRow || (row === previousRow && col < previousCol)) {
      return null
    }
    changes.push({
      kind: 'cell',
      address: { sheet: sheetId, row, col },
      sheetName,
      a1: trackedA1(row, col),
      newValue: readTrackedRuntimeCellValue(cellStore, cellIndex, strings),
    })
    previousRow = row
    previousCol = col
  }
  return changes
}

export function tryBuildDirectSingleLiteralTrackedChange(input: {
  readonly events: readonly TrackedEngineEvent[]
  readonly expected?: DirectSingleLiteralTrackedChangeExpectation
  readonly cellStore: WorkPaperTrackedAddressedCellStore
  readonly strings: WorkPaperTrackedStringPool
  readonly trackedA1: (row: number, col: number) => string
}): WorkPaperChange[] | null {
  const { cellStore, events, expected, strings, trackedA1 } = input
  if (expected === undefined || expected.cellIndex === undefined || events.length !== 1) {
    return null
  }
  const event = events[0]!
  if (
    event.invalidation === 'full' ||
    event.patches !== undefined ||
    event.changedCellIndices.length < 1 ||
    event.changedCellIndices[0] !== expected.cellIndex ||
    event.hasInvalidatedRanges ||
    event.hasInvalidatedRows ||
    event.hasInvalidatedColumns
  ) {
    return null
  }
  const literalChange: WorkPaperCellChange = {
    kind: 'cell',
    address: { sheet: expected.address.sheet, row: expected.address.row, col: expected.address.col },
    sheetName: expected.sheetName,
    a1: trackedA1(expected.address.row, expected.address.col),
    newValue: scalarValueFromLiteral(expected.value),
  }
  if (event.changedCellIndices.length === 1) {
    return [literalChange]
  }
  if (event.changedCellIndices.length > 2) {
    if (!expected.isPhysicalSheet) {
      return null
    }
    const changes: WorkPaperCellChange[] = []
    changes.length = event.changedCellIndices.length
    changes[0] = literalChange
    let previousRow = expected.address.row
    let previousCol = expected.address.col
    for (let index = 1; index < event.changedCellIndices.length; index += 1) {
      const formulaCellIndex = event.changedCellIndices[index]!
      if (cellStore.sheetIds[formulaCellIndex] !== expected.address.sheet) {
        return null
      }
      const formulaRow = cellStore.rows[formulaCellIndex]
      const formulaCol = cellStore.cols[formulaCellIndex]
      if (
        formulaRow === undefined ||
        formulaCol === undefined ||
        formulaRow < previousRow ||
        (formulaRow === previousRow && formulaCol < previousCol)
      ) {
        return null
      }
      changes[index] = {
        kind: 'cell',
        address: { sheet: expected.address.sheet, row: formulaRow, col: formulaCol },
        sheetName: expected.sheetName,
        a1: trackedA1(formulaRow, formulaCol),
        newValue: readTrackedRuntimeCellValue(cellStore, formulaCellIndex, strings),
      }
      previousRow = formulaRow
      previousCol = formulaCol
    }
    return changes
  }
  const formulaCellIndex = event.changedCellIndices[1]!
  if (cellStore.sheetIds[formulaCellIndex] !== expected.address.sheet || !expected.isPhysicalSheet) {
    return null
  }
  const formulaRow = cellStore.rows[formulaCellIndex]
  const formulaCol = cellStore.cols[formulaCellIndex]
  if (
    formulaRow === undefined ||
    formulaCol === undefined ||
    formulaRow < expected.address.row ||
    (formulaRow === expected.address.row && formulaCol < expected.address.col)
  ) {
    return null
  }
  return [
    literalChange,
    {
      kind: 'cell',
      address: { sheet: expected.address.sheet, row: formulaRow, col: formulaCol },
      sheetName: expected.sheetName,
      a1: trackedA1(formulaRow, formulaCol),
      newValue: readTrackedRuntimeCellValue(cellStore, formulaCellIndex, strings),
    },
  ]
}

export function tryBuildDirectExistingNumericTrackedChanges(input: {
  readonly result: EngineExistingNumericCellMutationResult
  readonly address: WorkPaperCellAddress
  readonly cellIndex: number
  readonly isPhysicalSheet: boolean
  readonly sheetName: string
  readonly value: LiteralInput
  readonly cellStore: WorkPaperTrackedAddressedCellStore
  readonly strings: WorkPaperTrackedStringPool
  readonly trackedA1: (row: number, col: number) => string
  readonly orderChanges: (changes: WorkPaperCellChange[], explicitChangedCount: number | undefined) => WorkPaperChange[]
}): WorkPaperChange[] | null {
  const { address, cellIndex, cellStore, isPhysicalSheet, orderChanges, result, sheetName, strings, trackedA1, value } = input
  const changedCellCount = result.changedCellIndices?.length ?? result.changedCellCount ?? 0
  const firstChangedCellIndex = result.changedCellIndices?.[0] ?? result.firstChangedCellIndex
  if (changedCellCount < 1 || firstChangedCellIndex !== cellIndex) {
    return null
  }
  const literalChange: WorkPaperCellChange = {
    kind: 'cell',
    address: { sheet: address.sheet, row: address.row, col: address.col },
    sheetName,
    a1: trackedA1(address.row, address.col),
    newValue: scalarValueFromLiteral(value),
  }
  if (changedCellCount === 1) {
    return [literalChange]
  }
  if (!isPhysicalSheet) {
    return null
  }
  if (changedCellCount > 2) {
    const changedCellIndices = result.changedCellIndices
    if (changedCellIndices === undefined) {
      return null
    }
    const changes: WorkPaperCellChange[] = []
    changes.length = changedCellCount
    changes[0] = literalChange
    let alreadySorted = true
    let previousRow = address.row
    let previousCol = address.col
    for (let index = 1; index < changedCellCount; index += 1) {
      const changedCellIndex = changedCellIndices[index]!
      if (cellStore.sheetIds[changedCellIndex] !== address.sheet) {
        return null
      }
      const row = cellStore.rows[changedCellIndex]
      const col = cellStore.cols[changedCellIndex]
      if (row === undefined || col === undefined) {
        return null
      }
      if (row < previousRow || (row === previousRow && col < previousCol)) {
        alreadySorted = false
      }
      changes[index] = {
        kind: 'cell',
        address: { sheet: address.sheet, row, col },
        sheetName,
        a1: trackedA1(row, col),
        newValue: readTrackedRuntimeCellValue(cellStore, changedCellIndex, strings),
      }
      previousRow = row
      previousCol = col
    }
    return alreadySorted ? changes : orderChanges(changes, result.explicitChangedCount)
  }
  const formulaCellIndex = result.changedCellIndices?.[1] ?? result.secondChangedCellIndex
  if (formulaCellIndex === undefined || cellStore.sheetIds[formulaCellIndex] !== address.sheet) {
    return null
  }
  const formulaRow = result.secondChangedRow ?? cellStore.rows[formulaCellIndex]
  const formulaCol = result.secondChangedCol ?? cellStore.cols[formulaCellIndex]
  if (
    formulaRow === undefined ||
    formulaCol === undefined ||
    formulaRow < address.row ||
    (formulaRow === address.row && formulaCol < address.col)
  ) {
    return null
  }
  const secondChangedValue = readCompactSecondChangedValue(result)
  return [
    literalChange,
    {
      kind: 'cell',
      address: { sheet: address.sheet, row: formulaRow, col: formulaCol },
      sheetName,
      a1: trackedA1(formulaRow, formulaCol),
      newValue:
        secondChangedValue ??
        (result.secondChangedNumericValue === undefined
          ? readTrackedRuntimeCellValue(cellStore, formulaCellIndex, strings)
          : { tag: ValueTag.Number, value: result.secondChangedNumericValue }),
    },
  ]
}

export function withEventChanges(event: QueuedEvent, changes: WorkPaperChange[]): QueuedEvent {
  switch (event.eventName) {
    case 'sheetAdded':
      return event
    case 'sheetRemoved':
      return {
        eventName: 'sheetRemoved',
        payload: {
          ...event.payload,
          changes,
        },
      }
    case 'sheetRenamed':
      return event
    case 'namedExpressionAdded':
      return {
        eventName: 'namedExpressionAdded',
        payload: {
          ...event.payload,
          changes,
        },
      }
    case 'namedExpressionRemoved':
      return {
        eventName: 'namedExpressionRemoved',
        payload: {
          ...event.payload,
          changes,
        },
      }
  }
}
