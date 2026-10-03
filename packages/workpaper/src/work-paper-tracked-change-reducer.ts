import { makeCellKey } from '@bilig/core/headless-runtime'
import { ValueTag, type CellValue } from '@bilig/protocol'
import { orderWorkPaperCellChanges } from './change-order.js'
import type { TrackedEngineEvent, TrackedPatch } from './tracked-engine-event-refs.js'
import { TINY_TRACKED_CHANGE_LIMIT } from './work-paper-tracked-event-helpers.js'
import type { WorkPaperCellChange, WorkPaperChange } from './work-paper-types.js'
import type { SheetStateSnapshot, VisibilitySnapshot } from './work-paper-visibility-snapshot.js'

export interface MaterializedTrackedEventChanges {
  readonly changes: WorkPaperCellChange[]
  readonly canReusePublicChanges: boolean
  readonly ordered: boolean
}

export interface MaterializedTrackedEventSources {
  readonly changes: WorkPaperCellChange[]
  readonly ordered: boolean
}

export interface WorkPaperTrackedChangeSheetRecord {
  readonly id: number
  readonly order: number
}

interface WorkPaperTrackedChangeReducerHooks {
  readonly listSheets: () => readonly WorkPaperTrackedChangeSheetRecord[]
  readonly materializeTrackedEventChanges: (event: TrackedEngineEvent, lazy?: boolean) => MaterializedTrackedEventChanges
  readonly materializeTrackedEventSources?: (
    events: readonly TrackedEngineEvent[],
    options: { readonly preferLazyPublicChanges?: boolean },
  ) => MaterializedTrackedEventSources | null
  readonly readSingleTrackedCellChange: (cellIndex: number) => WorkPaperCellChange | undefined
  readonly readTinySortedPhysicalTrackedEventChanges: (event: TrackedEngineEvent) => WorkPaperCellChange[] | null
  readonly sheetOrder: (sheetId: number) => number
}

export interface TryReadTinyTrackedEventChangesWithoutVisibilityInput extends WorkPaperTrackedChangeReducerHooks {
  readonly event: TrackedEngineEvent
}

export interface ComputeWorkPaperTrackedCellChangesInput extends WorkPaperTrackedChangeReducerHooks {
  readonly beforeVisibility: VisibilitySnapshot
  readonly events: readonly TrackedEngineEvent[]
  readonly preferLazyPublicChanges?: boolean
  readonly updateVisibility?: boolean
}

export function tryReadTinyTrackedEventChangesWithoutVisibility(
  input: TryReadTinyTrackedEventChangesWithoutVisibilityInput,
): WorkPaperChange[] | null {
  const event = input.event
  if (
    event.patches !== undefined &&
    event.invalidation !== 'full' &&
    event.patches.length <= TINY_TRACKED_CHANGE_LIMIT &&
    !event.hasInvalidatedRanges &&
    !event.hasInvalidatedRows &&
    !event.hasInvalidatedColumns
  ) {
    return tryReadTinyPatchTrackedEventChanges(input, event.patches, event.explicitChangedCount)
  }
  if (
    event.invalidation === 'full' ||
    event.patches !== undefined ||
    event.changedCellIndices.length > TINY_TRACKED_CHANGE_LIMIT ||
    event.hasInvalidatedRanges ||
    event.hasInvalidatedRows ||
    event.hasInvalidatedColumns
  ) {
    return null
  }
  if (event.changedCellIndices.length === 0) {
    return []
  }
  const sortedPhysicalChanges = input.readTinySortedPhysicalTrackedEventChanges(event)
  if (sortedPhysicalChanges) {
    return sortedPhysicalChanges
  }
  return (
    readSmallTrackedEventChanges({
      event,
      listSheets: input.listSheets,
      readSingleTrackedCellChange: input.readSingleTrackedCellChange,
      sheetOrder: input.sheetOrder,
      updateVisibility: false,
    })?.changes ?? null
  )
}

export function computeWorkPaperTrackedCellChangesFromEvents(
  input: ComputeWorkPaperTrackedCellChangesInput,
): { changes: WorkPaperChange[]; nextVisibility: VisibilitySnapshot } | null {
  const updateVisibility = input.updateVisibility ?? true
  if (input.events.some((event) => event.invalidation === 'full')) {
    return null
  }

  const nextVisibility = input.beforeVisibility
  const sheetOrders = new Map<number, number>()
  const sheetOrderFor = (sheetId: number): number => {
    const existing = sheetOrders.get(sheetId)
    if (existing !== undefined) {
      return existing
    }
    const order = input.sheetOrder(sheetId)
    sheetOrders.set(sheetId, order)
    return order
  }
  const ensureMutableSheet = (sheetId: number, sheetName: string): SheetStateSnapshot => {
    const existing = nextVisibility.get(sheetId)
    if (existing) {
      return existing
    }
    const created: SheetStateSnapshot = {
      sheetId,
      sheetName,
      order: input.sheetOrder(sheetId),
      cells: new Map<number, CellValue>(),
    }
    nextVisibility.set(sheetId, created)
    return created
  }

  if (input.events.length === 1) {
    const event = input.events[0]!
    if (!input.preferLazyPublicChanges) {
      const smallChanges = readSmallTrackedEventChanges({
        event,
        ensureMutableSheet,
        listSheets: input.listSheets,
        readSingleTrackedCellChange: input.readSingleTrackedCellChange,
        sheetOrder: sheetOrderFor,
        updateVisibility,
      })
      if (smallChanges) {
        return { changes: smallChanges.changes, nextVisibility }
      }
    }
    const materializedEventChanges = input.materializeTrackedEventChanges(event, !updateVisibility)
    const eventChanges = materializedEventChanges.changes
    if (!updateVisibility && materializedEventChanges.canReusePublicChanges && materializedEventChanges.ordered) {
      return {
        changes: eventChanges,
        nextVisibility,
      }
    }
    const directChanges = reduceSingleEventChanges({
      event,
      eventChanges,
      ensureMutableSheet,
      listSheets: input.listSheets,
      materializedEventChanges,
      sheetOrder: sheetOrderFor,
      updateVisibility,
    })
    if (directChanges) {
      return { changes: directChanges, nextVisibility }
    }
  }

  const materializedSources = updateVisibility
    ? null
    : (input.materializeTrackedEventSources?.(
        input.events,
        input.preferLazyPublicChanges === undefined ? {} : { preferLazyPublicChanges: input.preferLazyPublicChanges },
      ) ?? null)
  if (materializedSources) {
    return {
      changes: materializedSources.ordered
        ? materializedSources.changes
        : orderWorkPaperCellChanges(materializedSources.changes, input.listSheets()),
      nextVisibility,
    }
  }

  const latestChangesByKey = new Map<number, WorkPaperCellChange>()
  for (const event of input.events) {
    const eventChanges = input.materializeTrackedEventChanges(event).changes
    for (let index = 0; index < eventChanges.length; index += 1) {
      const change = eventChanges[index]!
      const cellKey = makeCellKey(change.address.sheet, change.address.row, change.address.col)
      latestChangesByKey.delete(cellKey)
      latestChangesByKey.set(cellKey, toPublicCellChange(change))
    }
  }
  if (updateVisibility) {
    for (const change of latestChangesByKey.values()) {
      applyTrackedChangeToVisibility(ensureMutableSheet(change.address.sheet, change.sheetName), change)
    }
  }
  const directChanges = [...latestChangesByKey.values()]
  return {
    changes: orderWorkPaperCellChanges(directChanges, input.listSheets()),
    nextVisibility,
  }
}

function tryReadTinyPatchTrackedEventChanges(
  input: WorkPaperTrackedChangeReducerHooks,
  patches: readonly TrackedPatch[],
  explicitChangedCount: number | undefined,
): WorkPaperChange[] | null {
  const changes: WorkPaperCellChange[] = []
  let alreadySorted = true
  let previousSheetId = -1
  let previousSheetOrder = -1
  let previousRow = -1
  let previousCol = -1
  for (let index = 0; index < patches.length; index += 1) {
    const patch = patches[index]
    if (!patch || patch.kind !== 'cell') {
      return null
    }
    const sheetOrder = patch.address.sheet === previousSheetId ? previousSheetOrder : input.sheetOrder(patch.address.sheet)
    if (
      sheetOrder < previousSheetOrder ||
      (sheetOrder === previousSheetOrder &&
        (patch.address.row < previousRow || (patch.address.row === previousRow && patch.address.col < previousCol)))
    ) {
      alreadySorted = false
    }
    changes.push(toPublicCellChange(patch))
    previousSheetId = patch.address.sheet
    previousSheetOrder = sheetOrder
    previousRow = patch.address.row
    previousCol = patch.address.col
  }
  return alreadySorted ? changes : orderWorkPaperCellChanges(changes, input.listSheets(), explicitChangedCount)
}

function readSmallTrackedEventChanges(input: {
  readonly ensureMutableSheet?: (sheetId: number, sheetName: string) => SheetStateSnapshot
  readonly event: TrackedEngineEvent
  readonly listSheets: () => readonly WorkPaperTrackedChangeSheetRecord[]
  readonly readSingleTrackedCellChange: (cellIndex: number) => WorkPaperCellChange | undefined
  readonly sheetOrder: (sheetId: number) => number
  readonly updateVisibility: boolean
}): { readonly changes: WorkPaperChange[] } | null {
  const event = input.event
  if (
    event.invalidation === 'full' ||
    event.patches !== undefined ||
    event.changedCellIndices.length > TINY_TRACKED_CHANGE_LIMIT ||
    event.hasInvalidatedRanges ||
    event.hasInvalidatedRows ||
    event.hasInvalidatedColumns
  ) {
    return null
  }
  if (event.changedCellIndices.length === 0) {
    return { changes: [] }
  }
  const changes: WorkPaperCellChange[] = []
  const cellKeys: number[] = []
  let alreadySorted = true
  let previousSheetOrder = -1
  let previousRow = -1
  let previousCol = -1
  for (let index = 0; index < event.changedCellIndices.length; index += 1) {
    const change = input.readSingleTrackedCellChange(event.changedCellIndices[index]!)
    if (!change) {
      continue
    }
    const cellKey = makeCellKey(change.address.sheet, change.address.row, change.address.col)
    for (let priorIndex = 0; priorIndex < cellKeys.length; priorIndex += 1) {
      if (cellKeys[priorIndex] === cellKey) {
        return null
      }
    }
    cellKeys.push(cellKey)
    const sheet = input.updateVisibility ? input.ensureMutableSheet?.(change.address.sheet, change.sheetName) : undefined
    const sheetOrder = sheet?.order ?? input.sheetOrder(change.address.sheet)
    if (
      sheetOrder < previousSheetOrder ||
      (sheetOrder === previousSheetOrder &&
        (change.address.row < previousRow || (change.address.row === previousRow && change.address.col < previousCol)))
    ) {
      alreadySorted = false
    }
    if (sheet) {
      applyTrackedChangeToVisibility(sheet, change)
    }
    changes.push(change)
    previousSheetOrder = sheetOrder
    previousRow = change.address.row
    previousCol = change.address.col
  }
  return {
    changes: alreadySorted ? changes : orderWorkPaperCellChanges(changes, input.listSheets(), event.explicitChangedCount),
  }
}

function reduceSingleEventChanges(input: {
  readonly ensureMutableSheet: (sheetId: number, sheetName: string) => SheetStateSnapshot
  readonly event: TrackedEngineEvent
  readonly eventChanges: readonly WorkPaperCellChange[]
  readonly listSheets: () => readonly WorkPaperTrackedChangeSheetRecord[]
  readonly materializedEventChanges: MaterializedTrackedEventChanges
  readonly sheetOrder: (sheetId: number) => number
  readonly updateVisibility: boolean
}): WorkPaperChange[] | null {
  const directChanges: WorkPaperCellChange[] = []
  const seenCellKeys = input.eventChanges.length > 4 && input.eventChanges.length <= 64 ? new Set<number>() : undefined
  const smallCellKeys: number[] | undefined = input.eventChanges.length > 1 && input.eventChanges.length <= 4 ? [] : undefined
  let hasDuplicateCellKey = false
  let alreadySorted = true
  let previousSheetOrder = -1
  let previousRow = -1
  let previousCol = -1
  for (let index = 0; index < input.eventChanges.length; index += 1) {
    const change = input.eventChanges[index]!
    const cellKey = makeCellKey(change.address.sheet, change.address.row, change.address.col)
    if (seenCellKeys) {
      if (seenCellKeys.has(cellKey)) {
        hasDuplicateCellKey = true
        break
      }
      seenCellKeys.add(cellKey)
    } else if (smallCellKeys) {
      for (let priorIndex = 0; priorIndex < index; priorIndex += 1) {
        if (smallCellKeys[priorIndex] === cellKey) {
          hasDuplicateCellKey = true
          break
        }
      }
      if (hasDuplicateCellKey) {
        break
      }
      smallCellKeys[index] = cellKey
    }
    const sheet = input.updateVisibility ? input.ensureMutableSheet(change.address.sheet, change.sheetName) : undefined
    const sheetOrder = sheet?.order ?? input.sheetOrder(change.address.sheet)
    if (
      sheetOrder < previousSheetOrder ||
      (sheetOrder === previousSheetOrder &&
        (change.address.row < previousRow || (change.address.row === previousRow && change.address.col < previousCol)))
    ) {
      alreadySorted = false
    }
    if (sheet) {
      applyTrackedChangeToVisibility(sheet, change)
    }
    directChanges[index] = input.materializedEventChanges.canReusePublicChanges ? change : toPublicCellChange(change)
    previousSheetOrder = sheetOrder
    previousRow = change.address.row
    previousCol = change.address.col
  }
  if (hasDuplicateCellKey) {
    return null
  }
  return alreadySorted ? directChanges : orderWorkPaperCellChanges(directChanges, input.listSheets(), input.event.explicitChangedCount)
}

function applyTrackedChangeToVisibility(sheet: SheetStateSnapshot, change: WorkPaperCellChange): void {
  const cellKey = makeCellKey(change.address.sheet, change.address.row, change.address.col)
  if (change.newValue.tag === ValueTag.Empty) {
    sheet.cells.delete(cellKey)
  } else {
    sheet.cells.set(cellKey, change.newValue)
  }
}

function toPublicCellChange(change: WorkPaperCellChange): WorkPaperCellChange {
  return {
    kind: 'cell',
    address: change.address,
    sheetName: change.sheetName,
    a1: change.a1,
    newValue: change.newValue,
  }
}
