import { Effect } from 'effect'
import type { EdgeArena, EdgeSlice } from '../../edge-arena.js'
import {
  makeExactLookupColumnEntity,
  makeSortedLookupColumnEntity,
  entityPayload,
  isExactLookupColumnEntity,
  isRangeEntity,
  isSortedLookupColumnEntity,
} from '../../entity-ids.js'
import { growUint32 } from '../../engine-buffer-utils.js'
import type { EngineRuntimeState, U32 } from '../runtime-state.js'
import { EngineTraversalError } from '../errors.js'
import type { RegionGraph } from '../../deps/region-graph.js'

export interface EngineTraversalService {
  readonly getEntityDependents: (entityId: number) => Effect.Effect<Uint32Array, EngineTraversalError>
  readonly collectFormulaDependents: (entityId: number) => Effect.Effect<Uint32Array, EngineTraversalError>
  readonly forEachFormulaDependencyCell: (
    cellIndex: number,
    fn: (dependencyCellIndex: number) => void,
  ) => Effect.Effect<void, EngineTraversalError>
  readonly forEachSheetCell: (
    sheetId: number,
    fn: (cellIndex: number, row: number, col: number) => void,
  ) => Effect.Effect<void, EngineTraversalError>
  readonly getEntityDependentsNow: (entityId: number) => Uint32Array
  readonly collectFormulaDependentsNow: (entityId: number) => Uint32Array
  readonly forEachFormulaDependencyCellNow: (cellIndex: number, fn: (dependencyCellIndex: number) => void) => void
  readonly forEachSheetCellNow: (sheetId: number, fn: (cellIndex: number, row: number, col: number) => void) => void
}

function traversalErrorMessage(message: string, cause: unknown): string {
  return cause instanceof Error && cause.message.length > 0 ? cause.message : message
}

export function createEngineTraversalService(args: {
  readonly state: Pick<EngineRuntimeState, 'workbook' | 'formulas' | 'ranges'>
  readonly regionGraph: Pick<RegionGraph, 'getRegion' | 'collectFormulaDependentsForCell'>
  readonly edgeArena: EdgeArena
  readonly reverseState: {
    reverseCellEdges: Array<EdgeSlice | undefined>
    reverseRangeEdges: Array<EdgeSlice | undefined>
    reverseExactLookupColumnEdges: Map<number, EdgeSlice>
    reverseSortedLookupColumnEdges: Map<number, EdgeSlice>
  }
}): EngineTraversalService {
  let topoFormulaBuffer: U32 = new Uint32Array(128)
  let topoEntityQueue: U32 = new Uint32Array(128)
  let topoFormulaSeenEpoch = 1
  let topoRangeSeenEpoch = 1
  let topoExactLookupSeenEpoch = 1
  let topoSortedLookupSeenEpoch = 1
  let topoFormulaSeen: U32 = new Uint32Array(128)
  let topoRangeSeen: U32 = new Uint32Array(128)
  const topoExactLookupSeen = new Map<number, number>()
  const topoSortedLookupSeen = new Map<number, number>()

  const ensureTraversalScratchCapacity = (cellSize: number, entitySize: number, rangeSize: number): void => {
    if (cellSize > topoFormulaBuffer.length) {
      topoFormulaBuffer = growUint32(topoFormulaBuffer, cellSize)
    }
    if (cellSize > topoFormulaSeen.length) {
      topoFormulaSeen = growUint32(topoFormulaSeen, cellSize)
    }
    if (entitySize > topoEntityQueue.length) {
      topoEntityQueue = growUint32(topoEntityQueue, entitySize)
    }
    if (rangeSize > topoRangeSeen.length) {
      topoRangeSeen = growUint32(topoRangeSeen, rangeSize)
    }
  }

  const ensureEntityQueueCapacity = (size: number): void => {
    if (size <= topoEntityQueue.length) {
      return
    }
    let capacity = topoEntityQueue.length
    while (capacity < size) {
      capacity *= 2
    }
    topoEntityQueue = growUint32(topoEntityQueue, capacity)
  }

  const ensureFormulaBufferCapacity = (size: number): void => {
    if (size <= topoFormulaBuffer.length) {
      return
    }
    let capacity = topoFormulaBuffer.length
    while (capacity < size) {
      capacity *= 2
    }
    topoFormulaBuffer = growUint32(topoFormulaBuffer, capacity)
  }

  const getReverseEdgeSlice = (entityId: number): EdgeSlice | undefined => {
    if (isRangeEntity(entityId)) {
      return args.reverseState.reverseRangeEdges[entityPayload(entityId)]
    }
    if (isExactLookupColumnEntity(entityId)) {
      return args.reverseState.reverseExactLookupColumnEdges.get(entityPayload(entityId))
    }
    if (isSortedLookupColumnEntity(entityId)) {
      return args.reverseState.reverseSortedLookupColumnEdges.get(entityPayload(entityId))
    }
    return args.reverseState.reverseCellEdges[entityPayload(entityId)]
  }

  const getEntityDependentsNow = (entityId: number): Uint32Array =>
    args.edgeArena.readView(getReverseEdgeSlice(entityId) ?? args.edgeArena.empty())

  const forEachFormulaDependencyCellNow = (cellIndex: number, fn: (dependencyCellIndex: number) => void): void => {
    const formula = args.state.formulas.get(cellIndex)
    if (!formula) {
      return
    }
    const seen = new Set<number>()
    const push = (dependencyCellIndex: number): void => {
      if (seen.has(dependencyCellIndex)) {
        return
      }
      seen.add(dependencyCellIndex)
      fn(dependencyCellIndex)
    }
    for (let index = 0; index < formula.dependencyIndices.length; index += 1) {
      push(formula.dependencyIndices[index]!)
    }
    for (let index = 0; index < formula.rangeDependencies.length; index += 1) {
      const members = args.state.ranges.expandToCells(formula.rangeDependencies[index]!)
      for (let memberIndex = 0; memberIndex < members.length; memberIndex += 1) {
        push(members[memberIndex]!)
      }
    }
    const pushDirectRegion = (regionId: number | undefined): void => {
      if (regionId === undefined) {
        return
      }
      const region = args.regionGraph.getRegion(regionId)
      if (!region) {
        return
      }
      const sheet = args.state.workbook.getSheet(region.sheetName)
      if (!sheet) {
        return
      }
      for (let row = region.rowStart; row <= region.rowEnd; row += 1) {
        const dependencyCellIndex = sheet.grid.get(row, region.col)
        if (dependencyCellIndex !== -1) {
          push(dependencyCellIndex)
        }
      }
    }
    const pushDirectLookupRange = (range: { sheetName: string; rowStart: number; rowEnd: number; col: number } | undefined): void => {
      if (!range) {
        return
      }
      const sheet = args.state.workbook.getSheet(range.sheetName)
      if (!sheet) {
        return
      }
      for (let row = range.rowStart; row <= range.rowEnd; row += 1) {
        const dependencyCellIndex = sheet.grid.get(row, range.col)
        if (dependencyCellIndex !== -1) {
          push(dependencyCellIndex)
        }
      }
    }
    pushDirectRegion(formula.directAggregate?.regionId)
    if (formula.directCriteria) {
      pushDirectRegion(formula.directCriteria.aggregateRange?.regionId)
      for (let index = 0; index < formula.directCriteria.criteriaPairs.length; index += 1) {
        const pair = formula.directCriteria.criteriaPairs[index]!
        pushDirectRegion(pair.range.regionId)
        if (pair.criterion.kind === 'cell') {
          push(pair.criterion.cellIndex)
        }
      }
    }
    const directLookup = formula.directLookup
    if (directLookup) {
      if (directLookup.kind === 'exact' || directLookup.kind === 'approximate') {
        pushDirectLookupRange({
          sheetName: directLookup.prepared.sheetName,
          rowStart: directLookup.prepared.rowStart,
          rowEnd: directLookup.prepared.rowEnd,
          col: directLookup.prepared.col,
        })
        push(directLookup.operandCellIndex)
      } else {
        pushDirectLookupRange({
          sheetName: directLookup.sheetName,
          rowStart: directLookup.rowStart,
          rowEnd: directLookup.rowEnd,
          col: directLookup.col,
        })
        push(directLookup.operandCellIndex)
      }
    }
  }

  const forEachSheetCellNow = (sheetId: number, fn: (cellIndex: number, row: number, col: number) => void): void => {
    const sheet = args.state.workbook.getSheetById(sheetId)
    if (!sheet) {
      return
    }
    sheet.grid.forEachCellEntry((cellIndex, row, col) => {
      fn(cellIndex, row, col)
    })
  }

  const collectFormulaDependentsNow = (entityId: number): Uint32Array => {
    ensureTraversalScratchCapacity(
      Math.max(args.state.workbook.cellStore.size + 1, 1),
      Math.max(args.state.workbook.cellStore.size + args.state.ranges.size + 1, 1),
      Math.max(args.state.ranges.size + 1, 1),
    )

    topoFormulaSeenEpoch += 1
    if (topoFormulaSeenEpoch === 0xffff_ffff) {
      topoFormulaSeenEpoch = 1
      topoFormulaSeen.fill(0)
    }
    topoRangeSeenEpoch += 1
    if (topoRangeSeenEpoch === 0xffff_ffff) {
      topoRangeSeenEpoch = 1
      topoRangeSeen.fill(0)
    }
    topoExactLookupSeenEpoch += 1
    if (topoExactLookupSeenEpoch === 0xffff_ffff) {
      topoExactLookupSeenEpoch = 1
      topoExactLookupSeen.clear()
    }
    topoSortedLookupSeenEpoch += 1
    if (topoSortedLookupSeenEpoch === 0xffff_ffff) {
      topoSortedLookupSeenEpoch = 1
      topoSortedLookupSeen.clear()
    }

    let entityQueueLength = 1
    let formulaCount = 0
    topoEntityQueue[0] = entityId

    for (let queueIndex = 0; queueIndex < entityQueueLength; queueIndex += 1) {
      const currentEntity = topoEntityQueue[queueIndex]!
      if (!isRangeEntity(currentEntity) && !isExactLookupColumnEntity(currentEntity) && !isSortedLookupColumnEntity(currentEntity)) {
        const cellIndex = entityPayload(currentEntity)
        const sheetId = args.state.workbook.cellStore.sheetIds[cellIndex]
        const position = args.state.workbook.getCellPosition(cellIndex)
        if (sheetId !== undefined && position) {
          const regionDependents = args.regionGraph.collectFormulaDependentsForCell(sheetId, position.row, position.col)
          for (let index = 0; index < regionDependents.length; index += 1) {
            const formulaCellIndex = regionDependents[index]!
            if (topoFormulaSeen[formulaCellIndex] === topoFormulaSeenEpoch) {
              continue
            }
            topoFormulaSeen[formulaCellIndex] = topoFormulaSeenEpoch
            ensureFormulaBufferCapacity(formulaCount + 1)
            topoFormulaBuffer[formulaCount] = formulaCellIndex
            formulaCount += 1
          }
        }
        if (sheetId !== undefined && position) {
          const exactLookupEntity = makeExactLookupColumnEntity(sheetId, position.col)
          const sortedLookupEntity = makeSortedLookupColumnEntity(sheetId, position.col)
          ensureEntityQueueCapacity(entityQueueLength + 2)
          topoEntityQueue[entityQueueLength] = exactLookupEntity
          entityQueueLength += 1
          topoEntityQueue[entityQueueLength] = sortedLookupEntity
          entityQueueLength += 1
        }
      }
      const dependents = getEntityDependentsNow(currentEntity)
      for (let index = 0; index < dependents.length; index += 1) {
        const dependent = dependents[index]!
        if (!(isRangeEntity(dependent) || isExactLookupColumnEntity(dependent) || isSortedLookupColumnEntity(dependent))) {
          const formulaCellIndex = entityPayload(dependent)
          if (topoFormulaSeen[formulaCellIndex] === topoFormulaSeenEpoch) {
            continue
          }
          topoFormulaSeen[formulaCellIndex] = topoFormulaSeenEpoch
          ensureFormulaBufferCapacity(formulaCount + 1)
          topoFormulaBuffer[formulaCount] = formulaCellIndex
          formulaCount += 1
          continue
        }
        if (isRangeEntity(dependent)) {
          const rangeIndex = entityPayload(dependent)
          if (topoRangeSeen[rangeIndex] === topoRangeSeenEpoch) {
            continue
          }
          topoRangeSeen[rangeIndex] = topoRangeSeenEpoch
        } else if (isExactLookupColumnEntity(dependent)) {
          const lookupColumnPayload = entityPayload(dependent)
          if (topoExactLookupSeen.get(lookupColumnPayload) === topoExactLookupSeenEpoch) {
            continue
          }
          topoExactLookupSeen.set(lookupColumnPayload, topoExactLookupSeenEpoch)
        } else {
          const lookupColumnPayload = entityPayload(dependent)
          if (topoSortedLookupSeen.get(lookupColumnPayload) === topoSortedLookupSeenEpoch) {
            continue
          }
          topoSortedLookupSeen.set(lookupColumnPayload, topoSortedLookupSeenEpoch)
        }
        ensureEntityQueueCapacity(entityQueueLength + 1)
        topoEntityQueue[entityQueueLength] = dependent
        entityQueueLength += 1
      }
    }

    return topoFormulaBuffer.subarray(0, formulaCount)
  }

  return {
    getEntityDependents(entityId) {
      return Effect.try({
        try: () => Uint32Array.from(getEntityDependentsNow(entityId)),
        catch: (cause) =>
          new EngineTraversalError({
            message: traversalErrorMessage('Failed to read entity dependents', cause),
            cause,
          }),
      })
    },
    collectFormulaDependents(entityId) {
      return Effect.try({
        try: () => Uint32Array.from(collectFormulaDependentsNow(entityId)),
        catch: (cause) =>
          new EngineTraversalError({
            message: traversalErrorMessage('Failed to collect formula dependents', cause),
            cause,
          }),
      })
    },
    forEachFormulaDependencyCell(cellIndex, fn) {
      return Effect.try({
        try: () => {
          forEachFormulaDependencyCellNow(cellIndex, fn)
        },
        catch: (cause) =>
          new EngineTraversalError({
            message: traversalErrorMessage('Failed to iterate formula dependencies', cause),
            cause,
          }),
      })
    },
    forEachSheetCell(sheetId, fn) {
      return Effect.try({
        try: () => {
          forEachSheetCellNow(sheetId, fn)
        },
        catch: (cause) =>
          new EngineTraversalError({
            message: traversalErrorMessage('Failed to iterate sheet cells', cause),
            cause,
          }),
      })
    },
    getEntityDependentsNow,
    collectFormulaDependentsNow,
    forEachFormulaDependencyCellNow,
    forEachSheetCellNow,
  }
}
