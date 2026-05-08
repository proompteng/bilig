import type { resolveRuntimeDirectLookupBinding } from '../direct-vector-lookup.js'
import {
  type MaterializedDependencies,
  type RuntimeDirectAggregateDescriptor,
  type RuntimeDirectScalarDescriptor,
  type RuntimeDirectScalarOperand,
  UNRESOLVED_WASM_OPERAND,
} from '../runtime-state.js'
import { makeCellEntity, makeRangeEntity } from '../../entity-ids.js'
import { tryParseDependencyCellAddress, tryParseDependencyRangeAddress } from './formula-binding-direct-scalar.js'
import type { ParsedCompiledFormula } from './formula-binding-direct-descriptors.js'
import {
  collectDynamicIndexDependencyPlan,
  formulaMayNeedDynamicIndexDependencyPlan,
} from './formula-binding-dynamic-index-dependencies.js'
import { getFormulaBindingReverseEdgeSlice } from './formula-binding-reverse-edges.js'
import type { CreateEngineFormulaBindingServiceArgs } from './formula-binding-service-types.js'

const EMPTY_DEPENDENCY_BUFFER = new Uint32Array(0)

interface FormulaBindingDependencyMaterializerArgs {
  readonly serviceArgs: CreateEngineFormulaBindingServiceArgs
  readonly hasFormulaColumnMembers: (sheetId: number, col: number) => boolean
  readonly isFormulaCell: (cellIndex: number) => boolean
  readonly ensureDependencyBuildCapacity: (
    cellCapacity: number,
    dependencyCapacity: number,
    symbolicRefCapacity?: number,
    symbolicRangeCapacity?: number,
  ) => void
}

export interface FormulaBindingDependencyMaterializer {
  readonly materializeDependencies: (
    currentSheetName: string,
    compiled: ParsedCompiledFormula,
    directAggregate: RuntimeDirectAggregateDescriptor | undefined,
    directLookupBinding: ReturnType<typeof resolveRuntimeDirectLookupBinding> | undefined,
  ) => MaterializedDependencies
  readonly materializeDirectScalarDependencies: (
    compiled: ParsedCompiledFormula,
    directScalar: RuntimeDirectScalarDescriptor | undefined,
  ) => MaterializedDependencies | undefined
  readonly materializeDirectAggregateDependencies: (
    compiled: ParsedCompiledFormula,
    directAggregate: RuntimeDirectAggregateDescriptor | undefined,
  ) => MaterializedDependencies | undefined
}

export function createFormulaBindingDependencyMaterializer(
  materializerArgs: FormulaBindingDependencyMaterializerArgs,
): FormulaBindingDependencyMaterializer {
  const args = materializerArgs.serviceArgs
  const { ensureDependencyBuildCapacity } = materializerArgs
  const isFormulaCell = materializerArgs.isFormulaCell

  const materializeDependencies = (
    currentSheetName: string,
    compiled: ParsedCompiledFormula,
    directAggregate: RuntimeDirectAggregateDescriptor | undefined,
    directLookupBinding: ReturnType<typeof resolveRuntimeDirectLookupBinding> | undefined,
  ): MaterializedDependencies => {
    const currentSheetId = args.state.workbook.getSheet(currentSheetName)?.id
    const deps = compiled.deps
    const parsedCellDeps = compiled.parsedDeps
    const dynamicIndexDependencyPlan = formulaMayNeedDynamicIndexDependencyPlan(compiled)
      ? collectDynamicIndexDependencyPlan({
          compiled,
          ownerSheetName: currentSheetName,
          workbook: args.state.workbook,
          strings: args.state.strings,
          getFormulaAst: (sheetName, address) => {
            const cellIndex = args.state.workbook.getCellIndex(sheetName, address)
            const formula = cellIndex === undefined ? undefined : args.state.formulas.get(cellIndex)
            if (cellIndex === undefined || !formula) {
              return undefined
            }
            const position = args.state.workbook.getCellPosition(cellIndex)
            const parsed = position ? undefined : tryParseDependencyCellAddress(address, sheetName)
            return {
              sheetName,
              address,
              row: position?.row ?? parsed?.row ?? args.state.workbook.cellStore.rows[cellIndex] ?? 0,
              col: position?.col ?? parsed?.col ?? args.state.workbook.cellStore.cols[cellIndex] ?? 0,
              ast: formula.compiled.optimizedAst,
            }
          },
        })
      : undefined
    if (
      dynamicIndexDependencyPlan === undefined &&
      compiled.symbolicRanges.length === 0 &&
      parsedCellDeps !== undefined &&
      parsedCellDeps.length === deps.length &&
      parsedCellDeps.length > 0 &&
      parsedCellDeps.length <= 2 &&
      parsedCellDeps.every((dependency) => dependency?.kind === 'cell')
    ) {
      ensureDependencyBuildCapacity(args.state.workbook.cellStore.size + 1, parsedCellDeps.length + 1, compiled.symbolicRefs.length + 1, 1)
      let dependencyIndexCount = 0
      let dependencyEntityCount = 0
      for (let depIndex = 0; depIndex < parsedCellDeps.length; depIndex += 1) {
        const parsedDep = parsedCellDeps[depIndex]!
        if (parsedDep.sheetName && !args.state.workbook.getSheet(parsedDep.sheetName)) {
          continue
        }
        const cellIndex =
          parsedDep.sheetName === undefined && parsedDep.row !== undefined && parsedDep.col !== undefined && currentSheetId !== undefined
            ? args.ensureCellTrackedByCoords(currentSheetId, parsedDep.row, parsedDep.col)
            : args.ensureCellTracked(parsedDep.sheetName ?? currentSheetName, parsedDep.address)
        let seen = false
        for (let existingIndex = 0; existingIndex < dependencyIndexCount; existingIndex += 1) {
          if (args.getDependencyBuildCells()[existingIndex] === cellIndex) {
            seen = true
            break
          }
        }
        if (!seen) {
          args.getDependencyBuildCells()[dependencyIndexCount] = cellIndex
          dependencyIndexCount += 1
        }
        args.getDependencyBuildEntities()[dependencyEntityCount] = makeCellEntity(cellIndex)
        dependencyEntityCount += 1
      }
      return {
        dependencyIndices: args.getDependencyBuildCells().slice(0, dependencyIndexCount),
        dependencyEntities: args.getDependencyBuildEntities().slice(0, dependencyEntityCount),
        rangeDependencies: args.getDependencyBuildRanges().slice(0, 0),
        graphRangeDependencies: args.getDependencyBuildRanges().slice(0, 0),
        symbolicRangeIndices: args.getSymbolicRangeBindings(),
        symbolicRangeCount: 0,
        newRangeIndices: args.getDependencyBuildNewRanges(),
        newRangeCount: 0,
      }
    }

    const extraDynamicCellDependencyCount = dynamicIndexDependencyPlan?.selectedCells.length ?? 0
    const sheetNamesInSpan = (startSheetName: string, endSheetName: string): string[] | undefined => {
      const sheetNames = [...args.state.workbook.sheetsByName.values()]
        .toSorted((left, right) => left.order - right.order)
        .map((sheet) => sheet.name)
      const startIndex = sheetNames.indexOf(startSheetName)
      const endIndex = sheetNames.indexOf(endSheetName)
      if (startIndex === -1 || endIndex === -1 || startIndex > endIndex) {
        return undefined
      }
      return sheetNames.slice(startIndex, endIndex + 1)
    }
    const sheetRangeDependencyCapacity =
      compiled.parsedDeps?.reduce((count, dependency) => {
        if (dependency.kind !== 'range' || dependency.sheetEndName === undefined) {
          return count
        }
        const sheetNames = sheetNamesInSpan(dependency.sheetName ?? currentSheetName, dependency.sheetEndName)
        return count + Math.max((sheetNames?.length ?? 1) - 1, 0)
      }, 0) ?? 0

    ensureDependencyBuildCapacity(
      args.state.workbook.cellStore.size + 1,
      deps.length + extraDynamicCellDependencyCount + sheetRangeDependencyCapacity + 1,
      compiled.symbolicRefs.length + 1,
      compiled.symbolicRanges.length + 1,
    )
    let epoch = args.getDependencyBuildEpoch() + 1
    if (epoch === 0xffff_ffff) {
      epoch = 1
      args.getDependencyBuildSeen().fill(0)
    }
    args.setDependencyBuildEpoch(epoch)

    let dependencyIndexCount = 0
    let dependencyEntityCount = 0
    let rangeDependencyCount = 0
    let newRangeCount = 0
    const graphRangeDependencies: number[] = []
    const symbolicRangeIndexByAddress =
      compiled.symbolicRanges.length > 0 ? new Map(compiled.symbolicRanges.map((range, index) => [range, index])) : undefined
    args.getSymbolicRangeBindings().fill(UNRESOLVED_WASM_OPERAND, 0, compiled.symbolicRanges.length)

    const appendCellDependency = (cellIndex: number): void => {
      if (args.getDependencyBuildSeen()[cellIndex] !== epoch) {
        args.getDependencyBuildSeen()[cellIndex] = epoch
        args.getDependencyBuildCells()[dependencyIndexCount] = cellIndex
        dependencyIndexCount += 1
      }
      args.getDependencyBuildEntities()[dependencyEntityCount] = makeCellEntity(cellIndex)
      dependencyEntityCount += 1
    }

    const rangeDependencySourceEdgesNeedSync = (rangeIndex: number): boolean => {
      const rangeEntity = makeRangeEntity(rangeIndex)
      const dependencySources = args.state.ranges.getDependencySourceEntities(rangeIndex)
      for (let sourceIndex = 0; sourceIndex < dependencySources.length; sourceIndex += 1) {
        const dependencyEntity = dependencySources[sourceIndex]!
        const slice = getFormulaBindingReverseEdgeSlice(args.reverseState, dependencyEntity)
        if (!slice) {
          return true
        }
        if (!args.edgeArena.readView(slice).includes(rangeEntity)) {
          return true
        }
      }
      return false
    }

    const appendRuntimeRangeDependency = (rangeIndex: number, needsSourceEdgeSync: boolean): void => {
      args.getDependencyBuildRanges()[rangeDependencyCount] = rangeIndex
      rangeDependencyCount += 1
      if (needsSourceEdgeSync) {
        args.getDependencyBuildNewRanges()[newRangeCount] = rangeIndex
        newRangeCount += 1
      }
    }

    const appendGraphRangeDependency = (rangeIndex: number): void => {
      args.getDependencyBuildEntities()[dependencyEntityCount] = makeRangeEntity(rangeIndex)
      dependencyEntityCount += 1
      graphRangeDependencies.push(rangeIndex)
    }

    const appendParsedRangeDependencyForSheet = (
      parsedRangeDep: Extract<NonNullable<ParsedCompiledFormula['parsedDeps']>[number], { kind: 'range' }>,
      sheetName: string,
    ): void => {
      const range = tryParseDependencyRangeAddress(`${parsedRangeDep.startAddress}:${parsedRangeDep.endAddress}`, sheetName)
      if (!range) {
        return
      }
      const sheet = args.state.workbook.getSheet(sheetName)
      if (!sheet) {
        return
      }
      const registered = args.state.ranges.intern(sheet.id, range, {
        ensureCell: (sheetId, row, col) => args.ensureCellTrackedByCoords(sheetId, row, col),
        forEachSheetCell: (sheetId, fn) => args.forEachSheetCell(sheetId, fn),
        isFormulaCell,
      })
      const needsSourceEdgeSync = registered.materialized || rangeDependencySourceEdgesNeedSync(registered.rangeIndex)
      appendRuntimeRangeDependency(registered.rangeIndex, needsSourceEdgeSync)
      appendGraphRangeDependency(registered.rangeIndex)
      const memberIndices = args.state.ranges.getFormulaMembersView(registered.rangeIndex)
      for (let memberIndex = 0; memberIndex < memberIndices.length; memberIndex += 1) {
        const cellIndex = memberIndices[memberIndex]!
        if (args.getDependencyBuildSeen()[cellIndex] === epoch) {
          continue
        }
        args.getDependencyBuildSeen()[cellIndex] = epoch
        args.getDependencyBuildCells()[dependencyIndexCount] = cellIndex
        dependencyIndexCount += 1
      }
    }

    dynamicIndexDependencyPlan?.selectedCells.forEach((cell) => {
      const sheet = args.state.workbook.getSheet(cell.sheetName)
      if (!sheet) {
        return
      }
      appendCellDependency(args.ensureCellTrackedByCoords(sheet.id, cell.row, cell.col))
    })

    for (let depIndex = 0; depIndex < deps.length; depIndex += 1) {
      const dep = deps[depIndex]!
      const parsedDep = compiled.parsedDeps?.[depIndex]
      if (parsedDep?.kind === 'range' && parsedDep.sheetEndName !== undefined) {
        const sheetNames = sheetNamesInSpan(parsedDep.sheetName ?? currentSheetName, parsedDep.sheetEndName)
        sheetNames?.forEach((sheetName) => {
          appendParsedRangeDependencyForSheet(parsedDep, sheetName)
        })
        continue
      }
      if (parsedDep?.kind === 'cell') {
        if (parsedDep.sheetName && !args.state.workbook.getSheet(parsedDep.sheetName)) {
          continue
        }
        const cellIndex =
          parsedDep.sheetName === undefined && parsedDep.row !== undefined && parsedDep.col !== undefined && currentSheetId !== undefined
            ? args.ensureCellTrackedByCoords(currentSheetId, parsedDep.row, parsedDep.col)
            : args.ensureCellTracked(parsedDep.sheetName ?? currentSheetName, parsedDep.address)
        appendCellDependency(cellIndex)
        continue
      }
      if (dep.includes(':')) {
        const parsedRangeDep =
          parsedDep?.kind === 'range' ? parsedDep : compiled.parsedSymbolicRanges?.find((range) => range.address === dep)
        const range =
          parsedRangeDep === undefined
            ? tryParseDependencyRangeAddress(dep, currentSheetName)
            : tryParseDependencyRangeAddress(
                `${parsedRangeDep.startAddress}:${parsedRangeDep.endAddress}`,
                parsedRangeDep.sheetName ?? currentSheetName,
              )
        if (!range) {
          continue
        }
        const sheetName = range.sheetName ?? currentSheetName
        const isDirectLookupColumn =
          directLookupBinding !== undefined &&
          range.kind === 'cells' &&
          range.start.col === range.end.col &&
          sheetName === directLookupBinding.lookupSheetName &&
          range.start.col === directLookupBinding.col &&
          range.start.row === directLookupBinding.rowStart &&
          range.end.row === directLookupBinding.rowEnd
        if (isDirectLookupColumn) {
          const sheet = args.state.workbook.getSheet(sheetName)
          if (sheet) {
            for (let row = range.start.row; row <= range.end.row; row += 1) {
              const cellIndex = sheet.grid.get(row, range.start.col)
              if (cellIndex === -1) {
                continue
              }
              if (!isFormulaCell(cellIndex)) {
                continue
              }
              if (args.getDependencyBuildSeen()[cellIndex] !== epoch) {
                args.getDependencyBuildSeen()[cellIndex] = epoch
                args.getDependencyBuildCells()[dependencyIndexCount] = cellIndex
                dependencyIndexCount += 1
              }
              args.getDependencyBuildEntities()[dependencyEntityCount] = makeCellEntity(cellIndex)
              dependencyEntityCount += 1
            }
          }
          continue
        }
        const symbolicRangeIndex = symbolicRangeIndexByAddress?.get(dep) ?? -1
        if (range.sheetName && !args.state.workbook.getSheet(sheetName)) {
          continue
        }
        const sheet = args.state.workbook.getSheet(sheetName)
        if (!sheet) {
          continue
        }
        const shouldCompactDynamicIndexRange = dynamicIndexDependencyPlan?.compactedRangeDependencies.has(dep) === true
        const compactDirectAggregateRange =
          directAggregate !== undefined &&
          range.kind === 'cells' &&
          range.start.col === directAggregate.col &&
          range.end.col === directAggregate.col &&
          range.start.row === directAggregate.rowStart &&
          range.end.row === directAggregate.rowEnd &&
          sheetName === directAggregate.sheetName
        if (compactDirectAggregateRange) {
          if (!materializerArgs.hasFormulaColumnMembers(sheet.id, range.start.col)) {
            continue
          }
          for (let row = range.start.row; row <= range.end.row; row += 1) {
            const cellIndex = sheet.grid.get(row, range.start.col)
            if (cellIndex === -1) {
              continue
            }
            if (!isFormulaCell(cellIndex)) {
              continue
            }
            if (args.getDependencyBuildSeen()[cellIndex] !== epoch) {
              args.getDependencyBuildSeen()[cellIndex] = epoch
              args.getDependencyBuildCells()[dependencyIndexCount] = cellIndex
              dependencyIndexCount += 1
            }
            args.getDependencyBuildEntities()[dependencyEntityCount] = makeCellEntity(cellIndex)
            dependencyEntityCount += 1
          }
          continue
        }
        const registered = args.state.ranges.intern(sheet.id, range, {
          ensureCell: (sheetId, row, col) => args.ensureCellTrackedByCoords(sheetId, row, col),
          forEachSheetCell: (sheetId, fn) => args.forEachSheetCell(sheetId, fn),
          isFormulaCell,
        })
        if (symbolicRangeIndex !== -1) {
          args.getSymbolicRangeBindings()[symbolicRangeIndex] = registered.rangeIndex
        }
        const needsSourceEdgeSync = registered.materialized || rangeDependencySourceEdgesNeedSync(registered.rangeIndex)
        if (shouldCompactDynamicIndexRange) {
          appendRuntimeRangeDependency(registered.rangeIndex, needsSourceEdgeSync)
          continue
        }
        appendRuntimeRangeDependency(registered.rangeIndex, needsSourceEdgeSync)
        appendGraphRangeDependency(registered.rangeIndex)
        const memberIndices = args.state.ranges.getFormulaMembersView(registered.rangeIndex)
        for (let memberIndex = 0; memberIndex < memberIndices.length; memberIndex += 1) {
          const cellIndex = memberIndices[memberIndex]!
          if (args.getDependencyBuildSeen()[cellIndex] === epoch) {
            continue
          }
          args.getDependencyBuildSeen()[cellIndex] = epoch
          args.getDependencyBuildCells()[dependencyIndexCount] = cellIndex
          dependencyIndexCount += 1
        }
        continue
      }
      const parsed = tryParseDependencyCellAddress(dep, currentSheetName)
      if (!parsed) {
        continue
      }
      const sheetName = parsed.sheetName ?? currentSheetName
      if (parsed.sheetName && !args.state.workbook.getSheet(sheetName)) {
        continue
      }
      const cellIndex = args.ensureCellTracked(sheetName, parsed.text)
      appendCellDependency(cellIndex)
    }
    return {
      dependencyIndices: args.getDependencyBuildCells().slice(0, dependencyIndexCount),
      dependencyEntities: args.getDependencyBuildEntities().slice(0, dependencyEntityCount),
      rangeDependencies: args.getDependencyBuildRanges().slice(0, rangeDependencyCount),
      graphRangeDependencies: Uint32Array.from(graphRangeDependencies),
      symbolicRangeIndices: args.getSymbolicRangeBindings(),
      symbolicRangeCount: compiled.symbolicRanges.length,
      newRangeIndices: args.getDependencyBuildNewRanges(),
      newRangeCount,
    }
  }

  const materializeDirectScalarDependencies = (
    compiled: ParsedCompiledFormula,
    directScalar: RuntimeDirectScalarDescriptor | undefined,
  ): MaterializedDependencies | undefined => {
    if (
      directScalar === undefined ||
      compiled.symbolicRanges.length !== 0 ||
      compiled.symbolicNames.length !== 0 ||
      compiled.symbolicTables.length !== 0 ||
      compiled.symbolicSpills.length !== 0 ||
      compiled.parsedSymbolicRefs?.some((ref) => ref.sheetName !== undefined || ref.explicitSheet === true) === true
    ) {
      return undefined
    }
    const dependencyIndices = new Uint32Array(Math.min(compiled.symbolicRefs.length, 2))
    const dependencyEntities = new Uint32Array(compiled.symbolicRefs.length)
    let dependencyIndexCount = 0
    let dependencyEntityCount = 0
    const appendOperand = (operand: RuntimeDirectScalarOperand): boolean => {
      if (operand.kind === 'literal-number') {
        return true
      }
      if (operand.kind === 'error') {
        return false
      }
      const cellIndex = operand.cellIndex
      let seen = false
      for (let existingIndex = 0; existingIndex < dependencyIndexCount; existingIndex += 1) {
        if (dependencyIndices[existingIndex] === cellIndex) {
          seen = true
          break
        }
      }
      if (!seen) {
        dependencyIndices[dependencyIndexCount] = cellIndex
        dependencyIndexCount += 1
      }
      dependencyEntities[dependencyEntityCount] = makeCellEntity(cellIndex)
      dependencyEntityCount += 1
      return true
    }
    const matched =
      directScalar.kind === 'abs'
        ? appendOperand(directScalar.operand)
        : appendOperand(directScalar.left) && appendOperand(directScalar.right)
    if (!matched || dependencyEntityCount !== compiled.symbolicRefs.length) {
      return undefined
    }
    return {
      dependencyIndices:
        dependencyIndexCount === dependencyIndices.length ? dependencyIndices : dependencyIndices.subarray(0, dependencyIndexCount),
      dependencyEntities,
      rangeDependencies: EMPTY_DEPENDENCY_BUFFER,
      graphRangeDependencies: EMPTY_DEPENDENCY_BUFFER,
      symbolicRangeIndices: EMPTY_DEPENDENCY_BUFFER,
      symbolicRangeCount: 0,
      newRangeIndices: EMPTY_DEPENDENCY_BUFFER,
      newRangeCount: 0,
    }
  }

  const materializeDirectAggregateDependencies = (
    compiled: ParsedCompiledFormula,
    directAggregate: RuntimeDirectAggregateDescriptor | undefined,
  ): MaterializedDependencies | undefined => {
    if (
      directAggregate === undefined ||
      compiled.deps.length !== 1 ||
      compiled.symbolicRefs.length !== 0 ||
      compiled.symbolicRanges.length !== 1 ||
      compiled.symbolicNames.length !== 0 ||
      compiled.symbolicTables.length !== 0 ||
      compiled.symbolicSpills.length !== 0
    ) {
      return undefined
    }
    const sheet = args.state.workbook.getSheet(directAggregate.sheetName)
    if (!sheet) {
      return undefined
    }
    for (let col = directAggregate.col; col <= directAggregate.colEnd; col += 1) {
      if (materializerArgs.hasFormulaColumnMembers(sheet.id, col)) {
        return undefined
      }
    }
    ensureDependencyBuildCapacity(args.state.workbook.cellStore.size + 1, 1, 1, compiled.symbolicRanges.length + 1)
    args.getSymbolicRangeBindings().fill(UNRESOLVED_WASM_OPERAND, 0, compiled.symbolicRanges.length)
    return {
      dependencyIndices: args.getDependencyBuildCells().slice(0, 0),
      dependencyEntities: args.getDependencyBuildEntities().slice(0, 0),
      rangeDependencies: args.getDependencyBuildRanges().slice(0, 0),
      graphRangeDependencies: args.getDependencyBuildRanges().slice(0, 0),
      symbolicRangeIndices: args.getSymbolicRangeBindings(),
      symbolicRangeCount: compiled.symbolicRanges.length,
      newRangeIndices: args.getDependencyBuildNewRanges(),
      newRangeCount: 0,
    }
  }

  return {
    materializeDependencies,
    materializeDirectScalarDependencies,
    materializeDirectAggregateDependencies,
  }
}
