import { Effect } from 'effect'
import { ValueTag, type SheetFormatRangeSnapshot, type SheetStyleRangeSnapshot } from '@bilig/protocol'
import {
  type CompiledFormula,
  formatAddress,
  parseCellAddress,
  rewriteAddressForStructuralTransform,
  rewriteCompiledFormulaForStructuralTransform,
  rewriteFormulaForStructuralTransform,
  rewriteRangeForStructuralTransform,
  type StructuralAxisTransform,
} from '@bilig/formula'
import type { EngineOp } from '@bilig/workbook-domain'
import { CellFlags } from '../../cell-store.js'
import { emptyValue } from '../../engine-value-utils.js'
import { mapStructuralAxisIndex, mapStructuralBoundary, structuralTransformForOp } from '../../engine-structural-utils.js'
import type { StructuralTransaction } from '../structural-transaction.js'
import type { FormulaTable } from '../../formula-table.js'
import {
  rewriteTemplateForStructuralTransform,
  retargetStructurallyRewrittenTemplateInstance,
  type StructurallyRewrittenTemplate,
} from '../../formula/structural-retargeting.js'
import type { RuntimeFormula } from '../runtime-state.js'
import { getRuntimeFormulaSource, getRuntimeFormulaStructuralCompiled } from '../runtime-formula-source.js'
import { EngineStructureError } from '../errors.js'
import { normalizeDefinedName, type WorkbookPivotRecord, type WorkbookStore } from '../../workbook-store.js'
import type { RangeRegistry } from '../../range-registry.js'
import { addEngineCounter, type EngineCounters } from '../../perf/engine-counters.js'

type StructuralAxisOp = Extract<
  EngineOp,
  {
    kind: 'insertRows' | 'deleteRows' | 'moveRows' | 'insertColumns' | 'deleteColumns' | 'moveColumns'
  }
>

interface EngineStructureState {
  readonly workbook: WorkbookStore
  readonly formulas: FormulaTable<RuntimeFormula>
  readonly ranges: RangeRegistry
  readonly pivotOutputOwners: Map<number, string>
  readonly counters?: EngineCounters
}

interface StructuralFormulaRebindInput {
  readonly cellIndex: number
  readonly ownerSheetName: string
  readonly ownerRow: number
  readonly ownerCol: number
  readonly source: string
  readonly compiled?: CompiledFormula
  readonly templateId?: number
  readonly preservesBinding?: boolean
  readonly preservesValue?: boolean
}

interface StructuralFormulaRebindResolution {
  readonly inputs: StructuralFormulaRebindInput[]
  readonly preservedCellIndices: readonly number[]
}

function dependencyTouchesSheet(dependency: string, sheetName: string): boolean {
  if (!dependency.includes('!')) {
    return false
  }
  const [qualifiedSheetName] = dependency.split('!')
  return qualifiedSheetName?.replace(/^'(.*)'$/, '$1') === sheetName
}

function rangeDependencyAxisAffected(
  rangeDescriptor: { sheetId: number; row1: number; row2: number; col1: number; col2: number },
  targetSheetId: number,
  transform: StructuralAxisTransform,
): boolean {
  if (rangeDescriptor.sheetId !== targetSheetId) {
    return false
  }
  const start = transform.axis === 'row' ? rangeDescriptor.row1 : rangeDescriptor.col1
  const end = transform.axis === 'row' ? rangeDescriptor.row2 : rangeDescriptor.col2
  return !(end < transform.start || start >= transform.start + transform.count)
}

function runtimeDirectRangeAxisAffected(
  targetSheetId: number | undefined,
  targetSheetName: string,
  transform: StructuralAxisTransform,
  range: { sheetName: string; rowStart: number; rowEnd: number; col: number } | undefined,
): boolean {
  if (!range || targetSheetId === undefined || range.sheetName !== targetSheetName) {
    return false
  }
  const descriptor =
    transform.axis === 'row'
      ? {
          sheetId: targetSheetId,
          row1: range.rowStart,
          row2: range.rowEnd,
          col1: range.col,
          col2: range.col,
        }
      : {
          sheetId: targetSheetId,
          row1: range.rowStart,
          row2: range.rowEnd,
          col1: range.col,
          col2: range.col,
        }
  return rangeDependencyAxisAffected(descriptor, targetSheetId, transform)
}

function isStructurallyStableSimpleFormulaNode(node: CompiledFormula['optimizedAst']): boolean {
  switch (node.kind) {
    case 'NumberLiteral':
    case 'BooleanLiteral':
    case 'StringLiteral':
    case 'ErrorLiteral':
    case 'CellRef':
      return true
    case 'UnaryExpr':
      return isStructurallyStableSimpleFormulaNode(node.argument)
    case 'BinaryExpr':
      return isStructurallyStableSimpleFormulaNode(node.left) && isStructurallyStableSimpleFormulaNode(node.right)
    case 'NameRef':
    case 'StructuredRef':
    case 'SpillRef':
    case 'RowRef':
    case 'ColumnRef':
    case 'RangeRef':
    case 'CallExpr':
    case 'InvokeExpr':
      return false
  }
}

function arrayValuesEqual(left: ArrayLike<number>, right: ArrayLike<number>): boolean {
  if (left.length !== right.length) {
    return false
  }
  for (let index = 0; index < left.length; index += 1) {
    if (left[index] !== right[index]) {
      return false
    }
  }
  return true
}

function hasStableStructuralSymbolicRangeLayout(
  current: Pick<CompiledFormula, 'symbolicRanges' | 'parsedSymbolicRanges'>,
  next: Pick<CompiledFormula, 'symbolicRanges' | 'parsedSymbolicRanges'>,
  rangeDependencyCount: number,
): boolean {
  if (current.symbolicRanges.length !== next.symbolicRanges.length) {
    return false
  }
  if (rangeDependencyCount !== current.symbolicRanges.length) {
    return false
  }
  if (current.parsedSymbolicRanges === undefined || next.parsedSymbolicRanges === undefined) {
    return current.parsedSymbolicRanges === next.parsedSymbolicRanges
  }
  if (current.parsedSymbolicRanges.length !== next.parsedSymbolicRanges.length) {
    return false
  }
  for (let index = 0; index < current.parsedSymbolicRanges.length; index += 1) {
    if (current.parsedSymbolicRanges[index]?.refKind !== next.parsedSymbolicRanges[index]?.refKind) {
      return false
    }
  }
  return true
}

function structuralRewritePreservesValue(
  formula: RuntimeFormula,
  rewritten: { compiled: CompiledFormula; reusedProgram: boolean },
  transform: StructuralAxisTransform,
): boolean {
  return (
    transform.kind !== 'delete' &&
    rewritten.reusedProgram &&
    !rewritten.compiled.volatile &&
    formula.compiled.symbolicNames.length === 0 &&
    formula.compiled.symbolicTables.length === 0 &&
    formula.compiled.symbolicSpills.length === 0 &&
    formula.compiled.symbolicRanges.length === 0 &&
    formula.compiled.deps.every((dependency) => !dependency.includes(':')) &&
    formula.directLookup === undefined &&
    formula.directAggregate === undefined &&
    formula.directCriteria === undefined &&
    (formula.directScalar === undefined || transform.kind === 'insert') &&
    isStructurallyStableSimpleFormulaNode(rewritten.compiled.optimizedAst)
  )
}

function structuralRewritePreservesBinding(
  formula: RuntimeFormula,
  rewritten: { compiled: CompiledFormula; reusedProgram: boolean },
  allowsRangeReuse: boolean,
): boolean {
  return (
    rewritten.reusedProgram &&
    (formula.directAggregate !== undefined ||
      formula.directScalar !== undefined ||
      (allowsRangeReuse &&
        hasStableStructuralSymbolicRangeLayout(formula.compiled, rewritten.compiled, formula.rangeDependencies.length))) &&
    formula.compiled.symbolicNames.length === 0 &&
    formula.compiled.symbolicTables.length === 0 &&
    formula.compiled.symbolicSpills.length === 0 &&
    rewritten.compiled.symbolicNames.length === 0 &&
    rewritten.compiled.symbolicTables.length === 0 &&
    rewritten.compiled.symbolicSpills.length === 0 &&
    formula.directLookup === undefined &&
    formula.directCriteria === undefined &&
    arrayValuesEqual(formula.compiled.program, rewritten.compiled.program) &&
    arrayValuesEqual(formula.compiled.constants, rewritten.compiled.constants) &&
    formula.compiled.mode === rewritten.compiled.mode &&
    (formula.directAggregate !== undefined || rewritten.compiled.symbolicRanges.length === formula.rangeDependencies.length)
  )
}

function structuralDirectAggregateRewritePreservesValue(
  formula: RuntimeFormula,
  rewritten: { compiled: CompiledFormula; reusedProgram: boolean },
  transform: StructuralAxisTransform,
): boolean {
  return (
    transform.kind === 'insert' &&
    formula.directAggregate !== undefined &&
    rewritten.reusedProgram &&
    !rewritten.compiled.volatile &&
    formula.compiled.symbolicNames.length === 0 &&
    formula.compiled.symbolicTables.length === 0 &&
    formula.compiled.symbolicSpills.length === 0 &&
    formula.directLookup === undefined &&
    formula.directCriteria === undefined
  )
}

export interface EngineStructureService {
  readonly captureSheetCellState: (sheetName: string) => Effect.Effect<EngineOp[], EngineStructureError>
  readonly captureRowRangeCellState: (sheetName: string, start: number, count: number) => Effect.Effect<EngineOp[], EngineStructureError>
  readonly captureColumnRangeCellState: (sheetName: string, start: number, count: number) => Effect.Effect<EngineOp[], EngineStructureError>
  readonly materializeDeferredStructuralFormulaSources: () => Effect.Effect<void, EngineStructureError>
  readonly applyStructuralAxisOp: (op: StructuralAxisOp) => Effect.Effect<
    {
      transaction: StructuralTransaction
      changedCellIndices: number[]
      formulaCellIndices: number[]
      topologyChanged: boolean
      graphRefreshRequired: boolean
    },
    EngineStructureError
  >
}

export function createEngineStructureService(args: {
  readonly state: EngineStructureState
  readonly captureStoredCellOps: (
    cellIndex: number,
    sheetName: string,
    address: string,
    sourceSheetName?: string,
    sourceAddress?: string,
  ) => EngineOp[]
  readonly removeFormula: (cellIndex: number) => boolean
  readonly clearOwnedPivot: (pivot: WorkbookPivotRecord) => number[]
  readonly refreshRangeDependencies: (rangeIndices: readonly number[]) => void
  readonly retargetRangeDependencies: (transaction: StructuralTransaction, rangeIndices: readonly number[]) => void
  readonly rebindFormulaCells: (inputs: readonly StructuralFormulaRebindInput[]) => void
  readonly collectFormulaCellsOwnedBySheet: (sheetName: string) => readonly number[]
  readonly collectFormulaCellsReferencingSheet: (sheetName: string) => readonly number[]
  readonly collectFormulaCellsForDefinedNames: (names: readonly string[]) => readonly number[]
  readonly collectFormulaCellsForTables: (tableNames: readonly string[]) => readonly number[]
}): EngineStructureService {
  let hasDeferredStructuralFormulaSources = false

  const shouldCaptureStoredCell = (cellIndex: number): boolean => {
    const value = args.state.workbook.cellStore.getValue(cellIndex, () => '')
    const explicitFormat = args.state.workbook.getCellFormat(cellIndex)
    const flags = args.state.workbook.cellStore.flags[cellIndex] ?? 0
    const formula = args.state.formulas.get(cellIndex)
    if ((flags & (CellFlags.SpillChild | CellFlags.PivotOutput)) !== 0) {
      return false
    }
    return !(
      formula === undefined &&
      explicitFormat === undefined &&
      (flags & CellFlags.AuthoredBlank) === 0 &&
      (value.tag === ValueTag.Empty || value.tag === ValueTag.Error)
    )
  }

  const captureStoredCellState = (
    cellIndex: number,
    sheetName: string,
    address: string,
    sourceSheetName?: string,
    sourceAddress?: string,
  ): EngineOp[] => args.captureStoredCellOps(cellIndex, sheetName, address, sourceSheetName, sourceAddress)

  const isSimpleStructuralFormulaSourceDeferrable = (formula: RuntimeFormula): boolean =>
    formula.rangeDependencies.length === 0 &&
    formula.dependencyIndices.every((dependencyCellIndex) => shouldCaptureStoredCell(dependencyCellIndex)) &&
    !formula.compiled.volatile &&
    formula.compiled.symbolicNames.length === 0 &&
    formula.compiled.symbolicTables.length === 0 &&
    formula.compiled.symbolicSpills.length === 0 &&
    formula.directLookup === undefined &&
    formula.directAggregate === undefined &&
    formula.directCriteria === undefined &&
    isStructurallyStableSimpleFormulaNode(formula.compiled.ast)

  const canDeferSimpleStructuralFormulaSource = (formula: RuntimeFormula, transform: StructuralAxisTransform): boolean =>
    transform.kind !== 'delete' && transform.axis === 'column' && isSimpleStructuralFormulaSourceDeferrable(formula)

  const canDeferSimpleDeleteStructuralFormulaSource = (
    formula: RuntimeFormula,
    targetSheetId: number | undefined,
    transform: StructuralAxisTransform,
  ): boolean => {
    if (
      transform.kind !== 'delete' ||
      transform.axis !== 'column' ||
      targetSheetId === undefined ||
      !isSimpleStructuralFormulaSourceDeferrable(formula)
    ) {
      return false
    }
    return formula.dependencyIndices.every((dependencyCellIndex) => {
      if (args.state.workbook.cellStore.sheetIds[dependencyCellIndex] !== targetSheetId) {
        return true
      }
      const dependencyPosition = args.state.workbook.getCellPosition(dependencyCellIndex)
      return dependencyPosition !== undefined && mapStructuralAxisIndex(dependencyPosition.col, transform) !== undefined
    })
  }

  const captureAxisRangeCellState = (sheetName: string, axis: 'row' | 'column', start: number, count: number): EngineOp[] => {
    const sheet = args.state.workbook.getSheet(sheetName)
    if (!sheet) {
      return []
    }
    const captured: Array<{ cellIndex: number; row: number; col: number }> = []
    sheet.grid.forEachCellEntry((cellIndex, row, col) => {
      if (!shouldCaptureStoredCell(cellIndex)) {
        return
      }
      const index = axis === 'row' ? row : col
      if (index >= start && index < start + count) {
        captured.push({ cellIndex, row, col })
      }
    })
    if (args.state.counters && captured.length > 0) {
      addEngineCounter(args.state.counters, 'structuralUndoCapturedCells', captured.length)
    }
    return captured
      .toSorted((left, right) => left.row - right.row || left.col - right.col)
      .flatMap(({ cellIndex, row, col }) => captureStoredCellState(cellIndex, sheetName, formatAddress(row, col)))
  }

  const captureSheetCellState = (sheetName: string): EngineOp[] => {
    const sheet = args.state.workbook.getSheet(sheetName)
    if (!sheet) {
      return []
    }
    const captured: Array<{ cellIndex: number; row: number; col: number }> = []
    sheet.grid.forEachCellEntry((cellIndex, row, col) => {
      if (!shouldCaptureStoredCell(cellIndex)) {
        return
      }
      captured.push({ cellIndex, row, col })
    })
    return captured
      .toSorted((left, right) => left.row - right.row || left.col - right.col)
      .flatMap(({ cellIndex }) => captureStoredCellState(cellIndex, sheetName, args.state.workbook.getAddress(cellIndex)))
  }

  const rewriteDefinedNamesForStructuralTransform = (sheetName: string, transform: StructuralAxisTransform): Set<string> => {
    const changedNames = new Set<string>()
    args.state.workbook.listDefinedNames().forEach((record) => {
      if (typeof record.value === 'string' && record.value.startsWith('=')) {
        const nextFormula = rewriteFormulaForStructuralTransform(record.value.slice(1), sheetName, sheetName, transform)
        if (`=${nextFormula}` !== record.value) {
          args.state.workbook.setDefinedName(record.name, `=${nextFormula}`)
        }
        return
      }
      if (typeof record.value !== 'object' || !record.value) {
        return
      }
      switch (record.value.kind) {
        case 'formula': {
          const nextFormula = rewriteFormulaForStructuralTransform(
            record.value.formula.startsWith('=') ? record.value.formula.slice(1) : record.value.formula,
            sheetName,
            sheetName,
            transform,
          )
          const nextValue = {
            ...record.value,
            formula: record.value.formula.startsWith('=') ? `=${nextFormula}` : nextFormula,
          }
          if (nextValue.formula !== record.value.formula) {
            args.state.workbook.setDefinedName(record.name, nextValue)
            changedNames.add(normalizeDefinedName(record.name))
          }
          return
        }
        case 'cell-ref': {
          if (record.value.sheetName !== sheetName) {
            return
          }
          const nextAddress = rewriteAddressForStructuralTransform(record.value.address, transform)
          if (!nextAddress) {
            args.state.workbook.deleteDefinedName(record.name)
            changedNames.add(normalizeDefinedName(record.name))
            return
          }
          if (nextAddress !== record.value.address) {
            args.state.workbook.setDefinedName(record.name, {
              ...record.value,
              address: nextAddress,
            })
            changedNames.add(normalizeDefinedName(record.name))
          }
          return
        }
        case 'range-ref': {
          if (record.value.sheetName !== sheetName) {
            return
          }
          const nextRange = rewriteRangeForStructuralTransform(record.value.startAddress, record.value.endAddress, transform)
          if (!nextRange) {
            args.state.workbook.deleteDefinedName(record.name)
            changedNames.add(normalizeDefinedName(record.name))
            return
          }
          if (nextRange.startAddress !== record.value.startAddress || nextRange.endAddress !== record.value.endAddress) {
            args.state.workbook.setDefinedName(record.name, {
              ...record.value,
              startAddress: nextRange.startAddress,
              endAddress: nextRange.endAddress,
            })
            changedNames.add(normalizeDefinedName(record.name))
          }
          return
        }
        case 'scalar':
        case 'structured-ref':
          return
      }
    })
    return changedNames
  }

  const rewriteStructuralFormulaCompiled = (
    formula: RuntimeFormula,
    ownerSheetName: string,
    sheetName: string,
    transform: StructuralAxisTransform,
  ): ReturnType<typeof rewriteCompiledFormulaForStructuralTransform> | undefined => {
    if (formula.directAggregate) {
      const directAggregateCandidate = formula.compiled.directAggregateCandidate
      const rangeIndex = directAggregateCandidate?.symbolicRangeIndex
      const parsedRange = rangeIndex === undefined ? undefined : formula.compiled.parsedSymbolicRanges?.[rangeIndex]
      if (
        directAggregateCandidate &&
        rangeIndex !== undefined &&
        parsedRange &&
        parsedRange.refKind === 'cells' &&
        (parsedRange.sheetName ?? ownerSheetName) === sheetName
      ) {
        const candidateRangeIndex = rangeIndex
        const nextRange = rewriteRangeForStructuralTransform(parsedRange.startAddress, parsedRange.endAddress, transform)
        if (nextRange) {
          const rangePrefix = parsedRange.address.includes('!')
            ? parsedRange.address.slice(0, parsedRange.address.lastIndexOf('!') + 1)
            : ''
          const nextAddress = `${rangePrefix}${nextRange.startAddress}:${nextRange.endAddress}`
          const nextStart = parseCellAddress(nextRange.startAddress, parsedRange.sheetName ?? ownerSheetName)
          const nextEnd = parseCellAddress(nextRange.endAddress, parsedRange.sheetName ?? ownerSheetName)
          const nextParsedRange = {
            ...parsedRange,
            address: nextAddress,
            startAddress: nextRange.startAddress,
            endAddress: nextRange.endAddress,
            startRow: nextStart.row,
            endRow: nextEnd.row,
            startCol: nextStart.col,
            endCol: nextEnd.col,
          }
          const nextParsedSymbolicRanges = formula.compiled.parsedSymbolicRanges?.slice()
          if (nextParsedSymbolicRanges) {
            nextParsedSymbolicRanges[candidateRangeIndex] = nextParsedRange
          }
          const nextParsedDeps = formula.compiled.parsedDeps?.map((dependency) =>
            dependency.kind === 'range' && dependency.address === parsedRange.address
              ? {
                  ...dependency,
                  address: nextAddress,
                  startAddress: nextRange.startAddress,
                  endAddress: nextRange.endAddress,
                  startRow: nextStart.row,
                  endRow: nextEnd.row,
                  startCol: nextStart.col,
                  endCol: nextEnd.col,
                }
              : dependency,
          )
          return {
            source: `${directAggregateCandidate.callee}(${nextAddress})`,
            compiled: {
              ...formula.compiled,
              source: `${directAggregateCandidate.callee}(${nextAddress})`,
              astMatchesSource: false,
              deps: formula.compiled.deps.map((dependency) => (dependency === parsedRange.address ? nextAddress : dependency)),
              symbolicRanges: formula.compiled.symbolicRanges.map((range, index) => (index === candidateRangeIndex ? nextAddress : range)),
              ...(nextParsedDeps ? { parsedDeps: nextParsedDeps } : {}),
              ...(nextParsedSymbolicRanges ? { parsedSymbolicRanges: nextParsedSymbolicRanges } : {}),
            },
            reusedProgram: true,
          }
        }
      }
    }
    const rewritten = rewriteCompiledFormulaForStructuralTransform(formula.compiled, ownerSheetName, sheetName, transform)
    return rewritten.source === formula.source ? undefined : rewritten
  }

  const rewriteFormulaSourceFallback = (
    source: string,
    ownerSheetName: string,
    sheetName: string,
    transform: StructuralAxisTransform,
  ): string => rewriteFormulaForStructuralTransform(source, ownerSheetName, sheetName, transform)

  const structuralRewritePreservesDirectCellDependencies = (
    formula: RuntimeFormula,
    rewritten: { compiled: CompiledFormula },
    ownerSheetName: string,
  ): boolean => {
    if (
      formula.directLookup !== undefined ||
      formula.directAggregate !== undefined ||
      formula.directCriteria !== undefined ||
      formula.rangeDependencies.length !== 0 ||
      formula.dependencyIndices.length === 0
    ) {
      return false
    }
    const parsedDeps = rewritten.compiled.parsedDeps
    if (!parsedDeps || parsedDeps.length !== formula.dependencyIndices.length) {
      return false
    }
    for (let index = 0; index < parsedDeps.length; index += 1) {
      const dependency = parsedDeps[index]!
      if (dependency.kind !== 'cell') {
        return false
      }
      const dependencySheetName = dependency.sheetName ?? ownerSheetName
      const dependencyCellIndex = args.state.workbook.getCellIndex(dependencySheetName, dependency.address)
      if (dependencyCellIndex === undefined || dependencyCellIndex !== formula.dependencyIndices[index]) {
        return false
      }
    }
    return true
  }

  const rewriteFormulaFromTemplate = (
    cache: Map<string, StructurallyRewrittenTemplate | null>,
    formula: RuntimeFormula,
    representative: {
      readonly templateId: number
      readonly ownerSheetName: string
      readonly targetSheetName: string
      readonly representativeRow: number
      readonly representativeCol: number
      readonly ownerRow: number
      readonly ownerCol: number
    },
    targetSheetName: string,
    transform: StructuralAxisTransform,
  ): { source: string; compiled: CompiledFormula; reusedProgram: boolean } | undefined => {
    if (formula.directAggregate !== undefined || formula.directCriteria !== undefined) {
      return undefined
    }
    const cacheKey =
      `${representative.templateId}:${representative.ownerSheetName}:${targetSheetName}:${transform.kind}:${transform.axis}:${transform.start}:${transform.count}:` +
      `${transform.kind === 'move' ? transform.target : ''}`
    let rewrittenTemplate = cache.get(cacheKey)
    if (rewrittenTemplate === undefined) {
      if (formula.compiled.astMatchesSource === false) {
        return undefined
      }
      rewrittenTemplate =
        rewriteTemplateForStructuralTransform({
          template: {
            id: representative.templateId,
            templateKey: `runtime-template-${representative.templateId}`,
            baseSource: formula.source,
            baseRow: representative.representativeRow,
            baseCol: representative.representativeCol,
            compiled: formula.compiled,
          },
          ownerSheetName: representative.ownerSheetName,
          targetSheetName,
          transform,
        }) ?? null
      cache.set(cacheKey, rewrittenTemplate)
    }
    if (rewrittenTemplate === null) {
      return undefined
    }

    try {
      return retargetStructurallyRewrittenTemplateInstance({
        rewrittenTemplate,
        ownerRow: representative.ownerRow,
        ownerCol: representative.ownerCol,
      })
    } catch {
      return undefined
    }
  }

  const resolveStructuralFormulaRebindInputs = (argsForResolve: {
    readonly formulaCellIndices: readonly number[]
    readonly sheetName: string
    readonly transform: StructuralAxisTransform
    readonly transaction: StructuralTransaction
    readonly changedDefinedNames: ReadonlySet<string>
    readonly changedTableNames: ReadonlySet<string>
    readonly ownerPositions: ReadonlyMap<number, { sheetName: string; row: number; col: number }>
  }): StructuralFormulaRebindResolution => {
    const inputs: StructuralFormulaRebindInput[] = []
    const preservedCellIndices: number[] = []
    const templateRewriteCache = new Map<string, StructurallyRewrittenTemplate | null>()
    const remappedCellsByIndex = new Map(argsForResolve.transaction.remappedCells.map((entry) => [entry.cellIndex, entry] as const))
    argsForResolve.formulaCellIndices.forEach((cellIndex) => {
      const formula = args.state.formulas.get(cellIndex)
      if (!formula) {
        return
      }
      const previousOwnerPosition = argsForResolve.ownerPositions.get(cellIndex)
      if (!previousOwnerPosition) {
        return
      }
      const ownerSheetName = previousOwnerPosition.sheetName
      const touchesChangedName = formula.compiled.symbolicNames.some((name) =>
        argsForResolve.changedDefinedNames.has(normalizeDefinedName(name)),
      )
      const touchesChangedTable = formula.compiled.symbolicTables.some((name) => argsForResolve.changedTableNames.has(name))
      const touchesTargetSheetDependency = formula.compiled.deps.some((dependency) =>
        dependencyTouchesSheet(dependency, argsForResolve.sheetName),
      )
      const shouldBypassTemplateStructuralRewrite = ownerSheetName !== argsForResolve.sheetName && touchesTargetSheetDependency
      const representative = remappedCellsByIndex.get(cellIndex)
      const previousOwnerRow = representative?.fromRow ?? previousOwnerPosition.row
      const previousOwnerCol = representative?.fromCol ?? previousOwnerPosition.col
      const ownerRow =
        representative?.toRow ??
        (ownerSheetName === argsForResolve.sheetName && argsForResolve.transform.axis === 'row'
          ? mapStructuralAxisIndex(previousOwnerRow, argsForResolve.transform)
          : previousOwnerRow)
      const ownerCol =
        representative?.toCol ??
        (ownerSheetName === argsForResolve.sheetName && argsForResolve.transform.axis === 'column'
          ? mapStructuralAxisIndex(previousOwnerCol, argsForResolve.transform)
          : previousOwnerCol)
      if (ownerRow === undefined || ownerCol === undefined) {
        return
      }
      if (!touchesChangedName && !touchesChangedTable && canDeferSimpleStructuralFormulaSource(formula, argsForResolve.transform)) {
        formula.structuralSourceTransform = {
          ownerSheetName,
          targetSheetName: argsForResolve.sheetName,
          transform: argsForResolve.transform,
          preservesValue: true,
        }
        hasDeferredStructuralFormulaSources = true
        preservedCellIndices.push(cellIndex)
        return
      }
      const templateRewrite =
        !touchesChangedName &&
        !touchesChangedTable &&
        !shouldBypassTemplateStructuralRewrite &&
        formula.templateId !== undefined &&
        previousOwnerRow !== undefined &&
        previousOwnerCol !== undefined
          ? rewriteFormulaFromTemplate(
              templateRewriteCache,
              formula,
              {
                templateId: formula.templateId,
                ownerSheetName,
                targetSheetName: argsForResolve.sheetName,
                representativeRow: previousOwnerRow,
                representativeCol: previousOwnerCol,
                ownerRow,
                ownerCol,
              },
              argsForResolve.sheetName,
              argsForResolve.transform,
            )
          : undefined
      const compiledRewrite =
        templateRewrite === undefined
          ? rewriteStructuralFormulaCompiled(formula, ownerSheetName, argsForResolve.sheetName, argsForResolve.transform)
          : undefined
      const rewritten = !touchesChangedName && !touchesChangedTable ? (compiledRewrite ?? templateRewrite) : compiledRewrite
      if (!rewritten) {
        if (!touchesChangedName && !touchesChangedTable && formula.directAggregate !== undefined) {
          return
        }
        const canReuseCompiled =
          formula.compiled.symbolicNames.length === 0 &&
          formula.compiled.symbolicTables.length === 0 &&
          formula.compiled.symbolicSpills.length === 0
        inputs.push(
          canReuseCompiled
            ? {
                cellIndex,
                ownerSheetName,
                ownerRow,
                ownerCol,
                source: formula.source,
                compiled: formula.compiled,
                ...(formula.templateId === undefined ? {} : { templateId: formula.templateId }),
              }
            : {
                cellIndex,
                ownerSheetName,
                ownerRow,
                ownerCol,
                source: formula.source,
              },
        )
        return
      }
      if (touchesChangedName || touchesChangedTable) {
        inputs.push({
          cellIndex,
          ownerSheetName,
          ownerRow,
          ownerCol,
          source: rewriteFormulaSourceFallback(formula.source, ownerSheetName, argsForResolve.sheetName, argsForResolve.transform),
        })
        return
      }
      const preservesDirectCellDependencies = structuralRewritePreservesDirectCellDependencies(formula, rewritten, ownerSheetName)
      const preservesBinding =
        structuralRewritePreservesBinding(
          formula,
          rewritten,
          formula.rangeDependencies.every((rangeIndex) => args.state.ranges.getFormulaMembersView(rangeIndex).length === 0),
        ) || preservesDirectCellDependencies
      const preservesValue =
        structuralRewritePreservesValue(formula, rewritten, argsForResolve.transform) ||
        structuralDirectAggregateRewritePreservesValue(formula, rewritten, argsForResolve.transform)
      const hasOnlyPlaceholderDirectDependencies =
        formula.dependencyIndices.length > 0 &&
        !formula.dependencyIndices.every((dependencyCellIndex) => shouldCaptureStoredCell(dependencyCellIndex))
      const rewrittenDirectDependenciesChanged =
        formula.compiled.deps.length !== rewritten.compiled.deps.length ||
        formula.compiled.deps.some((dependency, index) => dependency !== rewritten.compiled.deps[index])
      const rewrittenPlaceholderDependencyNeedsRebind =
        preservesBinding && rewrittenDirectDependenciesChanged && hasOnlyPlaceholderDirectDependencies
      inputs.push({
        cellIndex,
        ownerSheetName,
        ownerRow,
        ownerCol,
        source: rewritten.source,
        compiled: rewritten.compiled,
        ...(formula.templateId === undefined ? {} : { templateId: formula.templateId }),
        preservesBinding: preservesBinding && !rewrittenPlaceholderDependencyNeedsRebind,
        preservesValue,
      })
    })
    return { inputs, preservedCellIndices }
  }

  const collectStructuralRangeDependencies = (argsForCollect: { readonly formulaCellIndices: readonly number[] }): number[] => {
    const rangeIndices = new Set<number>()
    argsForCollect.formulaCellIndices.forEach((cellIndex) => {
      const formula = args.state.formulas.get(cellIndex)
      if (!formula) {
        return
      }
      formula.rangeDependencies.forEach((rangeIndex) => {
        rangeIndices.add(rangeIndex)
      })
    })
    return [...rangeIndices]
  }

  const clearSpillMetadataForSheet = (sheetName: string): void => {
    args.state.workbook.listSpills().forEach((spill) => {
      if (spill.sheetName !== sheetName) {
        return
      }
      args.state.workbook.deleteSpill(spill.sheetName, spill.address)
    })
  }

  const clearPivotOutputsForSheet = (sheetName: string): void => {
    args.state.workbook
      .listPivots()
      .filter((pivot) => pivot.sheetName === sheetName)
      .forEach((pivot) => {
        args.clearOwnedPivot(pivot)
      })
  }

  const clearDerivedCellArtifacts = (cellIndex: number): void => {
    args.state.pivotOutputOwners.delete(cellIndex)
  }

  const rewriteWorkbookMetadataForStructuralTransform = (
    sheetName: string,
    transform: StructuralAxisTransform,
  ): { changedTableNames: Set<string> } => {
    const changedTableNames = new Set<string>()
    args.state.workbook
      .listTables()
      .filter((table) => table.sheetName === sheetName)
      .forEach((table) => {
        const range = rewriteRangeForStructuralTransform(table.startAddress, table.endAddress, transform)
        if (!range) {
          changedTableNames.add(table.name)
          args.state.workbook.deleteTable(table.name)
          return
        }
        changedTableNames.add(table.name)
        args.state.workbook.setTable({
          ...table,
          startAddress: range.startAddress,
          endAddress: range.endAddress,
        })
      })
    args.state.workbook.listFilters(sheetName).forEach((filter) => {
      const range = rewriteRangeForStructuralTransform(filter.range.startAddress, filter.range.endAddress, transform)
      args.state.workbook.deleteFilter(sheetName, filter.range)
      if (range) {
        args.state.workbook.setFilter(sheetName, {
          ...filter.range,
          startAddress: range.startAddress,
          endAddress: range.endAddress,
        })
      }
    })
    args.state.workbook.listSorts(sheetName).forEach((sort) => {
      const range = rewriteRangeForStructuralTransform(sort.range.startAddress, sort.range.endAddress, transform)
      args.state.workbook.deleteSort(sheetName, sort.range)
      if (!range) {
        return
      }
      args.state.workbook.setSort(
        sheetName,
        { ...sort.range, startAddress: range.startAddress, endAddress: range.endAddress },
        sort.keys.map((key) => ({
          ...key,
          keyAddress: rewriteAddressForStructuralTransform(key.keyAddress, transform) ?? key.keyAddress,
        })),
      )
    })
    args.state.workbook.listDataValidations(sheetName).forEach((validation) => {
      const range = rewriteRangeForStructuralTransform(validation.range.startAddress, validation.range.endAddress, transform)
      args.state.workbook.deleteDataValidation(sheetName, validation.range)
      if (!range) {
        return
      }
      const nextValidation = structuredClone(validation)
      nextValidation.range = {
        ...validation.range,
        startAddress: range.startAddress,
        endAddress: range.endAddress,
      }
      if (nextValidation.rule.kind === 'list' && nextValidation.rule.source) {
        switch (nextValidation.rule.source.kind) {
          case 'cell-ref': {
            if (nextValidation.rule.source.sheetName !== sheetName) {
              break
            }
            const nextAddress = rewriteAddressForStructuralTransform(nextValidation.rule.source.address, transform)
            if (!nextAddress) {
              return
            }
            nextValidation.rule.source.address = nextAddress
            break
          }
          case 'range-ref': {
            if (nextValidation.rule.source.sheetName !== sheetName) {
              break
            }
            const nextSourceRange = rewriteRangeForStructuralTransform(
              nextValidation.rule.source.startAddress,
              nextValidation.rule.source.endAddress,
              transform,
            )
            if (!nextSourceRange) {
              return
            }
            nextValidation.rule.source.startAddress = nextSourceRange.startAddress
            nextValidation.rule.source.endAddress = nextSourceRange.endAddress
            break
          }
          case 'named-range':
          case 'structured-ref':
            break
        }
      }
      args.state.workbook.setDataValidation(nextValidation)
    })
    args.state.workbook.listConditionalFormats(sheetName).forEach((format) => {
      const range = rewriteRangeForStructuralTransform(format.range.startAddress, format.range.endAddress, transform)
      args.state.workbook.deleteConditionalFormat(format.id)
      if (!range) {
        return
      }
      args.state.workbook.setConditionalFormat({
        ...format,
        range: {
          ...format.range,
          startAddress: range.startAddress,
          endAddress: range.endAddress,
        },
      })
    })
    args.state.workbook.listRangeProtections(sheetName).forEach((protection) => {
      const range = rewriteRangeForStructuralTransform(protection.range.startAddress, protection.range.endAddress, transform)
      args.state.workbook.deleteRangeProtection(protection.id)
      if (!range) {
        return
      }
      args.state.workbook.setRangeProtection({
        ...protection,
        range: {
          ...protection.range,
          startAddress: range.startAddress,
          endAddress: range.endAddress,
        },
      })
    })
    args.state.workbook.listCommentThreads(sheetName).forEach((thread) => {
      const nextAddress = rewriteAddressForStructuralTransform(thread.address, transform)
      args.state.workbook.deleteCommentThread(sheetName, thread.address)
      if (!nextAddress) {
        return
      }
      args.state.workbook.setCommentThread({
        ...thread,
        address: nextAddress,
      })
    })
    args.state.workbook.listNotes(sheetName).forEach((note) => {
      const nextAddress = rewriteAddressForStructuralTransform(note.address, transform)
      args.state.workbook.deleteNote(sheetName, note.address)
      if (!nextAddress) {
        return
      }
      args.state.workbook.setNote({
        ...note,
        address: nextAddress,
      })
    })
    const rewrittenStyleRanges: SheetStyleRangeSnapshot[] = []
    const rewrittenFormatRanges: SheetFormatRangeSnapshot[] = []
    args.state.workbook.listStyleRanges(sheetName).forEach((record) => {
      const range = rewriteRangeForStructuralTransform(record.range.startAddress, record.range.endAddress, transform)
      if (!range) {
        return
      }
      rewrittenStyleRanges.push({
        range: {
          ...record.range,
          startAddress: range.startAddress,
          endAddress: range.endAddress,
        },
        styleId: record.styleId,
      })
    })
    args.state.workbook.setStyleRanges(sheetName, rewrittenStyleRanges)
    args.state.workbook.listFormatRanges(sheetName).forEach((record) => {
      const range = rewriteRangeForStructuralTransform(record.range.startAddress, record.range.endAddress, transform)
      if (!range) {
        return
      }
      rewrittenFormatRanges.push({
        range: {
          ...record.range,
          startAddress: range.startAddress,
          endAddress: range.endAddress,
        },
        formatId: record.formatId,
      })
    })
    args.state.workbook.setFormatRanges(sheetName, rewrittenFormatRanges)
    const freezePane = args.state.workbook.getFreezePane(sheetName)
    if (freezePane) {
      const nextRows = transform.axis === 'row' ? mapStructuralBoundary(freezePane.rows, transform) : freezePane.rows
      const nextCols = transform.axis === 'column' ? mapStructuralBoundary(freezePane.cols, transform) : freezePane.cols
      if (nextRows <= 0 && nextCols <= 0) {
        args.state.workbook.clearFreezePane(sheetName)
      } else {
        args.state.workbook.setFreezePane(sheetName, nextRows, nextCols)
      }
    }
    args.state.workbook.listPivots().forEach((pivot) => {
      const nextAddress = pivot.sheetName === sheetName ? rewriteAddressForStructuralTransform(pivot.address, transform) : pivot.address
      const nextSource =
        pivot.source.sheetName === sheetName
          ? rewriteRangeForStructuralTransform(pivot.source.startAddress, pivot.source.endAddress, transform)
          : { startAddress: pivot.source.startAddress, endAddress: pivot.source.endAddress }
      if (!nextAddress || !nextSource) {
        args.clearOwnedPivot(pivot)
        args.state.workbook.deletePivot(pivot.sheetName, pivot.address)
        return
      }
      if (nextAddress !== pivot.address) {
        args.clearOwnedPivot(pivot)
        args.state.workbook.deletePivot(pivot.sheetName, pivot.address)
      }
      args.state.workbook.setPivot({
        ...pivot,
        address: nextAddress,
        source: {
          ...pivot.source,
          startAddress: nextSource.startAddress,
          endAddress: nextSource.endAddress,
        },
      })
    })
    args.state.workbook.listCharts().forEach((chart) => {
      const nextAddress = chart.sheetName === sheetName ? rewriteAddressForStructuralTransform(chart.address, transform) : chart.address
      const nextSource =
        chart.source.sheetName === sheetName
          ? rewriteRangeForStructuralTransform(chart.source.startAddress, chart.source.endAddress, transform)
          : { startAddress: chart.source.startAddress, endAddress: chart.source.endAddress }
      if (!nextAddress || !nextSource) {
        args.state.workbook.deleteChart(chart.id)
        return
      }
      args.state.workbook.setChart({
        ...chart,
        address: nextAddress,
        source: {
          ...chart.source,
          startAddress: nextSource.startAddress,
          endAddress: nextSource.endAddress,
        },
      })
    })
    args.state.workbook.listImages().forEach((image) => {
      if (image.sheetName !== sheetName) {
        return
      }
      const nextAddress = rewriteAddressForStructuralTransform(image.address, transform)
      if (!nextAddress) {
        args.state.workbook.deleteImage(image.id)
        return
      }
      args.state.workbook.setImage({
        ...image,
        address: nextAddress,
      })
    })
    args.state.workbook.listShapes().forEach((shape) => {
      if (shape.sheetName !== sheetName) {
        return
      }
      const nextAddress = rewriteAddressForStructuralTransform(shape.address, transform)
      if (!nextAddress) {
        args.state.workbook.deleteShape(shape.id)
        return
      }
      args.state.workbook.setShape({
        ...shape,
        address: nextAddress,
      })
    })
    return { changedTableNames }
  }

  const isCellIndexMapped = (cellIndex: number): boolean => {
    const sheetId = args.state.workbook.cellStore.sheetIds[cellIndex]
    const position = args.state.workbook.getCellPosition(cellIndex)
    if (sheetId === undefined || !position || !Number.isFinite(position.row) || !Number.isFinite(position.col)) {
      return false
    }
    const sheet = args.state.workbook.getSheetById(sheetId)
    return (
      (sheet?.structureVersion === 1 ? sheet.grid.getPhysical(position.row, position.col) : sheet?.grid.get(position.row, position.col)) ===
      cellIndex
    )
  }

  const structuralAxisIndexAffected = (axisIndex: number, transform: StructuralAxisTransform): boolean => {
    const nextIndex = mapStructuralAxisIndex(axisIndex, transform)
    return nextIndex === undefined || nextIndex !== axisIndex
  }

  const collectStructuralFormulaImpacts = (argsForImpact: {
    readonly targetSheetId: number | undefined
    readonly transform: StructuralAxisTransform
    readonly sheetName: string
    readonly changedDefinedNames: ReadonlySet<string>
    readonly changedTableNames: ReadonlySet<string>
  }): {
    formulaCellIndices: number[]
    rebindCellIndices: number[]
    preservedCellIndices: number[]
    ownerPositions: Map<number, { sheetName: string; row: number; col: number }>
  } => {
    const formulaCellIndices = new Set<number>()
    const rebindCellIndices = new Set<number>()
    const preservedCellIndices = new Set<number>()
    const candidateCellIndices = new Set<number>()
    const ownerPositions = new Map<number, { sheetName: string; row: number; col: number }>()
    const tryDeferOwnedSimpleFormula = (cellIndex: number): boolean => {
      if (argsForImpact.changedDefinedNames.size > 0 || argsForImpact.changedTableNames.size > 0) {
        return false
      }
      const formula = args.state.formulas.get(cellIndex)
      if (!formula) {
        return false
      }
      const ownerPosition = args.state.workbook.getCellPosition(cellIndex)
      if (
        !ownerPosition ||
        mapStructuralAxisIndex(argsForImpact.transform.axis === 'row' ? ownerPosition.row : ownerPosition.col, argsForImpact.transform) ===
          undefined
      ) {
        return false
      }
      const preservesValue = canDeferSimpleStructuralFormulaSource(formula, argsForImpact.transform)
      if (!preservesValue && !canDeferSimpleDeleteStructuralFormulaSource(formula, argsForImpact.targetSheetId, argsForImpact.transform)) {
        return false
      }
      formula.structuralSourceTransform = {
        ownerSheetName: argsForImpact.sheetName,
        targetSheetName: argsForImpact.sheetName,
        transform: argsForImpact.transform,
        preservesValue,
      }
      hasDeferredStructuralFormulaSources = true
      if (preservesValue) {
        preservedCellIndices.add(cellIndex)
      } else {
        formulaCellIndices.add(cellIndex)
      }
      return true
    }
    args.collectFormulaCellsOwnedBySheet(argsForImpact.sheetName).forEach((cellIndex) => {
      if (tryDeferOwnedSimpleFormula(cellIndex)) {
        return
      }
      candidateCellIndices.add(cellIndex)
    })
    args.collectFormulaCellsReferencingSheet(argsForImpact.sheetName).forEach((cellIndex) => {
      candidateCellIndices.add(cellIndex)
    })
    if (argsForImpact.changedDefinedNames.size > 0) {
      args.collectFormulaCellsForDefinedNames([...argsForImpact.changedDefinedNames]).forEach((cellIndex) => {
        candidateCellIndices.add(cellIndex)
      })
    }
    if (argsForImpact.changedTableNames.size > 0) {
      args.collectFormulaCellsForTables([...argsForImpact.changedTableNames]).forEach((cellIndex) => {
        candidateCellIndices.add(cellIndex)
      })
    }
    if (args.state.counters && candidateCellIndices.size > 0) {
      addEngineCounter(args.state.counters, 'structuralFormulaImpactCandidates', candidateCellIndices.size)
    }
    candidateCellIndices.forEach((cellIndex) => {
      const formula = args.state.formulas.get(cellIndex)
      if (!formula) {
        return
      }
      if (!isCellIndexMapped(cellIndex)) {
        return
      }
      const ownerSheetName = args.state.workbook.getSheetNameById(args.state.workbook.cellStore.sheetIds[cellIndex]!)
      if (!ownerSheetName) {
        return
      }
      const ownerPosition = args.state.workbook.getCellPosition(cellIndex)
      if (!ownerPosition) {
        return
      }
      ownerPositions.set(cellIndex, { sheetName: ownerSheetName, row: ownerPosition.row, col: ownerPosition.col })
      const axisIndex = argsForImpact.transform.axis === 'row' ? ownerPosition?.row : ownerPosition?.col
      const ownerPositionAffected =
        ownerSheetName === argsForImpact.sheetName &&
        axisIndex !== undefined &&
        structuralAxisIndexAffected(axisIndex, argsForImpact.transform)
      const touchesChangedName =
        argsForImpact.changedDefinedNames.size > 0 &&
        formula.compiled.symbolicNames.some((name) => argsForImpact.changedDefinedNames.has(normalizeDefinedName(name)))
      const touchesChangedTable =
        argsForImpact.changedTableNames.size > 0 &&
        formula.compiled.symbolicTables.some((name) => argsForImpact.changedTableNames.has(name))
      if (!touchesChangedName && !touchesChangedTable && canDeferSimpleStructuralFormulaSource(formula, argsForImpact.transform)) {
        formula.structuralSourceTransform = {
          ownerSheetName,
          targetSheetName: argsForImpact.sheetName,
          transform: argsForImpact.transform,
          preservesValue: true,
        }
        hasDeferredStructuralFormulaSources = true
        preservedCellIndices.add(cellIndex)
        return
      }
      const dependencyPositionAffected =
        !ownerPositionAffected &&
        argsForImpact.targetSheetId !== undefined &&
        (formula.dependencyIndices.some((dependencyCellIndex) => {
          if (args.state.workbook.cellStore.sheetIds[dependencyCellIndex] !== argsForImpact.targetSheetId) {
            return false
          }
          const dependencyPosition = args.state.workbook.getCellPosition(dependencyCellIndex)
          const dependencyAxisIndex = argsForImpact.transform.axis === 'row' ? dependencyPosition?.row : dependencyPosition?.col
          return dependencyAxisIndex !== undefined && structuralAxisIndexAffected(dependencyAxisIndex, argsForImpact.transform)
        }) ||
          formula.rangeDependencies.some((rangeIndex) =>
            rangeDependencyAxisAffected(args.state.ranges.getDescriptor(rangeIndex), argsForImpact.targetSheetId!, argsForImpact.transform),
          ) ||
          runtimeDirectRangeAxisAffected(
            argsForImpact.targetSheetId,
            argsForImpact.sheetName,
            argsForImpact.transform,
            formula.directAggregate,
          ) ||
          runtimeDirectRangeAxisAffected(
            argsForImpact.targetSheetId,
            argsForImpact.sheetName,
            argsForImpact.transform,
            formula.directCriteria?.aggregateRange,
          ) ||
          formula.directCriteria?.criteriaPairs.some((pair) =>
            runtimeDirectRangeAxisAffected(argsForImpact.targetSheetId, argsForImpact.sheetName, argsForImpact.transform, pair.range),
          ) ||
          runtimeDirectRangeAxisAffected(
            argsForImpact.targetSheetId,
            argsForImpact.sheetName,
            argsForImpact.transform,
            formula.directLookup?.kind === 'exact' || formula.directLookup?.kind === 'approximate'
              ? {
                  sheetName: formula.directLookup.prepared.sheetName,
                  rowStart: formula.directLookup.prepared.rowStart,
                  rowEnd: formula.directLookup.prepared.rowEnd,
                  col: formula.directLookup.prepared.col,
                }
              : formula.directLookup?.kind === 'exact-uniform-numeric' || formula.directLookup?.kind === 'approximate-uniform-numeric'
                ? {
                    sheetName: formula.directLookup.sheetName,
                    rowStart: formula.directLookup.rowStart,
                    rowEnd: formula.directLookup.rowEnd,
                    col: formula.directLookup.col,
                  }
                : undefined,
          ))
      const touchesSheetDependency =
        !ownerPositionAffected &&
        !dependencyPositionAffected &&
        formula.compiled.deps.some((dependency) => dependencyTouchesSheet(dependency, argsForImpact.sheetName))
      if (!ownerPositionAffected && !dependencyPositionAffected && !touchesSheetDependency && !touchesChangedName && !touchesChangedTable) {
        return
      }
      formulaCellIndices.add(cellIndex)
      if (ownerPositionAffected || dependencyPositionAffected || touchesSheetDependency || touchesChangedName || touchesChangedTable) {
        rebindCellIndices.add(cellIndex)
      }
    })
    return {
      formulaCellIndices: [...formulaCellIndices],
      rebindCellIndices: [...rebindCellIndices],
      preservedCellIndices: [...preservedCellIndices],
      ownerPositions,
    }
  }

  const materializeDeferredStructuralFormulaSources = (): void => {
    if (!hasDeferredStructuralFormulaSources) {
      return
    }
    const inputs: StructuralFormulaRebindInput[] = []
    args.state.formulas.forEach((formula, cellIndex) => {
      if (
        formula.structuralSourceTransform === undefined ||
        formula.directLookup !== undefined ||
        formula.directAggregate !== undefined ||
        formula.directCriteria !== undefined ||
        formula.rangeDependencies.length !== 0 ||
        !isCellIndexMapped(cellIndex)
      ) {
        return
      }
      const sheetId = args.state.workbook.cellStore.sheetIds[cellIndex]
      const ownerPosition = args.state.workbook.getCellPosition(cellIndex)
      if (sheetId === undefined || !ownerPosition) {
        return
      }
      const ownerSheetName = args.state.workbook.getSheetNameById(sheetId)
      if (!ownerSheetName) {
        return
      }
      const source = getRuntimeFormulaSource(formula)
      const compiled = getRuntimeFormulaStructuralCompiled(formula)
      const preservesValue = formula.structuralSourceTransform.preservesValue
      inputs.push({
        cellIndex,
        ownerSheetName,
        ownerRow: ownerPosition.row,
        ownerCol: ownerPosition.col,
        source,
        ...(compiled
          ? {
              compiled,
              ...(formula.templateId === undefined ? {} : { templateId: formula.templateId }),
              preservesBinding: true,
              preservesValue,
            }
          : {}),
      })
    })
    if (inputs.length > 0) {
      args.rebindFormulaCells(inputs)
    }
    inputs.forEach(({ cellIndex }) => {
      const formula = args.state.formulas.get(cellIndex)
      if (formula) {
        formula.structuralSourceTransform = undefined
      }
    })
    hasDeferredStructuralFormulaSources = false
  }

  return {
    captureSheetCellState(sheetName) {
      return Effect.try({
        try: () => captureSheetCellState(sheetName),
        catch: (cause) =>
          new EngineStructureError({
            message: `Failed to capture sheet cell state for ${sheetName}`,
            cause,
          }),
      })
    },
    captureRowRangeCellState(sheetName, start, count) {
      return Effect.try({
        try: () => captureAxisRangeCellState(sheetName, 'row', start, count),
        catch: (cause) =>
          new EngineStructureError({
            message: `Failed to capture row state for ${sheetName}`,
            cause,
          }),
      })
    },
    captureColumnRangeCellState(sheetName, start, count) {
      return Effect.try({
        try: () => captureAxisRangeCellState(sheetName, 'column', start, count),
        catch: (cause) =>
          new EngineStructureError({
            message: `Failed to capture column state for ${sheetName}`,
            cause,
          }),
      })
    },
    materializeDeferredStructuralFormulaSources() {
      return Effect.try({
        try: () => materializeDeferredStructuralFormulaSources(),
        catch: (cause) =>
          new EngineStructureError({
            message: 'Failed to materialize deferred structural formula sources',
            cause,
          }),
      })
    },
    applyStructuralAxisOp(op) {
      return Effect.try({
        try: () => {
          materializeDeferredStructuralFormulaSources()
          const transform = structuralTransformForOp(op)
          const sheetName = op.sheetName
          const targetSheetId = args.state.workbook.getSheet(sheetName)?.id

          clearPivotOutputsForSheet(sheetName)
          const changedDefinedNames = rewriteDefinedNamesForStructuralTransform(sheetName, transform)
          const { changedTableNames } = rewriteWorkbookMetadataForStructuralTransform(sheetName, transform)
          const impactedFormulas = collectStructuralFormulaImpacts({
            targetSheetId,
            transform,
            sheetName,
            changedDefinedNames,
            changedTableNames,
          })

          const transaction =
            args.state.workbook.planStructuralAxisTransform(sheetName, transform) ??
            (() => {
              throw new Error(`Missing sheet for structural op: ${sheetName}`)
            })()

          switch (op.kind) {
            case 'insertRows':
              args.state.workbook.insertRows(sheetName, op.start, op.count, op.entries)
              break
            case 'deleteRows':
              args.state.workbook.deleteRows(sheetName, op.start, op.count)
              break
            case 'moveRows':
              args.state.workbook.moveRows(sheetName, op.start, op.count, op.target)
              break
            case 'insertColumns':
              args.state.workbook.insertColumns(sheetName, op.start, op.count, op.entries)
              break
            case 'deleteColumns':
              args.state.workbook.deleteColumns(sheetName, op.start, op.count)
              break
            case 'moveColumns':
              args.state.workbook.moveColumns(sheetName, op.start, op.count, op.target)
              break
          }

          args.state.workbook.applyPlannedStructuralTransaction(transaction)

          const structuralRangeDependencies = collectStructuralRangeDependencies({
            formulaCellIndices: impactedFormulas.formulaCellIndices,
          })

          const hadCycleFormulas = (() => {
            let found = false
            args.state.formulas.forEach((_formula, cellIndex) => {
              if (((args.state.workbook.cellStore.flags[cellIndex] ?? 0) & CellFlags.InCycle) !== 0) {
                found = true
              }
            })
            return found
          })()
          const removedFormulaCellIndices = transaction.removedCellIndices.filter((cellIndex) => args.state.formulas.has(cellIndex))
          const removedFormulaCellIndexSet = new Set<number>(removedFormulaCellIndices)
          transaction.removedCellIndices.forEach((cellIndex) => {
            clearDerivedCellArtifacts(cellIndex)
            args.removeFormula(cellIndex)
            args.state.workbook.setCellFormat(cellIndex, null)
            args.state.workbook.cellStore.setValue(cellIndex, emptyValue())
            args.state.workbook.cellStore.flags[cellIndex] =
              (args.state.workbook.cellStore.flags[cellIndex] ?? 0) &
              ~(CellFlags.HasFormula | CellFlags.JsOnly | CellFlags.InCycle | CellFlags.SpillChild | CellFlags.PivotOutput)
            args.state.workbook.detachCellIndex(cellIndex)
          })

          clearSpillMetadataForSheet(sheetName)
          args.retargetRangeDependencies(transaction, structuralRangeDependencies)
          const rebindResolution = resolveStructuralFormulaRebindInputs({
            formulaCellIndices: impactedFormulas.rebindCellIndices.filter((cellIndex) => isCellIndexMapped(cellIndex)),
            sheetName,
            transform,
            transaction,
            changedDefinedNames,
            changedTableNames,
            ownerPositions: impactedFormulas.ownerPositions,
          })
          const rebindInputs = rebindResolution.inputs
          if (args.state.counters && rebindInputs.length > 0) {
            addEngineCounter(args.state.counters, 'structuralFormulaRebindInputs', rebindInputs.length)
          }
          const formulaCellIndices = impactedFormulas.formulaCellIndices.filter((cellIndex) => isCellIndexMapped(cellIndex))
          const onlyDirectAggregateFormulaCells =
            formulaCellIndices.length > 0 &&
            formulaCellIndices.every((cellIndex) => args.state.formulas.get(cellIndex)?.directAggregate !== undefined)
          args.rebindFormulaCells(rebindInputs)
          const reboundFormulaCellIndices = new Set(rebindInputs.map((input) => input.cellIndex))
          const preservedFormulaCellIndices = new Set([
            ...impactedFormulas.preservedCellIndices,
            ...rebindResolution.preservedCellIndices,
            ...rebindInputs.filter((input) => input.preservesValue).map((input) => input.cellIndex),
          ])
          const lostSurvivingFormulaCells = impactedFormulas.formulaCellIndices.some(
            (cellIndex) =>
              !reboundFormulaCellIndices.has(cellIndex) && !isCellIndexMapped(cellIndex) && !removedFormulaCellIndexSet.has(cellIndex),
          )
          const hasNonPreservedRebind = rebindInputs.some((input) => input.preservesBinding !== true)
          const topologyChanged = removedFormulaCellIndices.length > 0 || hasNonPreservedRebind || lostSurvivingFormulaCells
          const graphRefreshRequired =
            ((hasNonPreservedRebind || lostSurvivingFormulaCells) && !onlyDirectAggregateFormulaCells) ||
            (removedFormulaCellIndices.length > 0 && hadCycleFormulas)
          return {
            transaction,
            changedCellIndices: [...transaction.removedCellIndices],
            formulaCellIndices: formulaCellIndices.filter((cellIndex) => !preservedFormulaCellIndices.has(cellIndex)),
            topologyChanged,
            graphRefreshRequired,
          }
        },
        catch: (cause) =>
          new EngineStructureError({
            message: `Failed to apply structural operation ${op.kind}`,
            cause,
          }),
      })
    },
  }
}
