import { Effect } from 'effect'
import type { CompiledFormula } from '@bilig/formula'
import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import type { EngineCellMutationRef, EngineFormulaSourceRef } from '../../cell-mutations-at.js'
import { CellFlags } from '../../cell-store.js'
import { buildFormulaFamilyShapeKey } from '../../formula/formula-family-deps.js'
import type { FormulaFamilyMember, FormulaFamilyRunUpsertArgs } from '../../formula/formula-family-store.js'
import type { FormulaTemplateResolution } from '../../formula/template-bank.js'
import { translateSimpleDirectScalarFormula } from '../../formula/simple-direct-scalar-compile.js'
import { addEngineCounter } from '../../perf/engine-counters.js'
import type {
  EngineRuntimeState,
  RuntimeDirectScalarDescriptor,
  RuntimeDirectScalarOperand,
  RuntimeFormula,
  U32,
} from '../runtime-state.js'
import { EngineMutationError } from '../errors.js'

const INITIAL_DIRECT_FORMULA_EVALUATION_LIMIT = 16_384
const EMPTY_U32 = new Uint32Array(0)

type InitialPrefixAggregateKind = 'sum' | 'count' | 'average' | 'min' | 'max'

interface InitialPrefixAggregateGroup {
  readonly sheetName: string
  readonly col: number
  readonly aggregateKind: InitialPrefixAggregateKind
  maxRowEnd: number
  lastRowEnd: number
  formulasAreOrdered: boolean
  readonly formulas: Array<{ cellIndex: number; rowEnd: number }>
}

type DeferredInitialFormulaFamilyRun = Omit<FormulaFamilyRunUpsertArgs, 'members'> & {
  members: FormulaFamilyMember[]
}

interface InitialTemplateFormulaCacheEntry {
  readonly resolution: FormulaTemplateResolution
  readonly anchorRow: number
  readonly anchorCol: number
  readonly anchorCompiled: CompiledFormula
}

function mutationErrorMessage(message: string, cause: unknown): string {
  return cause instanceof Error && cause.message.length > 0 ? cause.message : message
}

function canEvaluateInitialDirectRuntimeFormula(formula: RuntimeFormula | undefined): boolean {
  return (
    formula !== undefined &&
    !formula.compiled.volatile &&
    !formula.compiled.producesSpill &&
    (formula.directAggregate !== undefined ||
      formula.directCriteria !== undefined ||
      formula.directLookup !== undefined ||
      formula.directScalar !== undefined)
  )
}

function hasPendingFormulaDependency(formula: RuntimeFormula, pendingFormulaCells: Uint8Array): boolean {
  const dependencies = formula.dependencyIndices
  for (let index = 0; index < dependencies.length; index += 1) {
    if ((pendingFormulaCells[dependencies[index]!] ?? 0) !== 0) {
      return true
    }
  }
  return false
}

function initialColumnToIndex(column: string): number {
  let value = 0
  for (let index = 0; index < column.length; index += 1) {
    const code = column.charCodeAt(index)
    value = value * 26 + (code >= 97 && code <= 122 ? code - 96 : code - 64)
  }
  return value - 1
}

function initialReadColumn(source: string, start: number): { readonly column: string; readonly next: number } | undefined {
  let next = start
  while (next < source.length) {
    const code = source.charCodeAt(next)
    if (!((code >= 65 && code <= 90) || (code >= 97 && code <= 122))) {
      break
    }
    next += 1
  }
  return next === start ? undefined : { column: source.slice(start, next), next }
}

function initialReadRowNumber(source: string, start: number): { readonly row: number; readonly next: number } | undefined {
  let next = start
  let row = 0
  while (next < source.length) {
    const code = source.charCodeAt(next)
    if (code < 48 || code > 57) {
      break
    }
    row = row * 10 + (code - 48)
    next += 1
  }
  return next === start || row <= 0 ? undefined : { row, next }
}

function initialReadNumberLiteral(source: string, start: number): { readonly text: string; readonly next: number } | undefined {
  let next = start
  while (next < source.length) {
    const code = source.charCodeAt(next)
    if (code < 48 || code > 57) {
      break
    }
    next += 1
  }
  if (next < source.length && source.charCodeAt(next) === 46) {
    const fractionStart = next + 1
    next = fractionStart
    while (next < source.length) {
      const code = source.charCodeAt(next)
      if (code < 48 || code > 57) {
        break
      }
      next += 1
    }
    if (next === fractionStart) {
      return undefined
    }
  }
  return next === start ? undefined : { text: source.slice(start, next), next }
}

function initialReadRelativeCellToken(
  source: string,
  start: number,
  ownerRow: number,
  ownerCol: number,
): { readonly token: string; readonly next: number } | undefined {
  const column = initialReadColumn(source, start)
  if (!column) {
    return undefined
  }
  const row = initialReadRowNumber(source, column.next)
  if (!row || row.row - 1 !== ownerRow) {
    return undefined
  }
  const col = initialColumnToIndex(column.column)
  return col < 0 ? undefined : { token: `c${col - ownerCol}`, next: row.next }
}

function tryBuildInitialSimpleRowRelativeBinaryTemplateKey(source: string, ownerRow: number, ownerCol: number): string | undefined {
  let index = source.charCodeAt(0) === 61 ? 1 : 0
  const left = initialReadRelativeCellToken(source, index, ownerRow, ownerCol)
  if (!left) {
    return undefined
  }
  index = left.next
  const operator = source[index]
  if (operator !== '+' && operator !== '-' && operator !== '*' && operator !== '/') {
    return undefined
  }
  index += 1
  const rightCell = initialReadRelativeCellToken(source, index, ownerRow, ownerCol)
  if (rightCell) {
    return rightCell.next === source.length ? `${left.token}${operator}${rightCell.token}` : undefined
  }
  const rightNumber = initialReadNumberLiteral(source, index)
  return rightNumber && rightNumber.next === source.length ? `${left.token}${operator}n${rightNumber.text}` : undefined
}

function initialFormulaFamilyShapeKey(formula: RuntimeFormula): string {
  return buildFormulaFamilyShapeKey({
    compiled: formula.compiled,
    dependencyCount: formula.dependencyIndices.length,
    rangeDependencyCount: formula.rangeDependencies.length,
    directAggregateKind: formula.directAggregate?.aggregateKind,
    directLookupKind: formula.directLookup?.kind,
    directScalarKind: formula.directScalar?.kind,
    directCriteriaKind: formula.directCriteria?.aggregateKind,
  })
}

export interface EngineFormulaInitializationService {
  readonly initializeCellFormulasAt: (
    refs: readonly EngineCellMutationRef[],
    potentialNewCells?: number,
  ) => Effect.Effect<void, EngineMutationError>
  readonly initializeCellFormulasAtNow: (refs: readonly EngineCellMutationRef[], potentialNewCells?: number) => void
  readonly initializeFormulaSourcesAtNow: (refs: readonly EngineFormulaSourceRef[], potentialNewCells?: number) => void
  readonly initializePreparedCellFormulasAt: (
    refs: readonly PreparedFormulaInitializationRef[],
    potentialNewCells?: number,
  ) => Effect.Effect<void, EngineMutationError>
  readonly initializePreparedCellFormulasAtNow: (refs: readonly PreparedFormulaInitializationRef[], potentialNewCells?: number) => void
  readonly initializeHydratedPreparedCellFormulasAt: (
    refs: readonly HydratedPreparedFormulaInitializationRef[],
    potentialNewCells?: number,
  ) => Effect.Effect<void, EngineMutationError>
  readonly initializeHydratedPreparedCellFormulasAtNow: (
    refs: readonly HydratedPreparedFormulaInitializationRef[],
    potentialNewCells?: number,
  ) => void
}

export interface PreparedFormulaInitializationRef {
  readonly sheetId: number
  readonly row: number
  readonly col: number
  readonly source: string
  readonly compiled: CompiledFormula
  readonly templateId?: number
  readonly cellIndex?: number
}

export interface HydratedPreparedFormulaInitializationRef extends PreparedFormulaInitializationRef {
  readonly value: CellValue
}

interface InitialWrittenColumnTracker {
  columns: Uint8Array
  count: number
}

interface InitialFormulaValueWriter {
  readonly writeValue: (cellIndex: number, value: CellValue) => void
  readonly writeValueAt: (cellIndex: number, sheetId: number, col: number, value: CellValue) => void
  readonly writeNumber: (cellIndex: number, value: number) => void
  readonly writeNumberAt: (cellIndex: number, sheetId: number, col: number, value: number) => void
  readonly flush: () => void
}

function createInitialWrittenColumnTracker(): InitialWrittenColumnTracker {
  return {
    columns: new Uint8Array(8),
    count: 0,
  }
}

function markInitialWrittenColumn(tracker: InitialWrittenColumnTracker, col: number): void {
  if (col >= tracker.columns.length) {
    let nextLength = tracker.columns.length
    while (nextLength <= col) {
      nextLength *= 2
    }
    const nextColumns = new Uint8Array(nextLength)
    nextColumns.set(tracker.columns)
    tracker.columns = nextColumns
  }
  if (tracker.columns[col] !== 0) {
    return
  }
  tracker.columns[col] = 1
  tracker.count += 1
}

function materializeInitialWrittenColumns(tracker: InitialWrittenColumnTracker): Uint32Array {
  const columns = new Uint32Array(tracker.count)
  let writeIndex = 0
  for (let col = 0; col < tracker.columns.length; col += 1) {
    if (tracker.columns[col] !== 0) {
      columns[writeIndex] = col
      writeIndex += 1
    }
  }
  return columns
}

export function createEngineFormulaInitializationService(args: {
  readonly state: Pick<
    EngineRuntimeState,
    'workbook' | 'strings' | 'formulas' | 'ranges' | 'counters' | 'getLastMetrics' | 'setLastMetrics'
  >
  readonly beginMutationCollection: () => void
  readonly ensureRecalcScratchCapacity: (size: number) => void
  readonly ensureCellTrackedByCoords: (sheetId: number, row: number, col: number) => number
  readonly resetMaterializedCellScratch: (expectedSize: number) => void
  readonly bindFormula: (cellIndex: number, ownerSheetName: string, source: string) => void
  readonly bindPreparedFormula: (
    cellIndex: number,
    ownerSheetName: string,
    source: string,
    compiled: CompiledFormula,
    templateId?: number,
    options?: { readonly deferFamilyRegistration?: boolean },
  ) => boolean
  readonly upsertFormulaFamilyRun: (args: FormulaFamilyRunUpsertArgs) => void
  readonly compileTemplateFormula: (source: string, row: number, col: number) => FormulaTemplateResolution
  readonly clearTemplateFormulaCache: () => void
  readonly removeFormula: (cellIndex: number) => boolean
  readonly setInvalidFormulaValue: (cellIndex: number) => void
  readonly markInputChanged: (cellIndex: number, count: number) => number
  readonly markFormulaChanged: (cellIndex: number, count: number) => number
  readonly markVolatileFormulasChanged: (count: number) => number
  readonly syncDynamicRanges: (formulaChangedCount: number) => number
  readonly composeMutationRoots: (changedInputCount: number, formulaChangedCount: number) => U32
  readonly getChangedInputBuffer: () => U32
  readonly getChangedFormulaBuffer: () => U32
  readonly rebuildTopoRanks: () => void
  readonly repairTopoRanks: (changedFormulaCells: readonly number[] | U32) => boolean
  readonly detectCycles: () => void
  readonly recalculate: (changedRoots: readonly number[] | U32, kernelSyncRoots?: readonly number[] | U32) => U32
  readonly deferKernelSync: (cellIndices: readonly number[] | U32) => void
  readonly evaluateDirectFormula: (cellIndex: number) => readonly number[] | undefined
  readonly recalculatePreordered: (
    changedRoots: readonly number[] | U32,
    orderedFormulaCellIndices: readonly number[] | U32,
    orderedFormulaCount: number,
    kernelSyncRoots?: readonly number[] | U32,
  ) => U32
  readonly reconcilePivotOutputs: (baseChanged: U32, forceAllPivots?: boolean) => U32
  readonly getBatchMutationDepth: () => number
  readonly setBatchMutationDepth: (next: number) => void
  readonly prepareRegionQueryIndices: () => void
  readonly writeHydratedFormulaValue: (cellIndex: number, value: CellValue) => void
}): EngineFormulaInitializationService {
  const sheetNameById = new Map<number, string>()
  const hasCycleMembersNow = (): boolean => {
    addEngineCounter(args.state.counters, 'cycleFormulaScans')
    let found = false
    args.state.formulas.forEach((_formula, cellIndex) => {
      if (((args.state.workbook.cellStore.flags[cellIndex] ?? 0) & CellFlags.InCycle) !== 0) {
        found = true
      }
    })
    return found
  }
  const resolveSheetName = (sheetId: number): string => {
    const cached = sheetNameById.get(sheetId)
    if (cached !== undefined) {
      return cached
    }
    const sheet = args.state.workbook.getSheetById(sheetId)
    if (!sheet) {
      throw new Error(`Unknown sheet id: ${sheetId}`)
    }
    sheetNameById.set(sheetId, sheet.name)
    return sheet.name
  }

  const createInitialFormulaValueWriter = (): InitialFormulaValueWriter => {
    let writtenColumnsBySheetId: Map<number, InitialWrittenColumnTracker> | undefined
    const markKnownColumn = (sheetId: number, col: number): void => {
      writtenColumnsBySheetId ??= new Map()
      let tracker = writtenColumnsBySheetId.get(sheetId)
      if (!tracker) {
        tracker = createInitialWrittenColumnTracker()
        writtenColumnsBySheetId.set(sheetId, tracker)
      }
      markInitialWrittenColumn(tracker, col)
    }
    const markCellColumn = (cellIndex: number): void => {
      const sheetId = args.state.workbook.cellStore.sheetIds[cellIndex]
      const col = args.state.workbook.cellStore.cols[cellIndex]
      if (sheetId === undefined || col === undefined) {
        return
      }
      markKnownColumn(sheetId, col)
    }
    const clearDerivedFlags = (cellIndex: number): void => {
      args.state.workbook.cellStore.flags[cellIndex] =
        (args.state.workbook.cellStore.flags[cellIndex] ?? 0) & ~(CellFlags.SpillChild | CellFlags.PivotOutput)
    }
    const writeNumberCore = (cellIndex: number, value: number): void => {
      const cellStore = args.state.workbook.cellStore
      clearDerivedFlags(cellIndex)
      cellStore.tags[cellIndex] = ValueTag.Number
      cellStore.errors[cellIndex] = ErrorCode.None
      cellStore.stringIds[cellIndex] = 0
      cellStore.numbers[cellIndex] = value
      cellStore.versions[cellIndex] = (cellStore.versions[cellIndex] ?? 0) + 1
    }
    const writeValueCore = (cellIndex: number, value: CellValue): void => {
      const cellStore = args.state.workbook.cellStore
      clearDerivedFlags(cellIndex)
      cellStore.tags[cellIndex] = value.tag
      cellStore.errors[cellIndex] = value.tag === ValueTag.Error ? value.code : ErrorCode.None
      cellStore.stringIds[cellIndex] = value.tag === ValueTag.String ? args.state.strings.intern(value.value) : 0
      cellStore.numbers[cellIndex] =
        value.tag === ValueTag.Number ? value.value : value.tag === ValueTag.Boolean ? (value.value ? 1 : 0) : 0
      cellStore.versions[cellIndex] = (cellStore.versions[cellIndex] ?? 0) + 1
    }
    return {
      writeValue(cellIndex, value) {
        writeValueCore(cellIndex, value)
        markCellColumn(cellIndex)
      },
      writeValueAt(cellIndex, sheetId, col, value) {
        writeValueCore(cellIndex, value)
        markKnownColumn(sheetId, col)
      },
      writeNumber(cellIndex, value) {
        writeNumberCore(cellIndex, value)
        markCellColumn(cellIndex)
      },
      writeNumberAt(cellIndex, sheetId, col, value) {
        writeNumberCore(cellIndex, value)
        markKnownColumn(sheetId, col)
      },
      flush() {
        writtenColumnsBySheetId?.forEach((tracker, sheetId) => {
          if (tracker.count > 0) {
            args.state.workbook.notifyColumnsWritten(sheetId, materializeInitialWrittenColumns(tracker))
          }
        })
      },
    }
  }

  const canEvaluateInitialDirectFormula = (cellIndex: number): boolean => {
    return canEvaluateInitialDirectRuntimeFormula(args.state.formulas.get(cellIndex))
  }

  const createInitialTemplateFormulaResolver = (): ((source: string, row: number, col: number) => FormulaTemplateResolution) => {
    const simpleTemplateCache = new Map<string, InitialTemplateFormulaCacheEntry>()
    return (source, row, col) => {
      const templateKey = tryBuildInitialSimpleRowRelativeBinaryTemplateKey(source, row, col)
      const cached = templateKey === undefined ? undefined : simpleTemplateCache.get(templateKey)
      if (cached) {
        const anchorRowDelta = row - cached.anchorRow
        const anchorColDelta = col - cached.anchorCol
        const compiled = translateSimpleDirectScalarFormula(cached.anchorCompiled, anchorRowDelta, anchorColDelta, source)
        if (compiled) {
          return {
            ...cached.resolution,
            compiled,
            translated: cached.resolution.translated || anchorRowDelta !== 0 || anchorColDelta !== 0,
            rowDelta: cached.resolution.rowDelta + anchorRowDelta,
            colDelta: cached.resolution.colDelta + anchorColDelta,
          }
        }
      }
      const resolution = args.compileTemplateFormula(source, row, col)
      if (templateKey !== undefined) {
        simpleTemplateCache.set(templateKey, {
          resolution,
          anchorRow: row,
          anchorCol: col,
          anchorCompiled: resolution.compiled,
        })
      }
      return resolution
    }
  }

  const evaluateInitialPrefixAggregateGroups = (
    orderedCellIndices: readonly number[],
    pushChangedCellIndex: (cellIndex: number) => void,
    writeFormulaValue: (cellIndex: number, value: CellValue) => void,
  ): Set<number> | undefined => {
    const groups = new Map<string, InitialPrefixAggregateGroup>()
    for (let index = 0; index < orderedCellIndices.length; index += 1) {
      const cellIndex = orderedCellIndices[index]!
      const formula = args.state.formulas.get(cellIndex)
      const aggregate = formula?.directAggregate
      if (!formula || !aggregate || aggregate.rowStart !== 0 || formula.dependencyIndices.length !== 0) {
        continue
      }
      const key = `${aggregate.sheetName}\t${aggregate.col}\t${aggregate.aggregateKind}`
      let group = groups.get(key)
      if (!group) {
        group = {
          sheetName: aggregate.sheetName,
          col: aggregate.col,
          aggregateKind: aggregate.aggregateKind,
          maxRowEnd: aggregate.rowEnd,
          lastRowEnd: aggregate.rowEnd,
          formulasAreOrdered: true,
          formulas: [],
        }
        groups.set(key, group)
      } else {
        group.maxRowEnd = Math.max(group.maxRowEnd, aggregate.rowEnd)
        if (aggregate.rowEnd < group.lastRowEnd) {
          group.formulasAreOrdered = false
        }
        group.lastRowEnd = aggregate.rowEnd
      }
      group.formulas.push({ cellIndex, rowEnd: aggregate.rowEnd })
    }
    if (groups.size === 0) {
      return undefined
    }

    const handled = new Set<number>()
    groups.forEach((group) => {
      const sheet = args.state.workbook.getSheet(group.sheetName)
      if (!sheet) {
        return
      }
      const formulas = group.formulasAreOrdered ? group.formulas : group.formulas.toSorted((left, right) => left.rowEnd - right.rowEnd)
      let sum = 0
      let count = 0
      let averageCount = 0
      let errorCode = ErrorCode.None
      let errorCount = 0
      let minimum = Number.POSITIVE_INFINITY
      let maximum = Number.NEGATIVE_INFINITY
      let formulaIndex = 0
      for (let row = 0; row <= group.maxRowEnd; row += 1) {
        const memberCellIndex = sheet.structureVersion === 1 ? sheet.grid.getPhysical(row, group.col) : sheet.grid.get(row, group.col)
        if (memberCellIndex !== -1) {
          if (((args.state.workbook.cellStore.flags[memberCellIndex] ?? 0) & CellFlags.HasFormula) !== 0) {
            break
          }
          const tag = (args.state.workbook.cellStore.tags[memberCellIndex] as ValueTag | undefined) ?? ValueTag.Empty
          if (tag === ValueTag.Number) {
            const numeric = args.state.workbook.cellStore.numbers[memberCellIndex] ?? 0
            sum += numeric
            count += 1
            averageCount += 1
            minimum = Math.min(minimum, numeric)
            maximum = Math.max(maximum, numeric)
          } else if (tag === ValueTag.Boolean) {
            const numeric = (args.state.workbook.cellStore.numbers[memberCellIndex] ?? 0) !== 0 ? 1 : 0
            sum += numeric
            count += 1
            averageCount += 1
            minimum = Math.min(minimum, numeric)
            maximum = Math.max(maximum, numeric)
          } else if (tag === ValueTag.Empty) {
            averageCount += 1
            minimum = Math.min(minimum, 0)
            maximum = Math.max(maximum, 0)
          } else if (tag === ValueTag.Error) {
            errorCode ||= (args.state.workbook.cellStore.errors[memberCellIndex] as ErrorCode | undefined) ?? ErrorCode.None
            errorCount += 1
          }
        }
        while (formulaIndex < formulas.length && formulas[formulaIndex]!.rowEnd <= row) {
          const formula = formulas[formulaIndex]!
          const value =
            group.aggregateKind === 'sum'
              ? errorCount > 0 && errorCode !== ErrorCode.None
                ? { tag: ValueTag.Error as const, code: errorCode }
                : { tag: ValueTag.Number as const, value: sum }
              : group.aggregateKind === 'count'
                ? { tag: ValueTag.Number as const, value: count }
                : group.aggregateKind === 'average'
                  ? errorCount > 0 && errorCode !== ErrorCode.None
                    ? { tag: ValueTag.Error as const, code: errorCode }
                    : averageCount === 0
                      ? { tag: ValueTag.Error as const, code: ErrorCode.Div0 }
                      : { tag: ValueTag.Number as const, value: sum / averageCount }
                  : group.aggregateKind === 'min'
                    ? { tag: ValueTag.Number as const, value: minimum === Number.POSITIVE_INFINITY ? 0 : minimum }
                    : { tag: ValueTag.Number as const, value: maximum === Number.NEGATIVE_INFINITY ? 0 : maximum }
          writeFormulaValue(formula.cellIndex, value)
          handled.add(formula.cellIndex)
          pushChangedCellIndex(formula.cellIndex)
          formulaIndex += 1
        }
      }
    })
    return handled.size === 0 ? undefined : handled
  }

  const coerceInitialDirectScalarCell = (
    cellIndex: number,
  ): { kind: 'number'; value: number } | { kind: 'error'; code: ErrorCode } | undefined => {
    const cellStore = args.state.workbook.cellStore
    const tag = (cellStore.tags[cellIndex] as ValueTag | undefined) ?? ValueTag.Empty
    switch (tag) {
      case ValueTag.Number:
        return { kind: 'number', value: cellStore.numbers[cellIndex] ?? 0 }
      case ValueTag.Boolean:
        return { kind: 'number', value: (cellStore.numbers[cellIndex] ?? 0) !== 0 ? 1 : 0 }
      case ValueTag.Empty:
        return { kind: 'number', value: 0 }
      case ValueTag.Error:
        return { kind: 'error', code: (cellStore.errors[cellIndex] as ErrorCode | undefined) ?? ErrorCode.None }
      case ValueTag.String:
        return { kind: 'error', code: ErrorCode.Value }
      default:
        return undefined
    }
  }

  const readInitialDirectScalarOperand = (
    operand: RuntimeDirectScalarOperand,
  ): { kind: 'number'; value: number } | { kind: 'error'; code: ErrorCode } | undefined => {
    switch (operand.kind) {
      case 'literal-number':
        return { kind: 'number', value: operand.value }
      case 'error':
        return { kind: 'error', code: operand.code }
      case 'cell':
        return coerceInitialDirectScalarCell(operand.cellIndex)
    }
  }

  const evaluateInitialDirectScalar = (directScalar: RuntimeDirectScalarDescriptor): CellValue | undefined => {
    if (directScalar.kind === 'abs') {
      const operand = readInitialDirectScalarOperand(directScalar.operand)
      if (!operand) {
        return undefined
      }
      return operand.kind === 'error'
        ? { tag: ValueTag.Error, code: operand.code }
        : { tag: ValueTag.Number, value: Math.abs(operand.value) }
    }
    const left = readInitialDirectScalarOperand(directScalar.left)
    const right = readInitialDirectScalarOperand(directScalar.right)
    if (!left || !right) {
      return undefined
    }
    if (left.kind === 'error') {
      return { tag: ValueTag.Error, code: left.code }
    }
    if (right.kind === 'error') {
      return { tag: ValueTag.Error, code: right.code }
    }
    switch (directScalar.operator) {
      case '+':
        return { tag: ValueTag.Number, value: left.value + right.value }
      case '-':
        return { tag: ValueTag.Number, value: left.value - right.value }
      case '*':
        return { tag: ValueTag.Number, value: left.value * right.value }
      case '/':
        return right.value === 0 ? { tag: ValueTag.Error, code: ErrorCode.Div0 } : { tag: ValueTag.Number, value: left.value / right.value }
    }
  }

  const coerceInitialDirectScalarNumber = (cellIndex: number): number | undefined => {
    const cellStore = args.state.workbook.cellStore
    const tag = (cellStore.tags[cellIndex] as ValueTag | undefined) ?? ValueTag.Empty
    switch (tag) {
      case ValueTag.Number:
        return cellStore.numbers[cellIndex] ?? 0
      case ValueTag.Boolean:
        return (cellStore.numbers[cellIndex] ?? 0) !== 0 ? 1 : 0
      case ValueTag.Empty:
        return 0
      case ValueTag.String:
      case ValueTag.Error:
        return undefined
      default:
        return undefined
    }
  }

  const readInitialDirectScalarNumberOperand = (operand: RuntimeDirectScalarOperand): number | undefined => {
    switch (operand.kind) {
      case 'literal-number':
        return operand.value
      case 'cell':
        return coerceInitialDirectScalarNumber(operand.cellIndex)
      case 'error':
        return undefined
    }
  }

  const evaluateInitialDirectScalarNumber = (directScalar: RuntimeDirectScalarDescriptor): number | undefined => {
    if (directScalar.kind === 'abs') {
      const operand = readInitialDirectScalarNumberOperand(directScalar.operand)
      return operand === undefined ? undefined : Math.abs(operand)
    }
    const left = readInitialDirectScalarNumberOperand(directScalar.left)
    const right = readInitialDirectScalarNumberOperand(directScalar.right)
    if (left === undefined || right === undefined) {
      return undefined
    }
    switch (directScalar.operator) {
      case '+':
        return left + right
      case '-':
        return left - right
      case '*':
        return left * right
      case '/':
        return right === 0 ? undefined : left / right
    }
  }

  const evaluateInitialDirectFormulas = (
    orderedCellIndices: readonly number[],
    options?: { readonly alreadyValidated?: boolean; readonly hasPrefixAggregateCandidates?: boolean },
  ): U32 | undefined => {
    if (
      orderedCellIndices.length === 0 ||
      orderedCellIndices.length > INITIAL_DIRECT_FORMULA_EVALUATION_LIMIT ||
      (options?.alreadyValidated !== true && !orderedCellIndices.every(canEvaluateInitialDirectFormula))
    ) {
      return undefined
    }
    let changedCellBuffer = new Uint32Array(Math.max(orderedCellIndices.length, 1))
    let changedCellCount = 0
    const pushChangedCellIndex = (cellIndex: number): void => {
      if (changedCellCount === changedCellBuffer.length) {
        const next = new Uint32Array(changedCellBuffer.length * 2)
        next.set(changedCellBuffer)
        changedCellBuffer = next
      }
      changedCellBuffer[changedCellCount] = cellIndex
      changedCellCount += 1
    }
    const valueWriter = createInitialFormulaValueWriter()
    args.state.workbook.withBatchedColumnVersionUpdates(() => {
      const prefixAggregateHandled =
        options?.hasPrefixAggregateCandidates === true
          ? evaluateInitialPrefixAggregateGroups(orderedCellIndices, pushChangedCellIndex, valueWriter.writeValue)
          : undefined
      for (let index = 0; index < orderedCellIndices.length; index += 1) {
        const cellIndex = orderedCellIndices[index]!
        if (prefixAggregateHandled?.has(cellIndex)) {
          continue
        }
        const formula = args.state.formulas.get(cellIndex)
        if (formula?.directScalar !== undefined) {
          const numericValue = evaluateInitialDirectScalarNumber(formula.directScalar)
          if (numericValue !== undefined) {
            valueWriter.writeNumber(cellIndex, numericValue)
            pushChangedCellIndex(cellIndex)
            continue
          }
          const fallbackValue = evaluateInitialDirectScalar(formula.directScalar)
          if (fallbackValue !== undefined) {
            valueWriter.writeValue(cellIndex, fallbackValue)
            pushChangedCellIndex(cellIndex)
            continue
          }
        }
        const changedSpillIndices = args.evaluateDirectFormula(cellIndex)
        pushChangedCellIndex(cellIndex)
        if (changedSpillIndices) {
          for (let spillIndex = 0; spillIndex < changedSpillIndices.length; spillIndex += 1) {
            pushChangedCellIndex(changedSpillIndices[spillIndex]!)
          }
        }
      }
      valueWriter.flush()
    })
    const changedCellIndices = changedCellBuffer.subarray(0, changedCellCount)
    args.deferKernelSync(changedCellIndices)
    addEngineCounter(args.state.counters, 'directFormulaInitialEvaluations', orderedCellIndices.length)
    return changedCellIndices
  }

  const initializeFormulaEntriesNow = <Entry>(
    refs: readonly Entry[],
    potentialNewCells: number | undefined,
    resolveCellIndex: (ref: Entry) => number,
    resolveEntry: (
      ref: Entry,
      cellIndex: number,
    ) => {
      cellIndex: number
      sheetId: number
      row: number
      col: number
      ownerSheetName: string
      source: string
      compiled: CompiledFormula
      templateId?: number
    },
  ): void => {
    if (refs.length === 0) {
      return
    }

    args.beginMutationCollection()
    let hadCycleMembersBefore: boolean | undefined
    const hadCycleMembersBeforeNow = (): boolean => (hadCycleMembersBefore ??= hasCycleMembersNow())
    let changedInputCount = 0
    let formulaChangedCount = 0
    let topologyChanged = false
    let compileMs = 0
    const reservedNewCells = Math.max(potentialNewCells ?? refs.length, refs.length)
    const hadExistingFormulas = args.state.formulas.size > 0
    args.state.workbook.cellStore.ensureCapacity(args.state.workbook.cellStore.size + reservedNewCells)
    args.ensureRecalcScratchCapacity(args.state.workbook.cellStore.capacity + 1)
    args.resetMaterializedCellScratch(reservedNewCells)
    const targetCellIndices = hadExistingFormulas ? EMPTY_U32 : new Uint32Array(refs.length)
    let maxTargetCellIndex = 0
    if (!hadExistingFormulas) {
      for (let index = 0; index < refs.length; index += 1) {
        const cellIndex = resolveCellIndex(refs[index]!)
        targetCellIndices[index] = cellIndex
        if (cellIndex > maxTargetCellIndex) {
          maxTargetCellIndex = cellIndex
        }
      }
    }
    const pendingFormulaCells = hadExistingFormulas ? undefined : new Uint8Array(maxTargetCellIndex + 1)
    if (pendingFormulaCells) {
      for (let index = 0; index < targetCellIndices.length; index += 1) {
        pendingFormulaCells[targetCellIndices[index]!] = 1
      }
    }
    let canAssignTopoInBatch = !hadExistingFormulas
    let nextTopoRank = 0
    const orderedPreparedCellIndices: number[] = []
    let canUseInitialDirectEvaluation = false
    let allPreparedFormulasCanUseInitialDirectEvaluation = true
    let hasInitialPrefixAggregateCandidates = false
    let inlineInitialDirectScalarWriter: InitialFormulaValueWriter | undefined
    let inlineInitialDirectScalarCellBuffer = new Uint32Array(Math.max(refs.length, 1))
    let inlineInitialDirectScalarCellCount = 0
    const deferredFormulaFamilyRuns = hadExistingFormulas ? undefined : new Map<string, DeferredInitialFormulaFamilyRun>()

    const pushInlineInitialDirectScalarCell = (cellIndex: number): void => {
      if (inlineInitialDirectScalarCellCount === inlineInitialDirectScalarCellBuffer.length) {
        const next = new Uint32Array(inlineInitialDirectScalarCellBuffer.length * 2)
        next.set(inlineInitialDirectScalarCellBuffer)
        inlineInitialDirectScalarCellBuffer = next
      }
      inlineInitialDirectScalarCellBuffer[inlineInitialDirectScalarCellCount] = cellIndex
      inlineInitialDirectScalarCellCount += 1
    }

    const tryInlineInitialDirectScalarEvaluation = (
      prepared: { cellIndex: number; sheetId: number; col: number },
      runtimeFormula: RuntimeFormula | undefined,
    ): void => {
      if (
        hadExistingFormulas ||
        !canAssignTopoInBatch ||
        !runtimeFormula ||
        runtimeFormula.compiled.volatile ||
        runtimeFormula.compiled.producesSpill ||
        runtimeFormula.directScalar === undefined
      ) {
        return
      }
      const numericValue = evaluateInitialDirectScalarNumber(runtimeFormula.directScalar)
      inlineInitialDirectScalarWriter ??= createInitialFormulaValueWriter()
      if (numericValue !== undefined) {
        inlineInitialDirectScalarWriter.writeNumberAt(prepared.cellIndex, prepared.sheetId, prepared.col, numericValue)
        pushInlineInitialDirectScalarCell(prepared.cellIndex)
        return
      }
      const fallbackValue = evaluateInitialDirectScalar(runtimeFormula.directScalar)
      if (fallbackValue !== undefined) {
        inlineInitialDirectScalarWriter.writeValueAt(prepared.cellIndex, prepared.sheetId, prepared.col, fallbackValue)
        pushInlineInitialDirectScalarCell(prepared.cellIndex)
      }
    }

    const deferFormulaFamilyRegistration = (prepared: { cellIndex: number; sheetId: number; row: number; col: number }): void => {
      if (!deferredFormulaFamilyRuns) {
        return
      }
      const runtimeFormula = args.state.formulas.get(prepared.cellIndex)
      const templateId = runtimeFormula?.templateId
      if (runtimeFormula === undefined || templateId === undefined) {
        return
      }
      const familyKey = `${prepared.sheetId}\t${templateId}\t${prepared.col}`
      let run = deferredFormulaFamilyRuns.get(familyKey)
      if (!run) {
        const shapeKey = initialFormulaFamilyShapeKey(runtimeFormula)
        run = {
          sheetId: prepared.sheetId,
          templateId,
          shapeKey,
          members: [],
        }
        deferredFormulaFamilyRuns.set(familyKey, run)
      }
      run.members.push({
        cellIndex: prepared.cellIndex,
        row: prepared.row,
        col: prepared.col,
      })
    }

    args.setBatchMutationDepth(args.getBatchMutationDepth() + 1)
    try {
      args.clearTemplateFormulaCache()
      const compileStarted = performance.now()
      const bindFormulaEntries = (): void => {
        args.state.workbook.withBatchedColumnVersionUpdates(() => {
          for (let refIndex = 0; refIndex < refs.length; refIndex += 1) {
            const ref = refs[refIndex]!
            const cellIndex = hadExistingFormulas ? resolveCellIndex(ref) : targetCellIndices[refIndex]!
            try {
              const prepared = resolveEntry(ref, cellIndex)
              args.bindPreparedFormula(
                prepared.cellIndex,
                prepared.ownerSheetName,
                prepared.source,
                prepared.compiled,
                prepared.templateId,
                {
                  deferFamilyRegistration: deferredFormulaFamilyRuns !== undefined,
                },
              )
              deferFormulaFamilyRegistration(prepared)
              formulaChangedCount = args.markFormulaChanged(prepared.cellIndex, formulaChangedCount)
              topologyChanged = true
              orderedPreparedCellIndices.push(prepared.cellIndex)
              if (canAssignTopoInBatch && pendingFormulaCells) {
                const runtimeFormula = args.state.formulas.get(prepared.cellIndex)
                if (!canEvaluateInitialDirectRuntimeFormula(runtimeFormula)) {
                  allPreparedFormulasCanUseInitialDirectEvaluation = false
                }
                if (runtimeFormula?.directAggregate !== undefined) {
                  hasInitialPrefixAggregateCandidates = true
                }
                if (!runtimeFormula || hasPendingFormulaDependency(runtimeFormula, pendingFormulaCells)) {
                  canAssignTopoInBatch = false
                } else {
                  args.state.workbook.cellStore.topoRanks[prepared.cellIndex] = nextTopoRank
                  nextTopoRank += 1
                  tryInlineInitialDirectScalarEvaluation(prepared, runtimeFormula)
                }
              }
            } catch {
              topologyChanged = args.removeFormula(cellIndex) || topologyChanged
              args.setInvalidFormulaValue(cellIndex)
              changedInputCount = args.markInputChanged(cellIndex, changedInputCount)
            }
            if (pendingFormulaCells) {
              pendingFormulaCells[cellIndex] = 0
            }
          }
          if (args.state.ranges.size > 0) {
            const reboundCount = formulaChangedCount
            formulaChangedCount = args.syncDynamicRanges(formulaChangedCount)
            topologyChanged = topologyChanged || formulaChangedCount !== reboundCount
          }
          deferredFormulaFamilyRuns?.forEach((run) => {
            args.upsertFormulaFamilyRun(run)
          })
          inlineInitialDirectScalarWriter?.flush()
        })
      }
      bindFormulaEntries()
      canUseInitialDirectEvaluation =
        canAssignTopoInBatch &&
        !hadExistingFormulas &&
        changedInputCount === 0 &&
        allPreparedFormulasCanUseInitialDirectEvaluation &&
        orderedPreparedCellIndices.length <= INITIAL_DIRECT_FORMULA_EVALUATION_LIMIT &&
        orderedPreparedCellIndices.length === refs.length
      compileMs += performance.now() - compileStarted
    } finally {
      args.setBatchMutationDepth(args.getBatchMutationDepth() - 1)
    }

    if (topologyChanged && !(canAssignTopoInBatch && !hadExistingFormulas)) {
      const repaired =
        !hadCycleMembersBeforeNow() &&
        formulaChangedCount > 0 &&
        args.repairTopoRanks(args.getChangedFormulaBuffer().subarray(0, formulaChangedCount))
      if (!repaired) {
        args.rebuildTopoRanks()
        args.detectCycles()
        args.state.formulas.forEach((_formula, cellIndex) => {
          if (((args.state.workbook.cellStore.flags[cellIndex] ?? 0) & CellFlags.InCycle) !== 0) {
            changedInputCount = args.markInputChanged(cellIndex, changedInputCount)
          }
        })
      }
    }
    formulaChangedCount = args.markVolatileFormulasChanged(formulaChangedCount)
    const useInitialDirectEvaluation = canUseInitialDirectEvaluation && formulaChangedCount === orderedPreparedCellIndices.length
    if (!useInitialDirectEvaluation) {
      args.prepareRegionQueryIndices()
    }
    let recalculated: U32
    if (
      useInitialDirectEvaluation &&
      inlineInitialDirectScalarCellCount === orderedPreparedCellIndices.length &&
      !hasInitialPrefixAggregateCandidates
    ) {
      recalculated = inlineInitialDirectScalarCellBuffer.subarray(0, inlineInitialDirectScalarCellCount)
      args.deferKernelSync(recalculated)
    } else if (useInitialDirectEvaluation) {
      const direct = evaluateInitialDirectFormulas(orderedPreparedCellIndices, {
        alreadyValidated: true,
        hasPrefixAggregateCandidates: hasInitialPrefixAggregateCandidates,
      })
      if (direct) {
        recalculated = direct
      } else {
        const changedInputArray = args.getChangedInputBuffer().subarray(0, changedInputCount)
        const changedRoots = args.composeMutationRoots(changedInputCount, formulaChangedCount)
        recalculated = args.recalculate(changedRoots, changedInputArray)
      }
    } else {
      const changedInputArray = args.getChangedInputBuffer().subarray(0, changedInputCount)
      const changedRoots = args.composeMutationRoots(changedInputCount, formulaChangedCount)
      recalculated =
        canAssignTopoInBatch && !hadExistingFormulas && orderedPreparedCellIndices.length > 0
          ? args.recalculatePreordered(changedRoots, orderedPreparedCellIndices, orderedPreparedCellIndices.length, changedInputArray)
          : args.recalculate(changedRoots, changedInputArray)
    }
    recalculated = args.reconcilePivotOutputs(recalculated, false)
    void recalculated
    const lastMetrics = args.state.getLastMetrics()
    args.state.setLastMetrics({
      ...lastMetrics,
      batchId: lastMetrics.batchId + 1,
      changedInputCount: changedInputCount + formulaChangedCount,
      compileMs,
    })
  }

  const initializeCellFormulasAtNow = (refs: readonly EngineCellMutationRef[], potentialNewCells?: number): void => {
    const resolveInitialTemplateFormula = createInitialTemplateFormulaResolver()
    initializeFormulaEntriesNow(
      refs,
      potentialNewCells,
      (ref) => {
        if (ref.mutation.kind !== 'setCellFormula') {
          throw new Error('initializeCellFormulasAt only supports setCellFormula coordinate mutations')
        }
        return ref.cellIndex ?? args.ensureCellTrackedByCoords(ref.sheetId, ref.mutation.row, ref.mutation.col)
      },
      (ref, cellIndex) => {
        if (ref.mutation.kind !== 'setCellFormula') {
          throw new Error('initializeCellFormulasAt only supports setCellFormula coordinate mutations')
        }
        const ownerSheetName = resolveSheetName(ref.sheetId)
        const template = resolveInitialTemplateFormula(ref.mutation.formula, ref.mutation.row, ref.mutation.col)
        return {
          cellIndex,
          sheetId: ref.sheetId,
          row: ref.mutation.row,
          col: ref.mutation.col,
          ownerSheetName,
          source: ref.mutation.formula,
          compiled: template.compiled,
          templateId: template.templateId,
        }
      },
    )
  }

  const initializeFormulaSourcesAtNow = (refs: readonly EngineFormulaSourceRef[], potentialNewCells?: number): void => {
    const resolveInitialTemplateFormula = createInitialTemplateFormulaResolver()
    initializeFormulaEntriesNow(
      refs,
      potentialNewCells,
      (ref) => ref.cellIndex ?? args.ensureCellTrackedByCoords(ref.sheetId, ref.row, ref.col),
      (ref, cellIndex) => {
        const ownerSheetName = resolveSheetName(ref.sheetId)
        const template = resolveInitialTemplateFormula(ref.source, ref.row, ref.col)
        return {
          cellIndex,
          sheetId: ref.sheetId,
          row: ref.row,
          col: ref.col,
          ownerSheetName,
          source: ref.source,
          compiled: template.compiled,
          templateId: template.templateId,
        }
      },
    )
  }

  const initializePreparedCellFormulasAtNow = (refs: readonly PreparedFormulaInitializationRef[], potentialNewCells?: number): void => {
    initializeFormulaEntriesNow(
      refs,
      potentialNewCells,
      (ref) => ref.cellIndex ?? args.ensureCellTrackedByCoords(ref.sheetId, ref.row, ref.col),
      (ref, cellIndex) => ({
        cellIndex,
        sheetId: ref.sheetId,
        row: ref.row,
        col: ref.col,
        ownerSheetName: resolveSheetName(ref.sheetId),
        source: ref.source,
        compiled: ref.compiled,
        ...(ref.templateId !== undefined ? { templateId: ref.templateId } : {}),
      }),
    )
  }

  const initializeHydratedPreparedCellFormulasAtNow = (
    refs: readonly HydratedPreparedFormulaInitializationRef[],
    potentialNewCells?: number,
  ): void => {
    if (refs.length === 0) {
      return
    }

    args.beginMutationCollection()
    let hadCycleMembersBefore: boolean | undefined
    const hadCycleMembersBeforeNow = (): boolean => (hadCycleMembersBefore ??= hasCycleMembersNow())
    let topologyChanged = false
    let compileMs = 0
    const reservedNewCells = Math.max(potentialNewCells ?? refs.length, refs.length)
    const hadExistingFormulas = args.state.formulas.size > 0
    args.state.workbook.cellStore.ensureCapacity(args.state.workbook.cellStore.size + reservedNewCells)
    args.ensureRecalcScratchCapacity(args.state.workbook.cellStore.capacity + 1)
    args.resetMaterializedCellScratch(reservedNewCells)
    const targetCellIndices = hadExistingFormulas
      ? []
      : refs.map((ref) => ref.cellIndex ?? args.ensureCellTrackedByCoords(ref.sheetId, ref.row, ref.col))
    const pendingFormulaCells = hadExistingFormulas
      ? undefined
      : new Uint8Array(args.state.workbook.cellStore.capacity + reservedNewCells + 1)
    if (pendingFormulaCells) {
      for (let index = 0; index < targetCellIndices.length; index += 1) {
        pendingFormulaCells[targetCellIndices[index]!] = 1
      }
    }
    let canAssignTopoInBatch = !hadExistingFormulas
    let nextTopoRank = 0
    const deferredFormulaFamilyRuns = hadExistingFormulas ? undefined : new Map<string, DeferredInitialFormulaFamilyRun>()

    const deferHydratedFormulaFamilyRegistration = (prepared: {
      readonly cellIndex: number
      readonly sheetId: number
      readonly row: number
      readonly col: number
    }): void => {
      if (!deferredFormulaFamilyRuns) {
        return
      }
      const runtimeFormula = args.state.formulas.get(prepared.cellIndex)
      const templateId = runtimeFormula?.templateId
      if (runtimeFormula === undefined || templateId === undefined) {
        return
      }
      const familyKey = `${prepared.sheetId}\t${templateId}\t${prepared.col}`
      let run = deferredFormulaFamilyRuns.get(familyKey)
      if (!run) {
        run = {
          sheetId: prepared.sheetId,
          templateId,
          shapeKey: initialFormulaFamilyShapeKey(runtimeFormula),
          members: [],
        }
        deferredFormulaFamilyRuns.set(familyKey, run)
      }
      run.members.push({
        cellIndex: prepared.cellIndex,
        row: prepared.row,
        col: prepared.col,
      })
    }

    args.setBatchMutationDepth(args.getBatchMutationDepth() + 1)
    try {
      args.clearTemplateFormulaCache()
      const compileStarted = performance.now()
      const valueWriter = createInitialFormulaValueWriter()
      const bindFormulaEntries = (): void => {
        args.state.workbook.withBatchedColumnVersionUpdates(() => {
          for (let refIndex = 0; refIndex < refs.length; refIndex += 1) {
            const ref = refs[refIndex]!
            const cellIndex = hadExistingFormulas
              ? (ref.cellIndex ?? args.ensureCellTrackedByCoords(ref.sheetId, ref.row, ref.col))
              : targetCellIndices[refIndex]!
            const ownerSheetName = resolveSheetName(ref.sheetId)
            topologyChanged =
              args.bindPreparedFormula(cellIndex, ownerSheetName, ref.source, ref.compiled, ref.templateId, {
                deferFamilyRegistration: deferredFormulaFamilyRuns !== undefined,
              }) || topologyChanged
            deferHydratedFormulaFamilyRegistration({
              cellIndex,
              sheetId: ref.sheetId,
              row: ref.row,
              col: ref.col,
            })
            valueWriter.writeValueAt(cellIndex, ref.sheetId, ref.col, ref.value)
            if (canAssignTopoInBatch && pendingFormulaCells) {
              const runtimeFormula = args.state.formulas.get(cellIndex)
              if (!runtimeFormula || hasPendingFormulaDependency(runtimeFormula, pendingFormulaCells)) {
                canAssignTopoInBatch = false
              } else {
                args.state.workbook.cellStore.topoRanks[cellIndex] = nextTopoRank
                nextTopoRank += 1
              }
            }
            if (pendingFormulaCells) {
              pendingFormulaCells[cellIndex] = 0
            }
          }
          deferredFormulaFamilyRuns?.forEach((run) => {
            args.upsertFormulaFamilyRun(run)
          })
          valueWriter.flush()
        })
      }
      bindFormulaEntries()
      compileMs += performance.now() - compileStarted
    } finally {
      args.setBatchMutationDepth(args.getBatchMutationDepth() - 1)
    }

    if (topologyChanged && !(canAssignTopoInBatch && !hadExistingFormulas)) {
      const repaired =
        !hadCycleMembersBeforeNow() &&
        refs.length > 0 &&
        args.repairTopoRanks(
          Uint32Array.from(
            targetCellIndices.length > 0
              ? targetCellIndices
              : refs.map((ref) => ref.cellIndex ?? args.ensureCellTrackedByCoords(ref.sheetId, ref.row, ref.col)),
          ),
        )
      if (!repaired) {
        args.rebuildTopoRanks()
        args.detectCycles()
      }
    }
    const lastMetrics = args.state.getLastMetrics()
    args.state.setLastMetrics({
      ...lastMetrics,
      batchId: lastMetrics.batchId + 1,
      changedInputCount: 0,
      compileMs,
      recalcMs: 0,
    })
  }

  return {
    initializeCellFormulasAt(refs, potentialNewCells) {
      return Effect.try({
        try: () => {
          initializeCellFormulasAtNow(refs, potentialNewCells)
        },
        catch: (cause) =>
          new EngineMutationError({
            message: mutationErrorMessage('Failed to initialize cell formulas', cause),
            cause,
          }),
      })
    },
    initializePreparedCellFormulasAt(refs, potentialNewCells) {
      return Effect.try({
        try: () => {
          initializePreparedCellFormulasAtNow(refs, potentialNewCells)
        },
        catch: (cause) =>
          new EngineMutationError({
            message: mutationErrorMessage('Failed to initialize prepared cell formulas', cause),
            cause,
          }),
      })
    },
    initializeHydratedPreparedCellFormulasAt(refs, potentialNewCells) {
      return Effect.try({
        try: () => {
          initializeHydratedPreparedCellFormulasAtNow(refs, potentialNewCells)
        },
        catch: (cause) =>
          new EngineMutationError({
            message: mutationErrorMessage('Failed to initialize hydrated prepared cell formulas', cause),
            cause,
          }),
      })
    },
    initializeCellFormulasAtNow,
    initializeFormulaSourcesAtNow,
    initializePreparedCellFormulasAtNow,
    initializeHydratedPreparedCellFormulasAtNow,
  }
}
