import { FormulaMode } from '@bilig/protocol'
import type { EdgeArena } from '../../edge-arena.js'
import { appendDirectAggregateColumnReverseEdges } from './formula-binding-dependency-helpers.js'
import type { FormulaBindingMemberCounts } from './formula-binding-member-counts.js'
import { markFormulaCellBound } from './formula-binding-cell-flags.js'
import { makeUnmanagedCompiledPlan } from './formula-binding-plan-helpers.js'
import type {
  CreateEngineFormulaBindingServiceArgs,
  FreshDirectAggregateFormulaBindingInput,
  FreshDirectAggregateFormulaBindingMember,
} from './formula-binding-service-types.js'
import type { RuntimeDirectAggregateDescriptor, RuntimeFormula } from '../runtime-state.js'

const EMPTY_U32 = new Uint32Array(0)

export function bindFreshDirectAggregateFormulaRun(args: {
  readonly serviceArgs: CreateEngineFormulaBindingServiceArgs
  readonly edgeArena: EdgeArena
  readonly formulaMemberCounts: FormulaBindingMemberCounts
  readonly trackFormulaSheetIndexes: (
    cellIndex: number,
    ownerSheetName: string,
    compiled: Pick<RuntimeFormula['compiled'], 'deps' | 'parsedDeps'>,
  ) => void
  readonly run: FreshDirectAggregateFormulaBindingInput
}): void {
  if ('member' in args.run) {
    assertFreshDirectAggregateFormulaCell(args.serviceArgs, args.run.cellIndex, args.run.member)
    bindFreshDirectAggregateFormulaMember(
      args.serviceArgs,
      args.edgeArena,
      args.formulaMemberCounts,
      args.trackFormulaSheetIndexes,
      args.run.sheetId,
      args.run.ownerSheetName,
      args.run.cellIndex,
      args.run.member,
    )
    return
  }

  if (args.run.cellIndices.length !== args.run.members.length) {
    throw new Error('Expected fresh direct aggregate formula cell index count to match member count')
  }
  for (let index = 0; index < args.run.members.length; index += 1) {
    assertFreshDirectAggregateFormulaCell(args.serviceArgs, args.run.cellIndices[index]!, args.run.members[index]!)
  }

  for (let index = 0; index < args.run.members.length; index += 1) {
    bindFreshDirectAggregateFormulaMember(
      args.serviceArgs,
      args.edgeArena,
      args.formulaMemberCounts,
      args.trackFormulaSheetIndexes,
      args.run.sheetId,
      args.run.ownerSheetName,
      args.run.cellIndices[index]!,
      args.run.members[index]!,
    )
  }
}

function assertFreshDirectAggregateFormulaCell(
  serviceArgs: CreateEngineFormulaBindingServiceArgs,
  cellIndex: number,
  member: FreshDirectAggregateFormulaBindingMember,
): void {
  if ((serviceArgs.state.workbook.cellStore.formulaIds[cellIndex] ?? 0) !== 0) {
    throw new Error('Expected fresh direct aggregate formula cell')
  }
  if (
    member.aggregateRowStart > member.aggregateRowEnd ||
    member.aggregateColStart > member.aggregateColEnd ||
    (member.aggregateRowStart <= member.row &&
      member.row <= member.aggregateRowEnd &&
      member.aggregateColStart <= member.col &&
      member.col <= member.aggregateColEnd)
  ) {
    throw new Error('Expected non-recursive fresh direct aggregate formula')
  }
}

function bindFreshDirectAggregateFormulaMember(
  serviceArgs: CreateEngineFormulaBindingServiceArgs,
  edgeArena: EdgeArena,
  formulaMemberCounts: FormulaBindingMemberCounts,
  trackFormulaSheetIndexes: (
    cellIndex: number,
    ownerSheetName: string,
    compiled: Pick<RuntimeFormula['compiled'], 'deps' | 'parsedDeps'>,
  ) => void,
  sheetId: number,
  ownerSheetName: string,
  cellIndex: number,
  member: FreshDirectAggregateFormulaBindingMember,
): void {
  const directAggregate = buildFreshDirectAggregateDescriptor(serviceArgs, ownerSheetName, member)
  const runtimeFormula: RuntimeFormula = {
    cellIndex,
    formulaSlotId: 0,
    planId: 0,
    templateId: member.templateId,
    source: member.source,
    compiled: member.compiled,
    plan: makeUnmanagedCompiledPlan(member.source, member.compiled, member.templateId),
    dependencyIndices: EMPTY_U32,
    dependencyEntities: edgeArena.empty(),
    rangeDependencies: EMPTY_U32,
    graphRangeDependencies: EMPTY_U32,
    runtimeProgram: EMPTY_U32,
    constants: member.compiled.constants,
    structuralSourceTransform: undefined,
    programOffset: 0,
    programLength: 0,
    constNumberOffset: 0,
    constNumberLength: member.compiled.constants.length,
    rangeListOffset: 0,
    rangeListLength: 0,
    directLookup: undefined,
    directAggregate,
    directScalar: undefined,
    directCriteria: undefined,
  }
  const formulaSlotId = serviceArgs.state.formulas.set(cellIndex, runtimeFormula)
  runtimeFormula.formulaSlotId = formulaSlotId
  formulaMemberCounts.increment(sheetId, member.col)
  markFormulaCellBound(serviceArgs.state.workbook.cellStore, cellIndex, member.compiled.mode)
  appendDirectAggregateColumnReverseEdges(
    serviceArgs.reverseState.reverseAggregateColumnEdges,
    serviceArgs.state.workbook,
    directAggregate,
    cellIndex,
  )
  trackFormulaSheetIndexes(cellIndex, ownerSheetName, member.compiled)
  if (member.compiled.mode === FormulaMode.WasmFastPath && member.compiled.program.length > 0) {
    serviceArgs.scheduleWasmProgramSync()
  }
}

function buildFreshDirectAggregateDescriptor(
  serviceArgs: CreateEngineFormulaBindingServiceArgs,
  ownerSheetName: string,
  member: FreshDirectAggregateFormulaBindingMember,
): RuntimeDirectAggregateDescriptor {
  return {
    regionId: serviceArgs.regionGraph.internSingleColumnRegion({
      sheetName: ownerSheetName,
      rowStart: member.aggregateRowStart,
      rowEnd: member.aggregateRowEnd,
      col: member.aggregateColStart,
    }),
    aggregateKind: member.aggregateKind,
    sheetName: ownerSheetName,
    rowStart: member.aggregateRowStart,
    rowEnd: member.aggregateRowEnd,
    col: member.aggregateColStart,
    colEnd: member.aggregateColEnd,
    length: (member.aggregateRowEnd - member.aggregateRowStart + 1) * (member.aggregateColEnd - member.aggregateColStart + 1),
    ...(member.resultOffset !== undefined ? { resultOffset: member.resultOffset } : {}),
  }
}
