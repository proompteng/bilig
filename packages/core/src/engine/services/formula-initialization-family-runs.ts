import type {
  FormulaFamilyFreshUniformRunRegistrationArgs,
  FormulaFamilyRunAxis,
  FormulaFamilyMember,
  FormulaFamilyRunUpsertArgs,
  FormulaFamilyStore,
  FormulaFamilyStructuralSourceTransform,
} from '../../formula/formula-family-store.js'
import type { RuntimeFormula } from '../runtime-state.js'
import type { FormulaBindingFamilyShapeKeyCache } from './formula-binding-family-shape-key.js'
import type { InitialFormulaEntryRefSource } from './formula-initialization-refs.js'
import { initialFormulaFamilyShapeKey } from './formula-initialization-template-keys.js'

export type DeferredInitialFormulaFamilyRun = Omit<FormulaFamilyRunUpsertArgs, 'members'> & {
  axis: FormulaFamilyRunAxis
  fixedIndex: number
  start: number
  step: number
  lastIndex: number
  ordered: boolean
  cellIndices: number[]
  rows?: number[]
  cols?: number[]
}

export type DeferredInitialFormulaFamilyRunMap = Map<number, Map<number, Map<number, DeferredInitialFormulaFamilyRun>>>

export function createDeferredInitialFormulaFamilyRunMap(): DeferredInitialFormulaFamilyRunMap {
  return new Map()
}

function getDeferredFormulaFamilyRun(
  runs: DeferredInitialFormulaFamilyRunMap,
  sheetId: number,
  templateId: number,
  col: number,
): DeferredInitialFormulaFamilyRun | undefined {
  return runs.get(sheetId)?.get(templateId)?.get(col)
}

function setDeferredFormulaFamilyRun(
  runs: DeferredInitialFormulaFamilyRunMap,
  sheetId: number,
  templateId: number,
  col: number,
  run: DeferredInitialFormulaFamilyRun,
): void {
  let templateRuns = runs.get(sheetId)
  if (!templateRuns) {
    templateRuns = new Map()
    runs.set(sheetId, templateRuns)
  }
  let columnRuns = templateRuns.get(templateId)
  if (!columnRuns) {
    columnRuns = new Map()
    templateRuns.set(templateId, columnRuns)
  }
  columnRuns.set(col, run)
}

export function collectDeferredInitialFormulaFamilyRuns(runs: DeferredInitialFormulaFamilyRunMap): DeferredInitialFormulaFamilyRun[] {
  const result: DeferredInitialFormulaFamilyRun[] = []
  runs.forEach((templateRuns) => {
    templateRuns.forEach((columnRuns) => {
      columnRuns.forEach((run) => {
        result.push(run)
      })
    })
  })
  return result
}

export function flushDeferredInitialFormulaFamilyRuns(args: {
  readonly runs: DeferredInitialFormulaFamilyRunMap | undefined
  readonly shouldDeferFormulaFamilyIndex: boolean
  readonly deferFormulaFamilyIndexRuns?: ((runs: readonly DeferredInitialFormulaFamilyRun[]) => void) | undefined
  readonly deferFormulaFamilyIndexRebuild?: (() => void) | undefined
  readonly registerFormulaFamilyRun: (run: DeferredInitialFormulaFamilyRun) => void
  readonly checkEvaluationBudget?: (() => void) | undefined
}): void {
  if (args.shouldDeferFormulaFamilyIndex) {
    if (args.runs) {
      args.deferFormulaFamilyIndexRuns?.(collectDeferredInitialFormulaFamilyRuns(args.runs))
    } else {
      args.deferFormulaFamilyIndexRebuild?.()
    }
  } else if (args.runs) {
    collectDeferredInitialFormulaFamilyRuns(args.runs).forEach((run) => {
      args.checkEvaluationBudget?.()
      args.registerFormulaFamilyRun(run)
    })
  }
}

export function materializeDeferredFormulaFamilyRunMembers(run: DeferredInitialFormulaFamilyRun): FormulaFamilyMember[] {
  const step = run.cellIndices.length <= 1 ? 1 : run.step
  return run.cellIndices.map((cellIndex, index) => ({
    cellIndex,
    row: run.axis === 'row' ? (run.ordered ? run.start + step * index : run.rows![index]!) : run.fixedIndex,
    col: run.axis === 'row' ? run.fixedIndex : run.ordered ? run.start + step * index : run.cols![index]!,
  }))
}

export function readFreshFormulaFamilyRunsFromRefs<Entry>(refs: InitialFormulaEntryRefSource<Entry>):
  | {
      readonly runs: readonly DeferredInitialFormulaFamilyRun[] | undefined
      readonly fallbackCount: number
    }
  | undefined {
  if (!hasFreshFormulaFamilyRuns(refs)) {
    return undefined
  }
  return {
    runs: refs.freshFormulaFamilyRuns,
    fallbackCount: refs.freshFormulaFamilyRunFallbackCount ?? 0,
  }
}

function hasFreshFormulaFamilyRuns(value: unknown): value is {
  readonly freshFormulaFamilyRuns: readonly DeferredInitialFormulaFamilyRun[] | undefined
  readonly freshFormulaFamilyRunFallbackCount?: number
} {
  return typeof value === 'object' && value !== null && 'freshFormulaFamilyRuns' in value
}

export function noteDeferredFormulaFamilyRunMember(args: {
  readonly runs: DeferredInitialFormulaFamilyRunMap | undefined
  readonly runtimeFormula: RuntimeFormula | undefined
  readonly prepared: {
    readonly cellIndex: number
    readonly sheetId: number
    readonly row: number
    readonly col: number
    readonly templateId?: number
  }
}): void {
  const { prepared, runs } = args
  const templateId = prepared.templateId
  if (!runs || templateId === undefined) {
    return
  }
  let run = getDeferredFormulaFamilyRun(runs, prepared.sheetId, templateId, prepared.col)
  if (!run) {
    if (args.runtimeFormula === undefined) {
      return
    }
    run = {
      sheetId: prepared.sheetId,
      templateId,
      shapeKey: initialFormulaFamilyShapeKey(args.runtimeFormula),
      axis: 'row',
      fixedIndex: prepared.col,
      start: prepared.row,
      step: 0,
      lastIndex: prepared.row,
      ordered: true,
      cellIndices: [],
    }
    setDeferredFormulaFamilyRun(runs, prepared.sheetId, templateId, prepared.col, run)
  } else {
    const nextStep = prepared.row - run.lastIndex
    let breaksOrder = false
    if (run.cellIndices.length === 1) {
      run.step = nextStep
    } else if (run.step !== nextStep) {
      breaksOrder = true
    }
    if (prepared.row <= run.lastIndex || prepared.col !== run.fixedIndex) {
      breaksOrder = true
    }
    if (breaksOrder) {
      if (!run.rows) {
        const priorStep = run.cellIndices.length <= 1 ? 1 : run.step
        const start = run.start
        run.rows = Array.from({ length: run.cellIndices.length }, (_value, index) => start + priorStep * index)
      }
      run.ordered = false
    }
    run.lastIndex = prepared.row
  }
  run.cellIndices.push(prepared.cellIndex)
  run.rows?.push(prepared.row)
}

export function registerDeferredFormulaFamilyRunNow(args: {
  readonly run: DeferredInitialFormulaFamilyRun
  readonly registerFreshFormulaFamilyRun: (args: FormulaFamilyFreshUniformRunRegistrationArgs) => boolean
  readonly upsertFormulaFamilyRun: (args: FormulaFamilyRunUpsertArgs) => void
}): void {
  const { run } = args
  const step = run.cellIndices.length <= 1 ? 1 : run.step
  if (
    run.ordered &&
    step > 0 &&
    args.registerFreshFormulaFamilyRun({
      sheetId: run.sheetId,
      templateId: run.templateId,
      shapeKey: run.shapeKey,
      axis: run.axis,
      fixedIndex: run.fixedIndex,
      start: run.start,
      step,
      cellIndices: run.cellIndices,
    })
  ) {
    return
  }
  args.upsertFormulaFamilyRun({
    sheetId: run.sheetId,
    templateId: run.templateId,
    shapeKey: run.shapeKey,
    members: materializeDeferredFormulaFamilyRunMembers(run),
  })
}

export function registerDeferredFormulaFamilyIndexRunsNow(args: {
  readonly formulaFamilies: FormulaFamilyStore
  readonly formulaFamilyShapeKeyCache: FormulaBindingFamilyShapeKeyCache
  readonly runs: readonly DeferredInitialFormulaFamilyRun[]
  readonly structuralSourceTransforms?: ReadonlyMap<number, FormulaFamilyStructuralSourceTransform>
}): void {
  args.formulaFamilies.clear()
  args.formulaFamilyShapeKeyCache.clear()
  args.runs.forEach((run, runIndex) => {
    const step = run.cellIndices.length <= 1 ? 1 : run.step
    if (
      run.ordered &&
      step > 0 &&
      args.formulaFamilies.registerFreshUniformRun({
        sheetId: run.sheetId,
        templateId: run.templateId,
        shapeKey: run.shapeKey,
        axis: run.axis,
        fixedIndex: run.fixedIndex,
        start: run.start,
        step,
        cellIndices: run.cellIndices,
      })
    ) {
      const transform = args.structuralSourceTransforms?.get(runIndex)
      if (transform !== undefined) {
        const membership = args.formulaFamilies.getMembership(run.cellIndices[0]!)
        if (membership) {
          args.formulaFamilies.setStructuralSourceTransform(membership.familyId, transform)
        }
      }
      return
    }
    args.formulaFamilies.registerFormulaRun({
      sheetId: run.sheetId,
      templateId: run.templateId,
      shapeKey: run.shapeKey,
      members: materializeDeferredFormulaFamilyRunMembers(run),
    })
    const transform = args.structuralSourceTransforms?.get(runIndex)
    if (transform !== undefined) {
      const membership = args.formulaFamilies.getMembership(run.cellIndices[0]!)
      if (membership) {
        args.formulaFamilies.setStructuralSourceTransform(membership.familyId, transform)
      }
    }
  })
}
