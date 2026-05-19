import type {
  FormulaFamilyFreshUniformRunRegistrationArgs,
  FormulaFamilyMember,
  FormulaFamilyRunUpsertArgs,
  FormulaFamilyStore,
  FormulaFamilyStructuralSourceTransform,
} from '../../formula/formula-family-store.js'
import type { RuntimeFormula } from '../runtime-state.js'
import type { FormulaBindingFamilyShapeKeyCache } from './formula-binding-family-shape-key.js'
import { initialFormulaFamilyShapeKey } from './formula-initialization-template-keys.js'

export type DeferredInitialFormulaFamilyRun = Omit<FormulaFamilyRunUpsertArgs, 'members'> & {
  axis: 'row'
  fixedIndex: number
  start: number
  step: number
  lastIndex: number
  ordered: boolean
  cellIndices: number[]
  rows?: number[]
}

export function materializeDeferredFormulaFamilyRunMembers(run: DeferredInitialFormulaFamilyRun): FormulaFamilyMember[] {
  const step = run.cellIndices.length <= 1 ? 1 : run.step
  return run.cellIndices.map((cellIndex, index) => ({
    cellIndex,
    row: run.ordered ? run.start + step * index : run.rows![index]!,
    col: run.fixedIndex,
  }))
}

export function noteDeferredFormulaFamilyRunMember(args: {
  readonly runs: Map<string, DeferredInitialFormulaFamilyRun> | undefined
  readonly formulas: { readonly get: (cellIndex: number) => RuntimeFormula | undefined }
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
  const familyKey = `${prepared.sheetId}\t${templateId}\t${prepared.col}`
  let run = runs.get(familyKey)
  if (!run) {
    const runtimeFormula = args.formulas.get(prepared.cellIndex)
    if (runtimeFormula === undefined) {
      return
    }
    run = {
      sheetId: prepared.sheetId,
      templateId,
      shapeKey: initialFormulaFamilyShapeKey(runtimeFormula),
      axis: 'row',
      fixedIndex: prepared.col,
      start: prepared.row,
      step: 0,
      lastIndex: prepared.row,
      ordered: true,
      cellIndices: [],
    }
    runs.set(familyKey, run)
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
