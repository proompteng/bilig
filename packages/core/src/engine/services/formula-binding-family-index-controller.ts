import type {
  FormulaFamily,
  FormulaFamilyStore,
  FormulaFamilyStructuralSourceTransform,
  FormulaFamilyStructuralSourceTransformEntry,
} from '../../formula/formula-family-store.js'
import type { RuntimeFormula } from '../runtime-state.js'
import type { FormulaBindingFamilyShapeKeyCache } from './formula-binding-family-shape-key.js'
import type { FormulaOwnerPosition } from './formula-binding-service-types.js'
import { registerDeferredFormulaFamilyIndexRunsNow, type DeferredInitialFormulaFamilyRun } from './formula-initialization-family-runs.js'
import { queueDeferredFormulaFamilyStructuralSourceTransforms } from './formula-binding-deferred-family-transforms.js'

export interface FormulaBindingFamilyIndexController {
  readonly clearNow: () => void
  readonly registerFormulaFamilyNow: (cellIndex: number, formula: RuntimeFormula, ownerPosition?: FormulaOwnerPosition) => void
  readonly ensureNow: () => void
  readonly deferRebuildNow: () => void
  readonly deferRunsNow: (runs: readonly DeferredInitialFormulaFamilyRun[]) => void
  readonly canUseNow: () => boolean
  readonly isReadyNow: () => boolean
  readonly countSheetMembersNow: (sheetId: number) => number
  readonly tryDeferStructuralSourceTransformsNow: (
    sheetId: number,
    transform: FormulaFamilyStructuralSourceTransform,
    canDeferCellIndex: (cellIndex: number) => boolean,
  ) => number | undefined
  readonly forEachFamilyNow: (fn: (family: FormulaFamily) => void) => void
  readonly setStructuralSourceTransformNow: (familyId: number, transform: FormulaFamilyStructuralSourceTransform) => void
  readonly getStructuralSourceTransformNow: (cellIndex: number) => FormulaFamilyStructuralSourceTransform | undefined
  readonly hasStructuralSourceTransformsNow: () => boolean
  readonly consumeStructuralSourceTransformsNow: () => FormulaFamilyStructuralSourceTransformEntry[]
}

export function createFormulaBindingFamilyIndexController(args: {
  readonly formulaFamilies: FormulaFamilyStore
  readonly formulaFamilyShapeKeyCache: FormulaBindingFamilyShapeKeyCache
  readonly registerFormulaFamilyInStoreNow: (cellIndex: number, formula: RuntimeFormula, ownerPosition?: FormulaOwnerPosition) => void
  readonly countFormulaSheetMembersNow: (sheetId: number) => number
  readonly rebuildFormulaFamilyIndexNow: () => void
}): FormulaBindingFamilyIndexController {
  let needsRebuild = false
  let deferredRuns: readonly DeferredInitialFormulaFamilyRun[] | undefined
  let deferredStructuralSourceTransforms: Map<number, NonNullable<RuntimeFormula['structuralSourceTransform']>> | undefined

  const clearNow = (): void => {
    args.formulaFamilyShapeKeyCache.clear()
    deferredRuns = undefined
    deferredStructuralSourceTransforms = undefined
    needsRebuild = false
  }

  const registerFormulaFamilyNow = (cellIndex: number, formula: RuntimeFormula, ownerPosition?: FormulaOwnerPosition): void => {
    if (needsRebuild) {
      deferredRuns = undefined
      deferredStructuralSourceTransforms = undefined
      return
    }
    args.registerFormulaFamilyInStoreNow(cellIndex, formula, ownerPosition)
  }

  const ensureNow = (): void => {
    if (!needsRebuild) {
      return
    }
    if (deferredRuns) {
      registerDeferredFormulaFamilyIndexRunsNow({
        formulaFamilies: args.formulaFamilies,
        formulaFamilyShapeKeyCache: args.formulaFamilyShapeKeyCache,
        runs: deferredRuns,
        ...(deferredStructuralSourceTransforms === undefined ? {} : { structuralSourceTransforms: deferredStructuralSourceTransforms }),
      })
      deferredRuns = undefined
      deferredStructuralSourceTransforms = undefined
    } else {
      args.rebuildFormulaFamilyIndexNow()
      deferredStructuralSourceTransforms = undefined
    }
    needsRebuild = false
  }

  const deferRebuildNow = (): void => {
    args.formulaFamilyShapeKeyCache.clear()
    deferredRuns = undefined
    deferredStructuralSourceTransforms = undefined
    needsRebuild = true
  }

  const deferRunsNow = (runs: readonly DeferredInitialFormulaFamilyRun[]): void => {
    args.formulaFamilyShapeKeyCache.clear()
    deferredRuns = [...runs]
    deferredStructuralSourceTransforms = undefined
    needsRebuild = true
  }

  const tryDeferStructuralSourceTransformsNow = (
    sheetId: number,
    transform: FormulaFamilyStructuralSourceTransform,
    canDeferCellIndex: (cellIndex: number) => boolean,
  ): number | undefined => {
    if (!needsRebuild) {
      return undefined
    }
    const result = queueDeferredFormulaFamilyStructuralSourceTransforms({
      runs: deferredRuns,
      existingTransforms: deferredStructuralSourceTransforms,
      sheetId,
      transform,
      ownedFormulaCount: args.countFormulaSheetMembersNow(sheetId),
      canDeferCellIndex,
    })
    if (result === undefined) {
      return undefined
    }
    deferredStructuralSourceTransforms = result.transforms
    return result.memberCount
  }

  const getStructuralSourceTransformNow = (cellIndex: number): FormulaFamilyStructuralSourceTransform | undefined => {
    if (
      needsRebuild &&
      deferredRuns !== undefined &&
      !args.formulaFamilies.hasStructuralSourceTransforms() &&
      deferredStructuralSourceTransforms === undefined
    ) {
      return undefined
    }
    ensureNow()
    return args.formulaFamilies.getStructuralSourceTransform(cellIndex)
  }

  return {
    clearNow,
    registerFormulaFamilyNow,
    ensureNow,
    deferRebuildNow,
    deferRunsNow,
    canUseNow: () => !needsRebuild || deferredRuns !== undefined,
    isReadyNow: () => !needsRebuild,
    countSheetMembersNow(sheetId) {
      ensureNow()
      return args.formulaFamilies.countSheetMembers(sheetId)
    },
    tryDeferStructuralSourceTransformsNow,
    forEachFamilyNow(fn) {
      ensureNow()
      args.formulaFamilies.forEachFamily(fn)
    },
    setStructuralSourceTransformNow(familyId, transform) {
      ensureNow()
      args.formulaFamilies.setStructuralSourceTransform(familyId, transform)
    },
    getStructuralSourceTransformNow,
    hasStructuralSourceTransformsNow() {
      return (
        args.formulaFamilies.hasStructuralSourceTransforms() ||
        (deferredStructuralSourceTransforms !== undefined && deferredStructuralSourceTransforms.size > 0)
      )
    },
    consumeStructuralSourceTransformsNow() {
      ensureNow()
      return args.formulaFamilies.consumeStructuralSourceTransforms()
    },
  }
}
