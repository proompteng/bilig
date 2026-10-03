import type { EngineCellMutationRef } from '@bilig/core/headless-runtime'
import { translateFormulaReferences } from '@bilig/formula'
import { buildMatrixMutationPlan, type MatrixMutationDimensionImpact } from './matrix-mutation-plan.js'
import { workPaperFormulaMayResizeDynamically } from './work-paper-sheet-inspection.js'
import { stripLeadingEquals } from './work-paper-runtime-helpers.js'
import type { RawCellContent, WorkPaperCellAddress, WorkPaperSheet } from './work-paper-types.js'

export interface WorkPaperCellMutationApplyOptions {
  captureUndo?: boolean
  potentialNewCells?: number
  source?: 'local' | 'restore'
  returnUndoOps?: boolean
  reuseRefs?: boolean
  skipDimensionUpdate?: boolean
}

export interface WorkPaperMatrixApplyOptions {
  captureUndo?: boolean
  deferLiteralAddresses?: ReadonlySet<string>
  skipNulls?: boolean
}

type MatrixMutationPlanInput = Parameters<typeof buildMatrixMutationPlan>[0]

export function applyWorkPaperSerializedMatrix(input: {
  readonly applyCellMutationRefs: (refs: readonly EngineCellMutationRef[], options: WorkPaperCellMutationApplyOptions) => void
  readonly applyRawContent: (address: WorkPaperCellAddress, content: RawCellContent) => void
  readonly flushPendingBatchOps: () => void
  readonly rewriteFormulaForStorage: (formula: string, ownerSheetId: number) => string
  readonly serialized: RawCellContent[][]
  readonly sourceAnchor: WorkPaperCellAddress
  readonly targetLeftCorner: WorkPaperCellAddress
}): void {
  const { serialized, sourceAnchor, targetLeftCorner } = input
  input.flushPendingBatchOps()
  const { refs, potentialNewCells } = buildMatrixMutationPlan({
    target: targetLeftCorner,
    content: serialized,
    rewriteFormula: (formula, destination, rowOffset, columnOffset) =>
      input.rewriteFormulaForStorage(
        translateFormulaReferences(
          stripLeadingEquals(formula),
          destination.row - (sourceAnchor.row + rowOffset),
          destination.col - (sourceAnchor.col + columnOffset),
        ),
        destination.sheet,
      ),
  })
  if (refs.length === 0) {
    return
  }
  input.applyCellMutationRefs(refs, {
    captureUndo: true,
    potentialNewCells,
    source: 'local',
    returnUndoOps: false,
    reuseRefs: true,
  })
}

export function applyWorkPaperMatrixContents(input: {
  readonly address: WorkPaperCellAddress
  readonly applyCellMutationRefs: (refs: readonly EngineCellMutationRef[], options: WorkPaperCellMutationApplyOptions) => void
  readonly content: WorkPaperSheet
  readonly flushPendingBatchOps: () => void
  readonly isEvaluationSuspended?: () => boolean
  readonly options?: WorkPaperMatrixApplyOptions
  readonly rewriteFormulaForStorage: (formula: string, ownerSheetId: number) => string
  readonly updateSheetDimensionsAfterCellMutationRefs?: (refs: readonly EngineCellMutationRef[]) => void
  readonly updateSheetDimensionsAfterMatrixMutationImpact?: (impact: MatrixMutationDimensionImpact) => void
}): void {
  const options = input.options ?? {}
  input.flushPendingBatchOps()
  const planInput: MatrixMutationPlanInput = {
    target: input.address,
    content: input.content,
    includeCombinedRefs: false,
    rewriteFormula: (formula, destination) => input.rewriteFormulaForStorage(stripLeadingEquals(formula), destination.sheet),
  }
  if (options.deferLiteralAddresses !== undefined) {
    planInput.deferLiteralAddresses = options.deferLiteralAddresses
  }
  if (options.skipNulls !== undefined) {
    planInput.skipNulls = options.skipNulls
  }
  const phaseSource = options.captureUndo === false ? 'restore' : 'local'
  const createApplyOptions = (phasePotentialNewCells: number): WorkPaperCellMutationApplyOptions => {
    const applyOptions: WorkPaperCellMutationApplyOptions = {
      potentialNewCells: phasePotentialNewCells,
      source: phaseSource,
      returnUndoOps: false,
      reuseRefs: true,
    }
    if (options.captureUndo !== undefined) {
      applyOptions.captureUndo = options.captureUndo
    }
    return applyOptions
  }
  const updateSheetDimensionsAfterCellMutationRefs = input.updateSheetDimensionsAfterCellMutationRefs
  const canUpdateDimensionsOnce =
    updateSheetDimensionsAfterCellMutationRefs !== undefined && (phaseSource !== 'local' || input.isEvaluationSuspended?.() !== true)
  const updateDimensionsOnce = (refs: readonly EngineCellMutationRef[], impact: MatrixMutationDimensionImpact): void => {
    if (input.updateSheetDimensionsAfterMatrixMutationImpact) {
      input.updateSheetDimensionsAfterMatrixMutationImpact(impact)
    } else {
      updateSheetDimensionsAfterCellMutationRefs?.(refs)
    }
  }
  const freshNumericFormulaPlan = tryBuildFreshNumericFormulaColumnMatrixPlan(planInput)
  if (freshNumericFormulaPlan !== undefined) {
    const applyOptions = createApplyOptions(freshNumericFormulaPlan.potentialNewCells)
    if (canUpdateDimensionsOnce) {
      applyOptions.skipDimensionUpdate = true
    }
    input.applyCellMutationRefs(freshNumericFormulaPlan.refs, applyOptions)
    if (canUpdateDimensionsOnce) {
      updateDimensionsOnce(freshNumericFormulaPlan.refs, freshNumericFormulaPlan.dimensionImpact)
    }
    return
  }
  const {
    leadingRefs,
    leadingFreshNumericRefCount,
    leadingPotentialNewCells,
    canApplyFreshNumericAggregateMatrixInOnePass,
    formulaRefs,
    formulaPotentialNewCells,
    refCount,
    dimensionImpact,
    potentialNewCells,
    trailingLiteralRefs,
    trailingLiteralPotentialNewCells,
  } = buildMatrixMutationPlan(planInput)
  if (refCount === 0) {
    return
  }
  const applyPlannedRefs = (phaseRefs: readonly EngineCellMutationRef[], applyOptions: WorkPaperCellMutationApplyOptions): void => {
    if (phaseRefs.length === 0) {
      return
    }
    input.applyCellMutationRefs(phaseRefs, applyOptions)
  }

  if (formulaRefs.length === 0) {
    applyPlannedRefs(leadingRefs, createApplyOptions(potentialNewCells))
    return
  }

  const canApplyFormulaMatrixInOnePass =
    trailingLiteralRefs.length === 0 &&
    !dimensionImpact.hasDynamicFormula &&
    (!canApplyLeadingRefsThroughFreshNumericFastPath(leadingRefs.length, leadingFreshNumericRefCount, leadingPotentialNewCells) ||
      canApplyFreshNumericAggregateMatrixInOnePass)
  if (canApplyFormulaMatrixInOnePass) {
    const mergedRefs = mergeMatrixMutationRefPhases(leadingRefs, formulaRefs, trailingLiteralRefs)
    const applyOptions = createApplyOptions(potentialNewCells)
    if (canUpdateDimensionsOnce) {
      applyOptions.skipDimensionUpdate = true
    }
    applyPlannedRefs(mergedRefs, applyOptions)
    if (canUpdateDimensionsOnce) {
      updateDimensionsOnce(mergedRefs, dimensionImpact)
    }
    return
  }

  const createPhasedApplyOptions = (phasePotentialNewCells: number): WorkPaperCellMutationApplyOptions => {
    const phasedOptions = createApplyOptions(phasePotentialNewCells)
    if (canUpdateDimensionsOnce) {
      phasedOptions.skipDimensionUpdate = true
    }
    return phasedOptions
  }

  applyPlannedRefs(leadingRefs, createPhasedApplyOptions(leadingPotentialNewCells))
  applyPlannedRefs(formulaRefs, createPhasedApplyOptions(formulaPotentialNewCells))
  applyPlannedRefs(trailingLiteralRefs, createPhasedApplyOptions(trailingLiteralPotentialNewCells))
  if (canUpdateDimensionsOnce) {
    updateDimensionsOnce(mergeMatrixMutationRefPhases(leadingRefs, formulaRefs, trailingLiteralRefs), dimensionImpact)
  }
}

function isFormulaContent(content: RawCellContent): content is string {
  return typeof content === 'string' && content.trim().startsWith('=')
}

function tryBuildFreshNumericFormulaColumnMatrixPlan(args: MatrixMutationPlanInput):
  | {
      readonly refs: readonly EngineCellMutationRef[]
      readonly potentialNewCells: number
      readonly dimensionImpact: MatrixMutationDimensionImpact
    }
  | undefined {
  if (args.deferLiteralAddresses !== undefined || args.skipNulls === true || args.content.length === 0) {
    return undefined
  }
  const firstWidth = args.content[0]?.length ?? 0
  if (firstWidth < 3) {
    return undefined
  }
  const inputColCount = firstWidth - 1
  const rowCount = args.content.length
  const valueCount = rowCount * inputColCount
  const refs: EngineCellMutationRef[] = []
  refs.length = valueCount + rowCount
  let valueCursor = 0
  let formulaCursor = valueCount
  for (let rowOffset = 0; rowOffset < rowCount; rowOffset += 1) {
    const row = args.content[rowOffset]!
    if (row.length !== firstWidth) {
      return undefined
    }
    const destinationRow = args.target.row + rowOffset
    for (let columnOffset = 0; columnOffset < inputColCount; columnOffset += 1) {
      const raw = row[columnOffset]!
      if (typeof raw !== 'number' || Object.is(raw, -0)) {
        return undefined
      }
      refs[valueCursor] = {
        sheetId: args.target.sheet,
        mutation: {
          kind: 'setCellValue',
          row: destinationRow,
          col: args.target.col + columnOffset,
          value: raw,
        },
      }
      valueCursor += 1
    }
    const rawFormula = row[inputColCount]!
    if (!isFormulaContent(rawFormula)) {
      return undefined
    }
    const destination: WorkPaperCellAddress = {
      sheet: args.target.sheet,
      row: destinationRow,
      col: args.target.col + inputColCount,
    }
    const rewrittenFormula = args.rewriteFormula(rawFormula, destination, rowOffset, inputColCount)
    if (workPaperFormulaMayResizeDynamically(rewrittenFormula)) {
      return undefined
    }
    refs[formulaCursor] = {
      sheetId: args.target.sheet,
      mutation: {
        kind: 'setCellFormula',
        row: destinationRow,
        col: destination.col,
        formula: rewrittenFormula,
      },
    }
    formulaCursor += 1
  }
  return {
    refs,
    potentialNewCells: refs.length,
    dimensionImpact: {
      hasDynamicFormula: false,
      maxClearCol: -1,
      maxClearRow: -1,
      maxSetCol: args.target.col + inputColCount,
      maxSetRow: args.target.row + rowCount - 1,
      sheetId: args.target.sheet,
    },
  }
}

function mergeMatrixMutationRefPhases(
  leadingRefs: readonly EngineCellMutationRef[],
  formulaRefs: readonly EngineCellMutationRef[],
  trailingLiteralRefs: readonly EngineCellMutationRef[],
): readonly EngineCellMutationRef[] {
  if (leadingRefs.length === 0 && trailingLiteralRefs.length === 0) {
    return formulaRefs
  }
  if (formulaRefs.length === 0 && trailingLiteralRefs.length === 0) {
    return leadingRefs
  }
  if (leadingRefs.length === 0 && formulaRefs.length === 0) {
    return trailingLiteralRefs
  }
  return [...leadingRefs, ...formulaRefs, ...trailingLiteralRefs]
}

function canApplyLeadingRefsThroughFreshNumericFastPath(
  leadingRefCount: number,
  leadingFreshNumericRefCount: number,
  leadingPotentialNewCells: number,
): boolean {
  return leadingRefCount >= 32 && leadingPotentialNewCells === leadingRefCount && leadingFreshNumericRefCount === leadingRefCount
}
