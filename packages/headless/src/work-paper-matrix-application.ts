import type { EngineCellMutationRef } from '@bilig/core'
import { translateFormulaReferences } from '@bilig/formula'
import { buildMatrixMutationPlan } from './matrix-mutation-plan.js'
import { matrixContainsFormulaContent, stripLeadingEquals } from './work-paper-runtime-helpers.js'
import type { RawCellContent, WorkPaperCellAddress, WorkPaperSheet } from './work-paper-types.js'

export interface WorkPaperCellMutationApplyOptions {
  captureUndo?: boolean
  potentialNewCells?: number
  source?: 'local' | 'restore'
  returnUndoOps?: boolean
  reuseRefs?: boolean
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
  if (matrixContainsFormulaContent(serialized)) {
    serialized.forEach((row, rowOffset) => {
      row.forEach((raw, columnOffset) => {
        const destination = {
          sheet: targetLeftCorner.sheet,
          row: targetLeftCorner.row + rowOffset,
          col: targetLeftCorner.col + columnOffset,
        }
        let nextValue = raw
        if (typeof raw === 'string' && raw.startsWith('=')) {
          nextValue = `=${translateFormulaReferences(
            raw.slice(1),
            destination.row - (sourceAnchor.row + rowOffset),
            destination.col - (sourceAnchor.col + columnOffset),
          )}`
        }
        input.applyRawContent(destination, nextValue)
      })
    })
    return
  }

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
  readonly options?: WorkPaperMatrixApplyOptions
  readonly rewriteFormulaForStorage: (formula: string, ownerSheetId: number) => string
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
  const {
    leadingRefs,
    leadingPotentialNewCells,
    formulaRefs,
    formulaPotentialNewCells,
    refCount,
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
  const phaseSource = options.captureUndo === false ? 'restore' : 'local'
  const createApplyOptions = (phasePotentialNewCells = potentialNewCells): WorkPaperCellMutationApplyOptions => {
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

  if (formulaRefs.length === 0) {
    applyPlannedRefs(leadingRefs, createApplyOptions())
    return
  }

  applyPlannedRefs(leadingRefs, createApplyOptions(leadingPotentialNewCells))
  applyPlannedRefs(formulaRefs, createApplyOptions(formulaPotentialNewCells))
  applyPlannedRefs(trailingLiteralRefs, createApplyOptions(trailingLiteralPotentialNewCells))
}
