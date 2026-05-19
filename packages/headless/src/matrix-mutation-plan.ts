import type { EngineCellMutationRef } from '@bilig/core/headless-runtime'
import { isBlankRawCellContent } from './work-paper-runtime-helpers.js'
import { workPaperFormulaMayResizeDynamically } from './work-paper-sheet-inspection.js'
import type { WorkPaperCellAddress, WorkPaperSheet, RawCellContent } from './work-paper-types.js'

export type MatrixMutationRef = EngineCellMutationRef

export interface MatrixMutationPlan {
  canApplyFreshNumericAggregateMatrixInOnePass: boolean
  dimensionImpact: MatrixMutationDimensionImpact
  leadingRefs: MatrixMutationRef[]
  leadingFreshNumericRefCount: number
  leadingPotentialNewCells: number
  formulaRefs: MatrixMutationRef[]
  formulaPotentialNewCells: number
  refCount: number
  refs: MatrixMutationRef[]
  potentialNewCells: number
  trailingLiteralRefs: MatrixMutationRef[]
  trailingLiteralPotentialNewCells: number
}

export interface MatrixMutationDimensionImpact {
  hasDynamicFormula: boolean
  maxClearCol: number
  maxClearRow: number
  maxSetCol: number
  maxSetRow: number
  sheetId: number
}

interface BuildMatrixMutationPlanArgs {
  target: WorkPaperCellAddress
  content: WorkPaperSheet
  rewriteFormula: (formula: string, destination: WorkPaperCellAddress, rowOffset: number, columnOffset: number) => string
  deferLiteralAddresses?: ReadonlySet<string>
  includeCombinedRefs?: boolean
  skipNulls?: boolean
}

function isFormulaContent(content: RawCellContent): content is string {
  return typeof content === 'string' && content.trim().startsWith('=')
}

export function buildMatrixMutationPlan(args: BuildMatrixMutationPlanArgs): MatrixMutationPlan {
  const leadingRefs: MatrixMutationRef[] = []
  const formulaRefs: MatrixMutationRef[] = []
  const trailingLiteralRefs: MatrixMutationRef[] = []
  let leadingFreshNumericRefCount = 0
  let leadingPotentialNewCells = 0
  let formulaPotentialNewCells = 0
  let potentialNewCells = 0
  let trailingLiteralPotentialNewCells = 0
  let hasDynamicFormula = false
  let maxClearCol = -1
  let maxClearRow = -1
  let maxSetCol = -1
  let maxSetRow = -1
  const earliestFormulaRowByColumn: number[] = []
  let freshNumericAggregateMatrixCandidate = true
  let freshNumericAggregateInputColCount = -1
  let freshNumericAggregateRowCount = 0

  const shouldDeferLiteral = (row: number, col: number): boolean => {
    const earliestFormulaRow = earliestFormulaRowByColumn[col]
    return earliestFormulaRow !== undefined && row > earliestFormulaRow
  }

  const shouldDeferLiteralAddress = (row: number, col: number): boolean => {
    const explicitAddresses = args.deferLiteralAddresses
    return explicitAddresses !== undefined && explicitAddresses.has(formatMatrixPlanAddress(row, col))
  }

  for (let rowOffset = 0; rowOffset < args.content.length; rowOffset += 1) {
    const row = args.content[rowOffset]!
    const destinationRow = args.target.row + rowOffset
    let rowFreshNumericValueCount = 0
    let rowHasFreshAggregateFormula = false
    let rowHasFreshAggregateContent = false
    for (let columnOffset = 0; columnOffset < row.length; columnOffset += 1) {
      const raw = row[columnOffset]!
      const destinationCol = args.target.col + columnOffset

      if (isBlankRawCellContent(raw)) {
        freshNumericAggregateMatrixCandidate = false
        if (!args.skipNulls) {
          maxClearRow = Math.max(maxClearRow, destinationRow)
          maxClearCol = Math.max(maxClearCol, destinationCol)
          const ref = {
            sheetId: args.target.sheet,
            mutation: { kind: 'clearCell', row: destinationRow, col: destinationCol },
          } satisfies MatrixMutationRef
          if (shouldDeferLiteral(destinationRow, destinationCol) || shouldDeferLiteralAddress(destinationRow, destinationCol)) {
            trailingLiteralRefs.push(ref)
          } else {
            leadingRefs.push(ref)
          }
        }
        continue
      }

      potentialNewCells += 1
      maxSetRow = Math.max(maxSetRow, destinationRow)
      maxSetCol = Math.max(maxSetCol, destinationCol)

      if (isFormulaContent(raw)) {
        formulaPotentialNewCells += 1
        if (freshNumericAggregateMatrixCandidate) {
          if (rowHasFreshAggregateFormula || rowFreshNumericValueCount < 2 || columnOffset !== rowFreshNumericValueCount) {
            freshNumericAggregateMatrixCandidate = false
          } else {
            if (freshNumericAggregateInputColCount === -1) {
              freshNumericAggregateInputColCount = rowFreshNumericValueCount
            } else if (freshNumericAggregateInputColCount !== rowFreshNumericValueCount) {
              freshNumericAggregateMatrixCandidate = false
            }
            rowHasFreshAggregateFormula = true
            rowHasFreshAggregateContent = true
          }
        }
        const destination: WorkPaperCellAddress = {
          sheet: args.target.sheet,
          row: destinationRow,
          col: destinationCol,
        }
        const rewrittenFormula = args.rewriteFormula(raw, destination, rowOffset, columnOffset)
        hasDynamicFormula ||= workPaperFormulaMayResizeDynamically(rewrittenFormula)
        const earliestFormulaRow = earliestFormulaRowByColumn[destinationCol]
        if (earliestFormulaRow === undefined || destinationRow < earliestFormulaRow) {
          earliestFormulaRowByColumn[destinationCol] = destinationRow
        }
        formulaRefs.push({
          sheetId: args.target.sheet,
          mutation: {
            kind: 'setCellFormula',
            row: destinationRow,
            col: destinationCol,
            formula: rewrittenFormula,
          },
        })
        continue
      }

      const ref = {
        sheetId: args.target.sheet,
        mutation: {
          kind: 'setCellValue',
          row: destinationRow,
          col: destinationCol,
          value: raw,
        },
      } satisfies MatrixMutationRef
      if (freshNumericAggregateMatrixCandidate) {
        if (rowHasFreshAggregateFormula || typeof raw !== 'number' || Object.is(raw, -0) || columnOffset !== rowFreshNumericValueCount) {
          freshNumericAggregateMatrixCandidate = false
        } else {
          rowFreshNumericValueCount += 1
          rowHasFreshAggregateContent = true
        }
      }
      if (shouldDeferLiteral(destinationRow, destinationCol) || shouldDeferLiteralAddress(destinationRow, destinationCol)) {
        trailingLiteralPotentialNewCells += 1
        trailingLiteralRefs.push(ref)
      } else {
        if (typeof raw === 'number' && !Object.is(raw, -0)) {
          leadingFreshNumericRefCount += 1
        }
        leadingPotentialNewCells += 1
        leadingRefs.push(ref)
      }
    }
    if (freshNumericAggregateMatrixCandidate) {
      if (!rowHasFreshAggregateFormula) {
        freshNumericAggregateMatrixCandidate = false
      } else if (freshNumericAggregateInputColCount !== rowFreshNumericValueCount || !rowHasFreshAggregateContent) {
        freshNumericAggregateMatrixCandidate = false
      } else {
        freshNumericAggregateRowCount += 1
      }
    }
  }

  const canApplyFreshNumericAggregateMatrixInOnePass =
    freshNumericAggregateMatrixCandidate &&
    freshNumericAggregateInputColCount >= 2 &&
    freshNumericAggregateRowCount === formulaRefs.length &&
    leadingRefs.length === freshNumericAggregateRowCount * freshNumericAggregateInputColCount &&
    leadingPotentialNewCells === leadingRefs.length &&
    leadingFreshNumericRefCount === leadingRefs.length &&
    formulaPotentialNewCells === formulaRefs.length

  return {
    canApplyFreshNumericAggregateMatrixInOnePass,
    dimensionImpact: {
      hasDynamicFormula,
      maxClearCol,
      maxClearRow,
      maxSetCol,
      maxSetRow,
      sheetId: args.target.sheet,
    },
    leadingRefs,
    leadingFreshNumericRefCount,
    leadingPotentialNewCells,
    formulaRefs,
    formulaPotentialNewCells,
    refCount: leadingRefs.length + formulaRefs.length + trailingLiteralRefs.length,
    refs: args.includeCombinedRefs === false ? [] : [...leadingRefs, ...formulaRefs, ...trailingLiteralRefs],
    potentialNewCells,
    trailingLiteralRefs,
    trailingLiteralPotentialNewCells,
  }
}

function formatMatrixPlanAddress(row: number, col: number): string {
  let index = col
  let label = ''
  do {
    label = String.fromCharCode(65 + (index % 26)) + label
    index = Math.floor(index / 26) - 1
  } while (index >= 0)
  return `${label}${row + 1}`
}
