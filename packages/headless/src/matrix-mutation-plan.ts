import type { EngineCellMutationRef } from '@bilig/core'
import { isBlankRawCellContent } from './work-paper-runtime-helpers.js'
import type { WorkPaperCellAddress, WorkPaperSheet, RawCellContent } from './work-paper-types.js'

export type MatrixMutationRef = EngineCellMutationRef

export interface MatrixMutationPlan {
  leadingRefs: MatrixMutationRef[]
  leadingPotentialNewCells: number
  formulaRefs: MatrixMutationRef[]
  formulaPotentialNewCells: number
  refCount: number
  refs: MatrixMutationRef[]
  potentialNewCells: number
  trailingLiteralRefs: MatrixMutationRef[]
  trailingLiteralPotentialNewCells: number
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
  let leadingPotentialNewCells = 0
  let formulaPotentialNewCells = 0
  let potentialNewCells = 0
  let trailingLiteralPotentialNewCells = 0
  const earliestFormulaRowByColumn = new Map<number, number>()

  const shouldDeferLiteral = (row: number, col: number): boolean => {
    const earliestFormulaRow = earliestFormulaRowByColumn.get(col)
    return earliestFormulaRow !== undefined && row > earliestFormulaRow
  }

  const shouldDeferLiteralAddress = (row: number, col: number): boolean => {
    const explicitAddresses = args.deferLiteralAddresses
    return explicitAddresses !== undefined && explicitAddresses.has(formatMatrixPlanAddress(row, col))
  }

  args.content.forEach((row, rowOffset) => {
    row.forEach((raw, columnOffset) => {
      const destination: WorkPaperCellAddress = {
        sheet: args.target.sheet,
        row: args.target.row + rowOffset,
        col: args.target.col + columnOffset,
      }

      if (isBlankRawCellContent(raw)) {
        if (!args.skipNulls) {
          const ref = {
            sheetId: args.target.sheet,
            mutation: { kind: 'clearCell', row: destination.row, col: destination.col },
          } satisfies MatrixMutationRef
          if (shouldDeferLiteral(destination.row, destination.col) || shouldDeferLiteralAddress(destination.row, destination.col)) {
            trailingLiteralRefs.push(ref)
          } else {
            leadingRefs.push(ref)
          }
        }
        return
      }

      potentialNewCells += 1

      if (isFormulaContent(raw)) {
        formulaPotentialNewCells += 1
        const earliestFormulaRow = earliestFormulaRowByColumn.get(destination.col)
        if (earliestFormulaRow === undefined || destination.row < earliestFormulaRow) {
          earliestFormulaRowByColumn.set(destination.col, destination.row)
        }
        formulaRefs.push({
          sheetId: args.target.sheet,
          mutation: {
            kind: 'setCellFormula',
            row: destination.row,
            col: destination.col,
            formula: args.rewriteFormula(raw, destination, rowOffset, columnOffset),
          },
        })
        return
      }

      const ref = {
        sheetId: args.target.sheet,
        mutation: {
          kind: 'setCellValue',
          row: destination.row,
          col: destination.col,
          value: raw,
        },
      } satisfies MatrixMutationRef
      if (shouldDeferLiteral(destination.row, destination.col) || shouldDeferLiteralAddress(destination.row, destination.col)) {
        trailingLiteralPotentialNewCells += 1
        trailingLiteralRefs.push(ref)
      } else {
        leadingPotentialNewCells += 1
        leadingRefs.push(ref)
      }
    })
  })

  return {
    leadingRefs,
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
