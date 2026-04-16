import type { EngineCellMutationRef } from '@bilig/core'
import { formatAddress } from '@bilig/formula'
import type { WorkPaperCellAddress, WorkPaperSheet, RawCellContent } from './work-paper-types.js'

export type MatrixMutationRef = EngineCellMutationRef

export interface MatrixMutationPlan {
  leadingRefs: MatrixMutationRef[]
  formulaRefs: MatrixMutationRef[]
  refs: MatrixMutationRef[]
  potentialNewCells: number
  trailingLiteralRefs: MatrixMutationRef[]
}

interface BuildMatrixMutationPlanArgs {
  target: WorkPaperCellAddress
  content: WorkPaperSheet
  rewriteFormula: (formula: string, destination: WorkPaperCellAddress, rowOffset: number, columnOffset: number) => string
  deferLiteralAddresses?: ReadonlySet<string>
  skipNulls?: boolean
}

function isFormulaContent(content: RawCellContent): content is string {
  return typeof content === 'string' && content.trim().startsWith('=')
}

export function buildMatrixMutationPlan(args: BuildMatrixMutationPlanArgs): MatrixMutationPlan {
  const leadingRefs: MatrixMutationRef[] = []
  const formulaRefs: MatrixMutationRef[] = []
  const trailingLiteralRefs: MatrixMutationRef[] = []
  let potentialNewCells = 0
  const earliestFormulaRowByColumn = new Map<number, number>()

  args.content.forEach((row, rowOffset) => {
    row.forEach((raw, columnOffset) => {
      if (!isFormulaContent(raw)) {
        return
      }
      const destinationRow = args.target.row + rowOffset
      const destinationCol = args.target.col + columnOffset
      const earliestFormulaRow = earliestFormulaRowByColumn.get(destinationCol)
      if (earliestFormulaRow === undefined || destinationRow < earliestFormulaRow) {
        earliestFormulaRowByColumn.set(destinationCol, destinationRow)
      }
    })
  })

  const shouldDeferLiteral = (address: string, row: number, col: number): boolean =>
    args.deferLiteralAddresses?.has(address) === true || row > (earliestFormulaRowByColumn.get(col) ?? Number.POSITIVE_INFINITY)

  args.content.forEach((row, rowOffset) => {
    row.forEach((raw, columnOffset) => {
      const destination: WorkPaperCellAddress = {
        sheet: args.target.sheet,
        row: args.target.row + rowOffset,
        col: args.target.col + columnOffset,
      }
      const address = formatAddress(destination.row, destination.col)

      if (raw === null) {
        if (!args.skipNulls) {
          const ref = {
            sheetId: args.target.sheet,
            mutation: { kind: 'clearCell', row: destination.row, col: destination.col },
          } satisfies MatrixMutationRef
          if (shouldDeferLiteral(address, destination.row, destination.col)) {
            trailingLiteralRefs.push(ref)
          } else {
            leadingRefs.push(ref)
          }
        }
        return
      }

      potentialNewCells += 1

      if (isFormulaContent(raw)) {
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
      if (shouldDeferLiteral(address, destination.row, destination.col)) {
        trailingLiteralRefs.push(ref)
      } else {
        leadingRefs.push(ref)
      }
    })
  })

  return {
    leadingRefs,
    formulaRefs,
    refs: [...leadingRefs, ...formulaRefs, ...trailingLiteralRefs],
    potentialNewCells,
    trailingLiteralRefs,
  }
}
