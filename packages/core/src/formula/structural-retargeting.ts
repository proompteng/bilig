import {
  rewriteCompiledFormulaForStructuralTransform,
  translateCompiledFormula,
  type CompiledFormula,
  type StructuralAxisTransform,
} from '@bilig/formula'
import { mapStructuralAxisIndex } from '../engine-structural-utils.js'
import type { FormulaInstanceSnapshot } from './formula-instance-table.js'
import type { FormulaTemplateSnapshot } from './template-bank.js'

export function retargetFormulaInstance(
  record: FormulaInstanceSnapshot,
  next: {
    readonly sheetName?: string
    readonly row?: number
    readonly col?: number
    readonly source?: string
    readonly templateId?: number
  },
): FormulaInstanceSnapshot {
  return {
    cellIndex: record.cellIndex,
    sheetName: next.sheetName ?? record.sheetName,
    row: next.row ?? record.row,
    col: next.col ?? record.col,
    source: next.source ?? record.source,
    ...(next.templateId !== undefined
      ? { templateId: next.templateId }
      : record.templateId !== undefined
        ? { templateId: record.templateId }
        : {}),
  }
}

export interface StructurallyRewrittenTemplate {
  readonly templateId: number
  readonly baseRow: number
  readonly baseCol: number
  readonly source: string
  readonly compiled: CompiledFormula
  readonly reusedProgram: boolean
}

export function rewriteTemplateForStructuralTransform(args: {
  readonly template: FormulaTemplateSnapshot
  readonly ownerSheetName: string
  readonly targetSheetName: string
  readonly transform: StructuralAxisTransform
}): StructurallyRewrittenTemplate | undefined {
  const rewritten = rewriteCompiledFormulaForStructuralTransform(
    args.template.compiled,
    args.ownerSheetName,
    args.targetSheetName,
    args.transform,
  )
  if (rewritten.source === args.template.baseSource) {
    return undefined
  }

  const mappedBaseRow =
    args.ownerSheetName === args.targetSheetName && args.transform.axis === 'row'
      ? mapStructuralAxisIndex(args.template.baseRow, args.transform)
      : args.template.baseRow
  const mappedBaseCol =
    args.ownerSheetName === args.targetSheetName && args.transform.axis === 'column'
      ? mapStructuralAxisIndex(args.template.baseCol, args.transform)
      : args.template.baseCol
  if (mappedBaseRow === undefined || mappedBaseCol === undefined) {
    return undefined
  }

  return {
    templateId: args.template.id,
    baseRow: mappedBaseRow,
    baseCol: mappedBaseCol,
    source: rewritten.source,
    compiled: rewritten.compiled,
    reusedProgram: rewritten.reusedProgram,
  }
}

export function retargetStructurallyRewrittenTemplateInstance(args: {
  readonly rewrittenTemplate: StructurallyRewrittenTemplate
  readonly ownerRow: number
  readonly ownerCol: number
}): { readonly source: string; readonly compiled: CompiledFormula; readonly reusedProgram: boolean } {
  const rowDelta = args.ownerRow - args.rewrittenTemplate.baseRow
  const colDelta = args.ownerCol - args.rewrittenTemplate.baseCol
  if (rowDelta === 0 && colDelta === 0) {
    return {
      source: args.rewrittenTemplate.source,
      compiled: args.rewrittenTemplate.compiled,
      reusedProgram: args.rewrittenTemplate.reusedProgram,
    }
  }

  const translated = translateCompiledFormula(args.rewrittenTemplate.compiled, rowDelta, colDelta)
  return {
    source: translated.source,
    compiled: translated.compiled,
    reusedProgram: args.rewrittenTemplate.reusedProgram,
  }
}
