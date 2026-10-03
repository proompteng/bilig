import { parseFormula, serializeFormula, type CallExprNode, type FormulaNode, type NameRefNode } from '@bilig/formula'
import { WorkPaperParseError } from './work-paper-errors.js'
import { makeNamedExpressionKey, stripLeadingEquals, transformFormulaNode } from './work-paper-runtime-helpers.js'

export interface WorkPaperFormulaNamedExpressionBinding {
  readonly publicName: string
  readonly internalName: string
  readonly scope?: number
}

export interface WorkPaperFormulaFunctionBinding {
  readonly publicName: string
  readonly internalName: string
}

export function rewriteWorkPaperFormulaForStorage(input: {
  readonly formula: string
  readonly ownerSheetId: number
  readonly namedExpressions: ReadonlyMap<string, WorkPaperFormulaNamedExpressionBinding>
  readonly functionAliasLookup: ReadonlyMap<string, WorkPaperFormulaFunctionBinding>
  readonly messageOf: (error: unknown, fallback: string) => string
}): string {
  if (input.namedExpressions.size === 0 && input.functionAliasLookup.size === 0) {
    return input.formula
  }
  try {
    const transformed = transformFormulaNode(parseFormula(stripLeadingEquals(input.formula)), (node) => {
      if (node.kind === 'NameRef') {
        return rewriteNameRefForStorage(node, input.ownerSheetId, input.namedExpressions)
      }
      if (node.kind === 'CallExpr') {
        return rewriteCallForStorage(node, input.functionAliasLookup)
      }
      return node
    })
    return serializeFormula(transformed)
  } catch (error) {
    throw new WorkPaperParseError(input.messageOf(error, 'Unable to store formula'))
  }
}

export function restorePublicWorkPaperFormula(input: {
  readonly formula: string
  readonly ownerSheetId: number
  readonly namedExpressions: ReadonlyMap<string, WorkPaperFormulaNamedExpressionBinding>
  readonly internalFunctionLookup: ReadonlyMap<string, WorkPaperFormulaFunctionBinding>
}): string {
  if (input.namedExpressions.size === 0 && input.internalFunctionLookup.size === 0) {
    return input.formula
  }
  const transformed = transformFormulaNode(parseFormula(input.formula), (node) => {
    if (node.kind === 'NameRef') {
      return rewriteNameRefForPublic(node, input.ownerSheetId, input.namedExpressions)
    }
    if (node.kind === 'CallExpr') {
      return rewriteCallForPublic(node, input.internalFunctionLookup)
    }
    return node
  })
  return serializeFormula(transformed)
}

function rewriteNameRefForStorage(
  node: NameRefNode,
  ownerSheetId: number,
  namedExpressions: ReadonlyMap<string, WorkPaperFormulaNamedExpressionBinding>,
): FormulaNode {
  const scoped = namedExpressions.get(makeNamedExpressionKey(node.name, ownerSheetId))
  if (scoped) {
    return { ...node, name: scoped.internalName }
  }
  const workbookScoped = namedExpressions.get(makeNamedExpressionKey(node.name))
  if (workbookScoped) {
    return { ...node, name: workbookScoped.internalName }
  }
  return node
}

function rewriteNameRefForPublic(
  node: NameRefNode,
  ownerSheetId: number,
  namedExpressions: ReadonlyMap<string, WorkPaperFormulaNamedExpressionBinding>,
): FormulaNode {
  const exact = [...namedExpressions.values()].find(
    (expression) => expression.internalName === node.name && expression.scope === ownerSheetId,
  )
  if (exact) {
    return { ...node, name: exact.publicName }
  }
  return node
}

function rewriteCallForStorage(node: CallExprNode, functionAliasLookup: ReadonlyMap<string, WorkPaperFormulaFunctionBinding>): FormulaNode {
  const binding = functionAliasLookup.get(node.callee.trim().toUpperCase())
  if (!binding) {
    return node
  }
  return { ...node, callee: binding.internalName }
}

function rewriteCallForPublic(
  node: CallExprNode,
  internalFunctionLookup: ReadonlyMap<string, WorkPaperFormulaFunctionBinding>,
): FormulaNode {
  const binding = internalFunctionLookup.get(node.callee.trim().toUpperCase())
  if (!binding) {
    return node
  }
  return { ...node, callee: binding.publicName }
}
