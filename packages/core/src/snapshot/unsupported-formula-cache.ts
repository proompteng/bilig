import { hasBuiltin, parseFormula, type FormulaNode } from '@bilig/formula'
import type { WorkbookSnapshot } from '@bilig/protocol'

function normalizeFormulaName(name: string): string {
  return name.trim().toUpperCase()
}

function withLocalFormulaName(localNames: ReadonlySet<string>, name: string): ReadonlySet<string> {
  const normalized = normalizeFormulaName(name)
  if (normalized.length === 0 || localNames.has(normalized)) {
    return localNames
  }
  const next = new Set(localNames)
  next.add(normalized)
  return next
}

export function collectDefinedFormulaNames(snapshot: WorkbookSnapshot): ReadonlySet<string> {
  const definedNames = snapshot.workbook.metadata?.definedNames
  if (!definedNames || definedNames.length === 0) {
    return new Set()
  }
  const names = new Set<string>()
  for (const definedName of definedNames) {
    const normalized = normalizeFormulaName(definedName.name)
    if (normalized.length > 0) {
      names.add(normalized)
    }
  }
  return names
}

function isAvailableFormulaCall(callee: string, definedNames: ReadonlySet<string>, localNames: ReadonlySet<string>): boolean {
  const normalized = normalizeFormulaName(callee)
  return normalized.length > 0 && (hasBuiltin(normalized) || definedNames.has(normalized) || localNames.has(normalized))
}

function lambdaBodyHasUnavailableCall(
  args: readonly FormulaNode[],
  definedNames: ReadonlySet<string>,
  localNames: ReadonlySet<string>,
): boolean {
  if (args.length === 0) {
    return false
  }
  let lambdaLocals = localNames
  for (let index = 0; index < args.length - 1; index += 1) {
    const param = args[index]
    if (param?.kind === 'NameRef') {
      lambdaLocals = withLocalFormulaName(lambdaLocals, param.name)
    }
  }
  return formulaNodeHasUnavailableCall(args[args.length - 1]!, definedNames, lambdaLocals)
}

function letBodyHasUnavailableCall(
  args: readonly FormulaNode[],
  definedNames: ReadonlySet<string>,
  localNames: ReadonlySet<string>,
): boolean {
  if (args.length < 2) {
    return args.some((arg) => formulaNodeHasUnavailableCall(arg, definedNames, localNames))
  }
  let letLocals = localNames
  const finalArgIndex = args.length - 1
  for (let index = 0; index < finalArgIndex; index += 2) {
    const valueNode = args[index + 1]
    if (valueNode && formulaNodeHasUnavailableCall(valueNode, definedNames, letLocals)) {
      return true
    }
    const nameNode = args[index]
    if (nameNode?.kind === 'NameRef') {
      letLocals = withLocalFormulaName(letLocals, nameNode.name)
    }
  }
  return formulaNodeHasUnavailableCall(args[finalArgIndex]!, definedNames, letLocals)
}

function formulaNodeHasUnavailableCall(node: FormulaNode, definedNames: ReadonlySet<string>, localNames: ReadonlySet<string>): boolean {
  switch (node.kind) {
    case 'NumberLiteral':
    case 'BooleanLiteral':
    case 'StringLiteral':
    case 'ErrorLiteral':
    case 'OmittedArgument':
    case 'NameRef':
    case 'StructuredRef':
    case 'CellRef':
    case 'SpillRef':
    case 'RowRef':
    case 'ColumnRef':
    case 'RangeRef':
      return false
    case 'ArrayConstant':
      return node.rows.some((row) => row.some((entry) => formulaNodeHasUnavailableCall(entry, definedNames, localNames)))
    case 'UnaryExpr':
      return formulaNodeHasUnavailableCall(node.argument, definedNames, localNames)
    case 'BinaryExpr':
      return (
        formulaNodeHasUnavailableCall(node.left, definedNames, localNames) ||
        formulaNodeHasUnavailableCall(node.right, definedNames, localNames)
      )
    case 'CallExpr': {
      const normalized = normalizeFormulaName(node.callee)
      if (normalized === 'LAMBDA') {
        return lambdaBodyHasUnavailableCall(node.args, definedNames, localNames)
      }
      if (normalized === 'LET') {
        return letBodyHasUnavailableCall(node.args, definedNames, localNames)
      }
      if (!isAvailableFormulaCall(node.callee, definedNames, localNames)) {
        return true
      }
      return node.args.some((arg) => formulaNodeHasUnavailableCall(arg, definedNames, localNames))
    }
    case 'InvokeExpr':
      return (
        formulaNodeHasUnavailableCall(node.callee, definedNames, localNames) ||
        node.args.some((arg) => formulaNodeHasUnavailableCall(arg, definedNames, localNames))
      )
  }
}

export function formulaShouldUseCachedUnsupportedFunctionValue(source: string, definedNames: ReadonlySet<string>): boolean {
  try {
    return formulaNodeHasUnavailableCall(parseFormula(source), definedNames, new Set())
  } catch {
    return false
  }
}
