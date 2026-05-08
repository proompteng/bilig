import type { FormulaNode } from './ast.js'
import type { CompiledFormula, ParsedCellReferenceInfo, ParsedDependencyReference, ParsedRangeReferenceInfo } from './compiler.js'
import type { JsPlanInstruction, ReferenceOperand } from './js-evaluator.js'
import { parseFormula } from './parser.js'
import { serializeFormula } from './formula-serializer.js'
import { quoteSheetNameIfNeeded } from './translation-reference-utils.js'

export interface RenamedCompiledFormula {
  source: string
  compiled: CompiledFormula
  reusedProgram: boolean
}

export interface CompiledFormulaSheetRenameMetadataResult {
  compiled: CompiledFormula
  sourceChanged: boolean
}

export function renameFormulaSheetReferences(source: string, oldSheetName: string, newSheetName: string): string {
  const ast = parseFormula(source)
  return serializeFormula(renameNodeSheetReferences(ast, oldSheetName, newSheetName))
}

export function renameCompiledFormulaSheetReferences(
  compiled: CompiledFormula,
  oldSheetName: string,
  newSheetName: string,
): RenamedCompiledFormula {
  const currentAst = compiled.astMatchesSource === false ? parseFormula(compiled.source) : compiled.ast
  const currentOptimizedAst =
    compiled.astMatchesSource === false ? currentAst : compiled.optimizedAst === compiled.ast ? currentAst : compiled.optimizedAst
  const renamedAst = renameNodeSheetReferences(currentAst, oldSheetName, newSheetName)
  const renamedOptimizedAst =
    currentOptimizedAst === currentAst ? renamedAst : renameNodeSheetReferences(currentOptimizedAst, oldSheetName, newSheetName)
  const source = serializeFormula(renamedAst)
  return {
    source,
    compiled: {
      ...compiled,
      source,
      ast: renamedAst,
      optimizedAst: renamedOptimizedAst,
      astMatchesSource: true,
      deps: compiled.deps.map((dependency) => renameQualifiedReferenceSheet(dependency, oldSheetName, newSheetName)),
      ...(compiled.parsedDeps
        ? {
            parsedDeps: compiled.parsedDeps.map((dependency) => renameParsedDependencySheet(dependency, oldSheetName, newSheetName)),
          }
        : {}),
      jsPlan: renameJsPlanSheetReferences(compiled.jsPlan, oldSheetName, newSheetName),
      symbolicRefs: compiled.symbolicRefs.map((reference) => renameQualifiedReferenceSheet(reference, oldSheetName, newSheetName)),
      ...(compiled.parsedSymbolicRefs
        ? {
            parsedSymbolicRefs: compiled.parsedSymbolicRefs.map((reference) =>
              renameParsedCellReferenceSheet(reference, oldSheetName, newSheetName),
            ),
          }
        : {}),
      symbolicRanges: compiled.symbolicRanges.map((reference) => renameQualifiedReferenceSheet(reference, oldSheetName, newSheetName)),
      ...(compiled.parsedSymbolicRanges
        ? {
            parsedSymbolicRanges: compiled.parsedSymbolicRanges.map((reference) =>
              renameParsedRangeReferenceSheet(reference, oldSheetName, newSheetName),
            ),
          }
        : {}),
    },
    reusedProgram: true,
  }
}

function renameArraySheetReferences<T>(values: readonly T[], rename: (value: T) => T): { values: T[]; changed: boolean } {
  let changed = false
  const next = values.map((value) => {
    const renamed = rename(value)
    if (renamed !== value) {
      changed = true
    }
    return renamed
  })
  return { values: next, changed }
}

export function renameCompiledFormulaSheetReferenceMetadata(
  compiled: CompiledFormula,
  oldSheetName: string,
  newSheetName: string,
): CompiledFormulaSheetRenameMetadataResult {
  const deps = renameArraySheetReferences(compiled.deps, (dependency) =>
    renameQualifiedReferenceSheet(dependency, oldSheetName, newSheetName),
  )
  const symbolicRefs = renameArraySheetReferences(compiled.symbolicRefs, (reference) =>
    renameQualifiedReferenceSheet(reference, oldSheetName, newSheetName),
  )
  const symbolicRanges = renameArraySheetReferences(compiled.symbolicRanges, (reference) =>
    renameQualifiedReferenceSheet(reference, oldSheetName, newSheetName),
  )
  const parsedDeps = compiled.parsedDeps
    ? renameArraySheetReferences(compiled.parsedDeps, (dependency) => renameParsedDependencySheet(dependency, oldSheetName, newSheetName))
    : undefined
  const parsedSymbolicRefs = compiled.parsedSymbolicRefs
    ? renameArraySheetReferences(compiled.parsedSymbolicRefs, (reference) =>
        renameParsedCellReferenceSheet(reference, oldSheetName, newSheetName),
      )
    : undefined
  const parsedSymbolicRanges = compiled.parsedSymbolicRanges
    ? renameArraySheetReferences(compiled.parsedSymbolicRanges, (reference) =>
        renameParsedRangeReferenceSheet(reference, oldSheetName, newSheetName),
      )
    : undefined
  const jsPlan = renameJsPlanSheetReferences(compiled.jsPlan, oldSheetName, newSheetName)
  return {
    compiled: {
      ...compiled,
      astMatchesSource: false,
      deps: deps.values,
      ...(parsedDeps ? { parsedDeps: parsedDeps.values } : {}),
      jsPlan,
      symbolicRefs: symbolicRefs.values,
      ...(parsedSymbolicRefs ? { parsedSymbolicRefs: parsedSymbolicRefs.values } : {}),
      symbolicRanges: symbolicRanges.values,
      ...(parsedSymbolicRanges ? { parsedSymbolicRanges: parsedSymbolicRanges.values } : {}),
    },
    sourceChanged:
      deps.changed ||
      symbolicRefs.changed ||
      symbolicRanges.changed ||
      (parsedDeps?.changed ?? false) ||
      (parsedSymbolicRefs?.changed ?? false) ||
      (parsedSymbolicRanges?.changed ?? false),
  }
}

export function renameCompiledFormulaSheetReferenceMetadataInPlace(
  compiled: CompiledFormula,
  oldSheetName: string,
  newSheetName: string,
): boolean {
  let sourceChanged = false
  const renameStringArrayInPlace = (values: string[]): void => {
    for (let index = 0; index < values.length; index += 1) {
      const value = values[index]!
      const renamed = renameQualifiedReferenceSheet(value, oldSheetName, newSheetName)
      if (renamed !== value) {
        values[index] = renamed
        sourceChanged = true
      }
    }
  }
  renameStringArrayInPlace(compiled.deps)
  renameStringArrayInPlace(compiled.symbolicRefs)
  renameStringArrayInPlace(compiled.symbolicRanges)
  compiled.parsedDeps?.forEach((dependency) => {
    const previousSheetName = dependency.sheetName
    if (previousSheetName === oldSheetName) {
      dependency.sheetName = newSheetName
      sourceChanged = true
    }
    if (dependency.kind === 'range' && dependency.sheetEndName === oldSheetName) {
      dependency.sheetEndName = newSheetName
      sourceChanged = true
    }
  })
  compiled.parsedSymbolicRefs?.forEach((reference) => {
    if (reference.sheetName === oldSheetName) {
      reference.sheetName = newSheetName
      sourceChanged = true
    }
  })
  compiled.parsedSymbolicRanges?.forEach((reference) => {
    if (reference.sheetName === oldSheetName) {
      reference.sheetName = newSheetName
      sourceChanged = true
    }
    if (reference.sheetEndName === oldSheetName) {
      reference.sheetEndName = newSheetName
      sourceChanged = true
    }
  })
  if (compiled.jsPlan.length > 0) {
    compiled.jsPlan = renameJsPlanSheetReferences(compiled.jsPlan, oldSheetName, newSheetName)
  }
  if (sourceChanged) {
    compiled.astMatchesSource = false
  }
  return sourceChanged
}

function renameNodeSheetReferences(node: FormulaNode, oldSheetName: string, newSheetName: string): FormulaNode {
  switch (node.kind) {
    case 'NumberLiteral':
    case 'BooleanLiteral':
    case 'StringLiteral':
    case 'ErrorLiteral':
    case 'OmittedArgument':
    case 'NameRef':
    case 'StructuredRef':
      return node
    case 'ArrayConstant':
      return { ...node, rows: node.rows.map((row) => row.map((entry) => renameNodeSheetReferences(entry, oldSheetName, newSheetName))) }
    case 'CellRef':
    case 'SpillRef':
    case 'RowRef':
    case 'ColumnRef':
      return {
        ...node,
        ...(node.sheetName === oldSheetName ? { sheetName: newSheetName } : {}),
      }
    case 'RangeRef':
      return {
        ...node,
        ...(node.sheetName === oldSheetName ? { sheetName: newSheetName } : {}),
        ...(node.sheetEndName === oldSheetName ? { sheetEndName: newSheetName } : {}),
      }
    case 'UnaryExpr':
      return {
        ...node,
        argument: renameNodeSheetReferences(node.argument, oldSheetName, newSheetName),
      }
    case 'BinaryExpr':
      return {
        ...node,
        left: renameNodeSheetReferences(node.left, oldSheetName, newSheetName),
        right: renameNodeSheetReferences(node.right, oldSheetName, newSheetName),
      }
    case 'CallExpr':
      return {
        ...node,
        args: node.args.map((arg) => renameNodeSheetReferences(arg, oldSheetName, newSheetName)),
      }
    case 'InvokeExpr':
      return {
        ...node,
        callee: renameNodeSheetReferences(node.callee, oldSheetName, newSheetName),
        args: node.args.map((arg) => renameNodeSheetReferences(arg, oldSheetName, newSheetName)),
      }
  }
}

function renameQualifiedReferenceSheet(reference: string, oldSheetName: string, newSheetName: string): string {
  const bang = reference.lastIndexOf('!')
  if (bang <= 0) {
    return reference
  }
  const qualifier = reference.slice(0, bang)
  const suffix = reference.slice(bang)
  const sheetRange = splitSheetRangeQualifier(qualifier)
  if (sheetRange) {
    const start = renameSheetQualifierPart(sheetRange.start, oldSheetName, newSheetName)
    const end = renameSheetQualifierPart(sheetRange.end, oldSheetName, newSheetName)
    return start === sheetRange.start && end === sheetRange.end ? reference : `${start}:${end}${suffix}`
  }
  const renamed = renameSheetQualifierPart(qualifier, oldSheetName, newSheetName)
  return renamed === qualifier ? reference : `${renamed}${suffix}`
}

function splitSheetRangeQualifier(qualifier: string): { start: string; end: string } | undefined {
  let quoted = false
  for (let index = 0; index < qualifier.length; index += 1) {
    const char = qualifier[index]!
    if (char === "'") {
      if (quoted && qualifier[index + 1] === "'") {
        index += 1
        continue
      }
      quoted = !quoted
      continue
    }
    if (char === ':' && !quoted) {
      return {
        start: qualifier.slice(0, index),
        end: qualifier.slice(index + 1),
      }
    }
  }
  return undefined
}

function unquoteSheetQualifierPart(part: string): string {
  return part.startsWith("'") && part.endsWith("'") ? part.slice(1, -1).replace(/''/g, "'") : part
}

function renameSheetQualifierPart(part: string, oldSheetName: string, newSheetName: string): string {
  return unquoteSheetQualifierPart(part) === oldSheetName ? quoteSheetNameIfNeeded(newSheetName) : part
}

function renameParsedCellReferenceSheet<Reference extends ParsedCellReferenceInfo>(
  reference: Reference,
  oldSheetName: string,
  newSheetName: string,
): Reference {
  return reference.sheetName === oldSheetName ? { ...reference, sheetName: newSheetName } : reference
}

function renameParsedRangeReferenceSheet(
  reference: ParsedRangeReferenceInfo,
  oldSheetName: string,
  newSheetName: string,
): ParsedRangeReferenceInfo {
  const nextSheetName = reference.sheetName === oldSheetName ? newSheetName : reference.sheetName
  const nextSheetEndName = reference.sheetEndName === oldSheetName ? newSheetName : reference.sheetEndName
  return nextSheetName !== reference.sheetName || nextSheetEndName !== reference.sheetEndName
    ? {
        ...reference,
        ...(nextSheetName === undefined ? {} : { sheetName: nextSheetName }),
        ...(nextSheetEndName === undefined ? {} : { sheetEndName: nextSheetEndName }),
        address: formatQualifiedRangeReference(nextSheetName, nextSheetEndName, reference.startAddress, reference.endAddress),
      }
    : reference
}

function renameParsedDependencySheet(
  dependency: ParsedDependencyReference,
  oldSheetName: string,
  newSheetName: string,
): ParsedDependencyReference {
  return dependency.kind === 'cell'
    ? renameParsedCellReferenceSheet(dependency, oldSheetName, newSheetName)
    : renameParsedRangeReferenceSheet(dependency, oldSheetName, newSheetName)
}

function renameReferenceOperandSheet(
  operand: ReferenceOperand | undefined,
  oldSheetName: string,
  newSheetName: string,
): ReferenceOperand | undefined {
  if (!operand) {
    return operand
  }
  const nextSheetName = operand.sheetName === oldSheetName ? newSheetName : operand.sheetName
  const nextSheetEndName = operand.sheetEndName === oldSheetName ? newSheetName : operand.sheetEndName
  return nextSheetName !== operand.sheetName || nextSheetEndName !== operand.sheetEndName
    ? {
        ...operand,
        ...(nextSheetName === undefined ? {} : { sheetName: nextSheetName }),
        ...(nextSheetEndName === undefined ? {} : { sheetEndName: nextSheetEndName }),
      }
    : operand
}

function renameJsPlanSheetReferences(plan: readonly JsPlanInstruction[], oldSheetName: string, newSheetName: string): JsPlanInstruction[] {
  return plan.map((instruction) => {
    switch (instruction.opcode) {
      case 'push-cell':
      case 'lookup-exact-match':
      case 'lookup-approximate-match':
        return instruction.sheetName === oldSheetName ? { ...instruction, sheetName: newSheetName } : instruction
      case 'push-range': {
        const nextSheetName = instruction.sheetName === oldSheetName ? newSheetName : instruction.sheetName
        const nextSheetEndName = instruction.sheetEndName === oldSheetName ? newSheetName : instruction.sheetEndName
        return nextSheetName !== instruction.sheetName || nextSheetEndName !== instruction.sheetEndName
          ? {
              ...instruction,
              ...(nextSheetName === undefined ? {} : { sheetName: nextSheetName }),
              ...(nextSheetEndName === undefined ? {} : { sheetEndName: nextSheetEndName }),
            }
          : instruction
      }
      case 'push-lambda':
        return { ...instruction, body: renameJsPlanSheetReferences(instruction.body, oldSheetName, newSheetName) }
      case 'call':
        return instruction.argRefs
          ? {
              ...instruction,
              argRefs: instruction.argRefs.map((operand) => renameReferenceOperandSheet(operand, oldSheetName, newSheetName)),
            }
          : instruction
      case 'begin-scope':
      case 'binary':
      case 'bind-name':
      case 'end-scope':
      case 'invoke':
      case 'jump':
      case 'jump-if-false':
      case 'push-boolean':
      case 'push-error':
      case 'push-name':
      case 'push-number':
      case 'push-omitted':
      case 'push-string':
      case 'make-array':
      case 'return':
      case 'unary':
        return instruction
      default:
        return instruction
    }
  })
}

function formatQualifiedRangeReference(
  sheetName: string | undefined,
  sheetEndName: string | undefined,
  start: string,
  end: string,
): string {
  const prefix =
    sheetName && sheetEndName
      ? `${quoteSheetNameIfNeeded(sheetName)}:${quoteSheetNameIfNeeded(sheetEndName)}!`
      : sheetName
        ? `${quoteSheetNameIfNeeded(sheetName)}!`
        : ''
  return `${prefix}${start}:${end}`
}
