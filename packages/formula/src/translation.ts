import { FormulaMode } from '@bilig/protocol'
import type { FormulaNode } from './ast.js'
import { formatRangeAddress, parseRangeAddress } from './addressing.js'
import type { CompiledFormula, ParsedCellReferenceInfo, ParsedDependencyReference, ParsedRangeReferenceInfo } from './compiler.js'
import type { JsPlanInstruction, ReferenceOperand } from './js-evaluator.js'
import { parseFormula } from './parser.js'
import { serializeFormula } from './formula-serializer.js'
import { buildRelativeFormulaTemplateAstKey } from './formula-template-key.js'
import type { StructuralCompiledFormulaRewriteResult } from './formula-structural-rewrite.js'
import {
  renameCompiledFormulaSheetReferenceMetadata as renameCompiledFormulaSheetReferenceMetadataImpl,
  renameCompiledFormulaSheetReferenceMetadataInPlace as renameCompiledFormulaSheetReferenceMetadataInPlaceImpl,
  renameCompiledFormulaSheetReferences as renameCompiledFormulaSheetReferencesImpl,
  renameFormulaSheetReferences as renameFormulaSheetReferencesImpl,
} from './formula-sheet-rename.js'
import {
  buildTranslatedCellReferenceMap,
  buildTranslatedRangeReferenceMap,
  formatParsedCellReference,
  formatParsedLocalCellReference,
  formatParsedRangeReference,
  translatedCellInstructionKey,
  translatedRangeInstructionKey,
  translateCellReference,
  translateColumnReference,
  translateParsedCellReference,
  translateParsedDependencyReference,
  translateParsedRangeReference,
  translateQualifiedCellReference,
  translateQualifiedDependencyReference,
  translateQualifiedRangeReference,
  translateRowReference,
} from './formula-reference-translation.js'
import { quoteSheetNameIfNeeded } from './translation-reference-utils.js'
export {
  rewriteAddressForStructuralTransform,
  rewriteCompiledFormulaForStructuralTransform,
  rewriteFormulaForStructuralTransform,
  rewriteRangeForStructuralTransform,
} from './formula-structural-rewrite.js'
export type { StructuralCompiledFormulaRewriteResult } from './formula-structural-rewrite.js'
export { serializeFormula } from './formula-serializer.js'
export type { StructuralAxisKind, StructuralAxisTransform } from './translation-reference-utils.js'

export function translateFormulaReferences(source: string, rowDelta: number, colDelta: number): string {
  const ast = parseFormula(source)
  return serializeFormula(translateNode(ast, rowDelta, colDelta))
}

export function buildRelativeFormulaTemplateKey(source: string, ownerRow: number, ownerCol: number): string {
  return buildRelativeFormulaTemplateKeyFromAst(parseFormula(source), ownerRow, ownerCol)
}

export function buildRelativeFormulaTemplateKeyFromAst(node: FormulaNode, ownerRow: number, ownerCol: number): string {
  return buildRelativeFormulaTemplateAstKey(node, ownerRow, ownerCol)
}

export interface CompiledFormulaTranslationResult {
  source: string
  compiled: CompiledFormula
}

export function canTranslateCompiledFormulaWithoutAst(compiled: CompiledFormula): boolean {
  return (
    (compiled.symbolicRanges.length === 0 || compiled.directAggregateCandidate !== undefined) &&
    compiled.symbolicNames.length === 0 &&
    compiled.symbolicTables.length === 0 &&
    compiled.symbolicSpills.length === 0 &&
    !compiled.jsPlan.some((instruction) => instruction.opcode === 'lookup-exact-match' || instruction.opcode === 'lookup-approximate-match')
  )
}

export function translateCompiledFormulaWithoutAst(
  compiled: CompiledFormula,
  rowDelta: number,
  colDelta: number,
  sourceOverride?: string,
): CompiledFormulaTranslationResult {
  const translatedParsedDeps = compiled.parsedDeps?.map((dependency) => translateParsedDependencyReference(dependency, rowDelta, colDelta))
  const translatedParsedSymbolicRefs = compiled.parsedSymbolicRefs?.map((reference) =>
    translateParsedCellReference(reference, rowDelta, colDelta),
  )
  const translatedParsedSymbolicRanges = compiled.parsedSymbolicRanges?.map((range) =>
    translateParsedRangeReference(range, rowDelta, colDelta),
  )
  const source = sourceOverride ?? compiled.source
  const canReuseJsPlan = compiled.symbolicRanges.length === 0 && compiled.mode === FormulaMode.WasmFastPath
  const translatedCellMap = canReuseJsPlan
    ? undefined
    : buildTranslatedCellReferenceMap(compiled.parsedSymbolicRefs, translatedParsedSymbolicRefs)
  const translatedRangeMap = canReuseJsPlan
    ? undefined
    : buildTranslatedRangeReferenceMap(compiled.parsedSymbolicRanges, translatedParsedSymbolicRanges)

  return {
    source,
    compiled: {
      ...compiled,
      source,
      astMatchesSource: false,
      deps:
        translatedParsedDeps?.map((dependency) => formatCompiledDependencyReference(dependency)) ??
        compiled.deps.map((dependency) => translateQualifiedDependencyReference(dependency, rowDelta, colDelta)),
      symbolicRefs:
        translatedParsedSymbolicRefs?.map((reference) => formatParsedCellReference(reference)) ??
        compiled.symbolicRefs.map((ref) => translateQualifiedCellReference(ref, rowDelta, colDelta)),
      symbolicRanges:
        translatedParsedSymbolicRanges?.map((range) => formatParsedRangeReference(range)) ??
        compiled.symbolicRanges.map((range) => translateQualifiedRangeReference(range, rowDelta, colDelta)),
      jsPlan: canReuseJsPlan
        ? compiled.jsPlan
        : compiled.jsPlan.map((instruction) =>
            translateJsPlanInstructionWithoutAst(instruction, translatedCellMap!, translatedRangeMap!, rowDelta, colDelta),
          ),
      ...(translatedParsedDeps ? { parsedDeps: translatedParsedDeps } : {}),
      ...(translatedParsedSymbolicRefs ? { parsedSymbolicRefs: translatedParsedSymbolicRefs } : {}),
      ...(translatedParsedSymbolicRanges ? { parsedSymbolicRanges: translatedParsedSymbolicRanges } : {}),
    },
  }
}

export function translateCompiledFormula(
  compiled: CompiledFormula,
  rowDelta: number,
  colDelta: number,
  sourceOverride?: string,
): CompiledFormulaTranslationResult {
  const translatedAst = translateNode(compiled.ast, rowDelta, colDelta)
  const translatedOptimizedAst =
    compiled.optimizedAst === compiled.ast ? translatedAst : translateNode(compiled.optimizedAst, rowDelta, colDelta)
  const translatedParsedDeps = compiled.parsedDeps?.map((dependency) => translateParsedDependencyReference(dependency, rowDelta, colDelta))
  const translatedParsedSymbolicRefs = compiled.parsedSymbolicRefs?.map((reference) =>
    translateParsedCellReference(reference, rowDelta, colDelta),
  )
  const translatedParsedSymbolicRanges = compiled.parsedSymbolicRanges?.map((range) =>
    translateParsedRangeReference(range, rowDelta, colDelta),
  )
  const source = sourceOverride ?? serializeFormula(translatedAst)

  return {
    source,
    compiled: {
      ...compiled,
      source,
      ast: translatedAst,
      optimizedAst: translatedOptimizedAst,
      astMatchesSource: true,
      deps:
        translatedParsedDeps?.map((dependency) => formatCompiledDependencyReference(dependency)) ??
        compiled.deps.map((dependency) => translateQualifiedDependencyReference(dependency, rowDelta, colDelta)),
      symbolicRefs:
        translatedParsedSymbolicRefs?.map((reference) => formatParsedCellReference(reference)) ??
        compiled.symbolicRefs.map((ref) => translateQualifiedCellReference(ref, rowDelta, colDelta)),
      symbolicRanges:
        translatedParsedSymbolicRanges?.map((range) => formatParsedRangeReference(range)) ??
        compiled.symbolicRanges.map((range) => translateQualifiedRangeReference(range, rowDelta, colDelta)),
      jsPlan: compiled.jsPlan.map((instruction) => translateJsPlanInstruction(instruction, rowDelta, colDelta)),
      ...(translatedParsedDeps ? { parsedDeps: translatedParsedDeps } : {}),
      ...(translatedParsedSymbolicRefs ? { parsedSymbolicRefs: translatedParsedSymbolicRefs } : {}),
      ...(translatedParsedSymbolicRanges ? { parsedSymbolicRanges: translatedParsedSymbolicRanges } : {}),
    },
  }
}

export function renameFormulaSheetReferences(source: string, oldSheetName: string, newSheetName: string): string {
  return renameFormulaSheetReferencesImpl(source, oldSheetName, newSheetName)
}

export function renameCompiledFormulaSheetReferences(
  compiled: CompiledFormula,
  oldSheetName: string,
  newSheetName: string,
): StructuralCompiledFormulaRewriteResult {
  return renameCompiledFormulaSheetReferencesImpl(compiled, oldSheetName, newSheetName)
}

export interface CompiledFormulaSheetRenameMetadataResult {
  compiled: CompiledFormula
  sourceChanged: boolean
}

export function renameCompiledFormulaSheetReferenceMetadata(
  compiled: CompiledFormula,
  oldSheetName: string,
  newSheetName: string,
): CompiledFormulaSheetRenameMetadataResult {
  return renameCompiledFormulaSheetReferenceMetadataImpl(compiled, oldSheetName, newSheetName)
}

export function renameCompiledFormulaSheetReferenceMetadataInPlace(
  compiled: CompiledFormula,
  oldSheetName: string,
  newSheetName: string,
): boolean {
  return renameCompiledFormulaSheetReferenceMetadataInPlaceImpl(compiled, oldSheetName, newSheetName)
}

function translateNode(node: FormulaNode, rowDelta: number, colDelta: number): FormulaNode {
  switch (node.kind) {
    case 'NumberLiteral':
    case 'BooleanLiteral':
    case 'StringLiteral':
    case 'ErrorLiteral':
    case 'NameRef':
    case 'StructuredRef':
      return node
    case 'CellRef':
      return {
        ...node,
        ref: translateCellReference(node.ref, rowDelta, colDelta),
      }
    case 'SpillRef':
      return {
        ...node,
        ref: translateCellReference(node.ref, rowDelta, colDelta),
      }
    case 'ColumnRef':
      return {
        ...node,
        ref: translateColumnReference(node.ref, colDelta),
      }
    case 'RowRef':
      return {
        ...node,
        ref: translateRowReference(node.ref, rowDelta),
      }
    case 'RangeRef':
      return {
        ...node,
        start:
          node.refKind === 'cells'
            ? translateCellReference(node.start, rowDelta, colDelta)
            : node.refKind === 'cols'
              ? translateColumnReference(node.start, colDelta)
              : translateRowReference(node.start, rowDelta),
        end:
          node.refKind === 'cells'
            ? translateCellReference(node.end, rowDelta, colDelta)
            : node.refKind === 'cols'
              ? translateColumnReference(node.end, colDelta)
              : translateRowReference(node.end, rowDelta),
      }
    case 'UnaryExpr':
      return {
        ...node,
        argument: translateNode(node.argument, rowDelta, colDelta),
      }
    case 'BinaryExpr':
      return {
        ...node,
        left: translateNode(node.left, rowDelta, colDelta),
        right: translateNode(node.right, rowDelta, colDelta),
      }
    case 'CallExpr':
      return {
        ...node,
        args: node.args.map((arg) => translateNode(arg, rowDelta, colDelta)),
      }
    case 'InvokeExpr':
      return {
        ...node,
        callee: translateNode(node.callee, rowDelta, colDelta),
        args: node.args.map((arg) => translateNode(arg, rowDelta, colDelta)),
      }
  }
}

function translateJsPlanInstruction(instruction: JsPlanInstruction, rowDelta: number, colDelta: number): JsPlanInstruction {
  switch (instruction.opcode) {
    case 'push-cell':
      return {
        ...instruction,
        address: translateCellReference(instruction.address, rowDelta, colDelta),
      }
    case 'push-range': {
      const nextRange = translatePlanRangeInstruction(instruction.refKind, instruction.start, instruction.end, rowDelta, colDelta)
      return { ...instruction, ...nextRange }
    }
    case 'lookup-exact-match':
    case 'lookup-approximate-match': {
      const nextRange = translatePlanRangeInstruction(instruction.refKind, instruction.start, instruction.end, rowDelta, colDelta)
      const parsed = parseRangeAddress(formatQualifiedRangeReference(instruction.sheetName, nextRange.start, nextRange.end))
      if (parsed.kind !== 'cells') {
        return instruction
      }
      return {
        ...instruction,
        ...nextRange,
        startRow: parsed.start.row,
        endRow: parsed.end.row,
        startCol: parsed.start.col,
        endCol: parsed.end.col,
      }
    }
    case 'call':
      return instruction.argRefs
        ? {
            ...instruction,
            argRefs: instruction.argRefs.map((argRef) => (argRef ? translateReferenceOperand(argRef, rowDelta, colDelta) : argRef)),
          }
        : instruction
    case 'push-lambda':
      return {
        ...instruction,
        body: instruction.body.map((step) => translateJsPlanInstruction(step, rowDelta, colDelta)),
      }
    case 'push-number':
    case 'push-boolean':
    case 'push-string':
    case 'push-error':
    case 'push-name':
    case 'unary':
    case 'binary':
    case 'invoke':
    case 'begin-scope':
    case 'bind-name':
    case 'end-scope':
    case 'jump-if-false':
    case 'jump':
    case 'return':
      return instruction
  }
}

function translateJsPlanInstructionWithoutAst(
  instruction: JsPlanInstruction,
  translatedCellMap: ReadonlyMap<string, ParsedCellReferenceInfo>,
  translatedRangeMap: ReadonlyMap<string, ParsedRangeReferenceInfo>,
  rowDelta: number,
  colDelta: number,
): JsPlanInstruction {
  switch (instruction.opcode) {
    case 'push-cell': {
      const translated = translatedCellMap.get(translatedCellInstructionKey(instruction.sheetName, instruction.address))
      return translated
        ? {
            ...instruction,
            address: formatParsedLocalCellReference(translated),
          }
        : translateJsPlanInstruction(instruction, rowDelta, colDelta)
    }
    case 'push-range': {
      const translated = translatedRangeMap.get(
        translatedRangeInstructionKey(instruction.sheetName, instruction.refKind, instruction.start, instruction.end),
      )
      return translated
        ? {
            ...instruction,
            start: translated.startAddress,
            end: translated.endAddress,
          }
        : translateJsPlanInstruction(instruction, rowDelta, colDelta)
    }
    case 'lookup-exact-match':
    case 'lookup-approximate-match': {
      const translated = translatedRangeMap.get(
        translatedRangeInstructionKey(instruction.sheetName, instruction.refKind, instruction.start, instruction.end),
      )
      if (!translated || translated.refKind !== 'cells') {
        return translateJsPlanInstruction(instruction, rowDelta, colDelta)
      }
      return {
        ...instruction,
        start: translated.startAddress,
        end: translated.endAddress,
        startRow: translated.startRow,
        endRow: translated.endRow,
        startCol: translated.startCol,
        endCol: translated.endCol,
      }
    }
    case 'call':
      return instruction.argRefs
        ? {
            ...instruction,
            argRefs: instruction.argRefs.map((argRef) =>
              argRef ? translateReferenceOperandWithoutAst(argRef, translatedCellMap, translatedRangeMap, rowDelta, colDelta) : argRef,
            ),
          }
        : instruction
    case 'push-lambda':
      return {
        ...instruction,
        body: instruction.body.map((step) =>
          translateJsPlanInstructionWithoutAst(step, translatedCellMap, translatedRangeMap, rowDelta, colDelta),
        ),
      }
    case 'push-number':
    case 'push-boolean':
    case 'push-string':
    case 'push-error':
    case 'push-name':
    case 'unary':
    case 'binary':
    case 'invoke':
    case 'begin-scope':
    case 'bind-name':
    case 'end-scope':
    case 'jump-if-false':
    case 'jump':
    case 'return':
      return instruction
  }
}

function translateReferenceOperand(operand: ReferenceOperand, rowDelta: number, colDelta: number): ReferenceOperand {
  switch (operand.kind) {
    case 'cell':
      return operand.address
        ? {
            ...operand,
            address: translateCellReference(operand.address, rowDelta, colDelta),
          }
        : operand
    case 'range':
      if (!operand.start || !operand.end || !operand.refKind) {
        return operand
      }
      return {
        ...operand,
        ...translatePlanRangeInstruction(operand.refKind, operand.start, operand.end, rowDelta, colDelta),
      }
    case 'row':
      return operand.address
        ? {
            ...operand,
            address: translateRowReference(operand.address, rowDelta),
          }
        : operand
    case 'col':
      return operand.address
        ? {
            ...operand,
            address: translateColumnReference(operand.address, colDelta),
          }
        : operand
  }
}

function translateReferenceOperandWithoutAst(
  operand: ReferenceOperand,
  translatedCellMap: ReadonlyMap<string, ParsedCellReferenceInfo>,
  translatedRangeMap: ReadonlyMap<string, ParsedRangeReferenceInfo>,
  rowDelta: number,
  colDelta: number,
): ReferenceOperand {
  switch (operand.kind) {
    case 'cell': {
      if (!operand.address) {
        return operand
      }
      const translated = translatedCellMap.get(translatedCellInstructionKey(operand.sheetName, operand.address))
      return translated
        ? {
            ...operand,
            address: formatParsedLocalCellReference(translated),
          }
        : translateReferenceOperand(operand, rowDelta, colDelta)
    }
    case 'range': {
      if (!operand.start || !operand.end || !operand.refKind) {
        return operand
      }
      const translated = translatedRangeMap.get(
        translatedRangeInstructionKey(operand.sheetName, operand.refKind, operand.start, operand.end),
      )
      return translated
        ? {
            ...operand,
            start: translated.startAddress,
            end: translated.endAddress,
            refKind: translated.refKind,
          }
        : translateReferenceOperand(operand, rowDelta, colDelta)
    }
    case 'row':
    case 'col':
      return translateReferenceOperand(operand, rowDelta, colDelta)
  }
}

function translatePlanRangeInstruction(
  refKind: 'cells' | 'rows' | 'cols',
  start: string,
  end: string,
  rowDelta: number,
  colDelta: number,
): { start: string; end: string } {
  return translateRangeEndpoints(refKind, start, end, rowDelta, colDelta)
}

function formatQualifiedRangeReference(sheetName: string | undefined, start: string, end: string): string {
  const prefix = sheetName ? `${quoteSheetNameIfNeeded(sheetName)}!` : ''
  return `${prefix}${start}:${end}`
}

function formatCompiledDependencyReference(reference: ParsedDependencyReference): string {
  const formatted = reference.kind === 'range' ? formatParsedRangeReference(reference) : formatParsedCellReference(reference)
  if (reference.kind !== 'range') {
    return formatted
  }
  try {
    return formatRangeAddress(parseRangeAddress(formatted))
  } catch {
    return formatted
  }
}

function translateRangeEndpoints(
  refKind: 'cells' | 'rows' | 'cols',
  start: string,
  end: string,
  rowDelta: number,
  colDelta: number,
): { start: string; end: string } {
  if (refKind === 'cells') {
    return {
      start: translateCellReference(start, rowDelta, colDelta),
      end: translateCellReference(end, rowDelta, colDelta),
    }
  }
  if (refKind === 'rows') {
    return {
      start: translateRowReference(start, rowDelta),
      end: translateRowReference(end, rowDelta),
    }
  }
  return {
    start: translateColumnReference(start, colDelta),
    end: translateColumnReference(end, colDelta),
  }
}
