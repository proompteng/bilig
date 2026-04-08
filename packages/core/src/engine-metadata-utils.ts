import { ErrorCode, ValueTag, type CellValue, type LiteralInput } from "@bilig/protocol";
import type { FormulaNode } from "@bilig/formula";
import { parseCellAddress, parseFormula, renameFormulaSheetReferences } from "@bilig/formula";
import { StringPool } from "./string-pool.js";
import { normalizeDefinedName, normalizeWorkbookObjectName } from "./workbook-store.js";
import type { WorkbookDefinedNameValueSnapshot } from "@bilig/protocol";

function emptyValue(): CellValue {
  return { tag: ValueTag.Empty };
}

function errorValue(code: ErrorCode): CellValue {
  return { tag: ValueTag.Error, code };
}

function literalToValue(input: LiteralInput, stringPool: StringPool): CellValue {
  if (input === null) return emptyValue();
  if (typeof input === "number") return { tag: ValueTag.Number, value: input };
  if (typeof input === "boolean") return { tag: ValueTag.Boolean, value: input };
  return { tag: ValueTag.String, value: input, stringId: stringPool.intern(input) };
}

function literalToFormulaNode(input: LiteralInput): FormulaNode | null {
  if (typeof input === "number") {
    return { kind: "NumberLiteral", value: input };
  }
  if (typeof input === "string") {
    return { kind: "StringLiteral", value: input };
  }
  if (typeof input === "boolean") {
    return { kind: "BooleanLiteral", value: input };
  }
  return null;
}

function definedNameValueToFormulaNode(
  input: WorkbookDefinedNameValueSnapshot,
): FormulaNode | null {
  if (typeof input === "object" && input !== null && "kind" in input) {
    switch (input.kind) {
      case "scalar":
        return literalToFormulaNode(input.value);
      case "cell-ref":
        return { kind: "CellRef", ref: input.address, sheetName: input.sheetName };
      case "range-ref":
        return {
          kind: "RangeRef",
          refKind: "cells",
          start: input.startAddress,
          end: input.endAddress,
          sheetName: input.sheetName,
        };
      case "structured-ref":
        return {
          kind: "StructuredRef",
          tableName: input.tableName,
          columnName: input.columnName,
        };
      case "formula":
        try {
          return parseFormula(input.formula);
        } catch {
          return { kind: "ErrorLiteral", code: ErrorCode.Value };
        }
    }
  }
  if (typeof input === "string" && input.startsWith("=")) {
    try {
      return parseFormula(input);
    } catch {
      return { kind: "ErrorLiteral", code: ErrorCode.Value };
    }
  }
  return literalToFormulaNode(input);
}

function renameFormulaTextForSheet(
  input: string,
  oldSheetName: string,
  newSheetName: string,
): string {
  const hasLeadingEquals = input.startsWith("=");
  const source = hasLeadingEquals ? input.slice(1) : input;
  const rewritten = renameFormulaSheetReferences(source, oldSheetName, newSheetName);
  return hasLeadingEquals ? `=${rewritten}` : rewritten;
}

function unquoteFormulaSheetName(sheetName: string): string {
  if (!sheetName.startsWith("'") || !sheetName.endsWith("'")) {
    return sheetName;
  }
  return sheetName.slice(1, -1).replaceAll("''", "'");
}

export interface MetadataResolutionContext {
  resolveName: (name: string) => WorkbookDefinedNameValueSnapshot | undefined;
  resolveStructuredReference: (tableName: string, columnName: string) => FormulaNode | undefined;
  resolveSpillReference: (
    sheetName: string | undefined,
    address: string,
  ) => FormulaNode | undefined;
}

export function definedNameValuesEqual(
  left: WorkbookDefinedNameValueSnapshot,
  right: WorkbookDefinedNameValueSnapshot,
): boolean {
  if (left === right) {
    return true;
  }
  return JSON.stringify(left) === JSON.stringify(right);
}

export function definedNameValueToCellValue(
  input: WorkbookDefinedNameValueSnapshot,
  stringPool: StringPool,
): CellValue {
  if (typeof input === "object" && input !== null && "kind" in input) {
    if (input.kind === "scalar") {
      return literalToValue(input.value, stringPool);
    }
    return errorValue(ErrorCode.Value);
  }
  return literalToValue(input, stringPool);
}

export function renameDefinedNameValueSheet(
  input: WorkbookDefinedNameValueSnapshot,
  oldSheetName: string,
  newSheetName: string,
): WorkbookDefinedNameValueSnapshot {
  if (typeof input === "object" && input !== null && "kind" in input) {
    switch (input.kind) {
      case "scalar":
      case "structured-ref":
        return input;
      case "cell-ref":
        return input.sheetName === oldSheetName ? { ...input, sheetName: newSheetName } : input;
      case "range-ref":
        return input.sheetName === oldSheetName ? { ...input, sheetName: newSheetName } : input;
      case "formula":
        return {
          ...input,
          formula: renameFormulaTextForSheet(input.formula, oldSheetName, newSheetName),
        };
    }
  }
  if (typeof input === "string" && input.startsWith("=")) {
    return renameFormulaTextForSheet(input, oldSheetName, newSheetName);
  }
  return input;
}

export function resolveMetadataReferencesInAst(
  node: FormulaNode,
  context: MetadataResolutionContext,
  activeNames = new Set<string>(),
): { node: FormulaNode; fullyResolved: boolean; substituted: boolean } {
  switch (node.kind) {
    case "NumberLiteral":
    case "BooleanLiteral":
    case "StringLiteral":
    case "ErrorLiteral":
    case "CellRef":
    case "RowRef":
    case "ColumnRef":
    case "RangeRef":
      return { node, fullyResolved: true, substituted: false };
    case "NameRef": {
      const normalized = normalizeDefinedName(node.name);
      if (activeNames.has(normalized)) {
        return {
          node: { kind: "ErrorLiteral", code: ErrorCode.Cycle },
          fullyResolved: true,
          substituted: true,
        };
      }
      const literal = context.resolveName(node.name);
      const replacement =
        literal === undefined
          ? ({ kind: "ErrorLiteral", code: ErrorCode.Name } satisfies FormulaNode)
          : definedNameValueToFormulaNode(literal);
      if (!replacement) {
        return { node, fullyResolved: false, substituted: false };
      }
      const nextActiveNames = new Set(activeNames);
      nextActiveNames.add(normalized);
      const resolved = resolveMetadataReferencesInAst(replacement, context, nextActiveNames);
      return { node: resolved.node, fullyResolved: resolved.fullyResolved, substituted: true };
    }
    case "StructuredRef": {
      const replacement =
        context.resolveStructuredReference(node.tableName, node.columnName) ??
        ({ kind: "ErrorLiteral", code: ErrorCode.Ref } satisfies FormulaNode);
      return { node: replacement, fullyResolved: true, substituted: true };
    }
    case "SpillRef": {
      const replacement =
        context.resolveSpillReference(node.sheetName, node.ref) ??
        ({ kind: "ErrorLiteral", code: ErrorCode.Ref } satisfies FormulaNode);
      return { node: replacement, fullyResolved: true, substituted: true };
    }
    case "UnaryExpr": {
      const resolved = resolveMetadataReferencesInAst(node.argument, context, activeNames);
      return {
        node: resolved.substituted ? { ...node, argument: resolved.node } : node,
        fullyResolved: resolved.fullyResolved,
        substituted: resolved.substituted,
      };
    }
    case "BinaryExpr": {
      const left = resolveMetadataReferencesInAst(node.left, context, activeNames);
      const right = resolveMetadataReferencesInAst(node.right, context, activeNames);
      return {
        node:
          left.substituted || right.substituted
            ? { ...node, left: left.node, right: right.node }
            : node,
        fullyResolved: left.fullyResolved && right.fullyResolved,
        substituted: left.substituted || right.substituted,
      };
    }
    case "CallExpr": {
      let fullyResolved = true;
      let substituted = false;
      const args = node.args.map((arg) => {
        const resolved = resolveMetadataReferencesInAst(arg, context, activeNames);
        fullyResolved = fullyResolved && resolved.fullyResolved;
        substituted = substituted || resolved.substituted;
        return resolved.node;
      });
      return {
        node: substituted ? { ...node, args } : node,
        fullyResolved,
        substituted,
      };
    }
    case "InvokeExpr": {
      const callee = resolveMetadataReferencesInAst(node.callee, context, activeNames);
      let fullyResolved = callee.fullyResolved;
      let substituted = callee.substituted;
      const args = node.args.map((arg) => {
        const resolved = resolveMetadataReferencesInAst(arg, context, activeNames);
        fullyResolved = fullyResolved && resolved.fullyResolved;
        substituted = substituted || resolved.substituted;
        return resolved.node;
      });
      return {
        node: substituted ? { ...node, callee: callee.node, args } : node,
        fullyResolved,
        substituted,
      };
    }
  }
}

export function tableDependencyKey(name: string): string {
  return normalizeWorkbookObjectName(name, "Table");
}

export function spillDependencyKey(sheetName: string, address: string): string {
  return `${sheetName}!${parseCellAddress(address, sheetName).text}`;
}

export function spillDependencyKeyFromRef(ref: string, ownerSheetName: string): string {
  if (ref.includes("!")) {
    const separator = ref.indexOf("!");
    const sheetName = unquoteFormulaSheetName(ref.slice(0, separator));
    const address = ref.slice(separator + 1);
    return spillDependencyKey(sheetName, address);
  }
  return spillDependencyKey(ownerSheetName, ref);
}
