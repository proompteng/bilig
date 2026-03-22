import { ErrorCode } from "@bilig/protocol";
import type { BinaryExprNode, FormulaNode } from "./ast.js";
import { parseFormula } from "./parser.js";

const CELL_REF_RE = /^(\$?)([A-Z]+)(\$?)([1-9][0-9]*)$/;
const COLUMN_REF_RE = /^(\$?)([A-Z]+)$/;
const ROW_REF_RE = /^(\$?)([1-9][0-9]*)$/;

const BINARY_PRECEDENCE: Record<BinaryExprNode["operator"], number> = {
  "=": 1,
  "<>": 1,
  ">": 1,
  ">=": 1,
  "<": 1,
  "<=": 1,
  "&": 2,
  "+": 3,
  "-": 3,
  "*": 4,
  "/": 4,
  "^": 5
};

export type StructuralAxisKind = "row" | "column";

export type StructuralAxisTransform =
  | { kind: "insert"; axis: StructuralAxisKind; start: number; count: number }
  | { kind: "delete"; axis: StructuralAxisKind; start: number; count: number }
  | { kind: "move"; axis: StructuralAxisKind; start: number; count: number; target: number };

function assertNever(value: never): never {
  throw new Error(`Unexpected value: ${String(value)}`);
}

const ERROR_LITERAL_TEXT: Record<number, string> = {
  [ErrorCode.Ref]: "#REF!",
  [ErrorCode.Name]: "#NAME?",
  [ErrorCode.Div0]: "#DIV/0!",
  [ErrorCode.NA]: "#N/A",
  [ErrorCode.Value]: "#VALUE!",
  [ErrorCode.Cycle]: "#CYCLE!",
  [ErrorCode.Spill]: "#SPILL!",
  [ErrorCode.Blocked]: "#BLOCKED!"
};

export function translateFormulaReferences(source: string, rowDelta: number, colDelta: number): string {
  const ast = parseFormula(source);
  return serializeFormula(translateNode(ast, rowDelta, colDelta));
}

export function rewriteFormulaForStructuralTransform(
  source: string,
  ownerSheetName: string,
  targetSheetName: string,
  transform: StructuralAxisTransform
): string {
  const ast = parseFormula(source);
  return serializeFormula(rewriteNodeForStructuralTransform(ast, ownerSheetName, targetSheetName, transform));
}

export function rewriteAddressForStructuralTransform(
  address: string,
  transform: StructuralAxisTransform
): string | undefined {
  const parsed = parseCellReferenceParts(address);
  if (!parsed) {
    throw new Error(`Invalid cell reference '${address}'`);
  }
  const nextRow = transform.axis === "row" ? mapPointIndex(parsed.row, transform) : parsed.row;
  const nextCol = transform.axis === "column" ? mapPointIndex(parsed.col, transform) : parsed.col;
  if (nextRow === undefined || nextCol === undefined) {
    return undefined;
  }
  return formatCellReference(parsed, nextRow, nextCol);
}

export function rewriteRangeForStructuralTransform(
  startAddress: string,
  endAddress: string,
  transform: StructuralAxisTransform
): { startAddress: string; endAddress: string } | undefined {
  const start = parseCellReferenceParts(startAddress);
  const end = parseCellReferenceParts(endAddress);
  if (!start || !end) {
    throw new Error(`Invalid range reference '${startAddress}:${endAddress}'`);
  }
  const nextRows =
    transform.axis === "row"
      ? mapInterval(Math.min(start.row, end.row), Math.max(start.row, end.row), transform)
      : { start: Math.min(start.row, end.row), end: Math.max(start.row, end.row) };
  const nextCols =
    transform.axis === "column"
      ? mapInterval(Math.min(start.col, end.col), Math.max(start.col, end.col), transform)
      : { start: Math.min(start.col, end.col), end: Math.max(start.col, end.col) };
  if (!nextRows || !nextCols) {
    return undefined;
  }
  return {
    startAddress: formatCellReference(start, nextRows.start, nextCols.start),
    endAddress: formatCellReference(end, nextRows.end, nextCols.end)
  };
}

function translateNode(node: FormulaNode, rowDelta: number, colDelta: number): FormulaNode {
  switch (node.kind) {
    case "NumberLiteral":
    case "BooleanLiteral":
    case "StringLiteral":
    case "ErrorLiteral":
    case "NameRef":
    case "StructuredRef":
      return node;
    case "CellRef":
      return {
        ...node,
        ref: translateCellReference(node.ref, rowDelta, colDelta)
      };
    case "SpillRef":
      return {
        ...node,
        ref: translateCellReference(node.ref, rowDelta, colDelta)
      };
    case "ColumnRef":
      return {
        ...node,
        ref: translateColumnReference(node.ref, colDelta)
      };
    case "RowRef":
      return {
        ...node,
        ref: translateRowReference(node.ref, rowDelta)
      };
    case "RangeRef":
      return {
        ...node,
        start:
          node.refKind === "cells"
            ? translateCellReference(node.start, rowDelta, colDelta)
            : node.refKind === "cols"
              ? translateColumnReference(node.start, colDelta)
              : translateRowReference(node.start, rowDelta),
        end:
          node.refKind === "cells"
            ? translateCellReference(node.end, rowDelta, colDelta)
            : node.refKind === "cols"
              ? translateColumnReference(node.end, colDelta)
              : translateRowReference(node.end, rowDelta)
      };
    case "UnaryExpr":
      return {
        ...node,
        argument: translateNode(node.argument, rowDelta, colDelta)
      };
    case "BinaryExpr":
      return {
        ...node,
        left: translateNode(node.left, rowDelta, colDelta),
        right: translateNode(node.right, rowDelta, colDelta)
      };
    case "CallExpr":
      return {
        ...node,
        args: node.args.map((arg) => translateNode(arg, rowDelta, colDelta))
      };
  }
}

function rewriteNodeForStructuralTransform(
  node: FormulaNode,
  ownerSheetName: string,
  targetSheetName: string,
  transform: StructuralAxisTransform
): FormulaNode {
  switch (node.kind) {
    case "NumberLiteral":
    case "BooleanLiteral":
    case "StringLiteral":
    case "ErrorLiteral":
    case "NameRef":
    case "StructuredRef":
      return node;
    case "CellRef":
      return rewriteCellLikeNode(node, ownerSheetName, targetSheetName, transform);
    case "SpillRef":
      return rewriteCellLikeNode(node, ownerSheetName, targetSheetName, transform);
    case "ColumnRef":
      if (!targetsSheet(node.sheetName, ownerSheetName, targetSheetName) || transform.axis !== "column") {
        return node;
      }
      return rewriteAxisNode(node, transform);
    case "RowRef":
      if (!targetsSheet(node.sheetName, ownerSheetName, targetSheetName) || transform.axis !== "row") {
        return node;
      }
      return rewriteAxisNode(node, transform);
    case "RangeRef":
      return rewriteRangeNode(node, ownerSheetName, targetSheetName, transform);
    case "UnaryExpr":
      return {
        ...node,
        argument: rewriteNodeForStructuralTransform(node.argument, ownerSheetName, targetSheetName, transform)
      };
    case "BinaryExpr":
      return {
        ...node,
        left: rewriteNodeForStructuralTransform(node.left, ownerSheetName, targetSheetName, transform),
        right: rewriteNodeForStructuralTransform(node.right, ownerSheetName, targetSheetName, transform)
      };
    case "CallExpr":
      return {
        ...node,
        args: node.args.map((arg) => rewriteNodeForStructuralTransform(arg, ownerSheetName, targetSheetName, transform))
      };
  }
}

function rewriteCellLikeNode<T extends Extract<FormulaNode, { kind: "CellRef" | "SpillRef" }>>(
  node: T,
  ownerSheetName: string,
  targetSheetName: string,
  transform: StructuralAxisTransform
): FormulaNode {
  if (!targetsSheet(node.sheetName, ownerSheetName, targetSheetName)) {
    return node;
  }
  const parsed = parseCellReferenceParts(node.ref);
  if (!parsed) {
    return { kind: "ErrorLiteral", code: ErrorCode.Ref };
  }
  const nextRow = transform.axis === "row" ? mapPointIndex(parsed.row, transform) : parsed.row;
  const nextCol = transform.axis === "column" ? mapPointIndex(parsed.col, transform) : parsed.col;
  if (nextRow === undefined || nextCol === undefined) {
    return { kind: "ErrorLiteral", code: ErrorCode.Ref };
  }
  return {
    ...node,
    ref: formatCellReference(parsed, nextRow, nextCol)
  };
}

function rewriteAxisNode<T extends Extract<FormulaNode, { kind: "RowRef" | "ColumnRef" }>>(
  node: T,
  transform: StructuralAxisTransform
): FormulaNode {
  const parsed = parseAxisReferenceParts(node.ref, node.kind === "RowRef" ? "row" : "column");
  if (!parsed) {
    return { kind: "ErrorLiteral", code: ErrorCode.Ref };
  }
  const nextIndex = mapPointIndex(parsed.index, transform);
  if (nextIndex === undefined) {
    return { kind: "ErrorLiteral", code: ErrorCode.Ref };
  }
  return {
    ...node,
    ref: formatAxisReference(parsed.absolute, nextIndex, node.kind === "RowRef" ? "row" : "column")
  };
}

function rewriteRangeNode(
  node: Extract<FormulaNode, { kind: "RangeRef" }>,
  ownerSheetName: string,
  targetSheetName: string,
  transform: StructuralAxisTransform
): FormulaNode {
  if (!targetsSheet(node.sheetName, ownerSheetName, targetSheetName)) {
    return node;
  }
  if (
    (node.refKind === "rows" && transform.axis === "column")
    || (node.refKind === "cols" && transform.axis === "row")
  ) {
    return node;
  }
  if (node.refKind === "cells") {
    const start = parseCellReferenceParts(node.start);
    const end = parseCellReferenceParts(node.end);
    if (!start || !end) {
      return { kind: "ErrorLiteral", code: ErrorCode.Ref };
    }
    const nextRows =
      transform.axis === "row"
        ? mapInterval(Math.min(start.row, end.row), Math.max(start.row, end.row), transform)
        : { start: Math.min(start.row, end.row), end: Math.max(start.row, end.row) };
    const nextCols =
      transform.axis === "column"
        ? mapInterval(Math.min(start.col, end.col), Math.max(start.col, end.col), transform)
        : { start: Math.min(start.col, end.col), end: Math.max(start.col, end.col) };
    if (!nextRows || !nextCols) {
      return { kind: "ErrorLiteral", code: ErrorCode.Ref };
    }
    return {
      ...node,
      start: formatCellReference(start, nextRows.start, nextCols.start),
      end: formatCellReference(end, nextRows.end, nextCols.end)
    };
  }
  const start = parseAxisReferenceParts(node.start, node.refKind === "rows" ? "row" : "column");
  const end = parseAxisReferenceParts(node.end, node.refKind === "rows" ? "row" : "column");
  if (!start || !end) {
    return { kind: "ErrorLiteral", code: ErrorCode.Ref };
  }
  const nextInterval = mapInterval(Math.min(start.index, end.index), Math.max(start.index, end.index), transform);
  if (!nextInterval) {
    return { kind: "ErrorLiteral", code: ErrorCode.Ref };
  }
  return {
    ...node,
    start: formatAxisReference(start.absolute, nextInterval.start, node.refKind === "rows" ? "row" : "column"),
    end: formatAxisReference(end.absolute, nextInterval.end, node.refKind === "rows" ? "row" : "column")
  };
}

function translateCellReference(ref: string, rowDelta: number, colDelta: number): string {
  const parsed = parseCellReferenceParts(ref);
  if (!parsed) {
    throw new Error(`Invalid cell reference '${ref}'`);
  }
  const nextCol = parsed.colAbsolute ? parsed.col : parsed.col + colDelta;
  const nextRow = parsed.rowAbsolute ? parsed.row : parsed.row + rowDelta;
  if (nextCol < 0 || nextRow < 0) {
    throw new Error(`Translated reference moved outside worksheet bounds: ${ref}`);
  }
  return formatCellReference(parsed, nextRow, nextCol);
}

function translateColumnReference(ref: string, colDelta: number): string {
  const parsed = parseAxisReferenceParts(ref, "column");
  if (!parsed) {
    throw new Error(`Invalid column reference '${ref}'`);
  }
  const nextCol = parsed.absolute ? parsed.index : parsed.index + colDelta;
  if (nextCol < 0) {
    throw new Error(`Translated reference moved outside worksheet bounds: ${ref}`);
  }
  return formatAxisReference(parsed.absolute, nextCol, "column");
}

function translateRowReference(ref: string, rowDelta: number): string {
  const parsed = parseAxisReferenceParts(ref, "row");
  if (!parsed) {
    throw new Error(`Invalid row reference '${ref}'`);
  }
  const nextRow = parsed.absolute ? parsed.index : parsed.index + rowDelta;
  if (nextRow < 0) {
    throw new Error(`Translated reference moved outside worksheet bounds: ${ref}`);
  }
  return formatAxisReference(parsed.absolute, nextRow, "row");
}

export function serializeFormula(
  node: FormulaNode,
  parentPrecedence = 0,
  parentAssociativity: "left" | "right" | null = null
): string {
  switch (node.kind) {
    case "NumberLiteral":
      return String(node.value);
    case "BooleanLiteral":
      return node.value ? "TRUE" : "FALSE";
    case "StringLiteral":
      return `"${node.value.replaceAll("\"", "\"\"")}"`;
    case "ErrorLiteral":
      return ERROR_LITERAL_TEXT[node.code] ?? "#ERROR!";
    case "NameRef":
      return node.name;
    case "StructuredRef":
      return `${node.tableName}[${node.columnName}]`;
    case "CellRef":
      return `${formatSheetPrefix(node.sheetName)}${node.ref}`;
    case "SpillRef":
      return `${formatSheetPrefix(node.sheetName)}${node.ref}#`;
    case "ColumnRef":
      return `${formatSheetPrefix(node.sheetName)}${node.ref}`;
    case "RowRef":
      return `${formatSheetPrefix(node.sheetName)}${node.ref}`;
    case "RangeRef":
      return `${formatSheetPrefix(node.sheetName)}${node.start}:${node.end}`;
    case "UnaryExpr":
      return `${node.operator}${serializeFormula(node.argument, 6)}`;
    case "CallExpr":
      return `${node.callee}(${node.args.map((arg) => serializeFormula(arg)).join(",")})`;
    case "BinaryExpr": {
      const precedence = BINARY_PRECEDENCE[node.operator];
      const isRightAssociative = node.operator === "^";
      const left = serializeFormula(node.left, precedence, "left");
      const right = serializeFormula(node.right, precedence, "right");
      const output = `${left}${node.operator}${right}`;
      const needsParens =
        precedence < parentPrecedence ||
        (precedence === parentPrecedence &&
          ((parentAssociativity === "left" && isRightAssociative) ||
            (parentAssociativity === "right" && !isRightAssociative)));
      return needsParens ? `(${output})` : output;
    }
  }
}

function formatSheetPrefix(sheetName?: string): string {
  if (!sheetName) {
    return "";
  }
  return `${quoteSheetNameIfNeeded(sheetName)}!`;
}

function quoteSheetNameIfNeeded(sheetName: string): string {
  return /^[A-Za-z0-9_.$]+$/.test(sheetName) ? sheetName : `'${sheetName.replaceAll("'", "''")}'`;
}

function columnToIndex(column: string): number {
  let value = 0;
  for (const char of column) {
    value = value * 26 + (char.charCodeAt(0) - 64);
  }
  return value - 1;
}

function indexToColumn(index: number): string {
  let current = index + 1;
  let output = "";
  while (current > 0) {
    const remainder = (current - 1) % 26;
    output = String.fromCharCode(65 + remainder) + output;
    current = Math.floor((current - 1) / 26);
  }
  return output;
}

function targetsSheet(
  explicitSheetName: string | undefined,
  ownerSheetName: string,
  targetSheetName: string
): boolean {
  return (explicitSheetName ?? ownerSheetName) === targetSheetName;
}

interface ParsedCellReference {
  colAbsolute: boolean;
  rowAbsolute: boolean;
  col: number;
  row: number;
}

function parseCellReferenceParts(ref: string): ParsedCellReference | undefined {
  const match = CELL_REF_RE.exec(ref.toUpperCase());
  if (!match) {
    return undefined;
  }
  const [, colAbsolute, columnText, rowAbsolute, rowText] = match;
  return {
    colAbsolute: colAbsolute === "$",
    rowAbsolute: rowAbsolute === "$",
    col: columnToIndex(columnText!),
    row: Number.parseInt(rowText!, 10) - 1
  };
}

function formatCellReference(parts: ParsedCellReference, row: number, col: number): string {
  return `${parts.colAbsolute ? "$" : ""}${indexToColumn(col)}${parts.rowAbsolute ? "$" : ""}${row + 1}`;
}

interface ParsedAxisReference {
  absolute: boolean;
  index: number;
}

function parseAxisReferenceParts(ref: string, kind: StructuralAxisKind): ParsedAxisReference | undefined {
  const match = (kind === "row" ? ROW_REF_RE : COLUMN_REF_RE).exec(ref.toUpperCase());
  if (!match) {
    return undefined;
  }
  return kind === "row"
    ? {
        absolute: match[1] === "$",
        index: Number.parseInt(match[2]!, 10) - 1
      }
    : {
        absolute: match[1] === "$",
        index: columnToIndex(match[2]!)
      };
}

function formatAxisReference(absolute: boolean, index: number, kind: StructuralAxisKind): string {
  const prefix = absolute ? "$" : "";
  return kind === "row" ? `${prefix}${index + 1}` : `${prefix}${indexToColumn(index)}`;
}

function mapPointIndex(index: number, transform: StructuralAxisTransform): number | undefined {
  switch (transform.kind) {
    case "insert":
      return index >= transform.start ? index + transform.count : index;
    case "delete":
      if (index < transform.start) {
        return index;
      }
      if (index >= transform.start + transform.count) {
        return index - transform.count;
      }
      return undefined;
    case "move":
      if (transform.target < transform.start) {
        if (index >= transform.target && index < transform.start) {
          return index + transform.count;
        }
      } else if (transform.target > transform.start) {
        if (index >= transform.start + transform.count && index < transform.target + transform.count) {
          return index - transform.count;
        }
      }
      if (index >= transform.start && index < transform.start + transform.count) {
        return transform.target + (index - transform.start);
      }
      return index;
    default:
      return assertNever(transform);
  }
}

function mapInterval(start: number, end: number, transform: StructuralAxisTransform): { start: number; end: number } | undefined {
  switch (transform.kind) {
    case "insert": {
      if (transform.start <= start) {
        return { start: start + transform.count, end: end + transform.count };
      }
      if (transform.start <= end) {
        return { start, end: end + transform.count };
      }
      return { start, end };
    }
    case "delete": {
      const deleteEnd = transform.start + transform.count - 1;
      if (deleteEnd < start) {
        return { start: start - transform.count, end: end - transform.count };
      }
      if (transform.start > end) {
        return { start, end };
      }
      const survivingStart = start < transform.start ? start : deleteEnd + 1;
      const survivingEnd = end > deleteEnd ? end : transform.start - 1;
      if (survivingStart > survivingEnd) {
        return undefined;
      }
      const nextStart = mapPointIndex(survivingStart, transform);
      const nextEnd = mapPointIndex(survivingEnd, transform);
      return nextStart === undefined || nextEnd === undefined ? undefined : { start: nextStart, end: nextEnd };
    }
    case "move": {
      const segments =
        transform.target < transform.start
          ? [
              { start: 0, end: transform.target - 1, delta: 0 },
              { start: transform.target, end: transform.start - 1, delta: transform.count },
              { start: transform.start, end: transform.start + transform.count - 1, delta: transform.target - transform.start },
              { start: transform.start + transform.count, end: Number.MAX_SAFE_INTEGER, delta: 0 }
            ]
          : [
              { start: 0, end: transform.start - 1, delta: 0 },
              { start: transform.start, end: transform.start + transform.count - 1, delta: transform.target - transform.start },
              { start: transform.start + transform.count, end: transform.target + transform.count - 1, delta: -transform.count },
              { start: transform.target + transform.count, end: Number.MAX_SAFE_INTEGER, delta: 0 }
            ];
      let nextStart: number | undefined;
      let nextEnd: number | undefined;
      segments.forEach((segment) => {
        const overlapStart = Math.max(start, segment.start);
        const overlapEnd = Math.min(end, segment.end);
        if (overlapStart > overlapEnd) {
          return;
        }
        const mappedStart = overlapStart + segment.delta;
        const mappedEnd = overlapEnd + segment.delta;
        nextStart = nextStart === undefined ? mappedStart : Math.min(nextStart, mappedStart);
        nextEnd = nextEnd === undefined ? mappedEnd : Math.max(nextEnd, mappedEnd);
      });
      if (nextStart === undefined || nextEnd === undefined) {
        return undefined;
      }
      return { start: nextStart, end: nextEnd };
    }
    default:
      return assertNever(transform);
  }
}
