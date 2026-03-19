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

export function translateFormulaReferences(source: string, rowDelta: number, colDelta: number): string {
  const ast = parseFormula(source);
  return serializeFormula(translateNode(ast, rowDelta, colDelta));
}

function translateNode(node: FormulaNode, rowDelta: number, colDelta: number): FormulaNode {
  switch (node.kind) {
    case "CellRef":
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
    default:
      return node;
  }
}

function translateCellReference(ref: string, rowDelta: number, colDelta: number): string {
  const match = CELL_REF_RE.exec(ref.toUpperCase());
  if (!match) {
    throw new Error(`Invalid cell reference '${ref}'`);
  }
  const [, colAbsolute, columnText, rowAbsolute, rowText] = match;
  const currentCol = columnToIndex(columnText!);
  const currentRow = Number.parseInt(rowText!, 10) - 1;
  const nextCol = colAbsolute ? currentCol : currentCol + colDelta;
  const nextRow = rowAbsolute ? currentRow : currentRow + rowDelta;
  if (nextCol < 0 || nextRow < 0) {
    throw new Error(`Translated reference moved outside worksheet bounds: ${ref}`);
  }
  return `${colAbsolute}${indexToColumn(nextCol)}${rowAbsolute}${nextRow + 1}`;
}

function translateColumnReference(ref: string, colDelta: number): string {
  const match = COLUMN_REF_RE.exec(ref.toUpperCase());
  if (!match) {
    throw new Error(`Invalid column reference '${ref}'`);
  }
  const [, absolute, columnText] = match;
  const currentCol = columnToIndex(columnText!);
  const nextCol = absolute ? currentCol : currentCol + colDelta;
  if (nextCol < 0) {
    throw new Error(`Translated reference moved outside worksheet bounds: ${ref}`);
  }
  return `${absolute}${indexToColumn(nextCol)}`;
}

function translateRowReference(ref: string, rowDelta: number): string {
  const match = ROW_REF_RE.exec(ref);
  if (!match) {
    throw new Error(`Invalid row reference '${ref}'`);
  }
  const [, absolute, rowText] = match;
  const currentRow = Number.parseInt(rowText!, 10) - 1;
  const nextRow = absolute ? currentRow : currentRow + rowDelta;
  if (nextRow < 0) {
    throw new Error(`Translated reference moved outside worksheet bounds: ${ref}`);
  }
  return `${absolute}${nextRow + 1}`;
}

function serializeFormula(node: FormulaNode, parentPrecedence = 0, parentAssociativity: "left" | "right" | null = null): string {
  switch (node.kind) {
    case "NumberLiteral":
      return String(node.value);
    case "BooleanLiteral":
      return node.value ? "TRUE" : "FALSE";
    case "StringLiteral":
      return `"${node.value.replaceAll("\"", "\"\"")}"`;
    case "CellRef":
      return `${formatSheetPrefix(node.sheetName)}${node.ref}`;
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
