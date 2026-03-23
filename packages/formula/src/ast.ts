export type FormulaNode =
  | NumberLiteralNode
  | BooleanLiteralNode
  | StringLiteralNode
  | ErrorLiteralNode
  | NameRefNode
  | StructuredRefNode
  | CellRefNode
  | SpillRefNode
  | RowRefNode
  | ColumnRefNode
  | RangeRefNode
  | UnaryExprNode
  | BinaryExprNode
  | CallExprNode
  | InvokeExprNode;

export interface NumberLiteralNode {
  kind: "NumberLiteral";
  value: number;
}

export interface BooleanLiteralNode {
  kind: "BooleanLiteral";
  value: boolean;
}

export interface StringLiteralNode {
  kind: "StringLiteral";
  value: string;
}

export interface ErrorLiteralNode {
  kind: "ErrorLiteral";
  code: number;
}

export interface NameRefNode {
  kind: "NameRef";
  name: string;
}

export interface StructuredRefNode {
  kind: "StructuredRef";
  tableName: string;
  columnName: string;
}

export interface CellRefNode {
  kind: "CellRef";
  ref: string;
  sheetName?: string;
}

export interface SpillRefNode {
  kind: "SpillRef";
  ref: string;
  sheetName?: string;
}

export interface RowRefNode {
  kind: "RowRef";
  ref: string;
  sheetName?: string;
}

export interface ColumnRefNode {
  kind: "ColumnRef";
  ref: string;
  sheetName?: string;
}

export interface RangeRefNode {
  kind: "RangeRef";
  refKind: "cells" | "rows" | "cols";
  start: string;
  end: string;
  sheetName?: string;
}

export interface UnaryExprNode {
  kind: "UnaryExpr";
  operator: "+" | "-";
  argument: FormulaNode;
}

export interface BinaryExprNode {
  kind: "BinaryExpr";
  operator: "+" | "-" | "*" | "/" | "^" | "&" | "=" | "<>" | ">" | ">=" | "<" | "<=";
  left: FormulaNode;
  right: FormulaNode;
}

export interface CallExprNode {
  kind: "CallExpr";
  callee: string;
  args: FormulaNode[];
}

export interface InvokeExprNode {
  kind: "InvokeExpr";
  callee: FormulaNode;
  args: FormulaNode[];
}
