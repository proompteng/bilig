export type FormulaNode =
  | NumberLiteralNode
  | BooleanLiteralNode
  | StringLiteralNode
  | NameRefNode
  | CellRefNode
  | RowRefNode
  | ColumnRefNode
  | RangeRefNode
  | UnaryExprNode
  | BinaryExprNode
  | CallExprNode;

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

export interface NameRefNode {
  kind: "NameRef";
  name: string;
}

export interface CellRefNode {
  kind: "CellRef";
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
