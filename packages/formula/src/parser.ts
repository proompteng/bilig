import type {
  BinaryExprNode,
  CallExprNode,
  CellRefNode,
  ColumnRefNode,
  FormulaNode,
  InvokeExprNode,
  NameRefNode,
  RangeRefNode,
  RowRefNode,
  SpillRefNode,
  StructuredRefNode,
  UnaryExprNode
} from "./ast.js";
import { isCellReferenceText, isColumnReferenceText, isRowReferenceText } from "./addressing.js";
import { lexFormula, type Token } from "./lexer.js";

const PRECEDENCE: Record<string, number> = {
  eq: 1,
  neq: 1,
  gt: 1,
  gte: 1,
  lt: 1,
  lte: 1,
  ampersand: 2,
  plus: 3,
  minus: 3,
  star: 4,
  slash: 4,
  caret: 5,
  colon: 6
};

export function parseFormula(source: string): FormulaNode {
  const tokens = lexFormula(source.startsWith("=") ? source.slice(1) : source);
  let position = 0;

  function current(): Token {
    return tokens[position]!;
  }

  function eat(kind?: Token["kind"]): Token {
    const token = tokens[position]!;
    if (kind && token.kind !== kind) {
      throw new Error(`Expected ${kind}, received ${token.kind}`);
    }
    position += 1;
    return token;
  }

  function maybeParseReferenceValue(ref: string, sheetName?: string): CellRefNode | ColumnRefNode | RowRefNode | undefined {
    const normalized = ref.startsWith("$") && isRowReferenceText(ref) ? ref : ref.toUpperCase();
    const upper = normalized.toUpperCase();
    if (isCellReferenceText(upper)) {
      const result: CellRefNode = { kind: "CellRef", ref: upper };
      if (sheetName !== undefined) {
        result.sheetName = sheetName;
      }
      return result;
    }
    if (isColumnReferenceText(upper)) {
      const result: ColumnRefNode = { kind: "ColumnRef", ref: upper };
      if (sheetName !== undefined) {
        result.sheetName = sheetName;
      }
      return result;
    }
    if (isRowReferenceText(normalized)) {
      const result: RowRefNode = { kind: "RowRef", ref: normalized };
      if (sheetName !== undefined) {
        result.sheetName = sheetName;
      }
      return result;
    }
    return undefined;
  }

  function parseReferenceValue(ref: string, sheetName?: string): CellRefNode | ColumnRefNode | RowRefNode {
    const result = maybeParseReferenceValue(ref, sheetName);
    if (result) {
      return result;
    }
    throw new Error(`Unsupported reference '${ref}'`);
  }

  function parseIdentifierValue(identifier: string): CellRefNode | ColumnRefNode | RowRefNode | NameRefNode {
    const referenceValue = maybeParseReferenceValue(identifier);
    return referenceValue?.kind === "CellRef"
      ? referenceValue
      : ({ kind: "NameRef", name: identifier } satisfies NameRefNode);
  }

  function parseStructuredReference(tableName: string): StructuredRefNode {
    eat("lbracket");
    const token = current();
    if (
      token.kind !== "identifier"
      && token.kind !== "quotedIdentifier"
      && token.kind !== "string"
      && token.kind !== "number"
    ) {
      throw new Error(`Expected a structured reference column, received ${token.kind}`);
    }
    eat(token.kind);
    eat("rbracket");
    return {
      kind: "StructuredRef",
      tableName,
      columnName: token.value
    };
  }

  function parseCallArguments(): FormulaNode[] {
    eat("lparen");
    const args: FormulaNode[] = [];
    if (current().kind !== "rparen") {
      while (true) {
        args.push(parseExpression());
        if (current().kind !== "comma") {
          break;
        }
        eat("comma");
      }
    }
    eat("rparen");
    return args;
  }

  function parseSheetQualifiedReference(sheetName: string): CellRefNode | ColumnRefNode | RowRefNode {
    const token = current();
    if (token.kind === "identifier") {
      eat("identifier");
      return parseReferenceValue(token.value, sheetName);
    }
    if (token.kind === "number" && isRowReferenceText(token.value)) {
      eat("number");
      return { kind: "RowRef", ref: token.value, sheetName };
    }
    throw new Error(`Expected a sheet-qualified reference, received ${token.kind}`);
  }

  function buildRange(left: CellRefNode | ColumnRefNode | RowRefNode, right: CellRefNode | ColumnRefNode | RowRefNode): RangeRefNode {
    if (left.kind !== right.kind) {
      throw new Error("Range endpoints must use the same reference type");
    }

    const sheetName = left.sheetName ?? right.sheetName;
    if (left.sheetName && right.sheetName && left.sheetName !== right.sheetName) {
      throw new Error("Range endpoints must target the same sheet");
    }

    const range: RangeRefNode = {
      kind: "RangeRef",
      refKind: left.kind === "CellRef" ? "cells" : left.kind === "ColumnRef" ? "cols" : "rows",
      start: left.ref,
      end: right.ref
    };
    if (sheetName !== undefined) {
      range.sheetName = sheetName;
    }
    return range;
  }

  function toRangeEndpoint(node: FormulaNode): CellRefNode | ColumnRefNode | RowRefNode | undefined {
    switch (node.kind) {
      case "CellRef":
      case "ColumnRef":
      case "RowRef":
        return node;
      case "NumberLiteral":
        if (Number.isInteger(node.value) && node.value >= 1) {
          return { kind: "RowRef", ref: `${node.value}` };
        }
        return undefined;
      case "BooleanLiteral":
      case "StringLiteral":
      case "ErrorLiteral":
      case "UnaryExpr":
      case "BinaryExpr":
      case "CallExpr":
      case "InvokeExpr":
      case "StructuredRef":
      case "SpillRef":
      case "RangeRef":
        return undefined;
      case "NameRef":
        return maybeParseReferenceValue(node.name);
      default:
        return undefined;
    }
  }

  function assertNoStandaloneAxisRefs(node: FormulaNode): void {
    switch (node.kind) {
      case "NumberLiteral":
      case "BooleanLiteral":
      case "StringLiteral":
      case "ErrorLiteral":
      case "NameRef":
      case "StructuredRef":
      case "SpillRef":
      case "CellRef":
      case "RangeRef":
        return;
      case "ColumnRef":
      case "RowRef":
        throw new Error("Row and column references must appear inside a range");
      case "UnaryExpr":
        assertNoStandaloneAxisRefs(node.argument);
        return;
      case "BinaryExpr":
        assertNoStandaloneAxisRefs(node.left);
        assertNoStandaloneAxisRefs(node.right);
        return;
      case "CallExpr":
        node.args.forEach(assertNoStandaloneAxisRefs);
        return;
      case "InvokeExpr":
        assertNoStandaloneAxisRefs(node.callee);
        node.args.forEach(assertNoStandaloneAxisRefs);
        return;
    }
  }

  function parsePrimary(): FormulaNode {
    const token = current();
    let result: FormulaNode;

    if (token.kind === "number") {
      eat("number");
      if (isRowReferenceText(token.value) && current().kind === "colon") {
        result = { kind: "RowRef", ref: token.value };
      } else {
        result = { kind: "NumberLiteral", value: Number(token.value) };
      }
    } else if (token.kind === "string") {
      eat("string");
      result = { kind: "StringLiteral", value: token.value };
    } else if (token.kind === "quotedIdentifier") {
      const first = eat("quotedIdentifier").value;
      if (current().kind === "bang") {
        eat("bang");
        result = parseSheetQualifiedReference(first);
      } else {
        result = { kind: "StringLiteral", value: first };
      }
    } else if (token.kind === "plus" || token.kind === "minus") {
      eat(token.kind);
      result = {
        kind: "UnaryExpr",
        operator: token.kind === "plus" ? "+" : "-",
        argument: parseExpression(PRECEDENCE["caret"])
      } satisfies UnaryExprNode;
    } else if (token.kind === "lparen") {
      eat("lparen");
      result = parseExpression();
      eat("rparen");
    } else if (token.kind === "identifier") {
      const first = eat("identifier").value;

      if (current().kind === "bang") {
        eat("bang");
        result = parseSheetQualifiedReference(first);
      } else if (current().kind === "lparen") {
        result = { kind: "CallExpr", callee: first.toUpperCase(), args: parseCallArguments() } satisfies CallExprNode;
      } else if (current().kind === "lbracket") {
        result = parseStructuredReference(first);
      } else {
        const upper = first.toUpperCase();
        if (upper === "TRUE" || upper === "FALSE") {
          result = { kind: "BooleanLiteral", value: upper === "TRUE" };
        } else {
          result = parseIdentifierValue(first);
        }
      }
    } else {
      throw new Error(`Unexpected token ${token.kind}`);
    }

    while (current().kind === "lparen" || current().kind === "hash" || current().kind === "percent") {
      if (current().kind === "lparen") {
        result = { kind: "InvokeExpr", callee: result, args: parseCallArguments() } satisfies InvokeExprNode;
        continue;
      }
      if (current().kind === "hash") {
        eat("hash");
        if (result.kind !== "CellRef") {
          throw new Error("Spill references must target a single cell");
        }
        const spill: SpillRefNode = {
          kind: "SpillRef",
          ref: result.ref
        };
        if (result.sheetName !== undefined) {
          spill.sheetName = result.sheetName;
        }
        result = spill;
        continue;
      }

      eat("percent");
      result = {
        kind: "BinaryExpr",
        operator: "*",
        left: result,
        right: { kind: "NumberLiteral", value: 0.01 }
      };
    }

    return result;
  }

  function parseExpression(minPrecedence = 0): FormulaNode {
    let left = parsePrimary();

    while (true) {
      const token = current();
      const precedence = PRECEDENCE[token.kind];
      if (!precedence || precedence < minPrecedence) {
        break;
      }

      eat(token.kind);

      if (token.kind === "colon") {
        const start = toRangeEndpoint(left);
        if (!start) {
          throw new Error("Range start must be a cell reference");
        }
        const right = parsePrimary();
        const end = toRangeEndpoint(right);
        if (!end) {
          throw new Error("Range end must be a cell reference");
        }
        left = buildRange(start, end);
        continue;
      }

      const right = parseExpression(token.kind === "caret" ? precedence : precedence + 1);
      const operatorMap: Record<string, BinaryExprNode["operator"]> = {
        plus: "+",
        minus: "-",
        star: "*",
        slash: "/",
        caret: "^",
        ampersand: "&",
        eq: "=",
        neq: "<>",
        gt: ">",
        gte: ">=",
        lt: "<",
        lte: "<="
      };
      const operator = operatorMap[token.kind];
      if (!operator) {
        throw new Error(`Unsupported operator token ${token.kind}`);
      }
      left = {
        kind: "BinaryExpr",
        operator,
        left,
        right
      };
    }

    return left;
  }

  const result = parseExpression();
  assertNoStandaloneAxisRefs(result);
  eat("eof");
  return result;
}
