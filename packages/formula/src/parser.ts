import type { BinaryExprNode, CallExprNode, CellRefNode, FormulaNode, RangeRefNode, UnaryExprNode } from "./ast.js";
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

  function parsePrimary(): FormulaNode {
    const token = current();

    if (token.kind === "number") {
      eat("number");
      return { kind: "NumberLiteral", value: Number(token.value) };
    }

    if (token.kind === "string") {
      eat("string");
      return { kind: "StringLiteral", value: token.value };
    }

    if (token.kind === "quotedIdentifier") {
      const first = eat("quotedIdentifier").value;
      if (current().kind === "bang") {
        eat("bang");
        const ref = eat("identifier").value.toUpperCase();
        if (current().kind === "colon") {
          eat("colon");
          const end = eat("identifier").value.toUpperCase();
          return { kind: "RangeRef", sheetName: first, start: ref, end } satisfies RangeRefNode;
        }
        return { kind: "CellRef", sheetName: first, ref } satisfies CellRefNode;
      }
      return { kind: "StringLiteral", value: first };
    }

    if (token.kind === "plus" || token.kind === "minus") {
      eat(token.kind);
      return {
        kind: "UnaryExpr",
        operator: token.kind === "plus" ? "+" : "-",
        argument: parseExpression(PRECEDENCE.caret)
      } satisfies UnaryExprNode;
    }

    if (token.kind === "lparen") {
      eat("lparen");
      const inner = parseExpression();
      eat("rparen");
      return inner;
    }

    if (token.kind === "identifier") {
      const first = eat("identifier").value;

      if (current().kind === "bang") {
        eat("bang");
        const ref = eat("identifier").value.toUpperCase();
        if (current().kind === "colon") {
          eat("colon");
          const end = eat("identifier").value.toUpperCase();
          return { kind: "RangeRef", sheetName: first, start: ref, end } satisfies RangeRefNode;
        }
        return { kind: "CellRef", sheetName: first, ref } satisfies CellRefNode;
      }

      if (current().kind === "lparen") {
        eat("lparen");
        const args: FormulaNode[] = [];
        if (current().kind !== "rparen") {
          do {
            args.push(parseExpression());
            if (current().kind !== "comma") break;
            eat("comma");
          } while (true);
        }
        eat("rparen");
        return { kind: "CallExpr", callee: first.toUpperCase(), args } satisfies CallExprNode;
      }

      const upper = first.toUpperCase();
      if (upper === "TRUE" || upper === "FALSE") {
        return { kind: "BooleanLiteral", value: upper === "TRUE" };
      }

      return { kind: "CellRef", ref: upper } satisfies CellRefNode;
    }

    throw new Error(`Unexpected token ${token.kind}`);
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
        if (left.kind !== "CellRef") {
          throw new Error("Range start must be a cell reference");
        }
        const right = parsePrimary();
        if (right.kind !== "CellRef") {
          throw new Error("Range end must be a cell reference");
        }
        const nextRange: RangeRefNode = {
          kind: "RangeRef",
          start: left.ref,
          end: right.ref
        };
        const sheetName = left.sheetName ?? right.sheetName;
        if (sheetName !== undefined) {
          nextRange.sheetName = sheetName;
        }
        left = nextRange;
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
  eat("eof");
  return result;
}
