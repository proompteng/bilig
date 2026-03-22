import { ValueTag, type CellValue } from "@bilig/protocol";
import type { CallExprNode, FormulaNode } from "./ast.js";
import { evaluateAstResult, type EvaluationContext } from "./js-evaluator.js";
import { isArrayValue } from "./runtime-values.js";

const VOLATILE_BUILTINS = new Set(["TODAY", "NOW", "RAND"]);

function cellValueToAst(value: CellValue): FormulaNode | undefined {
  switch (value.tag) {
    case ValueTag.Number:
      return { kind: "NumberLiteral", value: value.value };
    case ValueTag.Boolean:
      return { kind: "BooleanLiteral", value: value.value };
    case ValueTag.String:
      return { kind: "StringLiteral", value: value.value };
    case ValueTag.Empty:
      return undefined;
    case ValueTag.Error:
      return { kind: "ErrorLiteral", code: value.code };
  }
}

function staticContext(): EvaluationContext {
  return {
    sheetName: "Sheet1",
    resolveCell: () => ({ tag: ValueTag.Empty }),
    resolveRange: () => []
  };
}

function isStaticNode(node: FormulaNode): boolean {
  switch (node.kind) {
    case "NumberLiteral":
    case "BooleanLiteral":
    case "StringLiteral":
    case "ErrorLiteral":
      return true;
    case "NameRef":
    case "CellRef":
    case "RowRef":
    case "ColumnRef":
    case "RangeRef":
      return false;
    case "UnaryExpr":
      return isStaticNode(node.argument);
    case "BinaryExpr":
      return isStaticNode(node.left) && isStaticNode(node.right);
    case "CallExpr":
      return node.args.every(isStaticNode);
  }
}

function tryEvaluateStatic(node: FormulaNode): CellValue | undefined {
  if (!isStaticNode(node)) {
    return undefined;
  }

  try {
    const result = evaluateAstResult(node, staticContext());
    return isArrayValue(result) ? undefined : result;
  } catch {
    return undefined;
  }
}

function flattenConcatArgs(args: FormulaNode[]): FormulaNode[] {
  const flattened: FormulaNode[] = [];
  args.forEach((arg) => {
    if (arg.kind === "CallExpr" && arg.callee.toUpperCase() === "CONCAT") {
      flattened.push(...flattenConcatArgs(arg.args));
      return;
    }
    flattened.push(arg);
  });
  return flattened;
}

function optimizeCall(node: CallExprNode): FormulaNode {
  const callee = node.callee.toUpperCase();
  let args = node.args.map(optimizeFormula);

  if (callee === "CONCAT") {
    args = flattenConcatArgs(args);
  }

  if (callee === "IF" && args.length === 3) {
    const conditionValue = tryEvaluateStatic(args[0]!);
    if (conditionValue && conditionValue.tag !== ValueTag.Error) {
      const truthy = conditionValue.tag === ValueTag.Boolean
        ? conditionValue.value
        : conditionValue.tag === ValueTag.Number
          ? conditionValue.value !== 0
          : conditionValue.tag === ValueTag.String
            ? conditionValue.value.length > 0
            : false;
      return optimizeFormula(truthy ? args[1]! : args[2]!);
    }
  }

  const candidate: CallExprNode = {
    kind: "CallExpr",
    callee,
    args
  };

  if (VOLATILE_BUILTINS.has(callee)) {
    return candidate;
  }

  if (args.every(isStaticNode)) {
    const folded = tryEvaluateStatic(candidate);
    const literal = folded ? cellValueToAst(folded) : undefined;
    if (literal) {
      return literal;
    }
  }

  return candidate;
}

export function optimizeFormula(node: FormulaNode): FormulaNode {
  switch (node.kind) {
    case "NumberLiteral":
    case "BooleanLiteral":
    case "StringLiteral":
    case "ErrorLiteral":
    case "NameRef":
    case "CellRef":
    case "RowRef":
    case "ColumnRef":
    case "RangeRef":
      return node;
    case "UnaryExpr": {
      const argument = optimizeFormula(node.argument);
      if (node.operator === "+") {
        return argument;
      }
      const folded = tryEvaluateStatic({ ...node, argument });
      return folded ? cellValueToAst(folded) ?? { ...node, argument } : { ...node, argument };
    }
    case "BinaryExpr": {
      const left = optimizeFormula(node.left);
      const right = optimizeFormula(node.right);
      const folded = tryEvaluateStatic({ ...node, left, right });
      return folded ? cellValueToAst(folded) ?? { ...node, left, right } : { ...node, left, right };
    }
    case "CallExpr":
      return optimizeCall(node);
  }
}
