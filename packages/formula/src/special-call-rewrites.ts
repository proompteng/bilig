import { ErrorCode } from "@bilig/protocol";
import type { BinaryExprNode, CallExprNode, FormulaNode } from "./ast.js";

function errorNode(code: ErrorCode): FormulaNode {
  return { kind: "ErrorLiteral", code };
}

function booleanNode(value: boolean): FormulaNode {
  return { kind: "BooleanLiteral", value };
}

function callNode(callee: string, args: FormulaNode[]): CallExprNode {
  return { kind: "CallExpr", callee: callee.toUpperCase(), args };
}

function binaryNode(operator: BinaryExprNode["operator"], left: FormulaNode, right: FormulaNode): FormulaNode {
  return { kind: "BinaryExpr", operator, left, right };
}

function coerceBooleanNode(node: FormulaNode): FormulaNode {
  return callNode("NOT", [callNode("NOT", [node])]);
}

function rewriteIfs(args: readonly FormulaNode[]): FormulaNode {
  if (args.length < 2 || args.length % 2 !== 0) {
    return errorNode(ErrorCode.Value);
  }

  let fallback: FormulaNode = errorNode(ErrorCode.NA);
  for (let index = args.length - 2; index >= 0; index -= 2) {
    fallback = callNode("IF", [args[index]!, args[index + 1]!, fallback]);
  }
  return fallback;
}

function rewriteSwitch(args: readonly FormulaNode[]): FormulaNode {
  if (args.length < 3) {
    return errorNode(ErrorCode.Value);
  }

  const expression = args[0]!;
  const entries = args.slice(1);
  if (entries.length < 2) {
    return errorNode(ErrorCode.Value);
  }

  const hasDefault = entries.length % 2 === 1;
  let fallback: FormulaNode = hasDefault ? entries[entries.length - 1]! : errorNode(ErrorCode.NA);
  const pairLimit = hasDefault ? entries.length - 1 : entries.length;
  for (let index = pairLimit - 2; index >= 0; index -= 2) {
    fallback = callNode("IF", [
      binaryNode("=", expression, entries[index]!),
      entries[index + 1]!,
      fallback
    ]);
  }
  return fallback;
}

function rewriteXor(args: readonly FormulaNode[]): FormulaNode {
  if (args.length === 0) {
    return errorNode(ErrorCode.Value);
  }

  let expression = coerceBooleanNode(args[0]!);
  for (let index = 1; index < args.length; index += 1) {
    expression = binaryNode("<>", expression, coerceBooleanNode(args[index]!));
  }
  return expression;
}

export function rewriteSpecialCall(node: CallExprNode): FormulaNode | undefined {
  switch (node.callee.toUpperCase()) {
    case "TRUE":
      return node.args.length === 0 ? booleanNode(true) : errorNode(ErrorCode.Value);
    case "FALSE":
      return node.args.length === 0 ? booleanNode(false) : errorNode(ErrorCode.Value);
    case "IFS":
      return rewriteIfs(node.args);
    case "SWITCH":
      return rewriteSwitch(node.args);
    case "XOR":
      return rewriteXor(node.args);
    default:
      return undefined;
  }
}
