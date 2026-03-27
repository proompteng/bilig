import { ErrorCode, ValueTag, type CellValue } from "@bilig/protocol";
import type { CallExprNode, FormulaNode, InvokeExprNode } from "./ast.js";
import { evaluateAstResult, type EvaluationContext } from "./js-evaluator.js";
import { isArrayValue } from "./runtime-values.js";
import { rewriteSpecialCall } from "./special-call-rewrites.js";

const VOLATILE_BUILTINS = new Set(["TODAY", "NOW", "RAND"]);
const CONTEXTUAL_BUILTINS = new Set(["CELL", "COLUMN", "FORMULATEXT", "ROW", "SHEET", "SHEETS"]);

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
    resolveRange: () => [],
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
    case "StructuredRef":
    case "CellRef":
    case "SpillRef":
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
    case "InvokeExpr":
      return isStaticNode(node.callee) && node.args.every(isStaticNode);
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

function normalizeName(name: string): string {
  return name.trim().toUpperCase();
}

function cloneFormulaNode(node: FormulaNode): FormulaNode {
  switch (node.kind) {
    case "NumberLiteral":
    case "BooleanLiteral":
    case "StringLiteral":
    case "ErrorLiteral":
    case "NameRef":
      return { ...node };
    case "StructuredRef":
      return { ...node };
    case "CellRef":
    case "SpillRef":
    case "RowRef":
    case "ColumnRef":
      return { ...node };
    case "RangeRef":
      return { ...node };
    case "UnaryExpr":
      return { ...node, argument: cloneFormulaNode(node.argument) };
    case "BinaryExpr":
      return {
        ...node,
        left: cloneFormulaNode(node.left),
        right: cloneFormulaNode(node.right),
      };
    case "CallExpr":
      return { ...node, args: node.args.map(cloneFormulaNode) };
    case "InvokeExpr":
      return {
        ...node,
        callee: cloneFormulaNode(node.callee),
        args: node.args.map(cloneFormulaNode),
      };
  }
}

function substituteNames(
  node: FormulaNode,
  replacements: ReadonlyMap<string, FormulaNode>,
  shadowed = new Set<string>(),
): FormulaNode {
  switch (node.kind) {
    case "NumberLiteral":
    case "BooleanLiteral":
    case "StringLiteral":
    case "ErrorLiteral":
    case "StructuredRef":
    case "CellRef":
    case "SpillRef":
    case "RowRef":
    case "ColumnRef":
    case "RangeRef":
      return cloneFormulaNode(node);
    case "NameRef": {
      const key = normalizeName(node.name);
      if (shadowed.has(key)) {
        return { ...node };
      }
      const replacement = replacements.get(key);
      return replacement ? cloneFormulaNode(replacement) : { ...node };
    }
    case "UnaryExpr":
      return { ...node, argument: substituteNames(node.argument, replacements, shadowed) };
    case "BinaryExpr":
      return {
        ...node,
        left: substituteNames(node.left, replacements, shadowed),
        right: substituteNames(node.right, replacements, shadowed),
      };
    case "CallExpr": {
      const callee = node.callee.toUpperCase();
      if (callee === "LAMBDA" && node.args.length >= 1) {
        const params = node.args.slice(0, -1);
        const body = node.args[node.args.length - 1]!;
        const innerShadowed = new Set(shadowed);
        params.forEach((param) => {
          if (param.kind === "NameRef") {
            innerShadowed.add(normalizeName(param.name));
          }
        });
        return {
          ...node,
          callee,
          args: [
            ...params.map(cloneFormulaNode),
            substituteNames(body, replacements, innerShadowed),
          ],
        };
      }

      const replacement = shadowed.has(callee) ? undefined : replacements.get(callee);
      if (replacement?.kind === "CallExpr" && replacement.callee.toUpperCase() === "LAMBDA") {
        return {
          kind: "InvokeExpr",
          callee: cloneFormulaNode(replacement),
          args: node.args.map((arg) => substituteNames(arg, replacements, shadowed)),
        };
      }

      return {
        ...node,
        callee,
        args: node.args.map((arg) => substituteNames(arg, replacements, shadowed)),
      };
    }
    case "InvokeExpr":
      return {
        ...node,
        callee: substituteNames(node.callee, replacements, shadowed),
        args: node.args.map((arg) => substituteNames(arg, replacements, shadowed)),
      };
  }
}

function rewriteLet(node: CallExprNode): FormulaNode | undefined {
  if (node.args.length < 3 || node.args.length % 2 === 0) {
    return { kind: "ErrorLiteral", code: ErrorCode.Value };
  }

  const replacements = new Map<string, FormulaNode>();
  const lastArgIndex = node.args.length - 1;
  for (let index = 0; index < lastArgIndex; index += 2) {
    const nameArg = node.args[index];
    const valueArg = node.args[index + 1];
    if (nameArg?.kind !== "NameRef" || valueArg === undefined) {
      return { kind: "ErrorLiteral", code: ErrorCode.Value };
    }
    replacements.set(normalizeName(nameArg.name), substituteNames(valueArg, replacements));
  }
  return substituteNames(node.args[lastArgIndex]!, replacements);
}

function rewriteImmediateLambdaInvoke(node: InvokeExprNode): FormulaNode | undefined {
  const callee = node.callee;
  if (
    callee.kind !== "CallExpr" ||
    callee.callee.toUpperCase() !== "LAMBDA" ||
    callee.args.length < 1
  ) {
    return undefined;
  }

  const params = callee.args.slice(0, -1);
  const body = callee.args[callee.args.length - 1]!;
  if (node.args.length > params.length) {
    return { kind: "ErrorLiteral", code: ErrorCode.Value };
  }
  if (node.args.length !== params.length) {
    return undefined;
  }

  const replacements = new Map<string, FormulaNode>();
  for (let index = 0; index < params.length; index += 1) {
    const param = params[index];
    if (param?.kind !== "NameRef") {
      return { kind: "ErrorLiteral", code: ErrorCode.Value };
    }
    replacements.set(normalizeName(param.name), node.args[index]!);
  }
  return substituteNames(body, replacements);
}

function rewriteMap(node: CallExprNode): FormulaNode | undefined {
  if (node.args.length < 2) {
    return { kind: "ErrorLiteral", code: ErrorCode.Value };
  }
  const lambda = node.args[node.args.length - 1]!;
  if (
    lambda.kind !== "CallExpr" ||
    lambda.callee.toUpperCase() !== "LAMBDA" ||
    lambda.args.length < 1
  ) {
    return undefined;
  }

  const params = lambda.args.slice(0, -1);
  const body = lambda.args[lambda.args.length - 1]!;
  const inputs = node.args.slice(0, -1);
  if (params.length !== inputs.length) {
    return { kind: "ErrorLiteral", code: ErrorCode.Value };
  }

  const replacements = new Map<string, FormulaNode>();
  for (let index = 0; index < params.length; index += 1) {
    const param = params[index];
    if (param?.kind !== "NameRef") {
      return { kind: "ErrorLiteral", code: ErrorCode.Value };
    }
    replacements.set(normalizeName(param.name), inputs[index]!);
  }

  return substituteNames(body, replacements);
}

function optimizeCall(node: CallExprNode): FormulaNode {
  const callee = node.callee.toUpperCase();
  let args = node.args.map(optimizeFormula);

  if (callee === "CONCAT") {
    args = flattenConcatArgs(args);
  }

  if (callee === "LET") {
    const rewritten = rewriteLet({ kind: "CallExpr", callee, args });
    if (rewritten) {
      return optimizeFormula(rewritten);
    }
  }

  if (callee === "MAP") {
    const rewritten = rewriteMap({ kind: "CallExpr", callee, args });
    if (rewritten) {
      return optimizeFormula(rewritten);
    }
  }

  if (callee === "IF" && args.length === 3) {
    const conditionValue = tryEvaluateStatic(args[0]!);
    if (conditionValue && conditionValue.tag !== ValueTag.Error) {
      const truthy =
        conditionValue.tag === ValueTag.Boolean
          ? conditionValue.value
          : conditionValue.tag === ValueTag.Number
            ? conditionValue.value !== 0
            : conditionValue.tag === ValueTag.String
              ? conditionValue.value.length > 0
              : false;
      return optimizeFormula(truthy ? args[1]! : args[2]!);
    }
  }

  const rewritten = rewriteSpecialCall({
    kind: "CallExpr",
    callee,
    args,
  });
  if (rewritten) {
    return optimizeFormula(rewritten);
  }

  const candidate: CallExprNode = {
    kind: "CallExpr",
    callee,
    args,
  };

  if (VOLATILE_BUILTINS.has(callee) || CONTEXTUAL_BUILTINS.has(callee)) {
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
    case "StructuredRef":
    case "CellRef":
    case "SpillRef":
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
      return folded ? (cellValueToAst(folded) ?? { ...node, argument }) : { ...node, argument };
    }
    case "BinaryExpr": {
      const left = optimizeFormula(node.left);
      const right = optimizeFormula(node.right);
      const folded = tryEvaluateStatic({ ...node, left, right });
      return folded
        ? (cellValueToAst(folded) ?? { ...node, left, right })
        : { ...node, left, right };
    }
    case "CallExpr":
      return optimizeCall(node);
    case "InvokeExpr": {
      const callee = optimizeFormula(node.callee);
      const args = node.args.map(optimizeFormula);
      const candidate = { ...node, callee, args };
      const rewritten = rewriteImmediateLambdaInvoke(candidate);
      if (rewritten) {
        return optimizeFormula(rewritten);
      }
      const folded = isStaticNode(candidate) ? tryEvaluateStatic(candidate) : undefined;
      return folded ? (cellValueToAst(folded) ?? candidate) : candidate;
    }
  }
}
