import { ErrorCode } from "@bilig/protocol";
import type { FormulaNode } from "./ast.js";
import { hasBuiltin } from "./builtins.js";
import type { JsPlanInstruction, ReferenceOperand } from "./js-evaluator-types.js";
import { rewriteSpecialCall } from "./special-call-rewrites.js";

function referenceOperandFromNode(node: FormulaNode): ReferenceOperand | undefined {
  switch (node.kind) {
    case "CellRef":
      return {
        kind: "cell",
        ...(node.sheetName === undefined ? {} : { sheetName: node.sheetName }),
        address: node.ref,
      };
    case "RangeRef":
      return {
        kind: "range",
        ...(node.sheetName === undefined ? {} : { sheetName: node.sheetName }),
        start: node.start,
        end: node.end,
        refKind: node.refKind,
      };
    case "RowRef":
      return {
        kind: "row",
        ...(node.sheetName === undefined ? {} : { sheetName: node.sheetName }),
        address: node.ref,
      };
    case "ColumnRef":
      return {
        kind: "col",
        ...(node.sheetName === undefined ? {} : { sheetName: node.sheetName }),
        address: node.ref,
      };
    case "BinaryExpr":
    case "BooleanLiteral":
    case "CallExpr":
    case "ErrorLiteral":
    case "InvokeExpr":
    case "NameRef":
    case "NumberLiteral":
    case "SpillRef":
    case "StringLiteral":
    case "StructuredRef":
    case "UnaryExpr":
      return undefined;
    default:
      return undefined;
  }
}

function lowerNode(node: FormulaNode, plan: JsPlanInstruction[]): void {
  switch (node.kind) {
    case "NumberLiteral":
      plan.push({ opcode: "push-number", value: node.value });
      return;
    case "BooleanLiteral":
      plan.push({ opcode: "push-boolean", value: node.value });
      return;
    case "StringLiteral":
      plan.push({ opcode: "push-string", value: node.value });
      return;
    case "ErrorLiteral":
      plan.push({ opcode: "push-error", code: node.code as ErrorCode });
      return;
    case "NameRef":
      plan.push({ opcode: "push-name", name: node.name });
      return;
    case "StructuredRef":
    case "SpillRef":
      plan.push({ opcode: "push-error", code: ErrorCode.Ref });
      return;
    case "CellRef":
      plan.push(
        node.sheetName
          ? { opcode: "push-cell", sheetName: node.sheetName, address: node.ref }
          : { opcode: "push-cell", address: node.ref },
      );
      return;
    case "RangeRef":
      plan.push(
        node.sheetName
          ? {
              opcode: "push-range",
              sheetName: node.sheetName,
              start: node.start,
              end: node.end,
              refKind: node.refKind,
            }
          : {
              opcode: "push-range",
              start: node.start,
              end: node.end,
              refKind: node.refKind,
            },
      );
      return;
    case "RowRef":
    case "ColumnRef":
      plan.push({ opcode: "push-number", value: Number.NaN });
      return;
    case "UnaryExpr":
      lowerNode(node.argument, plan);
      plan.push({ opcode: "unary", operator: node.operator });
      return;
    case "BinaryExpr":
      lowerNode(node.left, plan);
      lowerNode(node.right, plan);
      plan.push({ opcode: "binary", operator: node.operator });
      return;
    case "InvokeExpr":
      lowerNode(node.callee, plan);
      node.args.forEach((arg) => lowerNode(arg, plan));
      plan.push({ opcode: "invoke", argc: node.args.length });
      return;
    case "CallExpr": {
      const rewritten = rewriteSpecialCall(node);
      if (rewritten) {
        lowerNode(rewritten, plan);
        return;
      }
      const callee = node.callee.toUpperCase();
      if (callee === "LAMBDA") {
        if (node.args.length < 1) {
          plan.push({ opcode: "push-error", code: ErrorCode.Value });
          return;
        }
        const params: string[] = [];
        for (let index = 0; index < node.args.length - 1; index += 1) {
          const paramNode = node.args[index]!;
          if (paramNode.kind !== "NameRef") {
            plan.push({ opcode: "push-error", code: ErrorCode.Value });
            return;
          }
          params.push(paramNode.name);
        }
        const body: JsPlanInstruction[] = [];
        lowerNode(node.args[node.args.length - 1]!, body);
        body.push({ opcode: "return" });
        plan.push({ opcode: "push-lambda", params, body });
        return;
      }
      if (callee === "LET") {
        if (node.args.length < 3 || node.args.length % 2 === 0) {
          plan.push({ opcode: "push-error", code: ErrorCode.Value });
          return;
        }
        plan.push({ opcode: "begin-scope" });
        for (let index = 0; index < node.args.length - 1; index += 2) {
          const nameNode = node.args[index]!;
          if (nameNode.kind !== "NameRef") {
            plan.push({ opcode: "push-error", code: ErrorCode.Value });
            plan.push({ opcode: "end-scope" });
            return;
          }
          lowerNode(node.args[index + 1]!, plan);
          plan.push({ opcode: "bind-name", name: nameNode.name });
        }
        lowerNode(node.args[node.args.length - 1]!, plan);
        plan.push({ opcode: "end-scope" });
        return;
      }
      if (callee === "IF" && node.args.length === 3) {
        lowerNode(node.args[0]!, plan);
        const jumpIfFalseIndex = plan.push({ opcode: "jump-if-false", target: -1 }) - 1;
        lowerNode(node.args[1]!, plan);
        const jumpIndex = plan.push({ opcode: "jump", target: -1 }) - 1;
        const falseTarget = plan.length;
        lowerNode(node.args[2]!, plan);
        const endTarget = plan.length;
        plan[jumpIfFalseIndex] = { opcode: "jump-if-false", target: falseTarget };
        plan[jumpIndex] = { opcode: "jump", target: endTarget };
        return;
      }
      if (!hasBuiltin(callee)) {
        lowerNode({ kind: "NameRef", name: node.callee }, plan);
        node.args.forEach((arg) => lowerNode(arg, plan));
        plan.push({ opcode: "invoke", argc: node.args.length });
        return;
      }

      const aggregateArgumentIndex = callee === "GROUPBY" ? 2 : callee === "PIVOTBY" ? 3 : -1;
      node.args.forEach((arg, index) => {
        if (index === aggregateArgumentIndex && arg.kind === "NameRef") {
          plan.push({ opcode: "push-string", value: arg.name });
          return;
        }
        lowerNode(arg, plan);
      });
      plan.push({
        opcode: "call",
        callee,
        argc: node.args.length,
        argRefs: node.args.map((arg) => referenceOperandFromNode(arg)),
      });
    }
  }
}

export function lowerToPlan(node: FormulaNode): JsPlanInstruction[] {
  const plan: JsPlanInstruction[] = [];
  lowerNode(node, plan);
  plan.push({ opcode: "return" });
  return plan;
}
