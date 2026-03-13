import { BuiltinId, FormulaMode, type FormulaRecord } from "@bilig/protocol";
import type { FormulaNode } from "./ast.js";
import { getBuiltin } from "./builtins.js";

export interface BoundFormula {
  ast: FormulaNode;
  deps: string[];
  mode: FormulaMode;
}

const WASM_SAFE_BUILTINS = new Set(["ABS"]);

export function bindFormula(ast: FormulaNode): BoundFormula {
  const deps = new Set<string>();
  let wasmSafe = true;

  function visit(node: FormulaNode): void {
    switch (node.kind) {
      case "CellRef":
        deps.add(node.sheetName ? `${node.sheetName}!${node.ref}` : node.ref);
        break;
      case "RangeRef":
        wasmSafe = false;
        deps.add(node.sheetName ? `${node.sheetName}!${node.start}:${node.end}` : `${node.start}:${node.end}`);
        break;
      case "UnaryExpr":
        visit(node.argument);
        break;
      case "BinaryExpr":
        if (!["+", "-", "*", "/"].includes(node.operator)) {
          wasmSafe = false;
        }
        visit(node.left);
        visit(node.right);
        break;
      case "CallExpr":
        if (!getBuiltin(node.callee)) {
          wasmSafe = false;
          break;
        }
        if (!WASM_SAFE_BUILTINS.has(node.callee)) {
          wasmSafe = false;
        }
        node.args.forEach(visit);
        break;
      default:
        break;
    }
  }

  visit(ast);
  return {
    ast,
    deps: [...deps],
    mode: wasmSafe ? FormulaMode.WasmFastPath : FormulaMode.JsOnly
  };
}

export function isBuiltinAvailable(name: string): boolean {
  return getBuiltin(name) !== undefined;
}

export function encodeBuiltin(name: string): BuiltinId {
  const builtins: Record<string, BuiltinId> = {
    ABS: BuiltinId.Abs
  };
  const id = builtins[name.toUpperCase()];
  if (!id) {
    throw new Error(`Unsupported builtin for wasm: ${name}`);
  }
  return id;
}
