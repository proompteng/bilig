import { FormulaMode, Opcode, type FormulaRecord } from "@bilig/protocol";
import type { FormulaNode } from "./ast.js";
import { bindFormula, encodeBuiltin } from "./binder.js";
import { parseFormula } from "./parser.js";

function encodeInstruction(opcode: Opcode, operand = 0): number {
  return (opcode << 24) | (operand & 0x00ff_ffff);
}

interface CompilerState {
  program: number[];
  constants: number[];
  refs: string[];
}

function emitNode(node: FormulaNode, state: CompilerState): void {
  switch (node.kind) {
    case "NumberLiteral": {
      const index = state.constants.push(node.value) - 1;
      state.program.push(encodeInstruction(Opcode.PushNumber, index));
      return;
    }
    case "BooleanLiteral":
      state.program.push(encodeInstruction(Opcode.PushBoolean, node.value ? 1 : 0));
      return;
    case "CellRef": {
      const ref = node.sheetName ? `${node.sheetName}!${node.ref}` : node.ref;
      let index = state.refs.indexOf(ref);
      if (index === -1) index = state.refs.push(ref) - 1;
      state.program.push(encodeInstruction(Opcode.PushCell, index));
      return;
    }
    case "UnaryExpr":
      emitNode(node.argument, state);
      if (node.operator === "-") {
        state.program.push(encodeInstruction(Opcode.Neg));
      }
      return;
    case "BinaryExpr":
      emitNode(node.left, state);
      emitNode(node.right, state);
      state.program.push(
        encodeInstruction(
          {
            "+": Opcode.Add,
            "-": Opcode.Sub,
            "*": Opcode.Mul,
            "/": Opcode.Div,
            "^": Opcode.Pow,
            "&": Opcode.Concat,
            "=": Opcode.Eq,
            "<>": Opcode.Neq,
            ">": Opcode.Gt,
            ">=": Opcode.Gte,
            "<": Opcode.Lt,
            "<=": Opcode.Lte
          }[node.operator]
        )
      );
      return;
    case "CallExpr":
      node.args.forEach((arg) => emitNode(arg, state));
      state.program.push(encodeInstruction(Opcode.CallBuiltin, (encodeBuiltin(node.callee) << 8) | node.args.length));
      return;
    default:
      return;
  }
}

export function compileFormula(source: string): FormulaRecord & { ast: FormulaNode; deps: string[] } {
  const ast = parseFormula(source);
  const bound = bindFormula(ast);
  const state: CompilerState = { program: [], constants: [], refs: [] };

  if (bound.mode === FormulaMode.WasmFastPath) {
    emitNode(ast, state);
  }

  state.program.push(encodeInstruction(Opcode.Ret));
  return {
    id: 0,
    source,
    mode: bound.mode,
    program: Uint32Array.from(state.program),
    constants: state.constants,
    symbolicRefs: state.refs,
    ast,
    deps: bound.deps
  };
}
