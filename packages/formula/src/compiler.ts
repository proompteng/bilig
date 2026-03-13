import { FormulaMode, Opcode, type FormulaRecord } from "@bilig/protocol";
import type { FormulaNode } from "./ast.js";
import { formatAddress, parseRangeAddress } from "./addressing.js";
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

function emitCellRef(ref: string, sheetName: string | undefined, state: CompilerState): void {
  const qualifiedRef = sheetName ? `${sheetName}!${ref}` : ref;
  let index = state.refs.indexOf(qualifiedRef);
  if (index === -1) index = state.refs.push(qualifiedRef) - 1;
  state.program.push(encodeInstruction(Opcode.PushCell, index));
}

function emitArgument(node: FormulaNode, state: CompilerState): number {
  if (node.kind !== "RangeRef") {
    emitNode(node, state);
    return 1;
  }

  const prefix = node.sheetName ? `${node.sheetName}!` : "";
  const range = parseRangeAddress(`${prefix}${node.start}:${node.end}`);
  if (range.kind !== "cells") {
    throw new Error("Only bounded cell ranges are eligible for the wasm fast path");
  }
  let argc = 0;
  for (let row = range.start.row; row <= range.end.row; row += 1) {
    for (let col = range.start.col; col <= range.end.col; col += 1) {
      emitCellRef(formatAddress(row, col), range.sheetName, state);
      argc += 1;
    }
  }
  return argc;
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
      emitCellRef(node.ref, node.sheetName, state);
      return;
    }
    case "RowRef":
    case "ColumnRef":
      throw new Error("Row and column references must appear inside a range");
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
      {
        const callee = node.callee.toUpperCase();
        if (callee === "IF") {
          if (node.args.length !== 3) {
            throw new Error("IF requires exactly three arguments on the wasm fast path");
          }
          emitNode(node.args[0]!, state);
          const jumpIfFalseIndex = state.program.push(encodeInstruction(Opcode.JumpIfFalse, 0)) - 1;
          emitNode(node.args[1]!, state);
          const jumpIndex = state.program.push(encodeInstruction(Opcode.Jump, 0)) - 1;
          const falseBranchStart = state.program.length;
          emitNode(node.args[2]!, state);
          const end = state.program.length;
          state.program[jumpIfFalseIndex] = encodeInstruction(Opcode.JumpIfFalse, falseBranchStart);
          state.program[jumpIndex] = encodeInstruction(Opcode.Jump, end);
          return;
        }
        let argc = 0;
        node.args.forEach((arg) => {
          argc += emitArgument(arg, state);
        });
        state.program.push(encodeInstruction(Opcode.CallBuiltin, (encodeBuiltin(callee) << 8) | argc));
      }
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
