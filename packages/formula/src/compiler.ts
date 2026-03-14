import { FormulaMode, Opcode, type FormulaRecord } from "@bilig/protocol";
import type { FormulaNode } from "./ast.js";
import { formatRangeAddress, parseRangeAddress } from "./addressing.js";
import { bindFormula, encodeBuiltin } from "./binder.js";
import { lowerToPlan, type JsPlanInstruction } from "./js-evaluator.js";
import { optimizeFormula } from "./optimizer.js";
import { parseFormula } from "./parser.js";

function encodeInstruction(opcode: Opcode, operand = 0): number {
  return (opcode << 24) | (operand & 0x00ff_ffff);
}

interface CompilerState {
  program: number[];
  constants: number[];
  refs: string[];
  ranges: string[];
}

function emitCellRef(ref: string, sheetName: string | undefined, state: CompilerState): void {
  const qualifiedRef = sheetName ? `${sheetName}!${ref}` : ref;
  let index = state.refs.indexOf(qualifiedRef);
  if (index === -1) index = state.refs.push(qualifiedRef) - 1;
  state.program.push(encodeInstruction(Opcode.PushCell, index));
}

function emitRangeRef(node: Extract<FormulaNode, { kind: "RangeRef" }>, state: CompilerState): void {
  const qualifiedRange = formatRangeAddress(
    parseRangeAddress(node.sheetName ? `${node.sheetName}!${node.start}:${node.end}` : `${node.start}:${node.end}`)
  );
  let index = state.ranges.indexOf(qualifiedRange);
  if (index === -1) {
    index = state.ranges.push(qualifiedRange) - 1;
  }
  state.program.push(encodeInstruction(Opcode.PushRange, index));
}

function emitArgument(node: FormulaNode, state: CompilerState): number {
  if (node.kind === "RangeRef") {
    emitRangeRef(node, state);
    return 1;
  }

  emitNode(node, state);
  return 1;
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
    case "RangeRef":
      emitRangeRef(node, state);
      return;
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

export interface CompiledFormula extends FormulaRecord {
  ast: FormulaNode;
  optimizedAst: FormulaNode;
  deps: string[];
  jsPlan: JsPlanInstruction[];
}

function computeMaxStackDepth(plan: readonly JsPlanInstruction[]): number {
  let current = 0;
  let max = 0;
  for (const instruction of plan) {
    switch (instruction.opcode) {
      case "push-number":
      case "push-boolean":
      case "push-string":
      case "push-cell":
      case "push-range":
        current += 1;
        break;
      case "binary":
        current -= 1;
        break;
      case "call":
        current -= instruction.argc;
        current += 1;
        break;
      case "jump-if-false":
        current -= 1;
        break;
      case "unary":
      case "jump":
      case "return":
        break;
    }
    max = Math.max(max, current);
  }
  return max;
}

export function compileFormula(source: string): CompiledFormula {
  const ast = parseFormula(source);
  const optimizedAst = optimizeFormula(ast);
  const bound = bindFormula(optimizedAst);
  const state: CompilerState = { program: [], constants: [], refs: [], ranges: [] };
  const jsPlan = lowerToPlan(optimizedAst);

  if (bound.mode === FormulaMode.WasmFastPath) {
    emitNode(optimizedAst, state);
  }

  state.program.push(encodeInstruction(Opcode.Ret));
  return {
    id: 0,
    source,
    mode: bound.mode,
    depsPtr: 0,
    depsLen: 0,
    programOffset: 0,
    programLength: state.program.length,
    constNumberOffset: 0,
    constNumberLength: state.constants.length,
    rangeListOffset: 0,
    rangeListLength: state.ranges.length,
    program: Uint32Array.from(state.program),
    constants: state.constants,
    symbolicRefs: state.refs,
    symbolicRanges: state.ranges,
    ast,
    optimizedAst,
    deps: bound.deps,
    jsPlan,
    maxStackDepth: computeMaxStackDepth(jsPlan)
  };
}
