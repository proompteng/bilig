import { FormulaMode, Opcode, type FormulaRecord } from "@bilig/protocol";
import type { FormulaNode } from "./ast.js";
import { formatRangeAddress, parseRangeAddress } from "./addressing.js";
import { bindFormula, encodeBuiltin } from "./binder.js";
import { lowerToPlan, type JsPlanInstruction } from "./js-evaluator.js";
import { optimizeFormula } from "./optimizer.js";
import { parseFormula } from "./parser.js";
import { rewriteSpecialCall } from "./special-call-rewrites.js";

function encodeInstruction(opcode: Opcode, operand = 0): number {
  return (opcode << 24) | (operand & 0x00ff_ffff);
}

interface CompilerState {
  program: number[];
  constants: number[];
  refs: string[];
  ranges: string[];
  strings: string[];
}

const VOLATILE_BUILTINS = new Set(["TODAY", "NOW", "RAND", "RANDBETWEEN", "RANDARRAY"]);

interface VolatileMetadata {
  volatile: boolean;
  randCallCount: number;
}

function producesSpillResult(node: FormulaNode): boolean {
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
    case "InvokeExpr":
      return false;
    case "RangeRef":
      return true;
    case "UnaryExpr":
      return producesSpillResult(node.argument);
    case "BinaryExpr":
      return producesSpillResult(node.left) || producesSpillResult(node.right);
    case "CallExpr":
      return [
        "SEQUENCE",
        "OFFSET",
        "TAKE",
        "DROP",
        "CHOOSECOLS",
        "CHOOSEROWS",
        "SORT",
        "SORTBY",
        "TOCOL",
        "TOROW",
        "WRAPROWS",
        "WRAPCOLS",
        "FILTER",
        "UNIQUE",
        "MAKEARRAY",
        "MAP",
        "SCAN",
        "BYROW",
        "BYCOL",
        "RANDARRAY",
        "MUNIT",
        "MINVERSE",
        "MMULT",
      ].includes(node.callee.toUpperCase());
  }
}

function analyzeVolatileMetadata(node: FormulaNode): VolatileMetadata {
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
      return { volatile: false, randCallCount: 0 };
    case "UnaryExpr":
      return analyzeVolatileMetadata(node.argument);
    case "BinaryExpr": {
      const left = analyzeVolatileMetadata(node.left);
      const right = analyzeVolatileMetadata(node.right);
      return {
        volatile: left.volatile || right.volatile,
        randCallCount: left.randCallCount + right.randCallCount,
      };
    }
    case "CallExpr": {
      const rewritten = rewriteSpecialCall(node);
      if (rewritten) {
        return analyzeVolatileMetadata(rewritten);
      }
      const callee = node.callee.toUpperCase();
      let volatile = VOLATILE_BUILTINS.has(callee);
      let randCallCount = callee === "RAND" ? 1 : 0;
      node.args.forEach((arg) => {
        const child = analyzeVolatileMetadata(arg);
        volatile = volatile || child.volatile;
        randCallCount += child.randCallCount;
      });
      return { volatile, randCallCount };
    }
    case "InvokeExpr": {
      const callee = analyzeVolatileMetadata(node.callee);
      const args = node.args.map(analyzeVolatileMetadata);
      return {
        volatile: callee.volatile || args.some((child) => child.volatile),
        randCallCount:
          callee.randCallCount + args.reduce((sum, child) => sum + child.randCallCount, 0),
      };
    }
  }
}

function emitCellRef(ref: string, sheetName: string | undefined, state: CompilerState): void {
  const qualifiedRef = sheetName ? `${sheetName}!${ref}` : ref;
  let index = state.refs.indexOf(qualifiedRef);
  if (index === -1) index = state.refs.push(qualifiedRef) - 1;
  state.program.push(encodeInstruction(Opcode.PushCell, index));
}

function emitRangeRef(
  node: Extract<FormulaNode, { kind: "RangeRef" }>,
  state: CompilerState,
): void {
  const qualifiedRange = formatRangeAddress(
    parseRangeAddress(
      node.sheetName ? `${node.sheetName}!${node.start}:${node.end}` : `${node.start}:${node.end}`,
    ),
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
    case "StringLiteral":
      {
        let index = state.strings.indexOf(node.value);
        if (index === -1) index = state.strings.push(node.value) - 1;
        state.program.push(encodeInstruction(Opcode.PushString, index));
      }
      return;
    case "ErrorLiteral":
      state.program.push(encodeInstruction(Opcode.PushError, node.code));
      return;
    case "NameRef":
    case "StructuredRef":
    case "SpillRef":
      throw new Error("Defined names are not supported on the wasm fast path");
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
            "<=": Opcode.Lte,
          }[node.operator],
        ),
      );
      return;
    case "CallExpr":
      {
        const rewritten = rewriteSpecialCall(node);
        if (rewritten) {
          emitNode(rewritten, state);
          return;
        }
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
        state.program.push(
          encodeInstruction(Opcode.CallBuiltin, (encodeBuiltin(callee) << 8) | argc),
        );
      }
      return;
    case "InvokeExpr":
      throw new Error("Lambda invocation is not supported on the wasm fast path");
  }
}

export interface CompiledFormula extends FormulaRecord {
  ast: FormulaNode;
  optimizedAst: FormulaNode;
  deps: string[];
  symbolicNames: string[];
  symbolicTables: string[];
  symbolicSpills: string[];
  volatile: boolean;
  randCallCount: number;
  producesSpill: boolean;
  jsPlan: JsPlanInstruction[];
  program: Uint32Array;
  constants: Float64Array;
  symbolicRefs: string[];
  symbolicRanges: string[];
  symbolicStrings: string[];
}

interface CompileFormulaAstOptions {
  originalAst?: FormulaNode;
  symbolicNames?: string[];
  symbolicTables?: string[];
  symbolicSpills?: string[];
}

function computeMaxStackDepth(plan: readonly JsPlanInstruction[]): number {
  let current = 0;
  let max = 0;
  for (const instruction of plan) {
    switch (instruction.opcode) {
      case "push-number":
      case "push-boolean":
      case "push-string":
      case "push-error":
      case "push-name":
      case "push-cell":
      case "push-range":
      case "push-lambda":
        current += 1;
        break;
      case "binary":
        current -= 1;
        break;
      case "call":
        current -= instruction.argc;
        current += 1;
        break;
      case "invoke":
        current -= instruction.argc;
        current += 1;
        break;
      case "jump-if-false":
        current -= 1;
        break;
      case "bind-name":
        current -= 1;
        break;
      case "unary":
      case "begin-scope":
      case "end-scope":
      case "jump":
      case "return":
        break;
    }
    max = Math.max(max, current);
  }
  return max;
}

export function compileFormulaAst(
  source: string,
  ast: FormulaNode,
  options: CompileFormulaAstOptions = {},
): CompiledFormula {
  const optimizedAst = optimizeFormula(ast);
  const bound = bindFormula(optimizedAst);
  const state: CompilerState = { program: [], constants: [], refs: [], ranges: [], strings: [] };
  const jsPlan = lowerToPlan(optimizedAst);
  const volatileMetadata = analyzeVolatileMetadata(options.originalAst ?? ast);
  const spillResult = producesSpillResult(optimizedAst);

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
    constants: Float64Array.from(state.constants),
    symbolicRefs: state.refs,
    symbolicRanges: state.ranges,
    symbolicStrings: state.strings,
    ast: options.originalAst ?? ast,
    optimizedAst,
    deps: bound.deps,
    symbolicNames: options.symbolicNames ?? bound.symbolicNames,
    symbolicTables: options.symbolicTables ?? bound.symbolicTables,
    symbolicSpills: options.symbolicSpills ?? bound.symbolicSpills,
    volatile: volatileMetadata.volatile,
    randCallCount: volatileMetadata.randCallCount,
    producesSpill: spillResult,
    jsPlan,
    maxStackDepth: computeMaxStackDepth(jsPlan),
  };
}

export function compileFormula(source: string): CompiledFormula {
  return compileFormulaAst(source, parseFormula(source));
}
