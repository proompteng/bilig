import { BuiltinId, Opcode } from "./enums.js";

export interface BuiltinDescriptor {
  readonly id: BuiltinId;
  readonly name: string;
  readonly supportsWasm: boolean;
}

export const OPCODE_NAMES: Record<Opcode, string> = {
  [Opcode.PushNumber]: "PushNumber",
  [Opcode.PushBoolean]: "PushBoolean",
  [Opcode.PushCell]: "PushCell",
  [Opcode.PushRange]: "PushRange",
  [Opcode.Add]: "Add",
  [Opcode.Sub]: "Sub",
  [Opcode.Mul]: "Mul",
  [Opcode.Div]: "Div",
  [Opcode.Pow]: "Pow",
  [Opcode.Concat]: "Concat",
  [Opcode.Neg]: "Neg",
  [Opcode.Eq]: "Eq",
  [Opcode.Neq]: "Neq",
  [Opcode.Gt]: "Gt",
  [Opcode.Gte]: "Gte",
  [Opcode.Lt]: "Lt",
  [Opcode.Lte]: "Lte",
  [Opcode.Jump]: "Jump",
  [Opcode.JumpIfFalse]: "JumpIfFalse",
  [Opcode.CallBuiltin]: "CallBuiltin",
  [Opcode.Ret]: "Ret"
};

export const BUILTINS: BuiltinDescriptor[] = [
  { id: BuiltinId.Sum, name: "SUM", supportsWasm: false },
  { id: BuiltinId.Avg, name: "AVG", supportsWasm: false },
  { id: BuiltinId.Min, name: "MIN", supportsWasm: false },
  { id: BuiltinId.Max, name: "MAX", supportsWasm: false },
  { id: BuiltinId.Count, name: "COUNT", supportsWasm: false },
  { id: BuiltinId.CountA, name: "COUNTA", supportsWasm: false },
  { id: BuiltinId.Abs, name: "ABS", supportsWasm: true },
  { id: BuiltinId.Round, name: "ROUND", supportsWasm: false },
  { id: BuiltinId.Floor, name: "FLOOR", supportsWasm: false },
  { id: BuiltinId.Ceiling, name: "CEILING", supportsWasm: false },
  { id: BuiltinId.Mod, name: "MOD", supportsWasm: false },
  { id: BuiltinId.If, name: "IF", supportsWasm: false },
  { id: BuiltinId.And, name: "AND", supportsWasm: false },
  { id: BuiltinId.Or, name: "OR", supportsWasm: false },
  { id: BuiltinId.Not, name: "NOT", supportsWasm: false },
  { id: BuiltinId.Len, name: "LEN", supportsWasm: false },
  { id: BuiltinId.Concat, name: "CONCAT", supportsWasm: false }
];
