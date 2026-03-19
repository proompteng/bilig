// GENERATED FILE. DO NOT EDIT DIRECTLY.
// Source: scripts/gen-protocol.mjs

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
  { id: BuiltinId.Sum, name: "SUM", supportsWasm: true },
  { id: BuiltinId.Avg, name: "AVG", supportsWasm: true },
  { id: BuiltinId.Min, name: "MIN", supportsWasm: true },
  { id: BuiltinId.Max, name: "MAX", supportsWasm: true },
  { id: BuiltinId.Count, name: "COUNT", supportsWasm: true },
  { id: BuiltinId.CountA, name: "COUNTA", supportsWasm: true },
  { id: BuiltinId.Abs, name: "ABS", supportsWasm: true },
  { id: BuiltinId.Round, name: "ROUND", supportsWasm: true },
  { id: BuiltinId.Floor, name: "FLOOR", supportsWasm: true },
  { id: BuiltinId.Ceiling, name: "CEILING", supportsWasm: true },
  { id: BuiltinId.Mod, name: "MOD", supportsWasm: true },
  { id: BuiltinId.If, name: "IF", supportsWasm: false },
  { id: BuiltinId.And, name: "AND", supportsWasm: true },
  { id: BuiltinId.Or, name: "OR", supportsWasm: true },
  { id: BuiltinId.Not, name: "NOT", supportsWasm: true },
  { id: BuiltinId.Len, name: "LEN", supportsWasm: true },
  { id: BuiltinId.Concat, name: "CONCAT", supportsWasm: false },
  { id: BuiltinId.IsBlank, name: "ISBLANK", supportsWasm: true },
  { id: BuiltinId.IsNumber, name: "ISNUMBER", supportsWasm: true },
  { id: BuiltinId.IsText, name: "ISTEXT", supportsWasm: true },
  { id: BuiltinId.Date, name: "DATE", supportsWasm: true },
  { id: BuiltinId.Year, name: "YEAR", supportsWasm: true },
  { id: BuiltinId.Month, name: "MONTH", supportsWasm: true },
  { id: BuiltinId.Day, name: "DAY", supportsWasm: true },
  { id: BuiltinId.Edate, name: "EDATE", supportsWasm: true },
  { id: BuiltinId.Eomonth, name: "EOMONTH", supportsWasm: true },
  { id: BuiltinId.Exact, name: "EXACT", supportsWasm: true },
  { id: BuiltinId.Int, name: "INT", supportsWasm: true },
  { id: BuiltinId.RoundUp, name: "ROUNDUP", supportsWasm: true },
  { id: BuiltinId.RoundDown, name: "ROUNDDOWN", supportsWasm: true },
  { id: BuiltinId.Time, name: "TIME", supportsWasm: true },
  { id: BuiltinId.Hour, name: "HOUR", supportsWasm: true },
  { id: BuiltinId.Minute, name: "MINUTE", supportsWasm: true },
  { id: BuiltinId.Second, name: "SECOND", supportsWasm: true },
  { id: BuiltinId.Weekday, name: "WEEKDAY", supportsWasm: true }
];
