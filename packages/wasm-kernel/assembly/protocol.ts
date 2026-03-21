// GENERATED FILE. DO NOT EDIT DIRECTLY.
// Source: scripts/gen-protocol.mjs

export enum ValueTag {
  Empty = 0,
  Number = 1,
  Boolean = 2,
  String = 3,
  Error = 4
}

export enum ErrorCode {
  None = 0,
  Div0 = 1,
  Ref = 2,
  Value = 3,
  Name = 4,
  NA = 5,
  Cycle = 6,
  Spill = 7,
  Blocked = 8
}

export enum FormulaMode {
  JsOnly = 0,
  WasmFastPath = 1
}

export enum Opcode {
  PushNumber = 1,
  PushBoolean = 2,
  PushCell = 3,
  PushRange = 4,
  Add = 5,
  Sub = 6,
  Mul = 7,
  Div = 8,
  Pow = 9,
  Concat = 10,
  Neg = 11,
  Eq = 12,
  Neq = 13,
  Gt = 14,
  Gte = 15,
  Lt = 16,
  Lte = 17,
  Jump = 18,
  JumpIfFalse = 19,
  CallBuiltin = 20,
  Ret = 255
}

export enum BuiltinId {
  Sum = 1,
  Avg = 2,
  Min = 3,
  Max = 4,
  Count = 5,
  CountA = 6,
  Abs = 7,
  Round = 8,
  Floor = 9,
  Ceiling = 10,
  Mod = 11,
  If = 12,
  And = 13,
  Or = 14,
  Not = 15,
  Len = 16,
  Concat = 17,
  IsBlank = 18,
  IsNumber = 19,
  IsText = 20,
  Date = 21,
  Year = 22,
  Month = 23,
  Day = 24,
  Edate = 25,
  Eomonth = 26,
  Exact = 27,
  Int = 28,
  RoundUp = 29,
  RoundDown = 30,
  Time = 31,
  Hour = 32,
  Minute = 33,
  Second = 34,
  Weekday = 35
}
