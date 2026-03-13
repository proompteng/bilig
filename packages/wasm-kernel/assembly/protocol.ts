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
  Cycle = 6
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
  Abs = 7
}
