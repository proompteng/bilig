import { BuiltinId, ErrorCode, ValueTag } from "./protocol";

export function applyBuiltin(
  builtinId: i32,
  argc: i32,
  valueStack: Float64Array,
  tagStack: Uint8Array,
  sp: i32
): i32 {
  if (builtinId == BuiltinId.Abs && argc == 1) {
    valueStack[sp - 1] = Math.abs(valueStack[sp - 1]);
    tagStack[sp - 1] = ValueTag.Number;
    return sp;
  }

  valueStack[sp - 1] = ErrorCode.Value;
  tagStack[sp - 1] = ValueTag.Error;
  return sp;
}
