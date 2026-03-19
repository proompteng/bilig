import { ErrorCode, ValueTag, type CellValue } from "@bilig/protocol";

export interface TextFixtureCase {
  name: string;
  args: readonly CellValue[];
  expected: CellValue;
  note?: string;
}

export interface TextFixtureGroup {
  builtin: "LEN" | "CONCAT" | "LEFT" | "RIGHT" | "MID" | "TRIM" | "UPPER" | "LOWER" | "FIND" | "SEARCH";
  cases: readonly TextFixtureCase[];
}

export const TEXT_FIXTURE_METADATA = {
  source: "excel-web-like text builtin tranche",
  version: 1,
  builtins: ["LEN", "CONCAT", "LEFT", "RIGHT", "MID", "TRIM", "UPPER", "LOWER", "FIND", "SEARCH"] as const
} as const;

export const TEXT_FIXTURES: readonly TextFixtureGroup[] = [
  {
    builtin: "LEN",
    cases: [
      { name: "counts plain string length", args: [text("hello")], expected: number(5) },
      { name: "coerces booleans to text", args: [bool(true)], expected: number(4) },
      { name: "treats empty as empty string", args: [empty()], expected: number(0) }
    ]
  },
  {
    builtin: "CONCAT",
    cases: [
      { name: "joins mixed scalar values", args: [text("alpha"), number(2), empty()], expected: text("alpha2") },
      { name: "coerces booleans to uppercase logical text", args: [bool(false), text("-ok")], expected: text("FALSE-ok") }
    ]
  },
  {
    builtin: "LEFT",
    cases: [
      { name: "defaults to one character", args: [text("alpha")], expected: text("a") },
      { name: "takes requested prefix length", args: [text("alpha"), number(3)], expected: text("alp") },
      { name: "zero length returns empty string", args: [text("alpha"), empty()], expected: text("") }
    ]
  },
  {
    builtin: "RIGHT",
    cases: [
      { name: "defaults to one character", args: [text("alpha")], expected: text("a") },
      { name: "takes requested suffix length", args: [text("alpha"), number(2)], expected: text("ha") },
      { name: "large suffix returns whole string", args: [text("alpha"), number(99)], expected: text("alpha") }
    ]
  },
  {
    builtin: "MID",
    cases: [
      { name: "extracts substring from one-based start", args: [text("alphabet"), number(2), number(3)], expected: text("lph") },
      { name: "start beyond end returns empty string", args: [text("alpha"), number(9), number(2)], expected: text("") },
      { name: "zero count returns empty string", args: [text("alpha"), number(2), empty()], expected: text("") }
    ]
  },
  {
    builtin: "TRIM",
    cases: [
      { name: "collapses internal spaces", args: [text("  alpha   beta  ")], expected: text("alpha beta") },
      { name: "leaves clean strings alone", args: [text("alpha beta")], expected: text("alpha beta") }
    ]
  },
  {
    builtin: "UPPER",
    cases: [
      { name: "uppercases text", args: [text("Alpha beta")], expected: text("ALPHA BETA") }
    ]
  },
  {
    builtin: "LOWER",
    cases: [
      { name: "lowercases text", args: [text("Alpha BETA")], expected: text("alpha beta") }
    ]
  },
  {
    builtin: "FIND",
    cases: [
      { name: "finds first case-sensitive position", args: [text("ph"), text("alphabet")], expected: number(3) },
      { name: "respects one-based start", args: [text("a"), text("bananas"), number(3)], expected: number(4) },
      { name: "empty needle returns start position", args: [text(""), text("alpha"), number(3)], expected: number(3) }
    ]
  },
  {
    builtin: "SEARCH",
    cases: [
      { name: "searches case-insensitively", args: [text("PH"), text("alphabet")], expected: number(3) },
      { name: "supports wildcard question mark", args: [text("b?d"), text("ABCD")], expected: number(2) },
      { name: "supports escaped wildcard", args: [text("~*"), text("a*b")], expected: number(2) }
    ]
  }
];

function empty(): CellValue {
  return { tag: ValueTag.Empty };
}

function number(value: number): CellValue {
  return { tag: ValueTag.Number, value };
}

function bool(value: boolean): CellValue {
  return { tag: ValueTag.Boolean, value };
}

function text(value: string): CellValue {
  return { tag: ValueTag.String, value, stringId: 0 };
}

export function textValueError(): CellValue {
  return { tag: ValueTag.Error, code: ErrorCode.Value };
}
