import { ErrorCode, ValueTag, type CellValue } from "@bilig/protocol";

export interface LogicalFixtureCase {
  id: string;
  functionName: "IF" | "IFERROR" | "IFNA" | "AND" | "OR" | "NOT" | "ISBLANK" | "ISNUMBER" | "ISTEXT";
  args: CellValue[];
  expected: CellValue;
  notes: string;
  source: "excel-support-docs" | "bilig-contract";
}

const empty = (): CellValue => ({ tag: ValueTag.Empty });
const bool = (value: boolean): CellValue => ({ tag: ValueTag.Boolean, value });
const num = (value: number): CellValue => ({ tag: ValueTag.Number, value });
const text = (value: string): CellValue => ({ tag: ValueTag.String, value, stringId: 0 });
const err = (code: ErrorCode): CellValue => ({ tag: ValueTag.Error, code });

export const logicalFixtureMetadata = {
  group: "logical-info",
  baseline: "Excel for the web desktop semantics, normalized into bilig CellValue tags",
  coverage: ["IF", "IFERROR", "IFNA", "AND", "OR", "NOT", "ISBLANK", "ISNUMBER", "ISTEXT"],
  notes: [
    "Error handling stays explicit in bilig CellValue form rather than formatted strings.",
    "Empty arguments are represented as ValueTag.Empty instead of worksheet references."
  ]
} as const;

export const logicalFixtureCases: readonly LogicalFixtureCase[] = [
  {
    id: "if-true-branch",
    functionName: "IF",
    args: [bool(true), text("yes"), text("no")],
    expected: text("yes"),
    notes: "Boolean TRUE selects the true branch.",
    source: "excel-support-docs"
  },
  {
    id: "if-condition-error",
    functionName: "IF",
    args: [err(ErrorCode.Ref), num(1), num(2)],
    expected: err(ErrorCode.Ref),
    notes: "Condition errors propagate instead of masking.",
    source: "bilig-contract"
  },
  {
    id: "iferror-catches-any-error",
    functionName: "IFERROR",
    args: [err(ErrorCode.Div0), text("fallback")],
    expected: text("fallback"),
    notes: "IFERROR replaces non-NA and NA errors alike.",
    source: "excel-support-docs"
  },
  {
    id: "ifna-catches-na-only",
    functionName: "IFNA",
    args: [err(ErrorCode.NA), text("missing")],
    expected: text("missing"),
    notes: "IFNA handles #N/A and leaves other errors alone.",
    source: "excel-support-docs"
  },
  {
    id: "and-false-on-empty",
    functionName: "AND",
    args: [bool(true), empty()],
    expected: bool(false),
    notes: "Current bilig scalar coercion treats empty as FALSE.",
    source: "bilig-contract"
  },
  {
    id: "or-true-branch",
    functionName: "OR",
    args: [empty(), bool(true)],
    expected: bool(true),
    notes: "Any truthy logical argument yields TRUE.",
    source: "excel-support-docs"
  },
  {
    id: "not-number",
    functionName: "NOT",
    args: [num(2)],
    expected: bool(false),
    notes: "Non-zero numbers coerce to TRUE before inversion.",
    source: "bilig-contract"
  },
  {
    id: "isblank-empty",
    functionName: "ISBLANK",
    args: [empty()],
    expected: bool(true),
    notes: "Only the Empty tag counts as blank.",
    source: "excel-support-docs"
  },
  {
    id: "isnumber-number",
    functionName: "ISNUMBER",
    args: [num(42)],
    expected: bool(true),
    notes: "Numbers are detected without coercing text.",
    source: "excel-support-docs"
  },
  {
    id: "istext-string",
    functionName: "ISTEXT",
    args: [text("hello")],
    expected: bool(true),
    notes: "Text detection is based on the string tag.",
    source: "excel-support-docs"
  }
];
