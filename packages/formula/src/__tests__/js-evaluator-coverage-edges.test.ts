import { describe, expect, it } from "vitest";
import { ErrorCode, ValueTag, type CellValue } from "@bilig/protocol";
import {
  evaluatePlan,
  evaluatePlanResult,
  lowerToPlan,
  parseFormula,
  type JsPlanInstruction,
} from "../index.js";

const context = {
  sheetName: "Sheet1",
  resolveCell: (_sheetName: string, address: string): CellValue => {
    switch (address) {
      case "A1":
        return { tag: ValueTag.Number, value: 2 };
      case "B1":
        return { tag: ValueTag.Number, value: 3 };
      default:
        return { tag: ValueTag.Empty };
    }
  },
  resolveRange: (_sheetName: string, start: string, end: string): CellValue[] => {
    if (start === "A1" && end === "B2") {
      return [num(2), num(3), { tag: ValueTag.Boolean, value: true }, empty()];
    }
    if (start === "A1" && end === "A2") {
      return [num(2), num(3)];
    }
    if (start === "A1" && end === "A3") {
      return [num(2), num(3), num(4)];
    }
    return [];
  },
};

describe("js evaluator coverage edges", () => {
  it("covers stack guard rails for missing, omitted, lambda, range, and top-level lambda results", () => {
    expect(evaluatePlanResult([], context)).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });

    expect(
      evaluatePlan([{ opcode: "unary", operator: "+" }, { opcode: "return" }], context),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });

    expect(
      evaluatePlan(
        [
          { opcode: "push-lambda", params: ["x"], body: [{ opcode: "return" }] },
          { opcode: "unary", operator: "+" },
          { opcode: "return" },
        ],
        context,
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });

    expect(
      evaluatePlan(
        [
          { opcode: "push-range", start: "A1", end: "B2", refKind: "cells" },
          { opcode: "unary", operator: "+" },
          { opcode: "return" },
        ],
        context,
      ),
    ).toEqual(num(2));

    expect(
      evaluatePlanResult(
        [
          {
            opcode: "push-lambda",
            params: ["x"],
            body: [{ opcode: "push-name", name: "x" }, { opcode: "return" }],
          },
          { opcode: "invoke", argc: 0 },
          { opcode: "return" },
        ],
        context,
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });

    expect(
      evaluatePlanResult(
        [
          { opcode: "push-lambda", params: ["x"], body: [{ opcode: "return" }] },
          { opcode: "return" },
        ],
        context,
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });

    expect(
      evaluatePlan(
        [
          {
            opcode: "push-lambda",
            params: ["x"],
            body: [{ opcode: "push-name", name: "x" }, { opcode: "return" }],
          },
          { opcode: "invoke", argc: 0 },
          { opcode: "push-number", value: 1 },
          { opcode: "binary", operator: "+" },
          { opcode: "return" },
        ],
        context,
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });

    expect(
      evaluatePlan(
        [
          { opcode: "push-lambda", params: ["x"], body: [{ opcode: "return" }] },
          { opcode: "push-number", value: 1 },
          { opcode: "binary", operator: "+" },
          { opcode: "return" },
        ],
        context,
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });
  });

  it("covers incompatible broadcasts and manual lookup guard rails", () => {
    expect(evaluatePlan(lowerToPlan(parseFormula("SEQUENCE(2)+SEQUENCE(3,2)")), context)).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });

    expect(
      evaluatePlan(
        [
          { opcode: "push-range", start: "A1", end: "A2", refKind: "cells" },
          {
            opcode: "lookup-exact-match",
            callee: "MATCH",
            start: "A1",
            end: "A3",
            startRow: 0,
            endRow: 2,
            startCol: 0,
            endCol: 0,
            refKind: "cells",
            searchMode: 1,
          },
          { opcode: "return" },
        ],
        context,
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });

    expect(
      evaluatePlan(
        [
          { opcode: "push-range", start: "A1", end: "A2", refKind: "cells" },
          {
            opcode: "lookup-approximate-match",
            callee: "MATCH",
            start: "A1",
            end: "A3",
            startRow: 0,
            endRow: 2,
            startCol: 0,
            endCol: 0,
            refKind: "cells",
            matchMode: 1,
          },
          { opcode: "return" },
        ],
        context,
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });

    expect(
      evaluatePlan(
        [
          { opcode: "push-number", value: 3 },
          withInvalidLookupCallee(
            makeLookupExactMatchInstruction({
              start: "A1",
              end: "A3",
              startRow: 0,
              endRow: 2,
              startCol: 0,
              endCol: 0,
            }),
          ),
          { opcode: "return" },
        ],
        context,
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Name });

    expect(
      evaluatePlan(
        [
          { opcode: "push-number", value: 3 },
          withInvalidLookupCallee(
            makeLookupApproximateMatchInstruction({
              start: "A1",
              end: "A3",
              startRow: 0,
              endRow: 2,
              startCol: 0,
              endCol: 0,
            }),
          ),
          { opcode: "return" },
        ],
        context,
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Name });
  });

  it("covers GROUPBY aggregate fallbacks for empty subsets, missing builtins, and lambda totals", () => {
    const sumResult = evaluatePlanResult(
      lowerToPlan(parseFormula("GROUPBY(A1:A1,B1:B1,SUM,3,1)")),
      {
        ...context,
        resolveRange: (_sheetName: string, start: string, end: string): CellValue[] => {
          if (start === "A1" && end === "A1") {
            return [text("Group")];
          }
          if (start === "B1" && end === "B1") {
            return [text("Value")];
          }
          return [];
        },
      },
    );
    const averageResult = evaluatePlanResult(
      lowerToPlan(parseFormula("GROUPBY(A1:A1,B1:B1,AVERAGE,3,1)")),
      {
        ...context,
        resolveRange: (_sheetName: string, start: string, end: string): CellValue[] => {
          if (start === "A1" && end === "A1") {
            return [text("Group")];
          }
          if (start === "B1" && end === "B1") {
            return [text("Value")];
          }
          return [];
        },
      },
    );
    const defaultResult = evaluatePlanResult(
      lowerToPlan(parseFormula("GROUPBY(A1:A1,B1:B1,MIN,3,1)")),
      {
        ...context,
        resolveRange: (_sheetName: string, start: string, end: string): CellValue[] => {
          if (start === "A1" && end === "A1") {
            return [text("Group")];
          }
          if (start === "B1" && end === "B1") {
            return [text("Value")];
          }
          return [];
        },
      },
    );
    const missingBuiltinResult = evaluatePlanResult(
      lowerToPlan(parseFormula("GROUPBY(A1:A2,B1:B2,NOPE,3,1)")),
      {
        ...context,
        resolveRange: (_sheetName: string, start: string, end: string): CellValue[] => {
          if (start === "A1" && end === "A2") {
            return [text("Group"), text("East")];
          }
          if (start === "B1" && end === "B2") {
            return [text("Value"), num(10)];
          }
          return [];
        },
      },
    );
    const lambdaTotalResult = evaluatePlanResult(
      lowerToPlan(parseFormula("GROUPBY(A1:A3,B1:B3,LAMBDA(sub,total,COUNTA(total)),3,1)")),
      {
        ...context,
        resolveRange: (_sheetName: string, start: string, end: string): CellValue[] => {
          if (start === "A1" && end === "A3") {
            return [text("Group"), text("East"), text("West")];
          }
          if (start === "B1" && end === "B3") {
            return [text("Value"), num(10), num(7)];
          }
          return [];
        },
      },
    );

    expectArrayContains(sumResult, (value) => value.tag === ValueTag.Number && value.value === 0);
    expectArrayContains(
      averageResult,
      (value) => value.tag === ValueTag.Error && value.code === ErrorCode.Div0,
    );
    expectArrayContains(
      defaultResult,
      (value) => value.tag === ValueTag.Number && value.value === 0,
    );
    expectArrayContains(
      missingBuiltinResult,
      (value) => value.tag === ValueTag.Error && value.code === ErrorCode.Name,
    );
    expectArrayContains(
      lambdaTotalResult,
      (value) => value.tag === ValueTag.Number && value.value === 2,
    );
  });

  it("covers array-helper coercion failures that route through scalar and shape guards", () => {
    expect(evaluatePlan(lowerToPlan(parseFormula('TEXTSPLIT(A1:B2,",")')), context)).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });
    expect(
      evaluatePlan(lowerToPlan(parseFormula('TEXTSPLIT("a,b",",","",SEQUENCE(2))')), context),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });
    expect(
      evaluatePlan(
        lowerToPlan(parseFormula('TEXTSPLIT("a,b",",","",TRUE(),SEQUENCE(2))')),
        context,
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });
    expect(evaluatePlan(lowerToPlan(parseFormula("EXPAND(A1:B2,SEQUENCE(2),3)")), context)).toEqual(
      {
        tag: ValueTag.Error,
        code: ErrorCode.Value,
      },
    );
    expect(
      evaluatePlan(lowerToPlan(parseFormula("TRIMRANGE(A1:B2,SEQUENCE(2))")), context),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });
  });

  it("covers scalar coercion and comparison fallbacks through manual plans", () => {
    expect(
      evaluatePlan(
        [
          { opcode: "push-string", value: "x" },
          { opcode: "push-number", value: 1 },
          { opcode: "binary", operator: "=" },
          { opcode: "return" },
        ],
        context,
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });

    expect(
      evaluatePlan(
        [
          { opcode: "push-error", code: ErrorCode.Ref },
          { opcode: "push-error", code: ErrorCode.Div0 },
          { opcode: "binary", operator: "&" },
          { opcode: "return" },
        ],
        context,
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Ref });

    expect(
      evaluatePlan(
        [
          { opcode: "push-string", value: "x" },
          {
            opcode: "call",
            callee: "MATCH",
            argc: 1,
          },
          { opcode: "return" },
        ],
        context,
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });

    expect(
      evaluatePlan(
        [
          { opcode: "push-lambda", params: ["x"], body: [{ opcode: "return" }] },
          {
            opcode: "call",
            callee: "MATCH",
            argc: 1,
          },
          { opcode: "return" },
        ],
        context,
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });

    expect(
      evaluatePlan(
        [
          { opcode: "push-string", value: "x" },
          {
            opcode: "call",
            callee: "TEXTSPLIT",
            argc: 1,
          },
          { opcode: "return" },
        ],
        context,
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });

    expect(
      evaluatePlan(
        [
          { opcode: "push-number", value: Number.NaN },
          {
            opcode: "call",
            callee: "EXPAND",
            argc: 1,
          },
          { opcode: "return" },
        ],
        context,
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });

    expect(
      evaluatePlan(
        [
          { opcode: "push-error", code: ErrorCode.Ref },
          {
            opcode: "call",
            callee: "TRIMRANGE",
            argc: 2,
          },
          { opcode: "return" },
        ],
        context,
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Ref });

    expect(
      evaluatePlan(
        [
          { opcode: "push-string", value: "x" },
          {
            opcode: "call",
            callee: "TEXTSPLIT",
            argc: 2,
          },
          { opcode: "return" },
        ],
        context,
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });

    expect(
      evaluatePlan(
        [
          { opcode: "push-error", code: ErrorCode.Ref },
          {
            opcode: "call",
            callee: "TEXTSPLIT",
            argc: 5,
          },
          { opcode: "return" },
        ],
        context,
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });

    expect(
      evaluatePlan(
        [
          { opcode: "push-error", code: ErrorCode.Ref },
          {
            opcode: "call",
            callee: "EXPAND",
            argc: 2,
          },
          { opcode: "return" },
        ],
        context,
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Ref });

    expect(
      evaluatePlan(
        [
          { opcode: "push-lambda", params: ["x"], body: [{ opcode: "return" }] },
          {
            opcode: "call",
            callee: "GETPIVOTDATA",
            argc: 1,
          },
          { opcode: "return" },
        ],
        context,
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });
  });
});

type ExactMatchInstruction = Extract<JsPlanInstruction, { opcode: "lookup-exact-match" }>;
type ApproximateMatchInstruction = Extract<
  JsPlanInstruction,
  { opcode: "lookup-approximate-match" }
>;

function makeLookupExactMatchInstruction(
  overrides: Partial<ExactMatchInstruction> = {},
): ExactMatchInstruction {
  return {
    opcode: "lookup-exact-match",
    callee: "MATCH",
    start: "A1",
    end: "A1",
    startRow: 0,
    endRow: 0,
    startCol: 0,
    endCol: 0,
    refKind: "cells",
    searchMode: 1,
    ...overrides,
  };
}

function makeLookupApproximateMatchInstruction(
  overrides: Partial<ApproximateMatchInstruction> = {},
): ApproximateMatchInstruction {
  return {
    opcode: "lookup-approximate-match",
    callee: "MATCH",
    start: "A1",
    end: "A1",
    startRow: 0,
    endRow: 0,
    startCol: 0,
    endCol: 0,
    refKind: "cells",
    matchMode: 1,
    ...overrides,
  };
}

function withInvalidLookupCallee<T extends ExactMatchInstruction | ApproximateMatchInstruction>(
  instruction: T,
): T {
  Reflect.set(instruction, "callee", "NOPE");
  return instruction;
}

function expectArrayContains(
  result: ReturnType<typeof evaluatePlanResult>,
  predicate: (value: CellValue) => boolean,
): void {
  expect(result).toMatchObject({ kind: "array" });
  if (!("kind" in result) || result.kind !== "array") {
    throw new Error("Expected array evaluation result");
  }
  expect(result.values.some(predicate)).toBe(true);
}

function num(value: number): CellValue {
  return { tag: ValueTag.Number, value };
}

function text(value: string): CellValue {
  return { tag: ValueTag.String, value, stringId: 0 };
}

function empty(): CellValue {
  return { tag: ValueTag.Empty };
}
