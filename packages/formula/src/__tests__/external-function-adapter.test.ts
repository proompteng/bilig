import { afterEach, describe, expect, it } from "vitest";
import { ErrorCode, FormulaMode, ValueTag, type CellValue } from "@bilig/protocol";
import {
  bindFormula,
  clearExternalFunctionAdapters,
  evaluatePlan,
  getBuiltinId,
  installExternalFunctionAdapter,
  isBuiltinAvailable,
  listExternalFunctionAdapterSurfaces,
  lowerToPlan,
  parseFormula
} from "../index.js";

const context = {
  sheetName: "Sheet1",
  resolveCell: (): CellValue => ({ tag: ValueTag.Empty }),
  resolveRange: (_sheetName: string, start: string, end: string): CellValue[] => {
    if (start === "A1" && end === "A3") {
      return [numberValue(1), numberValue(2), numberValue(3)];
    }
    return [];
  }
};

describe("external function adapters", () => {
  afterEach(() => {
    clearExternalFunctionAdapters();
  });

  it("registers host-backed scalar functions without treating them as native wasm builtins", () => {
    installExternalFunctionAdapter({
      surface: "host",
      resolveFunction(name) {
        if (name !== "HOSTDOUBLE") {
          return undefined;
        }
        return {
          kind: "scalar",
          implementation: (value = { tag: ValueTag.Empty }) =>
            value.tag === ValueTag.Number ? numberValue(value.value * 2) : { tag: ValueTag.Error, code: ErrorCode.Value }
        };
      }
    });

    const ast = parseFormula("HOSTDOUBLE(21)");
    expect(listExternalFunctionAdapterSurfaces()).toEqual(["host"]);
    expect(isBuiltinAvailable("HOSTDOUBLE")).toBe(true);
    expect(getBuiltinId("HOSTDOUBLE")).toBeUndefined();
    expect(bindFormula(ast).mode).toBe(FormulaMode.JsOnly);
    expect(evaluatePlan(lowerToPlan(ast), context)).toEqual(numberValue(42));
  });

  it("routes range-aware external functions through the JS lookup path", () => {
    installExternalFunctionAdapter({
      surface: "external-data",
      resolveFunction(name) {
        if (name !== "EXTERNALPICK") {
          return undefined;
        }
        return {
          kind: "lookup",
          implementation: (...args) => {
            const range = args[0];
            if (!range || typeof range !== "object" || !("kind" in range) || range.kind !== "range") {
              return { tag: ValueTag.Error, code: ErrorCode.Value };
            }
            return range.values.at(-1) ?? { tag: ValueTag.Empty };
          }
        };
      }
    });

    const ast = parseFormula("EXTERNALPICK(A1:A3)");
    expect(isBuiltinAvailable("EXTERNALPICK")).toBe(true);
    expect(bindFormula(ast).mode).toBe(FormulaMode.JsOnly);
    expect(evaluatePlan(lowerToPlan(ast), context)).toEqual(numberValue(3));
  });
});

function numberValue(value: number): CellValue {
  return { tag: ValueTag.Number, value };
}
