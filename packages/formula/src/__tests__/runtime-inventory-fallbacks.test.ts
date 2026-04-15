import { describe, expect, it } from "vitest";
import {
  getFormulaRuntimeJsStatus,
  getFormulaRuntimeLookupNames,
  getFormulaRuntimeStatus,
  getFormulaRuntimeWasmStatus,
  isLookupBuiltinRuntime,
  isPlaceholderBuiltinRuntime,
  isScalarBuiltinRuntime,
  normalizeFormulaName,
} from "../runtime-inventory.js";

describe("runtime inventory fallbacks", () => {
  it("normalizes names and resolves runtime aliases for non-inventory formulas", () => {
    expect(normalizeFormulaName("  avg  ")).toBe("AVG");
    expect(getFormulaRuntimeLookupNames(" average ")).toEqual(["AVERAGE", "AVG"]);
    expect(getFormulaRuntimeLookupNames(" use.the.countif ")).toEqual([
      "USE.THE.COUNTIF",
      "COUNTIF",
    ]);
    expect(getFormulaRuntimeLookupNames("sum")).toEqual(["SUM"]);
  });

  it("detects scalar, lookup, placeholder, and missing runtime implementations without inventory entries", () => {
    expect(isScalarBuiltinRuntime("avg")).toBe(true);
    expect(isScalarBuiltinRuntime(" totally_missing ")).toBe(false);

    expect(isLookupBuiltinRuntime(" forecast.linear ")).toBe(true);
    expect(isLookupBuiltinRuntime("avg")).toBe(false);

    expect(isPlaceholderBuiltinRuntime("copilot")).toBe(true);
    expect(isPlaceholderBuiltinRuntime("totally_missing")).toBe(false);

    expect(getFormulaRuntimeStatus("AVG")).toBe("implemented");
    expect(getFormulaRuntimeStatus("FORECAST.LINEAR")).toBe("implemented");
    expect(getFormulaRuntimeStatus("copilot")).toBe("placeholder");
    expect(getFormulaRuntimeStatus("totally_missing")).toBe("missing");
  });

  it("derives js and wasm statuses for fallback names outside the generated inventory", () => {
    expect(getFormulaRuntimeJsStatus("AVG")).toBe("implemented");
    expect(getFormulaRuntimeJsStatus("copilot")).toBe("placeholder");
    expect(getFormulaRuntimeJsStatus("totally_missing")).toBe("missing");

    expect(getFormulaRuntimeWasmStatus("AVG")).toBe("production");
    expect(getFormulaRuntimeWasmStatus("FORECAST.LINEAR")).toBe("production");
    expect(getFormulaRuntimeWasmStatus("copilot")).toBe("not-started");
    expect(getFormulaRuntimeWasmStatus("totally_missing")).toBe("not-started");
  });
});
