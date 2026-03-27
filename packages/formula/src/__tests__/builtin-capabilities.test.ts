import { describe, expect, it } from "vitest";
import {
  builtinJsSpecialNames,
  builtinWasmEnabledNames,
  getBuiltinCapability,
} from "../builtin-capabilities.js";

describe("builtin capabilities", () => {
  it("tracks native production coverage for the current promoted builtin set", () => {
    expect(builtinWasmEnabledNames.has("SUM")).toBe(true);
    expect(builtinWasmEnabledNames.has("COUNTIFS")).toBe(true);
    expect(builtinWasmEnabledNames.has("LET")).toBe(false);
  });

  it("tracks JS-only higher-order builtins separately from native coverage", () => {
    expect(builtinJsSpecialNames.has("LET")).toBe(true);
    expect(builtinJsSpecialNames.has("LAMBDA")).toBe(true);
    expect(builtinJsSpecialNames.has("MAP")).toBe(true);
    expect(builtinJsSpecialNames.has("INDIRECT")).toBe(true);
    expect(builtinJsSpecialNames.has("TEXTSPLIT")).toBe(true);
    expect(builtinJsSpecialNames.has("EXPAND")).toBe(false);
    expect(builtinJsSpecialNames.has("TRIMRANGE")).toBe(false);
  });

  it("exposes array-runtime backlog metadata for non-native families", () => {
    expect(getBuiltinCapability("FILTER")).toMatchObject({
      category: "dynamic-array",
      jsStatus: "implemented",
      wasmStatus: "production",
      needsArrayRuntime: true,
    });
    expect(getBuiltinCapability("BYROW")).toMatchObject({
      category: "lambda",
      jsStatus: "special-js-only",
      wasmStatus: "not-started",
      needsArrayRuntime: true,
    });
    expect(getBuiltinCapability("EXPAND")).toMatchObject({
      category: "dynamic-array",
      jsStatus: "implemented",
      wasmStatus: "production",
      needsArrayRuntime: true,
    });
    expect(getBuiltinCapability("TRIMRANGE")).toMatchObject({
      category: "dynamic-array",
      jsStatus: "implemented",
      wasmStatus: "production",
      needsArrayRuntime: true,
    });
  });
});
