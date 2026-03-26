import { ErrorCode, ValueTag } from "@bilig/protocol";
import { afterEach, describe, expect, it, vi } from "vitest";
import {
  clearExternalFunctionAdapters,
  getExternalLookupFunction,
  getExternalScalarFunction,
  hasExternalFunction,
  installExternalFunctionAdapter,
  listExternalFunctionAdapterSurfaces,
  removeExternalFunctionAdapter,
} from "../external-function-adapter.js";

afterEach(() => {
  clearExternalFunctionAdapters();
});

describe("external function adapter registry", () => {
  it("installs, resolves, lists, removes, and clears external adapters", () => {
    const scalar = vi.fn(() => ({ tag: ValueTag.Number, value: 7 }));
    const lookup = vi.fn(() => ({ tag: ValueTag.Error, code: ErrorCode.Blocked }));

    installExternalFunctionAdapter({
      surface: "web",
      resolveFunction(name) {
        if (name === "WEBVALUE") {
          return { kind: "scalar", implementation: scalar };
        }
        return undefined;
      },
    });
    installExternalFunctionAdapter({
      surface: "cube",
      resolveFunction(name) {
        if (name === "CUBEVALUE") {
          return { kind: "lookup", implementation: lookup };
        }
        return undefined;
      },
    });

    expect(hasExternalFunction("")).toBe(false);
    expect(hasExternalFunction(" webvalue ")).toBe(true);
    expect(getExternalScalarFunction("WEBVALUE")).toBe(scalar);
    expect(getExternalLookupFunction("cubevalue")).toBe(lookup);
    expect(listExternalFunctionAdapterSurfaces().toSorted()).toEqual(["cube", "web"]);

    removeExternalFunctionAdapter("cube");
    expect(getExternalLookupFunction("CUBEVALUE")).toBeUndefined();

    clearExternalFunctionAdapters();
    expect(listExternalFunctionAdapterSurfaces()).toEqual([]);
    expect(getExternalScalarFunction("WEBVALUE")).toBeUndefined();
  });
});
