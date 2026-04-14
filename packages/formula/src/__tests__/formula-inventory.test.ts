import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { formulaInventory, formulaInventorySummary } from "../generated/formula-inventory.js";
import {
  getFormulaRuntimeJsStatus,
  getFormulaRuntimeStatus,
  getFormulaRuntimeWasmStatus,
  normalizeFormulaName,
} from "../runtime-inventory.js";

interface FormulaInventorySource {
  version: number;
  entries: Array<{ name: string; odfStatus: string; inOfficeList: boolean }>;
}

describe("formula inventory", () => {
  it("tracks the canonical unified formula count", () => {
    expect(formulaInventorySummary.total).toBe(formulaInventory.length);
    expect(formulaInventory.length).toBeGreaterThan(500);
  });

  it("keeps the generated inventory aligned with the canonical source inventory", () => {
    const source = readFormulaInventorySource();
    expect(source.version).toBe(1);
    expect(formulaInventory).toHaveLength(source.entries.length);
    expect(formulaInventory.map((entry) => entry.name)).toEqual(
      source.entries.map((entry) => normalizeFormulaName(entry.name)),
    );
  });

  it("keeps generated runtime statuses aligned with the live formula runtime", () => {
    formulaInventory.forEach((entry) => {
      const name = normalizeFormulaName(entry.name);
      expect(entry.runtimeStatus).toBe(getFormulaRuntimeStatus(name));
      expect(entry.jsStatus).toBe(getFormulaRuntimeJsStatus(name));
      expect(entry.wasmStatus).toBe(getFormulaRuntimeWasmStatus(name));
      expect(entry.placeholder).toBe(entry.runtimeStatus === "placeholder");
      expect(entry.registeredInCodebase).toBe(entry.runtimeStatus !== "missing");
    });
  });

  it("tracks the full unified formula inventory as registered runtime coverage", () => {
    expect(formulaInventorySummary.registeredInCodebase).toBe(formulaInventorySummary.total);
    expect(formulaInventorySummary.missingInCodebase).toBe(0);
    expect(formulaInventorySummary.placeholders).toBe(0);
  });

  it("keeps runtime and protocol reporting for key formulas", () => {
    const letEntry = formulaInventory.find((entry) => entry.name === "LET");
    const sumEntry = formulaInventory.find((entry) => entry.name === "SUM");
    const imageEntry = formulaInventory.find((entry) => entry.name === "IMAGE");

    expect(letEntry).toMatchObject({
      registeredInCodebase: true,
      protocolId: undefined,
      deterministic: "deterministic",
    });
    expect(sumEntry).toMatchObject({
      registeredInCodebase: true,
      protocolSupportsWasm: true,
      runtimeStatus: "implemented",
    });
    expect(imageEntry).toMatchObject({
      deterministic: "provider-backed",
      protocolId: undefined,
    });
  });
});

function readFormulaInventorySource(): FormulaInventorySource {
  const parsed = JSON.parse(
    readFileSync(new URL("../formula-inventory-source.json", import.meta.url), "utf8"),
  ) as unknown;
  if (!isFormulaInventorySource(parsed)) {
    throw new Error("Invalid formula inventory source fixture");
  }
  return parsed;
}

function isFormulaInventorySource(value: unknown): value is FormulaInventorySource {
  if (typeof value !== "object" || value === null) {
    return false;
  }
  const candidate = value as { version?: unknown; entries?: unknown };
  return (
    candidate.version === 1 &&
    Array.isArray(candidate.entries) &&
    candidate.entries.every(
      (entry) =>
        typeof entry === "object" &&
        entry !== null &&
        typeof Reflect.get(entry, "name") === "string" &&
        typeof Reflect.get(entry, "odfStatus") === "string" &&
        typeof Reflect.get(entry, "inOfficeList") === "boolean",
    )
  );
}
