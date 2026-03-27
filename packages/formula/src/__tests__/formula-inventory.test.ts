import { describe, expect, it } from "vitest";
import { formulaInventory, formulaInventorySummary } from "../generated/formula-inventory.js";

describe("formula inventory", () => {
  it("tracks the canonical unified formula count", () => {
    expect(formulaInventorySummary.total).toBe(formulaInventory.length);
    expect(formulaInventory.length).toBeGreaterThan(500);
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
