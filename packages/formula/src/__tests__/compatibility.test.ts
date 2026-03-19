import { describe, expect, it } from "vitest";
import {
  compatibilityFamilies,
  compatibilityStatuses,
  getCompatibilityEntry,
  top50CompatibilityRegistry
} from "../compatibility.js";
import {
  excelFixtureFamilies,
  excelFixtureIdPattern,
  excelTop50StarterFixtures
} from "../../../excel-fixtures/src/index.js";

describe("formula compatibility registry", () => {
  it("keeps the starter registry aligned with the Excel fixtures pack", () => {
    expect(excelTop50StarterFixtures).toHaveLength(50);
    expect(top50CompatibilityRegistry).toHaveLength(50);

    const fixtureIds = new Set(excelTop50StarterFixtures.map((fixture) => fixture.id));
    const registryIds = new Set(top50CompatibilityRegistry.map((entry) => entry.id));

    expect([...registryIds].sort()).toEqual([...fixtureIds].sort());
  });

  it("uses unique, status-friendly fixture ids", () => {
    const ids = excelTop50StarterFixtures.map((fixture) => fixture.id);
    expect(new Set(ids).size).toBe(ids.length);
    ids.forEach((id) => expect(excelFixtureIdPattern.test(id)).toBe(true));
  });

  it("keeps family and formula metadata consistent across fixtures and registry", () => {
    const fixtureMap = new Map(excelTop50StarterFixtures.map((fixture) => [fixture.id, fixture]));

    top50CompatibilityRegistry.forEach((entry) => {
      const fixture = fixtureMap.get(entry.id);
      expect(fixture).toBeDefined();
      expect(entry.family).toBe(fixture?.family);
      expect(entry.formula).toBe(fixture?.formula);
      expect(compatibilityStatuses).toContain(entry.status);
    });
  });

  it("covers every declared family in the fixtures and registry layers", () => {
    const fixtureFamiliesSeen = new Set(excelTop50StarterFixtures.map((fixture) => fixture.family));
    const registryFamiliesSeen = new Set(top50CompatibilityRegistry.map((entry) => entry.family));

    expect([...fixtureFamiliesSeen].sort()).toEqual([...excelFixtureFamilies].sort());
    expect([...registryFamiliesSeen].sort()).toEqual([...compatibilityFamilies].sort());
  });

  it("exposes lookup by id", () => {
    const entry = getCompatibilityEntry("aggregation:sum-range");
    expect(entry).toMatchObject({
      id: "aggregation:sum-range",
      family: "aggregation",
      status: "implemented-js-and-wasm"
    });
  });

  it("includes all registry statuses in the starter set", () => {
    const statuses = new Set(top50CompatibilityRegistry.map((entry) => entry.status));
    expect([...statuses].sort()).toEqual(["implemented-js", "implemented-js-and-wasm", "unsupported"]);
    compatibilityStatuses.forEach((status) => {
      if (status !== "seeded") {
        expect(statuses.has(status)).toBe(true);
      }
    });
  });
});
