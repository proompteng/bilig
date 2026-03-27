import { describe, expect, it } from "vitest";
import {
  compatibilityFamilies,
  compatibilityScopes,
  compatibilityStatuses,
  deriveWasmStatus,
  formulaCompatibilityRegistry,
  getCompatibilityEntry,
  isCompatibilityStatus,
  isWasmCompatibilityStatus,
  wasmCompatibilityStatuses,
} from "../compatibility.js";
import {
  canonicalFormulaFixtures,
  canonicalFormulaSmokeSuite,
  canonicalWorkbookSemanticsFixtures,
  excelFixtureFamilies,
  excelFixtureIdPattern,
} from "../../../excel-fixtures/src/index.js";

describe("formula compatibility registry", () => {
  it("keeps the canonical formula fixture corpus and registry in exact lockstep", () => {
    const canonicalRegistryEntries = formulaCompatibilityRegistry.filter(
      (entry) => entry.scope === "canonical",
    );

    expect(canonicalFormulaFixtures).toHaveLength(102);
    expect(canonicalRegistryEntries).toHaveLength(canonicalFormulaFixtures.length);

    const fixtureIds = new Set(canonicalFormulaFixtures.map((fixture) => fixture.id));
    const registryIds = new Set(canonicalRegistryEntries.map((entry) => entry.id));

    expect([...fixtureIds].toSorted()).toEqual([...registryIds].toSorted());
  });

  it("tracks workbook semantics fixtures as extended coverage without redefining the canonical corpus", () => {
    const workbookSemanticsIds = new Set(
      canonicalWorkbookSemanticsFixtures.map((fixture) => fixture.id),
    );
    const extendedRegistryEntries = formulaCompatibilityRegistry.filter(
      (entry) => entry.scope === "extended",
    );

    expect(extendedRegistryEntries).toHaveLength(canonicalWorkbookSemanticsFixtures.length);
    expect(new Set(extendedRegistryEntries.map((entry) => entry.id))).toEqual(workbookSemanticsIds);
  });

  it("uses unique, status-friendly fixture ids in the canonical formula corpus", () => {
    const ids = canonicalFormulaFixtures.map((fixture) => fixture.id);
    expect(new Set(ids).size).toBe(ids.length);
    ids.forEach((id) => expect(excelFixtureIdPattern.test(id)).toBe(true));
  });

  it("keeps family, formula, and metadata aligned across fixtures and registry", () => {
    const trackedFixtures = [...canonicalFormulaFixtures, ...canonicalWorkbookSemanticsFixtures];
    const fixtureMap = new Map(trackedFixtures.map((fixture) => [fixture.id, fixture]));

    formulaCompatibilityRegistry.forEach((entry) => {
      const fixture = fixtureMap.get(entry.id);
      expect(fixture).toBeDefined();
      expect(entry.family).toBe(fixture?.family);
      expect(entry.formula).toBe(fixture?.formula);
      expect(compatibilityStatuses).toContain(entry.status);
      expect(compatibilityScopes).toContain(entry.scope);
      expect(entry.fixtureIds).toEqual([entry.id]);
      expect(entry.owner.length).toBeGreaterThan(0);
      expect(entry.prerequisites.length).toBeGreaterThan(0);
      expect(isWasmCompatibilityStatus(entry.wasmStatus)).toBe(true);
    });
  });

  it("covers every declared family in both the fixture and registry layers", () => {
    const fixtureFamiliesSeen = new Set(canonicalFormulaFixtures.map((fixture) => fixture.family));
    const registryFamiliesSeen = new Set(formulaCompatibilityRegistry.map((entry) => entry.family));

    expect([...fixtureFamiliesSeen].toSorted()).toEqual([...excelFixtureFamilies].toSorted());
    expect([...registryFamiliesSeen].toSorted()).toEqual([...compatibilityFamilies].toSorted());
  });

  it("validates compatibility status helper values", () => {
    expect(isCompatibilityStatus("implemented-wasm-production")).toBe(true);
    expect(isCompatibilityStatus("implemented-js-and-wasm-shadow")).toBe(true);
    expect(isCompatibilityStatus("nope")).toBe(false);
  });

  it("derives wasm compatibility status for shadowed and blocked entries", () => {
    expect(deriveWasmStatus("implemented-js-and-wasm-shadow")).toBe("shadow");
    expect(deriveWasmStatus("blocked")).toBe("blocked");
  });

  it("exposes lookup by id", () => {
    const entry = getCompatibilityEntry("aggregation:sum-range");
    expect(entry).toMatchObject({
      id: "aggregation:sum-range",
      family: "aggregation",
      status: "implemented-wasm-production",
      wasmStatus: "production",
    });
  });

  it("retains a canonical smoke suite derived from the canonical export", () => {
    expect(canonicalFormulaSmokeSuite.id).toBe("canonical-smoke");
    expect(canonicalFormulaSmokeSuite.cases).toHaveLength(5);
    expect(canonicalFormulaSmokeSuite.cases).toEqual(canonicalFormulaFixtures.slice(0, 5));
  });

  it("uses only the expanded compatibility status and wasm status contracts", () => {
    const statuses = new Set(formulaCompatibilityRegistry.map((entry) => entry.status));
    const wasmStatuses = new Set(formulaCompatibilityRegistry.map((entry) => entry.wasmStatus));

    expect([...statuses].toSorted()).toEqual(
      compatibilityStatuses.filter((status) => statuses.has(status)).toSorted(),
    );
    expect([...wasmStatuses].toSorted()).toEqual(
      wasmCompatibilityStatuses.filter((status) => wasmStatuses.has(status)).toSorted(),
    );
  });
});
