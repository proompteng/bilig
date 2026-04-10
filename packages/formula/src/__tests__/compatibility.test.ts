import { readFileSync } from "node:fs";
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

interface FormulaDominanceSnapshotFixture {
  canonical: {
    nonProductionRows: Array<{ id: string }>;
    summary: { percent: number; production: number; total: number };
  };
  formulaBreadth: {
    officeListed: { percent: number; production: number; total: number };
    tracked: { percent: number; production: number; total: number };
  };
}

describe("formula compatibility registry", () => {
  it("keeps the generated formula dominance snapshot aligned with inventory and canonical status", () => {
    const snapshot = readFormulaDominanceSnapshot();
    const canonicalRegistryEntries = formulaCompatibilityRegistry.filter(
      (entry) => entry.scope === "canonical",
    );
    const canonicalProductionEntries = canonicalRegistryEntries.filter(
      (entry) => entry.status === "implemented-wasm-production",
    );
    const canonicalOpenRows = canonicalRegistryEntries
      .filter((entry) => entry.status !== "implemented-wasm-production")
      .map((entry) => entry.id)
      .toSorted();

    expect(snapshot.formulaBreadth.officeListed).toEqual({
      production: 487,
      total: 508,
      percent: 95.9,
    });
    expect(snapshot.formulaBreadth.tracked).toEqual({
      production: 487,
      total: 525,
      percent: 92.8,
    });
    expect(snapshot.canonical.summary).toEqual({
      production: canonicalProductionEntries.length,
      total: canonicalRegistryEntries.length,
      percent: 99.3,
    });
    expect(snapshot.canonical.nonProductionRows.map((row) => row.id).toSorted()).toEqual(
      canonicalOpenRows,
    );
    expect(canonicalOpenRows).toEqual([
      "dynamic-array:groupby-basic",
      "dynamic-array:pivotby-basic",
    ]);
  });

  it("keeps the canonical formula fixture corpus and registry in exact lockstep", () => {
    const canonicalRegistryEntries = formulaCompatibilityRegistry.filter(
      (entry) => entry.scope === "canonical",
    );

    expect(canonicalFormulaFixtures.length).toBeGreaterThan(0);
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
    expect(deriveWasmStatus("implemented-js")).toBe("not-started");
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

function readFormulaDominanceSnapshot(): FormulaDominanceSnapshotFixture {
  const parsed = JSON.parse(
    readFileSync(new URL("./fixtures/formula-dominance-snapshot.json", import.meta.url), "utf8"),
  ) as unknown;

  if (!isFormulaDominanceSnapshotFixture(parsed)) {
    throw new Error("Invalid formula dominance snapshot fixture shape");
  }

  return parsed;
}

function isFormulaDominanceSnapshotFixture(
  value: unknown,
): value is FormulaDominanceSnapshotFixture {
  if (!isRecord(value)) {
    return false;
  }

  return (
    isRatioRecord(value["formulaBreadth"]) &&
    isRatioRecord(value["canonical"]) &&
    Array.isArray(value["canonical"]["nonProductionRows"]) &&
    value["canonical"]["nonProductionRows"].every(
      (row) => isRecord(row) && typeof row["id"] === "string",
    )
  );
}

function isRatioRecord(value: unknown): value is {
  officeListed?: { percent: number; production: number; total: number };
  tracked?: { percent: number; production: number; total: number };
  summary?: { percent: number; production: number; total: number };
  nonProductionRows?: unknown[];
} {
  if (!isRecord(value)) {
    return false;
  }

  const summary = value["summary"];
  const officeListed = value["officeListed"];
  const tracked = value["tracked"];
  return (
    (summary === undefined || isRatio(summary)) &&
    (officeListed === undefined || isRatio(officeListed)) &&
    (tracked === undefined || isRatio(tracked))
  );
}

function isRatio(value: unknown): value is { percent: number; production: number; total: number } {
  return (
    isRecord(value) &&
    typeof value["percent"] === "number" &&
    typeof value["production"] === "number" &&
    typeof value["total"] === "number"
  );
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}
