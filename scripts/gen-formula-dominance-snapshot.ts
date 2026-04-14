#!/usr/bin/env bun

import { existsSync, mkdtempSync, mkdirSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join, resolve } from "node:path";
import {
  compatibilityFamilies,
  formulaCompatibilityRegistry,
  type CompatibilityFamily,
  type CompatibilityStatus,
  type FormulaCompatibilityEntry,
} from "../packages/formula/src/compatibility.ts";
import {
  formulaInventory,
  formulaInventorySummary,
} from "../packages/formula/src/generated/formula-inventory.ts";

interface FormulaDominanceSnapshot {
  schemaVersion: 1;
  formulaBreadth: {
    officeListed: FormulaDominanceRatio;
    tracked: FormulaDominanceRatio;
    missingOfficeFunctions: string[];
  };
  canonical: {
    nonProductionRows: FormulaDominanceRow[];
    statusCounts: Record<CompatibilityStatus, number>;
    summary: FormulaDominanceRatio;
  };
  families: FormulaDominanceFamily[];
  strategicFamilies: FormulaDominanceFamily[];
}

interface FormulaDominanceRatio {
  percent: number;
  production: number;
  total: number;
}

interface FormulaDominanceRow {
  family: CompatibilityFamily;
  formula: string;
  id: string;
  notes?: string;
  status: CompatibilityStatus;
  wasmStatus: FormulaCompatibilityEntry["wasmStatus"];
}

interface FormulaDominanceFamily {
  family: CompatibilityFamily;
  nonProductionRows: FormulaDominanceRow[];
  statusCounts: Record<CompatibilityStatus, number>;
  summary: FormulaDominanceRatio;
}

const rootDir = resolve(new URL("..", import.meta.url).pathname);
const outputPath = join(
  rootDir,
  "packages",
  "formula",
  "src",
  "__tests__",
  "fixtures",
  "formula-dominance-snapshot.json",
);
const isCheckMode = process.argv.includes("--check");

const snapshot = buildSnapshot();
const serializedSnapshot = formatJsonForRepo(`${JSON.stringify(snapshot, null, 2)}\n`);

if (isCheckMode) {
  if (!existsSync(outputPath)) {
    throw new Error(
      `Missing generated formula dominance snapshot at ${outputPath}. Run: bun scripts/gen-formula-dominance-snapshot.ts`,
    );
  }

  const currentSnapshot = readFileSync(outputPath, "utf8");
  if (currentSnapshot !== serializedSnapshot) {
    throw new Error(
      `Generated formula dominance snapshot is out of date. Run: bun scripts/gen-formula-dominance-snapshot.ts`,
    );
  }
} else {
  mkdirSync(dirname(outputPath), { recursive: true });
  writeFileSync(outputPath, serializedSnapshot);
}

console.log(
  JSON.stringify(
    {
      outputPath,
      mode: isCheckMode ? "check" : "write",
      canonicalRows: snapshot.canonical.summary.total,
      canonicalProductionRows: snapshot.canonical.summary.production,
      openCanonicalRows: snapshot.canonical.nonProductionRows.map((row) => row.id),
      officeBreadth: snapshot.formulaBreadth.officeListed,
      trackedBreadth: snapshot.formulaBreadth.tracked,
    },
    null,
    2,
  ),
);

function buildSnapshot(): FormulaDominanceSnapshot {
  const officeListed = formulaInventory.filter((entry) => entry.inOfficeList);
  const canonicalRows = formulaCompatibilityRegistry.filter((entry) => entry.scope === "canonical");
  const canonicalProduction = canonicalRows.filter(
    (entry) => entry.status === "implemented-wasm-production",
  );
  const canonicalNonProduction = canonicalRows
    .filter((entry) => entry.status !== "implemented-wasm-production")
    .map(toDominanceRow);

  return {
    schemaVersion: 1,
    formulaBreadth: {
      officeListed: ratio(
        formulaInventorySummary.officeListedRegisteredInCodebase,
        officeListed.length,
      ),
      tracked: ratio(formulaInventorySummary.registeredInCodebase, formulaInventorySummary.total),
      missingOfficeFunctions: officeListed
        .filter((entry) => !entry.registeredInCodebase)
        .map((entry) => entry.name),
    },
    canonical: {
      summary: ratio(canonicalProduction.length, canonicalRows.length),
      statusCounts: countStatuses(canonicalRows),
      nonProductionRows: canonicalNonProduction,
    },
    families: compatibilityFamilies.map((family) => buildFamilySummary(family, canonicalRows)),
    strategicFamilies: [
      buildFamilySummary("dynamic-array", canonicalRows),
      buildFamilySummary("names", canonicalRows),
      buildFamilySummary("tables", canonicalRows),
      buildFamilySummary("structured-reference", canonicalRows),
      buildFamilySummary("lambda", canonicalRows),
    ],
  };
}

function buildFamilySummary(
  family: CompatibilityFamily,
  canonicalRows: readonly FormulaCompatibilityEntry[],
): FormulaDominanceFamily {
  const familyRows = canonicalRows.filter((entry) => entry.family === family);
  const familyProduction = familyRows.filter(
    (entry) => entry.status === "implemented-wasm-production",
  );

  return {
    family,
    summary: ratio(familyProduction.length, familyRows.length),
    statusCounts: countStatuses(familyRows),
    nonProductionRows: familyRows
      .filter((entry) => entry.status !== "implemented-wasm-production")
      .map(toDominanceRow),
  };
}

function countStatuses(
  entries: readonly FormulaCompatibilityEntry[],
): Record<CompatibilityStatus, number> {
  const counts: Record<CompatibilityStatus, number> = {
    unsupported: 0,
    seeded: 0,
    "implemented-js": 0,
    "implemented-js-and-wasm-shadow": 0,
    "implemented-wasm-production": 0,
    blocked: 0,
  };

  for (const entry of entries) {
    counts[entry.status] += 1;
  }

  return counts;
}

function ratio(production: number, total: number): FormulaDominanceRatio {
  return {
    production,
    total,
    percent: total === 0 ? 0 : Number(((production / total) * 100).toFixed(1)),
  };
}

function toDominanceRow(entry: FormulaCompatibilityEntry): FormulaDominanceRow {
  return {
    id: entry.id,
    family: entry.family,
    formula: entry.formula,
    status: entry.status,
    wasmStatus: entry.wasmStatus,
    notes: entry.notes,
  };
}

function formatJsonForRepo(serializedJson: string): string {
  const tempDir = mkdtempSync(join(tmpdir(), "formula-dominance-"));
  const tempFilePath = join(tempDir, "snapshot.json");
  writeFileSync(tempFilePath, serializedJson);
  const oxfmtPath = join(rootDir, "node_modules", ".bin", "oxfmt");

  const formatResult = Bun.spawnSync([oxfmtPath, "--write", tempFilePath], {
    stdin: "ignore",
    stdout: "pipe",
    stderr: "pipe",
  });
  if (formatResult.exitCode !== 0) {
    rmSync(tempDir, { recursive: true, force: true });
    throw new Error(
      `Unable to format generated snapshot: ${new TextDecoder().decode(formatResult.stderr).trim()}`,
    );
  }

  const formattedJson = readFileSync(tempFilePath, "utf8");
  rmSync(tempDir, { recursive: true, force: true });
  return formattedJson;
}
