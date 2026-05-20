# Semantic Correctness Program Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Make Bilig's workbook runtime semantically correct by defining one reusable semantic contract, expanding oracle-backed coverage, and gating formula, recalc, import/export, sync, and assistant-visible readback against that contract.

**Architecture:** Keep `@bilig/formula` as the parser/evaluator semantic source, `@bilig/core` as workbook runtime owner, and `@bilig/excel-import` as package fidelity owner. Add a shared semantic projection layer under `packages/core/src/semantics/` and route tests, scorecards, corpus checks, and assistant verification through it instead of comparing incidental snapshot shape.

**Tech Stack:** TypeScript, ESM, pnpm monorepo, Bun scripts, Vitest, fast-check, Microsoft Excel/Google Sheets live scorecard scripts, `@bilig/core`, `@bilig/formula`, `@bilig/excel-import`, `@bilig/headless`, `apps/bilig`, `apps/web`.

---

## Audit Baseline

Run these from:

```bash
cd /Users/gregkonush/.codex/worktrees/0c3e/bilig3
```

Current audit evidence from this checkout:

- `git status --short --branch` reports `## HEAD (no branch)`. Execution should move onto `main` before committing, because repo guidance says to stay on `main`.
- `pnpm install --frozen-lockfile` completed. It warned about unbuilt workspace package bins before `dist` existed.
- `pnpm build` passed.
- `pnpm source-size:check` passed, with max non-test source file `packages/core/src/engine/services/operation-fresh-direct-aggregate-formula-batch-fast-path.ts` at 996 lines.
- `pnpm formula-inventory:check` passed.
- `pnpm calculation:semantics:check` passed with 301/301 canonical fixtures and 11/11 workbook-semantics fixtures covered.
- `pnpm import-export:fidelity:check` passed with 42 covered features, 0 unsupported features, and 1 declined runtime feature: `xlsx.macros.execution`.
- `pnpm workpaper:parity:check` passed against HyperFormula 3.2.0 surface metadata.
- `pnpm dominance:check` passed, but still reports `goalStatus: active-not-achieved` and `blanketTenXClaimAllowed: false`.
- `pnpm test:correctness:formula` was started after build but stayed silent for multiple minutes and was terminated with exit 143. Treat that as an execution follow-up, not as a semantic failure.

Key code findings:

- Semantic comparison logic exists in at least three places:
  - `packages/core/src/__tests__/engine-fuzz-helpers.ts`
  - `scripts/import-export-fidelity-projection.ts`
  - `packages/excel-import/src/__tests__/excel-import.test.ts`
- The codebase already has strong mechanics: replay fixtures, fuzz tests, differential JS/WASM recalc, import/export scorecards, live Excel/Sheets scorecards, and workbook-agent mutation receipts.
- The main semantic gap is not absence of tests. It is that "semantic meaning" is not a first-class shared API, so different gates can pass while comparing different projections of the same workbook.
- Workbook-semantics fixture count is only 11. That is too narrow for a correctness claim that covers names, structured references, dynamic arrays/spills, volatile context, calculation settings, structural rewrites, import/export, sync projection, and assistant-rendered verification.

## File Structure

Create or modify these files:

- Create: `packages/core/src/semantics/workbook-semantic-projection.ts`
  - Owns the shared semantic projection for workbook snapshots.
- Create: `packages/core/src/semantics/workbook-semantic-comparison.ts`
  - Owns equality, mismatch reporting, and stable serialization helpers.
- Create: `packages/core/src/semantics/index.ts`
  - Exports the public semantic helpers from `@bilig/core`.
- Create: `packages/core/src/__tests__/workbook-semantic-projection.test.ts`
  - Locks projection behavior and examples.
- Modify: `packages/core/src/index.ts`
  - Re-export semantic helpers.
- Modify: `packages/core/src/__tests__/engine-fuzz-helpers.ts`
  - Replace local normalization with shared semantic projection.
- Modify: `scripts/import-export-fidelity-projection.ts`
  - Replace script-local projection with shared semantic projection or a thin import/export-specific view built from it.
- Modify: `packages/excel-import/src/__tests__/excel-import.test.ts`
  - Remove local duplicated projection helper and import the shared helper.
- Modify: `packages/excel-fixtures/src/workbook-semantics-fixtures.ts`
  - Add high-risk workbook-semantics fixtures.
- Modify: `packages/formula/src/compatibility.ts`
  - Register new workbook-semantics fixture IDs outside the canonical formula corpus.
- Modify: `scripts/gen-calculation-semantics-scorecard.ts`
  - Report workbook-semantics categories, not just counts.
- Modify: `scripts/__tests__/calculation-semantics-scorecard.test.ts`
  - Lock the new category coverage.
- Modify: `packages/core/src/__tests__/formula-runtime-correctness.test.ts`
  - Run all production-routed canonical fixtures that the engine can execute through `recalculateDifferential()`.
- Create: `packages/core/src/__tests__/engine-semantic-invariants.test.ts`
  - Adds deterministic semantic invariant checks across edit, rebuild, snapshot, and full recalc paths.
- Modify: `scripts/gen-import-export-fidelity-scorecard.ts`
  - Include semantic-loss accounting for every unsupported or intentionally declined runtime feature.
- Create: `scripts/import-export-semantic-loss-ledger.ts`
  - Produces a stable ledger of preserved, unsupported, external, and declined semantics.
- Modify: `apps/bilig/src/codex-app/workbook-agent-mutation-proof.ts`
  - Use semantic comparison helpers for authoritative proof when comparing values, formulas, formats, and metadata.
- Modify: `apps/bilig/src/codex-app/workbook-agent-mutation-receipt.ts`
  - Add a semantic proof summary to mutation receipts.
- Modify: `apps/bilig/src/codex-app/__tests__/workbook-agent-mutation-receipt.test.ts`
  - Assert that applied receipts include semantic proof and never call stale rendered readback success.
- Modify: `apps/web/src/__tests__/worker-runtime-authoritative-bootstrap.test.ts`
  - Add a browser-side authoritative/rendered semantic consistency case.
- Modify: `package.json`
  - Add `test:semantic:fast`, `test:semantic:medium`, and `test:semantic:deep`.
- Modify: `scripts/run-ci.ts`
  - Wire semantic gates into the existing fast CI profile without live Excel/Sheets dependencies.

## Task 1: Add Shared Workbook Semantic Projection

**Files:**

- Create: `packages/core/src/semantics/workbook-semantic-projection.ts`
- Create: `packages/core/src/semantics/workbook-semantic-comparison.ts`
- Create: `packages/core/src/semantics/index.ts`
- Create: `packages/core/src/__tests__/workbook-semantic-projection.test.ts`
- Modify: `packages/core/src/index.ts`

- [ ] **Step 1: Write the projection test first**

Create `packages/core/src/__tests__/workbook-semantic-projection.test.ts`:

```ts
import { describe, expect, it } from "vitest";
import type { WorkbookSnapshot } from "@bilig/protocol";
import {
  projectWorkbookSemanticSnapshot,
  workbookSemanticSnapshotsEqual,
} from "../semantics/index.js";

describe("workbook semantic projection", () => {
  it("ignores incidental ids while preserving workbook meaning", () => {
    const left: WorkbookSnapshot = {
      version: 1,
      workbook: {
        name: "semantic-left",
        metadata: {
          definedNames: [{ name: "TaxRate", value: { kind: "scalar", value: 0.085 } }],
          styles: [{ id: "style-b", fill: { backgroundColor: "#fff2cc" } }],
        },
      },
      sheets: [
        {
          id: 100,
          name: "Sheet1",
          order: 0,
          metadata: {
            styleRanges: [
              {
                range: { sheetName: "Sheet1", startAddress: "B2", endAddress: "B2" },
                styleId: "style-b",
              },
            ],
          },
          cells: [{ address: "B2", value: 10, formula: "A1*2", format: "0.00" }],
        },
      ],
    };
    const right: WorkbookSnapshot = {
      version: 1,
      workbook: {
        name: "semantic-right",
        metadata: {
          styles: [{ id: "style-a", fill: { backgroundColor: "#fff2cc" } }],
          definedNames: [{ name: "TaxRate", value: { kind: "scalar", value: 0.085 } }],
        },
      },
      sheets: [
        {
          id: 200,
          name: "Sheet1",
          order: 0,
          metadata: {
            styleRanges: [
              {
                range: { sheetName: "Sheet1", startAddress: "B2", endAddress: "B2" },
                styleId: "style-a",
              },
            ],
          },
          cells: [{ address: "B2", formula: "A1*2", value: 10, format: "0.00" }],
        },
      ],
    };

    expect(projectWorkbookSemanticSnapshot(left)).toEqual(projectWorkbookSemanticSnapshot(right));
    expect(workbookSemanticSnapshotsEqual(left, right)).toBe(true);
  });

  it("keeps formula, value, format, metadata, and sheet-order differences visible", () => {
    const base: WorkbookSnapshot = {
      version: 1,
      workbook: { name: "base" },
      sheets: [{ name: "Sheet1", order: 0, cells: [{ address: "A1", value: 1 }] }],
    };
    const changed: WorkbookSnapshot = {
      version: 1,
      workbook: { name: "base" },
      sheets: [{ name: "Sheet1", order: 0, cells: [{ address: "A1", value: 2 }] }],
    };

    expect(workbookSemanticSnapshotsEqual(base, changed)).toBe(false);
  });
});
```

- [ ] **Step 2: Run the failing test**

Run:

```bash
pnpm exec vitest run packages/core/src/__tests__/workbook-semantic-projection.test.ts
```

Expected: fail because `../semantics/index.js` does not exist.

- [ ] **Step 3: Create the semantic projection module**

Create `packages/core/src/semantics/workbook-semantic-projection.ts`:

```ts
import { parseCellAddress } from "@bilig/formula";
import type {
  CellRangeRef,
  SheetMetadataSnapshot,
  WorkbookSnapshot,
} from "@bilig/protocol";

export interface WorkbookSemanticCell {
  readonly address: string;
  readonly value?: unknown;
  readonly formula?: string;
  readonly format?: string;
  readonly style?: unknown;
}

export interface WorkbookSemanticSheet {
  readonly name: string;
  readonly order: number;
  readonly cells: readonly WorkbookSemanticCell[];
  readonly metadata: {
    readonly rows: readonly unknown[];
    readonly columns: readonly unknown[];
    readonly styleRanges: readonly unknown[];
    readonly formatRanges: readonly unknown[];
    readonly merges: readonly CellRangeRef[];
    readonly filters: readonly CellRangeRef[];
    readonly sorts: readonly unknown[];
    readonly validations: readonly unknown[];
    readonly conditionalFormats: readonly unknown[];
    readonly freezePane: unknown | null;
    readonly sheetProtection: unknown | null;
  };
}

export interface WorkbookSemanticSnapshot {
  readonly version: 1;
  readonly workbook: {
    readonly calculationSettings: unknown | null;
    readonly volatileContext: unknown | null;
    readonly definedNames: readonly unknown[];
    readonly tables: readonly unknown[];
    readonly spills: readonly unknown[];
    readonly pivots: readonly unknown[];
    readonly charts: readonly unknown[];
    readonly externalWorkbookReferences: readonly unknown[];
    readonly unsupportedFormulaDependencies: readonly unknown[];
  };
  readonly sheets: readonly WorkbookSemanticSheet[];
}

function stableJson(value: unknown): string {
  return JSON.stringify(value, (_key, entry) => {
    if (entry && typeof entry === "object" && !Array.isArray(entry)) {
      return Object.fromEntries(Object.entries(entry).sort(([left], [right]) => left.localeCompare(right)));
    }
    return entry;
  });
}

function byStableJson(left: unknown, right: unknown): number {
  return stableJson(left).localeCompare(stableJson(right));
}

function compareRange(left: CellRangeRef, right: CellRangeRef): number {
  const leftStart = parseCellAddress(left.startAddress, left.sheetName);
  const rightStart = parseCellAddress(right.startAddress, right.sheetName);
  const leftEnd = parseCellAddress(left.endAddress, left.sheetName);
  const rightEnd = parseCellAddress(right.endAddress, right.sheetName);
  return (
    left.sheetName.localeCompare(right.sheetName) ||
    leftStart.row - rightStart.row ||
    leftStart.col - rightStart.col ||
    leftEnd.row - rightEnd.row ||
    leftEnd.col - rightEnd.col
  );
}

function sorted<T>(values: readonly T[] | undefined, compare: (left: T, right: T) => number = byStableJson): readonly T[] {
  return [...(values ?? [])].sort(compare);
}

function projectSheetMetadata(
  metadata: SheetMetadataSnapshot | undefined,
  styles: ReadonlyMap<string, unknown>,
): WorkbookSemanticSheet["metadata"] {
  return {
    rows: sorted((metadata?.rows ?? []).map(({ id: _id, ...row }) => row)),
    columns: sorted((metadata?.columns ?? []).map(({ id: _id, ...column }) => column)),
    styleRanges: sorted(
      (metadata?.styleRanges ?? []).map((record) => ({
        range: record.range,
        style: styles.get(record.styleId) ?? record.styleId,
      })),
    ),
    formatRanges: sorted(metadata?.formatRanges),
    merges: sorted(metadata?.merges, compareRange),
    filters: sorted(metadata?.filters, compareRange),
    sorts: sorted(metadata?.sorts),
    validations: sorted(metadata?.validations),
    conditionalFormats: sorted(metadata?.conditionalFormats),
    freezePane: metadata?.freezePane ?? null,
    sheetProtection: metadata?.sheetProtection ?? null,
  };
}

function styleById(snapshot: WorkbookSnapshot): ReadonlyMap<string, unknown> {
  return new Map((snapshot.workbook.metadata?.styles ?? []).map((style) => [style.id, style]));
}

export function projectWorkbookSemanticSnapshot(snapshot: WorkbookSnapshot): WorkbookSemanticSnapshot {
  const styles = styleById(snapshot);
  return {
    version: 1,
    workbook: {
      calculationSettings: snapshot.workbook.metadata?.calculationSettings ?? null,
      volatileContext: snapshot.workbook.metadata?.volatileContext ?? null,
      definedNames: sorted(snapshot.workbook.metadata?.definedNames),
      tables: sorted(snapshot.workbook.metadata?.tables),
      spills: sorted(snapshot.workbook.metadata?.spills),
      pivots: sorted(snapshot.workbook.metadata?.pivots),
      charts: sorted(snapshot.workbook.metadata?.charts),
      externalWorkbookReferences: sorted(snapshot.workbook.metadata?.externalWorkbookReferences),
      unsupportedFormulaDependencies: sorted(snapshot.workbook.metadata?.unsupportedFormulaDependencies),
    },
    sheets: snapshot.sheets
      .toSorted((left, right) => left.order - right.order || left.name.localeCompare(right.name))
      .map((sheet) => ({
        name: sheet.name,
        order: sheet.order,
        cells: sheet.cells
          .filter((cell) => cell.value !== null || cell.formula !== undefined || cell.format !== undefined)
          .map((cell) => ({
            address: cell.address,
            ...(cell.value !== undefined ? { value: cell.value } : {}),
            ...(cell.formula !== undefined ? { formula: cell.formula } : {}),
            ...(cell.format !== undefined ? { format: cell.format } : {}),
          }))
          .toSorted((left, right) => left.address.localeCompare(right.address)),
        metadata: projectSheetMetadata(sheet.metadata, styles),
      })),
  };
}
```

- [ ] **Step 4: Create comparison helpers**

Create `packages/core/src/semantics/workbook-semantic-comparison.ts`:

```ts
import type { WorkbookSnapshot } from "@bilig/protocol";
import { projectWorkbookSemanticSnapshot, type WorkbookSemanticSnapshot } from "./workbook-semantic-projection.js";

export interface WorkbookSemanticMismatch {
  readonly path: string;
  readonly left: unknown;
  readonly right: unknown;
}

function stableJson(value: unknown): string {
  return JSON.stringify(value, null, 2);
}

export function serializeWorkbookSemanticSnapshot(snapshot: WorkbookSemanticSnapshot): string {
  return `${stableJson(snapshot)}\n`;
}

export function workbookSemanticSnapshotsEqual(left: WorkbookSnapshot, right: WorkbookSnapshot): boolean {
  return serializeWorkbookSemanticSnapshot(projectWorkbookSemanticSnapshot(left)) === serializeWorkbookSemanticSnapshot(projectWorkbookSemanticSnapshot(right));
}

export function diffWorkbookSemanticSnapshots(left: WorkbookSnapshot, right: WorkbookSnapshot): readonly WorkbookSemanticMismatch[] {
  const leftProjected = projectWorkbookSemanticSnapshot(left);
  const rightProjected = projectWorkbookSemanticSnapshot(right);
  if (serializeWorkbookSemanticSnapshot(leftProjected) === serializeWorkbookSemanticSnapshot(rightProjected)) {
    return [];
  }
  return [{ path: "$", left: leftProjected, right: rightProjected }];
}
```

- [ ] **Step 5: Export helpers**

Create `packages/core/src/semantics/index.ts`:

```ts
export {
  projectWorkbookSemanticSnapshot,
  type WorkbookSemanticCell,
  type WorkbookSemanticSheet,
  type WorkbookSemanticSnapshot,
} from "./workbook-semantic-projection.js";
export {
  diffWorkbookSemanticSnapshots,
  serializeWorkbookSemanticSnapshot,
  workbookSemanticSnapshotsEqual,
  type WorkbookSemanticMismatch,
} from "./workbook-semantic-comparison.js";
```

Modify `packages/core/src/index.ts`:

```ts
export * from "./semantics/index.js";
```

- [ ] **Step 6: Run the test**

Run:

```bash
pnpm exec vitest run packages/core/src/__tests__/workbook-semantic-projection.test.ts
```

Expected: pass.

- [ ] **Step 7: Commit**

```bash
git add packages/core/src/semantics packages/core/src/__tests__/workbook-semantic-projection.test.ts packages/core/src/index.ts
git commit -m "feat(core): add workbook semantic projection"
```

## Task 2: Replace Duplicated Semantic Normalizers

**Files:**

- Modify: `packages/core/src/__tests__/engine-fuzz-helpers.ts`
- Modify: `scripts/import-export-fidelity-projection.ts`
- Modify: `packages/excel-import/src/__tests__/excel-import.test.ts`
- Test: existing fuzz, import/export, and scorecard tests

- [ ] **Step 1: Replace core test normalization**

In `packages/core/src/__tests__/engine-fuzz-helpers.ts`, change `normalizeSnapshotForSemanticComparison` to delegate to the shared projection:

```ts
import { projectWorkbookSemanticSnapshot } from "../semantics/index.js";

export function normalizeSnapshotForSemanticComparison(snapshot: WorkbookSnapshot): ReturnType<typeof projectWorkbookSemanticSnapshot> {
  return projectWorkbookSemanticSnapshot(snapshot);
}
```

- [ ] **Step 2: Run core semantic consumers**

Run:

```bash
pnpm exec vitest run packages/core/src/__tests__/engine-snapshot.fuzz.test.ts packages/core/src/__tests__/engine-history.fuzz.test.ts packages/core/src/__tests__/engine-replay-fixtures.test.ts
```

Expected: pass.

- [ ] **Step 3: Replace import/export projection**

In `scripts/import-export-fidelity-projection.ts`, preserve the exported function name but route through shared semantics:

```ts
import { projectWorkbookSemanticSnapshot } from "../packages/core/src/semantics/index.js";
import type { WorkbookSnapshot } from "../packages/protocol/src/types.js";

export function projectSupportedSnapshotSemantics(snapshot: WorkbookSnapshot) {
  return projectWorkbookSemanticSnapshot(snapshot);
}
```

- [ ] **Step 4: Remove the local copy in excel-import tests**

In `packages/excel-import/src/__tests__/excel-import.test.ts`, replace the local `projectSupportedSnapshotSemantics` helper with:

```ts
import { projectWorkbookSemanticSnapshot as projectSupportedSnapshotSemantics } from "../../../core/src/semantics/index.js";
```

- [ ] **Step 5: Run import/export gates**

Run:

```bash
pnpm exec vitest run packages/excel-import/src/__tests__/excel-import.test.ts packages/excel-import/src/__tests__/xlsx-roundtrip-semantics.test.ts
pnpm import-export:fidelity:check
```

Expected: both commands pass and `packages/benchmarks/baselines/import-export-fidelity-scorecard.json` stays unchanged unless the new projection exposes a real semantic gap.

- [ ] **Step 6: Commit**

```bash
git add packages/core/src/__tests__/engine-fuzz-helpers.ts scripts/import-export-fidelity-projection.ts packages/excel-import/src/__tests__/excel-import.test.ts
git commit -m "refactor(core): use shared workbook semantic comparison"
```

## Task 3: Expand Workbook-Semantics Fixture Coverage

**Files:**

- Modify: `packages/excel-fixtures/src/workbook-semantics-fixtures.ts`
- Modify: `packages/formula/src/compatibility.ts`
- Modify: `scripts/gen-calculation-semantics-scorecard.ts`
- Modify: `scripts/__tests__/calculation-semantics-scorecard.test.ts`

- [ ] **Step 1: Add fixture categories to the scorecard model**

In `scripts/gen-calculation-semantics-scorecard.ts`, add this summary field:

```ts
readonly workbookSemanticsFamilies: string[]
```

Populate it with:

```ts
const workbookSemanticsFamilies = [...new Set(canonicalWorkbookSemanticsFixtures.map((fixture) => fixture.family))].toSorted();
```

Add it to `summary`:

```ts
workbookSemanticsFamilies,
```

Parse it with:

```ts
workbookSemanticsFamilies: stringArrayField(summary, "workbookSemanticsFamilies"),
```

- [ ] **Step 2: Let the local workbook-semantics helper attach tables**

In `packages/excel-fixtures/src/workbook-semantics-fixtures.ts`, widen the local `fixture()` helper options:

```ts
function fixture(
  family: ExcelFixtureFamily,
  slug: string,
  title: string,
  formula: string,
  inputs: ExcelFixtureInputCell[],
  outputs: ExcelFixtureExpectedOutput[],
  options: {
    notes?: string;
    definedNames?: ExcelFixtureDefinedName[];
    tables?: ExcelFixtureCase["tables"];
  } = {},
): ExcelFixtureCase {
  const base: ExcelFixtureCase = {
    id: createExcelFixtureId(family, slug),
    family,
    title,
    formula,
    inputs,
    outputs,
    sheetName: "Sheet1",
  };
  if (options.notes !== undefined) {
    base.notes = options.notes;
  }
  if (options.definedNames !== undefined) {
    base.definedNames = options.definedNames;
  }
  if (options.tables !== undefined) {
    base.tables = options.tables;
  }
  return base;
}
```

- [ ] **Step 3: Add high-risk workbook semantic fixtures**

Append fixtures in `packages/excel-fixtures/src/workbook-semantics-fixtures.ts` for these exact behaviors:

```ts
fixture(
  "structured-reference",
  "table-column-sum",
  "Structured references resolve table columns",
  "=SUM(Sales[Amount])",
  [input("A1", "Amount"), input("A2", 10), input("A3", 15)],
  [output("B1", numberExpected(25))],
  {
    notes: "The engine must bind table column names semantically, not by incidental range text.",
    tables: [
      {
        name: "Sales",
        startAddress: "A1",
        endAddress: "A3",
        columnNames: ["Amount"],
        headerRow: true,
        totalsRow: false,
      },
    ],
  },
),
fixture(
  "structured-reference",
  "table-missing-column-ref-error",
  "Missing structured-reference columns surface #REF!",
  "=SUM(Sales[Missing])",
  [input("A1", "Amount"), input("A2", 10), input("A3", 15)],
  [output("B1", errorExpected(ErrorCode.Ref, "#REF!"))],
  {
    notes: "Structured-reference misses must be explicit semantic errors.",
    tables: [
      {
        name: "Sales",
        startAddress: "A1",
        endAddress: "A3",
        columnNames: ["Amount"],
        headerRow: true,
        totalsRow: false,
      },
    ],
  },
),
fixture(
  "structured-reference",
  "table-totals-row-excluded",
  "Structured references exclude totals rows",
  "=SUM(Sales[Amount])",
  [input("A1", "Amount"), input("A2", 10), input("A3", 15), input("A4", 999)],
  [output("B1", numberExpected(25))],
  {
    notes: "Table totals rows are metadata, not normal data rows for structured-reference column binding.",
    tables: [
      {
        name: "Sales",
        startAddress: "A1",
        endAddress: "A4",
        columnNames: ["Amount"],
        headerRow: true,
        totalsRow: true,
      },
    ],
  },
),
```

- [ ] **Step 4: Register the new fixtures**

In `packages/formula/src/compatibility.ts`, add `extended` entries matching the new IDs and mark them `implemented-wasm-production` only when the fixture harness and engine runtime both pass. Use `implemented-js` for a newly captured semantic fixture that is not yet production-routed.

```ts
entry("structured-reference:table-column-sum", "structured-reference", "=SUM(Sales[Amount])", "implemented-js", {
  scope: "extended",
  notes: "Workbook semantics fixture for table column binding.",
}),
entry("structured-reference:table-missing-column-ref-error", "structured-reference", "=SUM(Sales[Missing])", "implemented-js", {
  scope: "extended",
  notes: "Workbook semantics fixture for structured-reference missing-column errors.",
}),
entry("structured-reference:table-totals-row-excluded", "structured-reference", "=SUM(Sales[Amount])", "implemented-js", {
  scope: "extended",
  notes: "Workbook semantics fixture for header and totals row table bounds.",
}),
```

- [ ] **Step 5: Run the fixture harness and scorecard**

Run:

```bash
pnpm exec vitest run packages/formula/src/__tests__/fixture-harness.test.ts packages/formula/src/__tests__/compatibility.test.ts scripts/__tests__/calculation-semantics-scorecard.test.ts
pnpm calculation:semantics:generate
pnpm calculation:semantics:check
```

Expected: the tests pass, the generated scorecard includes the new workbook semantics categories, and no missing workbook-semantics fixture IDs remain.

- [ ] **Step 6: Commit**

```bash
git add packages/excel-fixtures/src/workbook-semantics-fixtures.ts packages/formula/src/compatibility.ts scripts/gen-calculation-semantics-scorecard.ts scripts/__tests__/calculation-semantics-scorecard.test.ts packages/benchmarks/baselines/calculation-semantics-scorecard.json
git commit -m "test(formula): expand workbook semantic fixtures"
```

## Task 4: Strengthen JS/WASM Runtime Differential Coverage

**Files:**

- Modify: `packages/core/src/__tests__/formula-runtime-correctness.test.ts`
- Test: `packages/core/src/__tests__/formula-runtime-differential.fuzz.test.ts`

- [ ] **Step 1: Replace family-limited production fixture selection**

In `packages/core/src/__tests__/formula-runtime-correctness.test.ts`, replace the current `text` and `lookup-reference` filter with:

```ts
const engineRunnableProductionFixtures = canonicalFormulaFixtures.filter((fixture) => {
  const entry = getCompatibilityEntry(fixture.id);
  return (
    entry?.wasmStatus === "production" &&
    !runtimeJsOnlyFixtureIds.has(fixture.id) &&
    fixture.multipleOperations === undefined
  );
});
```

- [ ] **Step 2: Keep explicit skip reasons as data**

Add:

```ts
const engineRuntimeSkipReasons = new Map<string, string>([
  ["lookup-reference:offset-basic", "OFFSET is contextual and verified through dedicated engine tests."],
]);
```

Filter with:

```ts
!engineRuntimeSkipReasons.has(fixture.id)
```

- [ ] **Step 3: Assert skipped fixtures stay documented**

Add a test:

```ts
it("documents every production fixture excluded from engine differential parity", () => {
  const skipped = canonicalFormulaFixtures.filter((fixture) => {
    const entry = getCompatibilityEntry(fixture.id);
    return entry?.wasmStatus === "production" && !engineRunnableProductionFixtures.includes(fixture);
  });

  for (const fixture of skipped) {
    const documented = engineRuntimeSkipReasons.has(fixture.id) || runtimeJsOnlyFixtureIds.has(fixture.id);
    expect(documented, `Missing runtime skip reason for ${fixture.id}`).toBe(true);
  }
});
```

- [ ] **Step 4: Run formula runtime tests**

Run:

```bash
pnpm exec vitest run packages/core/src/__tests__/formula-runtime-correctness.test.ts packages/core/src/__tests__/formula-runtime-differential.fuzz.test.ts
```

Expected: pass. If runtime takes too long, split the test into smaller describe blocks by fixture family and keep each block under the existing Vitest timeout.

- [ ] **Step 5: Commit**

```bash
git add packages/core/src/__tests__/formula-runtime-correctness.test.ts
git commit -m "test(core): broaden production formula runtime parity"
```

## Task 5: Add Engine Semantic Invariant Harness

**Files:**

- Create: `packages/core/src/__tests__/engine-semantic-invariants.test.ts`
- Modify: `package.json`

- [ ] **Step 1: Create a deterministic invariant test**

Create `packages/core/src/__tests__/engine-semantic-invariants.test.ts`:

```ts
import { describe, expect, it } from "vitest";
import { SpreadsheetEngine } from "../engine.js";
import {
  diffWorkbookSemanticSnapshots,
  workbookSemanticSnapshotsEqual,
} from "../semantics/index.js";

describe("engine semantic invariants", () => {
  it("keeps edit, snapshot restore, and full recalc semantically equivalent", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "semantic-invariants", replicaId: "semantic-primary" });
    await engine.ready();
    engine.createSheet("Sheet1");
    engine.setRangeValues({ sheetName: "Sheet1", startAddress: "A1", endAddress: "B3" }, [
      [1, 10],
      [2, 20],
      [3, 30],
    ]);
    engine.setCellFormula("Sheet1", "C1", "SUM(A1:A3)");
    engine.setCellFormula("Sheet1", "C2", "SUM(B1:B3)");
    engine.insertRows("Sheet1", 1, 1);
    engine.setCellValue("Sheet1", "A2", 4);
    engine.setCellValue("Sheet1", "B2", 40);

    const beforeRestore = engine.exportSnapshot();

    const restored = new SpreadsheetEngine({ workbookName: "semantic-invariants", replicaId: "semantic-restored" });
    await restored.ready();
    restored.importSnapshot(beforeRestore);
    restored.recalculateNow();

    const afterRestore = restored.exportSnapshot();
    expect(diffWorkbookSemanticSnapshots(beforeRestore, afterRestore)).toEqual([]);
    expect(workbookSemanticSnapshotsEqual(beforeRestore, afterRestore)).toBe(true);
  });
});
```

- [ ] **Step 2: Add semantic test scripts**

Modify `package.json`:

```json
"test:semantic:fast": "pnpm source-size:check && pnpm calculation:semantics:check && pnpm import-export:fidelity:check && tsx scripts/run-vitest.ts --run packages/core/src/__tests__/workbook-semantic-projection.test.ts packages/core/src/__tests__/engine-semantic-invariants.test.ts",
"test:semantic:medium": "pnpm test:semantic:fast && pnpm test:correctness:formula && pnpm test:correctness:core",
"test:semantic:deep": "pnpm test:semantic:medium && pnpm test:correctness:corpus && pnpm test:fuzz:main"
```

- [ ] **Step 3: Run semantic fast gate**

Run:

```bash
pnpm test:semantic:fast
```

Expected: pass.

- [ ] **Step 4: Commit**

```bash
git add package.json packages/core/src/__tests__/engine-semantic-invariants.test.ts
git commit -m "test(core): add semantic invariant gate"
```

## Task 6: Add Import/Export Semantic Loss Ledger

**Files:**

- Create: `scripts/import-export-semantic-loss-ledger.ts`
- Modify: `scripts/gen-import-export-fidelity-scorecard.ts`
- Modify: `scripts/__tests__/import-export-fidelity-scorecard.test.ts`
- Modify: `packages/benchmarks/baselines/import-export-fidelity-scorecard.json`

- [ ] **Step 1: Create the ledger module**

Create `scripts/import-export-semantic-loss-ledger.ts`:

```ts
export type ImportExportSemanticDisposition = "preserved" | "unsupported" | "external" | "declined-runtime";

export interface ImportExportSemanticLedgerEntry {
  readonly feature: string;
  readonly disposition: ImportExportSemanticDisposition;
  readonly reason: string;
}

export const importExportSemanticLossLedger: readonly ImportExportSemanticLedgerEntry[] = [
  {
    feature: "xlsx.macros.execution",
    disposition: "declined-runtime",
    reason: "Bilig preserves macro payload metadata but intentionally never executes workbook macros.",
  },
];

export function importExportUnsupportedFeatures(): readonly string[] {
  return importExportSemanticLossLedger.filter((entry) => entry.disposition === "unsupported").map((entry) => entry.feature).toSorted();
}

export function importExportDeclinedRuntimeFeatures(): readonly string[] {
  return importExportSemanticLossLedger
    .filter((entry) => entry.disposition === "declined-runtime")
    .map((entry) => entry.feature)
    .toSorted();
}
```

- [ ] **Step 2: Use the ledger in the scorecard generator**

In `scripts/gen-import-export-fidelity-scorecard.ts`, replace local constants with:

```ts
import {
  importExportDeclinedRuntimeFeatures,
  importExportUnsupportedFeatures,
} from "./import-export-semantic-loss-ledger.ts";

const unsupportedFeatures = importExportUnsupportedFeatures();
const declinedRuntimeFeatures = importExportDeclinedRuntimeFeatures();
```

- [ ] **Step 3: Test the ledger output**

Add to `scripts/__tests__/import-export-fidelity-scorecard.test.ts`:

```ts
it("keeps unsupported and declined import/export semantics explicit", async () => {
  const scorecard = await buildImportExportFidelityScorecard("test-generated");

  expect(scorecard.summary.unsupportedFeatures).toEqual([]);
  expect(scorecard.summary.declinedRuntimeFeatures).toEqual(["xlsx.macros.execution"]);
});
```

- [ ] **Step 4: Regenerate and check**

Run:

```bash
pnpm import-export:fidelity:generate
pnpm import-export:fidelity:check
```

Expected: pass.

- [ ] **Step 5: Commit**

```bash
git add scripts/import-export-semantic-loss-ledger.ts scripts/gen-import-export-fidelity-scorecard.ts scripts/__tests__/import-export-fidelity-scorecard.test.ts packages/benchmarks/baselines/import-export-fidelity-scorecard.json
git commit -m "test(excel-import): track semantic loss ledger"
```

## Task 7: Add Assistant-Visible Semantic Proof

**Files:**

- Modify: `apps/bilig/src/codex-app/workbook-agent-mutation-proof.ts`
- Modify: `apps/bilig/src/codex-app/workbook-agent-mutation-receipt.ts`
- Modify: `apps/bilig/src/codex-app/workbook-agent-mutation-receipt.test.ts`
- Modify: `apps/web/src/__tests__/worker-runtime-authoritative-bootstrap.test.ts`

- [ ] **Step 1: Add semantic proof to mutation proof types**

In `apps/bilig/src/codex-app/workbook-agent-mutation-proof.ts`, add:

```ts
export interface WorkbookSemanticReadbackProof {
  readonly requested: boolean;
  readonly matched: boolean | null;
  readonly incompleteReason: string | null;
}
```

- [ ] **Step 2: Add semantic proof to receipts**

In `apps/bilig/src/codex-app/workbook-agent-mutation-receipt.ts`, extend `WorkbookToolMutationReceipt`:

```ts
readonly semanticReadback: WorkbookSemanticReadbackProof;
```

Build the proof from the authoritative and rendered proof:

```ts
const semanticReadback: WorkbookSemanticReadbackProof = {
  requested: authoritativeReadback.requested || renderedReadback.requested,
  matched: authoritativeReadback.matched === true && (!renderedReadback.requested || renderedReadback.matched === true),
  incompleteReason:
    authoritativeReadback.matched !== true
      ? authoritativeReadback.incompleteReason ?? "Authoritative semantic readback did not match."
      : renderedReadback.requested && renderedReadback.matched !== true
        ? renderedReadback.incompleteReason ?? "Rendered semantic readback did not match."
        : null,
};
```

Include `semanticReadback` in the returned receipt and require it in `hasAppliedProof`.

- [ ] **Step 3: Add receipt tests**

In `apps/bilig/src/codex-app/workbook-agent-mutation-receipt.test.ts`, assert applied receipts include:

```ts
semanticReadback: z.object({
  requested: z.literal(true),
  matched: z.literal(true),
  incompleteReason: z.null(),
}),
```

Also add a stale-rendered test case:

```ts
expect(payload.mutationReceipt.status).toBe("verification_incomplete");
expect(payload.mutationReceipt.semanticReadback.matched).toBe(false);
expect(payload.mutationReceipt.semanticReadback.incompleteReason).toContain("Rendered");
```

- [ ] **Step 4: Run assistant proof tests**

Run:

```bash
pnpm exec vitest run apps/bilig/src/codex-app/workbook-agent-mutation-receipt.test.ts apps/bilig/src/codex-app/workbook-agent-rendered-freshness.test.ts apps/web/src/__tests__/worker-runtime-authoritative-bootstrap.test.ts
```

Expected: pass.

- [ ] **Step 5: Commit**

```bash
git add apps/bilig/src/codex-app/workbook-agent-mutation-proof.ts apps/bilig/src/codex-app/workbook-agent-mutation-receipt.ts apps/bilig/src/codex-app/workbook-agent-mutation-receipt.test.ts apps/web/src/__tests__/worker-runtime-authoritative-bootstrap.test.ts
git commit -m "feat(workbook-agent): add semantic mutation proof"
```

## Task 8: Wire Semantic Gates Into CI

**Files:**

- Modify: `scripts/run-ci.ts`
- Modify: `package.json`
- Test: `scripts/__tests__/run-ci*.test.ts` if present

- [ ] **Step 1: Add semantic fast gate to CI**

In `scripts/run-ci.ts`, add `pnpm test:semantic:fast` to the fast profile after generated artifact checks and before browser checks.

Use the existing command runner style in that file. The intended command entry is:

```ts
{
  label: "semantic correctness fast gate",
  command: "pnpm",
  args: ["test:semantic:fast"],
}
```

- [ ] **Step 2: Keep live Excel/Sheets checks out of default CI**

Do not add these commands to fast CI:

```bash
pnpm calculation:excel-live:check
pnpm calculation:google-sheets-live:check
pnpm structural:excel-live:check
pnpm structural:google-sheets-live:check
```

Add them only to release or manually refreshed scorecard workflows because they depend on local app automation or external services.

- [ ] **Step 3: Run CI script tests and semantic fast gate**

Run:

```bash
pnpm exec vitest run scripts/__tests__
pnpm test:semantic:fast
```

Expected: pass.

- [ ] **Step 4: Run final preflight**

Run:

```bash
pnpm typecheck
pnpm lint
pnpm run ci
```

Expected: all pass.

- [ ] **Step 5: Commit**

```bash
git add scripts/run-ci.ts package.json
git commit -m "ci: gate semantic correctness"
```

## Exit Gate

The semantic correctness program is complete only when all of these are true:

- `projectWorkbookSemanticSnapshot()` is the only shared semantic comparison contract used by core, import/export, corpus, and assistant proof code.
- Workbook-semantics fixture coverage is category-based, not just a count of 11 fixtures.
- All production-routed canonical formula fixtures that can run in the engine are covered by JS/WASM differential recalc or have a checked-in skip reason.
- Import/export scorecards distinguish preserved, unsupported, external, and intentionally declined semantics.
- Applied workbook-agent mutation receipts require authoritative proof, rendered proof when requested, semantic proof, and undo availability before reporting `status: "applied"`.
- `pnpm build` passes.
- `pnpm test:semantic:fast` passes.
- `pnpm test:semantic:medium` passes before merging.
- `pnpm run ci` passes before pushing or merging.
