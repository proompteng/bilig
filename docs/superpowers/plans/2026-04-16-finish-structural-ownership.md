# Structural Ownership Finish Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Finish structural ownership end to end so row and column insert, delete, and move flow through one authoritative structural transaction, eliminate the remaining broad remap and rebuild paths, and flip the structural benchmark family against HyperFormula without breaking undo, replica parity, or snapshot/export semantics.

**Architecture:** Introduce one shared `StructuralTransaction` substrate that owns axis remap scope, moved and removed cell coordinates, touched ranges, preserved-formula candidates, and inverse-history payloads. `WorkbookStore`, `SheetGrid`, `RangeRegistry`, `StructureService`, `MutationService`, and the event layer all consume that transaction instead of each recomputing structure impact independently.

**Tech Stack:** TypeScript, Effect, Vitest, fast-check fuzz suites, `pnpm`, Vite browser tests, competitive benchmark harness, existing formula rewrite helpers in `@bilig/formula`.

---

## Scope Lock

This plan finishes **structural ownership only**. It does not include:

- family-compressed dependency graph work
- Pearce/Kelly topo repair
- exact lookup-after-write cleanup
- sorted lookup-after-write cleanup
- sliding-window aggregate reuse cleanup

Those come after structural ownership. Do not dilute this tranche.

## One-Turn / Non-Stop Execution Rules

This is the production-safe way to do “one turn, non stop”:

1. Execute all tasks in order without waiting for user confirmation.
2. Stop only for:
   - red CI that cannot be root-caused locally
   - an external remote move that requires rebase
   - a benchmark regression that proves the current task design is wrong
3. Commit after each task group once the local slice is green.
4. Run full `pnpm run ci` at three checkpoints:
   - after Task 3
   - after Task 5
   - after Task 7
5. Push only after the final committed tree is CI-clean.

## Current Structural Reality To Beat

Current structural lane reality from the broader competitive suite on `main`:

- `structural-insert-rows`: about `10.14 ms` vs HyperFormula `5.95 ms`
- `structural-delete-rows`: about `56-92 ms` depending on tree state, still badly red
- `structural-move-rows`: about `66-101 ms`, still badly red
- `structural-insert-columns`: about `21-23 ms` vs HyperFormula `0.55-0.64 ms`
- `structural-delete-columns`: about `32-41 ms` vs HyperFormula `8.15-8.19 ms`
- `structural-move-columns`: about `11.88-12.12 ms` vs HyperFormula `6.23-8.45 ms`

Key measured fact already established in this repo:

- row delete and row move are **not** dominated by recalc
- direct engine timing showed `recalcMs` around `1.5-1.9 ms` inside operations that still cost `~60 ms`
- the remaining bottleneck is structural ownership itself

## File Map

### Create

- `packages/core/src/engine/structural-transaction.ts`
  - shared `StructuralTransaction` types and helpers
- `packages/core/src/__tests__/structural-transaction.test.ts`
  - unit tests for scope calculation, move windows, and deleted-span bookkeeping

### Modify

- `packages/core/src/sheet-grid.ts`
  - produce remap results as reusable structural entries instead of ad hoc changed-entry arrays
- `packages/core/src/workbook-store.ts`
  - plan and apply axis transactions
  - remap only touched cells
  - update sheet structure versions and row/column invalidation from transaction output
- `packages/core/src/range-registry.ts`
  - retarget range nodes in place from the structural transaction
- `packages/core/src/engine/services/structure-service.ts`
  - consume `StructuralTransaction`
  - classify preserve-binding vs rebind vs delete
  - stop recomputing structure effects from scratch
- `packages/core/src/engine/services/formula-binding-service.ts`
  - preserve bindings and compiled plans from transaction-owned rewrite data
- `packages/core/src/engine/services/mutation-service.ts`
  - use transaction-owned inverse payloads and narrow stored-cell capture
- `packages/core/src/engine/live.ts`
  - wire transaction helpers through runtime construction
- `packages/core/src/engine/services/operation-service.ts`
  - emit row/column invalidations from transaction result instead of remapped-cell floods
- `packages/headless/src/work-paper-runtime.ts`
  - keep structural public semantics aligned with preserved-value structural edits

### Tests To Extend

- `packages/core/src/__tests__/sheet-grid.test.ts`
- `packages/core/src/__tests__/workbook-store.test.ts`
- `packages/core/src/__tests__/structure-service.test.ts`
- `packages/core/src/__tests__/mutation-service.test.ts`
- `packages/core/src/__tests__/engine.test.ts`
- `packages/core/src/__tests__/engine-history.fuzz.test.ts`
- `packages/core/src/__tests__/engine-structure.fuzz.test.ts`
- `packages/core/src/__tests__/engine-correctness.test.ts`
- `packages/headless/src/__tests__/work-paper-runtime.test.ts`
- `packages/benchmarks/src/__tests__/expanded-workloads.test.ts`

## Target API Shape

Task 1 defines this API. Later tasks must use these names exactly.

```ts
export interface StructuralRemappedCell {
  readonly cellIndex: number;
  readonly fromRow: number;
  readonly fromCol: number;
  readonly toRow: number | undefined;
  readonly toCol: number | undefined;
}

export interface StructuralInvalidationSpan {
  readonly axis: "row" | "column";
  readonly start: number;
  readonly end: number;
}

export interface StructuralTransaction {
  readonly sheetName: string;
  readonly sheetId: number;
  readonly transform: StructuralAxisTransform;
  readonly scope: SheetGridAxisRemapScope;
  readonly remappedCells: readonly StructuralRemappedCell[];
  readonly removedCellIndices: readonly number[];
  readonly invalidationSpans: readonly StructuralInvalidationSpan[];
}
```

Do not rename these types mid-plan.

### Task 1: Introduce The Structural Transaction Substrate

**Files:**
- Create: `packages/core/src/engine/structural-transaction.ts`
- Test: `packages/core/src/__tests__/structural-transaction.test.ts`
- Modify: `packages/core/src/engine/services/structure-service.ts`

- [ ] **Step 1: Write the failing transaction-shape tests**

```ts
import { describe, expect, it } from "vitest";
import { buildStructuralTransaction, structuralScopeForTransform } from "../engine/structural-transaction.js";

describe("StructuralTransaction", () => {
  it("builds a bounded move scope that covers both source and target windows", () => {
    expect(
      structuralScopeForTransform({
        axis: "row",
        kind: "move",
        start: 10,
        count: 3,
        target: 4,
      }),
    ).toEqual({ start: 4, end: 13 });
  });

  it("records deleted cells separately from remapped survivors", () => {
    const transaction = buildStructuralTransaction({
      sheetName: "Sheet1",
      sheetId: 1,
      transform: {
        axis: "column",
        kind: "delete",
        start: 2,
        count: 1,
      },
      remappedCells: [
        { cellIndex: 10, fromRow: 0, fromCol: 4, toRow: 0, toCol: 3 },
        { cellIndex: 11, fromRow: 5, fromCol: 2, toRow: undefined, toCol: undefined },
      ],
    });

    expect(transaction.removedCellIndices).toEqual([11]);
    expect(transaction.remappedCells).toHaveLength(2);
  });
});
```

- [ ] **Step 2: Run the new tests and verify they fail**

Run:

```bash
pnpm exec vitest run packages/core/src/__tests__/structural-transaction.test.ts
```

Expected:

- FAIL with missing module or missing export errors for `structural-transaction.ts`

- [ ] **Step 3: Add the shared transaction module**

```ts
import type { StructuralAxisTransform } from "@bilig/formula";
import type { SheetGridAxisRemapScope } from "../sheet-grid.js";

export interface StructuralRemappedCell {
  readonly cellIndex: number;
  readonly fromRow: number;
  readonly fromCol: number;
  readonly toRow: number | undefined;
  readonly toCol: number | undefined;
}

export interface StructuralInvalidationSpan {
  readonly axis: "row" | "column";
  readonly start: number;
  readonly end: number;
}

export interface StructuralTransaction {
  readonly sheetName: string;
  readonly sheetId: number;
  readonly transform: StructuralAxisTransform;
  readonly scope: SheetGridAxisRemapScope;
  readonly remappedCells: readonly StructuralRemappedCell[];
  readonly removedCellIndices: readonly number[];
  readonly invalidationSpans: readonly StructuralInvalidationSpan[];
}

export function structuralScopeForTransform(transform: StructuralAxisTransform): SheetGridAxisRemapScope {
  if (transform.kind === "move") {
    if (transform.target < transform.start) {
      return { start: transform.target, end: transform.start + transform.count };
    }
    if (transform.target > transform.start) {
      return { start: transform.start, end: transform.target + transform.count };
    }
    return { start: transform.start, end: transform.start + transform.count };
  }
  return { start: transform.start };
}

export function buildStructuralTransaction(input: {
  sheetName: string;
  sheetId: number;
  transform: StructuralAxisTransform;
  remappedCells: readonly StructuralRemappedCell[];
}): StructuralTransaction {
  const scope = structuralScopeForTransform(input.transform);
  const removedCellIndices = input.remappedCells
    .filter((entry) => entry.toRow === undefined || entry.toCol === undefined)
    .map((entry) => entry.cellIndex);
  return {
    sheetName: input.sheetName,
    sheetId: input.sheetId,
    transform: input.transform,
    scope,
    remappedCells: input.remappedCells,
    removedCellIndices,
    invalidationSpans: [
      {
        axis: input.transform.axis,
        start: scope.start,
        end:
          input.transform.kind === "move"
            ? (scope.end ?? input.transform.start + input.transform.count)
            : input.transform.start + input.transform.count,
      },
    ],
  };
}
```

- [ ] **Step 4: Run the transaction tests and typecheck**

Run:

```bash
pnpm exec vitest run packages/core/src/__tests__/structural-transaction.test.ts
pnpm exec tsc -p packages/core/tsconfig.json --noEmit
```

Expected:

- PASS
- no type errors

- [ ] **Step 5: Commit**

```bash
git add packages/core/src/engine/structural-transaction.ts packages/core/src/__tests__/structural-transaction.test.ts
git commit -m "refactor(engine): add structural transaction substrate"
```

### Task 2: Make SheetGrid And WorkbookStore Produce One Structural Transaction

**Files:**
- Modify: `packages/core/src/sheet-grid.ts`
- Modify: `packages/core/src/workbook-store.ts`
- Test: `packages/core/src/__tests__/sheet-grid.test.ts`
- Test: `packages/core/src/__tests__/workbook-store.test.ts`

- [ ] **Step 1: Write failing remap tests**

```ts
it("returns stable remap entries for deleteRows without touching unrelated blocks", () => {
  const grid = new SheetGrid();
  grid.set(0, 0, 1);
  grid.set(10, 2, 2);
  grid.set(11, 2, 3);

  const remapped = grid.remapAxis("row", (row) => {
    if (row === 10) return undefined;
    if (row > 10) return row - 1;
    return row;
  }, { start: 10 });

  expect(remapped).toEqual([
    { cellIndex: 2, row: 10, col: 2, nextRow: undefined, nextCol: 2 },
    { cellIndex: 3, row: 11, col: 2, nextRow: 10, nextCol: 2 },
  ]);
  expect(grid.get(0, 0)).toBe(1);
  expect(grid.get(10, 2)).toBe(3);
});

it("builds one transaction from workbook remap and tracks removed cell indices", () => {
  const workbook = new WorkbookStore("structural");
  workbook.createSheet("Sheet1");
  workbook.setCellValue("Sheet1", "C11", 7);
  workbook.setCellValue("Sheet1", "C12", 8);

  const transaction = workbook.applyStructuralAxisTransform("Sheet1", {
    axis: "row",
    kind: "delete",
    start: 10,
    count: 1,
  });

  expect(transaction.removedCellIndices).toHaveLength(1);
  expect(transaction.remappedCells.some((entry) => entry.toRow === 10)).toBe(true);
});
```

- [ ] **Step 2: Run the targeted tests and verify failure**

Run:

```bash
pnpm exec vitest run packages/core/src/__tests__/sheet-grid.test.ts packages/core/src/__tests__/workbook-store.test.ts
```

Expected:

- FAIL because `WorkbookStore` does not yet expose an authoritative structural transaction

- [ ] **Step 3: Refactor `SheetGrid.remapAxis` to return reusable remap entries**

```ts
export interface SheetGridRemapEntry {
  readonly cellIndex: number;
  readonly row: number;
  readonly col: number;
  readonly nextRow: number | undefined;
  readonly nextCol: number | undefined;
}

remapAxis(...): SheetGridRemapEntry[] {
  const remapped: SheetGridRemapEntry[] = [];
  for (const key of this.blocks.keys()) {
    if (!blockIntersectsScope(axis, key, scope)) continue;
    // preserve the current sparse-block walk
    // but record each moved or deleted cell as a reusable remap entry
  }
  return remapped;
}
```

- [ ] **Step 4: Add `WorkbookStore.applyStructuralAxisTransform` that returns `StructuralTransaction`**

```ts
applyStructuralAxisTransform(sheetName: string, transform: StructuralAxisTransform): StructuralTransaction {
  const sheet = this.requireSheet(sheetName);
  const remappedCells = sheet.grid.remapAxis(
    transform.axis,
    (index) => mapStructuralAxisIndex(index, transform),
    structuralScopeForTransform(transform),
  ).map((entry) => ({
    cellIndex: entry.cellIndex,
    fromRow: entry.row,
    fromCol: entry.col,
    toRow: entry.nextRow,
    toCol: entry.nextCol,
  }));

  for (const entry of remappedCells) {
    this.remapCellIndex(sheet, entry);
  }

  this.bumpStructureVersion(sheet);
  return buildStructuralTransaction({
    sheetName,
    sheetId: sheet.id,
    transform,
    remappedCells,
  });
}
```

- [ ] **Step 5: Re-run the remap tests and typecheck**

Run:

```bash
pnpm exec vitest run packages/core/src/__tests__/sheet-grid.test.ts packages/core/src/__tests__/workbook-store.test.ts
pnpm exec tsc -p packages/core/tsconfig.json --noEmit
```

Expected:

- PASS

- [ ] **Step 6: Benchmark the insert lanes immediately**

Run:

```bash
pnpm bench:workpaper:competitive -- --sample-count 2 --warmup-count 0
```

Expected movement:

- `structural-insert-rows` materially down from current level
- `structural-insert-columns` materially down from current level
- no regression in `build-dense-literals`, `rebuild-config-toggle`, or `lookup-with-column-index`

- [ ] **Step 7: Commit**

```bash
git add packages/core/src/sheet-grid.ts packages/core/src/workbook-store.ts packages/core/src/__tests__/sheet-grid.test.ts packages/core/src/__tests__/workbook-store.test.ts
git commit -m "refactor(store): drive axis remaps through structural transactions"
```

### Task 3: Move RangeRegistry And Metadata Retargeting Onto The Transaction

**Files:**
- Modify: `packages/core/src/range-registry.ts`
- Modify: `packages/core/src/workbook-store.ts`
- Modify: `packages/core/src/engine/services/structure-service.ts`
- Test: `packages/core/src/__tests__/structure-service.test.ts`
- Test: `packages/core/src/__tests__/engine.test.ts`

- [ ] **Step 1: Add failing tests for in-place range retargeting**

```ts
it("retargets range dependencies from a structural transaction without full rebuild", async () => {
  const engine = new SpreadsheetEngine();
  await engine.ready();
  engine.createSheet("Sheet1");
  engine.setCellValue("Sheet1", "A1", 1);
  engine.setCellValue("Sheet1", "A2", 2);
  engine.setCellFormula("Sheet1", "B1", "SUM(A1:A2)");

  engine.deleteRows("Sheet1", 0, 1);

  expect(engine.getCellFormula("Sheet1", "B1")).toBe("SUM(A1:A1)");
  expect(engine.getCellValue("Sheet1", "B1")).toEqual({ tag: ValueTag.Number, value: 2 });
});
```

- [ ] **Step 2: Run the structural tests and verify failure**

Run:

```bash
pnpm exec vitest run packages/core/src/__tests__/structure-service.test.ts packages/core/src/__tests__/engine.test.ts
```

Expected:

- FAIL because range maintenance still depends on broader refresh behavior

- [ ] **Step 3: Add `RangeRegistry.applyStructuralTransaction`**

```ts
applyStructuralTransaction(transaction: StructuralTransaction): number[] {
  const touchedRangeIndices: number[] = [];
  for (const range of this.ranges.values()) {
    if (range.sheetId !== transaction.sheetId) continue;
    const rewritten = rewriteRangeForStructuralTransform(range.descriptor, transaction.transform);
    if (!rewritten.changed) continue;
    range.descriptor = rewritten.range;
    touchedRangeIndices.push(range.index);
  }
  return touchedRangeIndices;
}
```

- [ ] **Step 4: Update `StructureService` to consume touched ranges from the transaction**

```ts
const transaction = workbook.applyStructuralAxisTransform(sheetName, transform);
const touchedRangeIndices = ranges.applyStructuralTransaction(transaction);
refreshRangeDependencies(touchedRangeIndices);
```

- [ ] **Step 5: Run full structural core slice**

Run:

```bash
pnpm exec vitest run packages/core/src/__tests__/sheet-grid.test.ts packages/core/src/__tests__/workbook-store.test.ts packages/core/src/__tests__/structure-service.test.ts packages/core/src/__tests__/engine.test.ts
pnpm run ci
```

Expected:

- green structural core slice
- full CI green at this checkpoint

- [ ] **Step 6: Commit**

```bash
git add packages/core/src/range-registry.ts packages/core/src/workbook-store.ts packages/core/src/engine/services/structure-service.ts packages/core/src/__tests__/structure-service.test.ts packages/core/src/__tests__/engine.test.ts
git commit -m "refactor(engine): retarget ranges from structural transactions"
```

### Task 4: Make Formula Rewrite Ownership Transaction-Driven

**Files:**
- Modify: `packages/core/src/engine/services/structure-service.ts`
- Modify: `packages/core/src/engine/services/formula-binding-service.ts`
- Modify: `packages/core/src/engine/services/compiled-plan-service.ts`
- Test: `packages/core/src/__tests__/engine.test.ts`
- Test: `packages/core/src/__tests__/structure-service.test.ts`

- [ ] **Step 1: Add failing preserve-path tests for delete and move**

```ts
it("keeps plan ids stable for shape-preserving moveColumns rewrites", async () => {
  const engine = new SpreadsheetEngine();
  await engine.ready();
  engine.createSheet("Sheet1");
  engine.setCellFormula("Sheet1", "C1", "A1+B1");
  const before = readRuntimeFormula(engine, engine.getCellIndex("Sheet1", "C1"));

  engine.moveColumns("Sheet1", 0, 1, 2);

  const after = readRuntimeFormula(engine, engine.getCellIndex("Sheet1", "C1"));
  expect(after.planId).toBe(before.planId);
});

it("marks topology changed only when preserve binding falls off", async () => {
  const engine = new SpreadsheetEngine();
  await engine.ready();
  engine.createSheet("Sheet1");
  engine.setCellValue("Sheet1", "A1", 1);
  engine.setCellValue("Sheet1", "A2", 2);
  engine.setCellFormula("Sheet1", "B1", "SUM(A1:A2)");
  engine.defineName("NamedInput", { sheetName: "Sheet1", address: "A1" });
  engine.setCellFormula("Sheet1", "C1", "NamedInput+1");

  const event = captureSingleEngineEvent(() => engine.insertRows("Sheet1", 0, 1));

  expect(event.invalidatedRows).toEqual([{ start: 0, end: 1 }]);
  expect(event.topologyChanged).toBe(true);
});
```

- [ ] **Step 2: Run the structural formula tests and verify failure**

Run:

```bash
pnpm exec vitest run packages/core/src/__tests__/engine.test.ts packages/core/src/__tests__/structure-service.test.ts
```

Expected:

- FAIL on plan-id stability or topology-change assertions

- [ ] **Step 3: Refactor structure service to classify formulas from the transaction**

```ts
interface StructuralFormulaEffect {
  readonly cellIndex: number;
  readonly action: "preserve" | "rebind" | "remove";
  readonly source: string;
  readonly compiled?: CompiledFormula;
}

const effects = classifyStructuralFormulaEffects(transaction, formulas, workbook);
```

- [ ] **Step 4: Reuse compiled plans in place for preserve-path rewrites**

```ts
if (effect.action === "preserve") {
  formulaBinding.updateFormulaDependenciesInPlaceNow({
    cellIndex: effect.cellIndex,
    source: effect.source,
    compiled: effect.compiled,
  });
}
```

- [ ] **Step 5: Re-run tests and benchmark the move lanes**

Run:

```bash
pnpm exec vitest run packages/core/src/__tests__/engine.test.ts packages/core/src/__tests__/structure-service.test.ts
pnpm bench:workpaper:competitive -- --sample-count 2 --warmup-count 0
```

Expected movement:

- `structural-move-rows` down materially
- `structural-move-columns` down materially
- `topologyChanged` no longer flips on preserved inserts and moves

- [ ] **Step 6: Commit**

```bash
git add packages/core/src/engine/services/structure-service.ts packages/core/src/engine/services/formula-binding-service.ts packages/core/src/engine/services/compiled-plan-service.ts packages/core/src/__tests__/engine.test.ts packages/core/src/__tests__/structure-service.test.ts
git commit -m "perf(engine): make structural formula rewrites transaction-driven"
```

### Task 5: Move Undo/Redo And Structural Inverse Capture Onto The Same Transaction

**Files:**
- Modify: `packages/core/src/engine/services/mutation-service.ts`
- Modify: `packages/core/src/engine/live.ts`
- Test: `packages/core/src/__tests__/mutation-service.test.ts`
- Test: `packages/core/src/__tests__/engine-history.fuzz.test.ts`
- Test: `packages/core/src/__tests__/engine-structure.fuzz.test.ts`
- Test: `packages/core/src/__tests__/engine-correctness.test.ts`

- [ ] **Step 1: Add failing history tests for transaction-owned inverse replay**

```ts
it("restores deleted row spans without restoring unrelated literal cells", async () => {
  const engine = new SpreadsheetEngine();
  await engine.ready();
  engine.createSheet("Sheet1");
  engine.setCellValue("Sheet1", "A1", 1);
  engine.setCellValue("Sheet1", "A20", 20);
  engine.deleteRows("Sheet1", 0, 1);
  engine.undo();

  expect(engine.getCellValue("Sheet1", "A1")).toEqual({ tag: ValueTag.Number, value: 1 });
  expect(engine.getCellValue("Sheet1", "A20")).toEqual({ tag: ValueTag.Number, value: 20 });
});

it("preserves explicit authored blanks through structural undo replay", async () => {
  const engine = new SpreadsheetEngine();
  await engine.ready();
  engine.createSheet("Sheet1");
  engine.setCellFormat("Sheet1", "A2", "0.00");
  engine.setCellValue("Sheet1", "A2", null);
  engine.deleteRows("Sheet1", 1, 1);
  engine.undo();

  expect(engine.exportSnapshot().sheets[0]?.cells).toContainEqual({
    address: "A2",
    value: null,
    format: "0.00",
  });
});
```

- [ ] **Step 2: Run history and fuzz slices**

Run:

```bash
pnpm exec vitest run packages/core/src/__tests__/mutation-service.test.ts packages/core/src/__tests__/engine-history.fuzz.test.ts packages/core/src/__tests__/engine-structure.fuzz.test.ts packages/core/src/__tests__/engine-correctness.test.ts
```

Expected:

- FAIL in at least one structural inverse replay path until the transaction drives inverse capture

- [ ] **Step 3: Refactor mutation service to capture inverse payloads from the transaction**

```ts
const transaction = structure.applyStructuralAxisOp(op);
const inverse = captureStructuralInverseFromTransaction(transaction, workbook, formulas);
pushHistoryFrame({
  forward: { kind: "single-op", op },
  inverse,
});
```

- [ ] **Step 4: Use stored-cell capture only for touched spans and required formula owners**

```ts
function captureStructuralInverseFromTransaction(
  transaction: StructuralTransaction,
  workbook: WorkbookStore,
  formulas: FormulaTable<RuntimeFormula>,
): EngineOp[] {
  return [
    ...captureDeletedSpanCells(transaction),
    ...captureImpactedFormulaOwners(transaction, formulas),
    ...captureStructuralMetadata(transaction),
  ];
}
```

- [ ] **Step 5: Run the history, fuzz, and full CI checkpoint**

Run:

```bash
pnpm exec vitest run packages/core/src/__tests__/mutation-service.test.ts packages/core/src/__tests__/engine-history.fuzz.test.ts packages/core/src/__tests__/engine-structure.fuzz.test.ts packages/core/src/__tests__/engine-correctness.test.ts
pnpm run ci
```

Expected:

- green structural history slice
- full CI green at this checkpoint

- [ ] **Step 6: Commit**

```bash
git add packages/core/src/engine/services/mutation-service.ts packages/core/src/engine/live.ts packages/core/src/__tests__/mutation-service.test.ts packages/core/src/__tests__/engine-history.fuzz.test.ts packages/core/src/__tests__/engine-structure.fuzz.test.ts packages/core/src/__tests__/engine-correctness.test.ts
git commit -m "fix(engine): drive structural undo replay from transactions"
```

### Task 6: Replace Remapped-Cell Flooding With Axis Invalidations Everywhere

**Files:**
- Modify: `packages/core/src/engine/services/operation-service.ts`
- Modify: `packages/core/src/engine/services/structure-service.ts`
- Modify: `packages/headless/src/work-paper-runtime.ts`
- Test: `packages/headless/src/__tests__/work-paper-runtime.test.ts`
- Test: `packages/core/src/__tests__/engine.test.ts`

- [ ] **Step 1: Add failing event-payload tests**

```ts
it("emits row invalidations instead of a changed-cell flood for preserved insertRows", async () => {
  const engine = new SpreadsheetEngine();
  await engine.ready();
  engine.createSheet("Sheet1");
  engine.setCellFormula("Sheet1", "A2", "1+2");

  const event = captureSingleEngineEvent(() => engine.insertRows("Sheet1", 0, 1));

  expect(event.changedCells).toEqual([]);
  expect(event.invalidatedRows).toEqual([{ start: 0, end: 1 }]);
});
```

- [ ] **Step 2: Run the operation/headless tests**

Run:

```bash
pnpm exec vitest run packages/core/src/__tests__/engine.test.ts packages/headless/src/__tests__/work-paper-runtime.test.ts
```

Expected:

- FAIL because current payloads still mix remapped-cell lists with structural invalidation

- [ ] **Step 3: Emit invalidation spans from the transaction**

```ts
const transactionResult = structure.applyStructuralAxisOp(op);
events.push({
  changedCells: transactionResult.topologyChanged ? changedCells : [],
  invalidatedRows: collectRowInvalidations(transactionResult),
  invalidatedColumns: collectColumnInvalidations(transactionResult),
});
```

- [ ] **Step 4: Align headless structural semantics**

```ts
if (!result.topologyChanged && result.changedCellIndices.length === 0) {
  return [];
}
```

- [ ] **Step 5: Run targeted tests and competitive benchmark**

Run:

```bash
pnpm exec vitest run packages/core/src/__tests__/engine.test.ts packages/headless/src/__tests__/work-paper-runtime.test.ts
pnpm bench:workpaper:competitive -- --sample-count 2 --warmup-count 0
```

Expected movement:

- insert lanes stay green or improve
- no event-payload regressions in headless/runtime semantics

- [ ] **Step 6: Commit**

```bash
git add packages/core/src/engine/services/operation-service.ts packages/core/src/engine/services/structure-service.ts packages/headless/src/work-paper-runtime.ts packages/headless/src/__tests__/work-paper-runtime.test.ts packages/core/src/__tests__/engine.test.ts
git commit -m "perf(engine): emit structural axis invalidations from transactions"
```

### Task 7: Delete The Displaced Structural Paths And Lock The Finish Criteria

**Files:**
- Modify: `packages/core/src/engine/services/structure-service.ts`
- Modify: `packages/core/src/workbook-store.ts`
- Modify: `packages/core/src/engine/services/mutation-service.ts`
- Modify: `packages/benchmarks/src/__tests__/expanded-workloads.test.ts`
- Modify: `docs/workpaper-ultra-performance-engine-architecture-2026-04-12.md`
- Modify: `docs/workpaper-ultra-performance-engine-delivery-2026-04-12.md`

- [ ] **Step 1: Remove displaced broad structural paths**

Delete any path that still:

- recomputes structural scope outside `StructuralTransaction`
- captures whole-sheet structural inverse cell state
- emits remapped-cell floods when only row/column invalidation is required
- retargets ranges outside `RangeRegistry.applyStructuralTransaction`

- [ ] **Step 2: Add regression assertions to the benchmark workload tests**

```ts
it("keeps all six structural workloads registered in the expanded suite", () => {
  expect(workloadNames).toEqual(
    expect.arrayContaining([
      "structural-insert-rows",
      "structural-delete-rows",
      "structural-move-rows",
      "structural-insert-columns",
      "structural-delete-columns",
      "structural-move-columns",
    ]),
  );
});
```

- [ ] **Step 3: Run the final full verification**

Run:

```bash
pnpm exec vitest run packages/core/src/__tests__/sheet-grid.test.ts packages/core/src/__tests__/workbook-store.test.ts packages/core/src/__tests__/structure-service.test.ts packages/core/src/__tests__/mutation-service.test.ts packages/core/src/__tests__/engine.test.ts packages/core/src/__tests__/engine-history.fuzz.test.ts packages/core/src/__tests__/engine-structure.fuzz.test.ts packages/core/src/__tests__/engine-correctness.test.ts packages/headless/src/__tests__/work-paper-runtime.test.ts packages/benchmarks/src/__tests__/expanded-workloads.test.ts
pnpm bench:workpaper:competitive -- --sample-count 2 --warmup-count 0
pnpm bench:workpaper:competitive -- --sample-count 5 --warmup-count 0
pnpm run ci
```

Expected:

- all targeted structural tests green
- both benchmark artifacts saved on the committed tree
- full CI green

- [ ] **Step 4: Update the architecture docs with the new structural truth**

Record:

- final structural lane numbers
- whether delete and move are now green
- any residual red that is no longer structural-ownership related

- [ ] **Step 5: Final commit**

```bash
git add packages/core/src/engine/services/structure-service.ts packages/core/src/workbook-store.ts packages/core/src/engine/services/mutation-service.ts packages/benchmarks/src/__tests__/expanded-workloads.test.ts docs/workpaper-ultra-performance-engine-architecture-2026-04-12.md docs/workpaper-ultra-performance-engine-delivery-2026-04-12.md
git commit -m "perf(engine): finish structural ownership cutover"
```

## Benchmark Gates

These are the only numbers that matter for calling the phase done.

### Midpoint Gate After Task 4

- `structural-insert-rows` under `8 ms`
- `structural-insert-columns` under `5 ms`
- `structural-move-columns` under `10 ms`
- no regression in `rebuild-config-toggle`, `lookup-with-column-index`, or `aggregate-overlapping-ranges`

### Final Structural Ownership Gate

On the `sample-count 2` competitive suite:

- `structural-insert-rows` <= HyperFormula
- `structural-delete-rows` <= `2x` HyperFormula on the first finish pass, then keep pushing
- `structural-move-rows` <= `2x` HyperFormula on the first finish pass, then keep pushing
- `structural-insert-columns` <= `4x` HyperFormula on the first finish pass
- `structural-delete-columns` <= `1.5x` HyperFormula
- `structural-move-columns` <= `1.25x` HyperFormula

On the `sample-count 5` suite:

- no structural lane regresses back above the current pre-plan baselines
- at least `4/6` structural lanes are green

If those gates are missed because only row delete and move remain red, structural ownership is
**mostly complete but not finished**. Do not declare it done.

## Risks And Rollback Criteria

### Risk 1: Undo replay regresses again

Rollback rule:

- revert only the latest structural inverse-capture change
- keep transaction substrate and store/range ownership

### Risk 2: Preserve-binding overreaches and returns stale values

Rollback rule:

- narrow preserve-path eligibility
- do **not** reintroduce full broad structural rebinding everywhere

### Risk 3: Event payload changes break headless or browser expectations

Rollback rule:

- keep row/column invalidation spans
- temporarily emit a compatibility shim
- remove it before final completion

## Done Criteria

Structural ownership is done only when all of the following are true:

1. `WorkbookStore`, `RangeRegistry`, `StructureService`, `MutationService`, and runtime event emission all consume one `StructuralTransaction`.
2. Structural inverse replay no longer captures unrelated sheet state.
3. Preserve-binding structural rewrites reuse compiled plans in place when shape and semantics allow it.
4. Structural insert no longer floods `changedCells`.
5. Structural delete and move are no longer dominated by broad structural bookkeeping.
6. Full `pnpm run ci` is green on the committed push candidate.
7. Competitive benchmark artifacts are rerun and saved on the committed push candidate.

## Final Notes

- Do not start TACO, topo repair, or lookup-after-write cleanup until this plan is complete.
- If row delete and row move remain structurally red after Task 7, the next move is **not**
  another generic optimization pass. It is a deeper row/column address-mapping core under the same
  transaction model.
