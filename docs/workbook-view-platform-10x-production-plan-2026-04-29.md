# Workbook View Platform 10x Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Make Bilig's workbook surface measurably more trustworthy and more capable than Excel and Google Sheets for AI-assisted financial work: visible state must match authoritative state, selection must remain exact under scroll and virtualization, assistant mutations must prove both model and rendered outcomes, and analytics views must be built on a stable view-window contract.

**Architecture:** Keep Bilig's existing workbook engine, React shell, worker transport, and TypeGPU renderer. Add a Perspective-inspired view-window layer, render commit barrier, stable sheet identity rules, plugin-grade view contracts, and assistant semantic workflows around the existing system instead of replacing the grid with Perspective.

**Tech Stack:** TypeScript, React, Vite, pnpm monorepo, `apps/web`, `apps/bilig`, `packages/grid`, `packages/worker-transport`, `packages/workbook-domain`, TypeGPU renderer V3, Vitest, Playwright.

---

## Product Bar

"10x better than Excel and Google Sheets" is not a slogan for this work. It means Bilig must make classes of spreadsheet failure impossible or immediately observable:

- No silent blank grid when authoritative workbook data exists.
- No stale sheet names or table metadata after rename, create, delete, or table operations.
- No selection drift after vertical or horizontal scroll.
- No assistant claim of success until the mutation is applied, recalculated, reflected in context, rendered in the visible grid, and verified by invariant checks.
- No coordinate-only authoring burden for normal financial templates.
- No ambiguity between real formula errors and compatibility warnings.
- No UI success state without a render batch acknowledgement.
- No plugin or analytics view reading from private renderer internals.

## Current System Facts

Bilig already has several strong primitives that should be preserved:

- `packages/worker-transport/src/workbook-delta-v3.ts` carries revision tuples through `WorkbookDeltaBatchV3`, including value, style, axis, freeze, and calc sequence fields.
- `packages/worker-transport/src/viewport-patch.ts` already transports viewport cells, styles, rows, columns, freeze panes, metrics, and sheet name.
- `apps/web/src/worker-runtime-viewport.ts` builds `ViewportPatch` from the engine and can emit an empty viewport when the target sheet is missing.
- `apps/web/src/projected-viewport-store.ts` owns projected cell, axis, patch, tile-scene, and sheet-channel state, but still has sheet-name-sensitive paths.
- `packages/grid/src/gridGeometry.ts` defines `GridGeometrySnapshot` with camera, axis indexes, hit testing, cell rects, range rects, editor rects, and header hit testing.
- `packages/grid/src/renderer-v3/typegpu-workbook-backend-v3.ts` already separates TypeGPU tile panes, headers, overlays, atlas state, and tile residency.

The production gap is not that Bilig lacks building blocks. The gap is that authoritative state, projected viewport state, renderer state, assistant context, table metadata, and browser-visible state are not tied together by one audited commit contract.

## Perspective Ideas Worth Applying

The Perspective codebase has several ideas that map cleanly to Bilig:

- **Table/View separation:** Perspective keeps a typed table as source data and exposes `View` objects for specific projections. Bilig should keep the workbook engine as source of truth and expose `WorkbookViewWindow` snapshots for grid, assistant, chart, pivot, audit, and export views.
- **Explicit `ViewWindow`:** Perspective's `ViewWindow` includes row and column bounds, formatting flags, compression options, and request identity. Bilig needs the same explicit viewport read contract, with workbook revisions and sheet identity added.
- **Renderer plugin lifecycle:** Perspective plugins implement lifecycle methods like draw, update, resize, restyle, save, restore, and clear. Bilig should expose a view-plugin lifecycle for future chart, pivot, audit, and assistant rendered-state consumers.
- **Update then render synchronization:** Perspective's viewer update path validates session state before rendering a created view. Bilig needs an apply barrier that validates model state, recalculation, context, render acknowledgement, and invariants before reporting success.
- **Columnar typed reads:** Perspective chart readers use typed arrays during render. Bilig should add typed columnar reads for analytics and charting, while keeping cell editing and formulas in the workbook engine.
- **Stable identity and cleanup:** Perspective tracks request ids and listener lifecycle. Bilig needs stable sheet ids, window ids, render batch ids, and lifecycle cleanup across rename, delete, scroll, and remount.

## Perspective Ideas Not To Copy

These choices should stay out of the Bilig migration:

- Do not replace Bilig's workbook grid with `regular-table`; Bilig needs spreadsheet editing, formulas, selection, fill, formatting, frozen panes, and assistant workflows.
- Do not adopt Perspective's immutable schema rule for workbook cells; spreadsheets need flexible per-cell types.
- Do not replace the formula stack with ExprTK or a Perspective expression model.
- Do not rewrite the React shell into custom elements.
- Do not make Perspective the rendering engine. Use its contracts as architecture reference, not as a runtime dependency.

## Target Architecture

### WorkbookViewWindow

Create a transport-level snapshot type that becomes the common contract for visible grid reads, assistant readback, charts, pivots, audits, and browser verification.

File:

- `packages/worker-transport/src/workbook-view-window.ts`

Shape:

```ts
export interface WorkbookViewWindow {
  version: 1;
  windowId: string;
  documentId: string;
  documentRevision: number;
  calcRevision: number;
  renderRevision: number;
  sheet: {
    id: string;
    name: string;
    ordinal: number;
  };
  viewport: {
    rowStart: number;
    rowEnd: number;
    colStart: number;
    colEnd: number;
  };
  selection: WorkbookViewSelection;
  cells: WorkbookViewCell[];
  styles: WorkbookViewStyleDictionary;
  rows: WorkbookViewAxisSegment[];
  columns: WorkbookViewAxisSegment[];
  freeze: WorkbookViewFreezeState;
  tables: WorkbookViewTableSummary[];
  formulaStatus: WorkbookFormulaStatus;
  renderStatus: WorkbookRenderStatus;
  invariants: WorkbookInvariantStatus;
  source: "authoritative" | "projected" | "rendered";
}
```

Cell entries must separate the values needed by different surfaces:

```ts
export interface WorkbookViewCell {
  row: number;
  col: number;
  address: string;
  snapshot: WorkbookCellSnapshot;
  displayText: string;
  editorText: string;
  copyText: string;
  formatId: string | null;
  styleId: string | null;
}
```

### Apply And Verify Barrier

Every assistant-visible mutation must run through one barrier:

```text
apply mutation
-> recalculate
-> refresh workbook context
-> publish authoritative WorkbookViewWindow
-> sync projected viewport
-> sync TypeGPU tile scene
-> wait for visible render batch acknowledgement
-> read authoritative state
-> read rendered state
-> run formula diagnostics
-> run workbook invariants
-> return verified result
```

The assistant must never report a write as complete if the barrier cannot prove visible state.

### Stable Sheet Identity

All internal transport, table metadata, selection state, viewport patches, and assistant operations must use `sheetId` as primary identity. Sheet names remain labels and URL affordances.

Required behavior:

- Renaming `Sheet3` to `Prepaid Template` updates sheet labels without changing sheet identity.
- Table metadata stores `sheetId` and resolves current sheet name at read time.
- Deleting a sheet removes or invalidates dependent tables, selections, view windows, and tile packets in the same committed revision.
- Reading a missing sheet must return a typed not-found status, not an empty successful viewport.

### View Plugin Contract

Define a small plugin contract inspired by Perspective but shaped for Bilig:

File:

- `packages/workbook-domain/src/workbook-view-plugin.ts`

Contract:

```ts
export interface WorkbookViewPlugin {
  readonly id: string;
  draw(window: WorkbookViewWindow): Promise<void> | void;
  update(window: WorkbookViewWindow): Promise<void> | void;
  resize(bounds: WorkbookViewBounds): Promise<void> | void;
  restyle(theme: WorkbookThemeSnapshot): Promise<void> | void;
  save(): Promise<WorkbookViewPluginState> | WorkbookViewPluginState;
  restore(state: WorkbookViewPluginState): Promise<void> | void;
  clear(): Promise<void> | void;
}
```

The grid remains the first plugin-grade consumer. Charts, pivots, audit trails, and assistant rendered-read tools should consume the same view-window contract.

## Measurable Acceptance Bars

- Deep link `http://localhost:5173/?sheet=Prepaid+Template&cell=E46` renders existing backend data when that sheet has authoritative data.
- If a sheet does not exist, the UI shows a typed missing-sheet state and available sheet recovery actions instead of a blank grid.
- Local mutation to visible acknowledgement p95 is under 200 ms for a viewport-sized edit on the development machine.
- Vertical scroll selection remains exact across at least 100,000 rows and 500 columns in Playwright pointer tests.
- `rename Sheet3 -> Prepaid Template` followed by context, sheets, workbook, tables, and rendered visible reads contains no stale `Sheet3` reference.
- Creating tables after sheet rename cannot produce a table pointing at a missing sheet.
- `write_range` preserves numbers, dates, blank cells, formulas, and strings without converting numeric-looking values into strings.
- `apply_and_verify` returns applied revision, recalculation status, authoritative readback, rendered readback, formula issues, and invariant status.
- Formula diagnostics separate actionable errors from compatibility warnings and fallback notices.
- No successful assistant mutation result is returned before render acknowledgement.

## Implementation Plan

### Phase 0: Baseline Regressions And Instrumentation

- [ ] Add a regression test for blank-visible-grid with existing authoritative data.
  - Create `apps/web/src/__tests__/workbook-visible-state-regression.test.tsx`.
  - Cover `?sheet=Prepaid Template&cell=E46`, engine data present, visible grid must contain expected text.
  - Assert missing sheets produce a typed missing-sheet state, not a successful empty viewport.
- [ ] Extend selection and scroll browser tests.
  - Modify `e2e/tests/web-shell-selection.pw.ts`.
  - Modify `e2e/tests/web-shell-scroll-performance.pw.ts`.
  - Add pointer tests after large vertical scroll and large horizontal scroll.
  - Assert selected address equals the cell under pointer from rendered geometry.
- [ ] Add render commit counters.
  - Modify `apps/web/src/projected-viewport-store.ts`.
  - Modify `packages/grid/src/renderer-v3/typegpu-workbook-backend-v3.ts`.
  - Track last authoritative revision, projected revision, tile-scene revision, and visible render revision.

Commands:

```bash
pnpm exec vitest run apps/web/src/__tests__/workbook-visible-state-regression.test.tsx
pnpm exec playwright test e2e/tests/web-shell-selection.pw.ts e2e/tests/web-shell-scroll-performance.pw.ts
```

### Phase 1: Define WorkbookViewWindow Transport

- [ ] Create `packages/worker-transport/src/workbook-view-window.ts`.
  - Define `WorkbookViewWindow`, `WorkbookViewCell`, `WorkbookFormulaStatus`, `WorkbookRenderStatus`, and `WorkbookInvariantStatus`.
  - Include explicit document, calc, render, sheet, viewport, selection, table, and source metadata.
- [ ] Export the type from `packages/worker-transport/src/index.ts`.
- [ ] Add runtime validation helpers for impossible windows.
  - Missing `sheet.id` fails validation.
  - Empty cells are allowed only when the source range is truly empty or the sheet is missing with a typed status.
  - Table summaries must not reference unknown sheets.
- [ ] Add `packages/worker-transport/src/__tests__/workbook-view-window.test.ts`.

Commands:

```bash
pnpm exec vitest run packages/worker-transport/src/__tests__/workbook-view-window.test.ts
pnpm exec vitest run packages/worker-transport/src/__tests__/viewport-patch.test.ts packages/worker-transport/src/__tests__/workbook-delta-v3.test.ts
```

### Phase 2: Produce Authoritative View Windows

- [ ] Create `apps/web/src/worker-runtime-view-window.ts`.
  - Build `WorkbookViewWindow` from the workbook engine.
  - Include typed cell values, display text, editor text, styles, axis segments, freeze panes, table summaries, formula status, and invariant status.
- [ ] Modify `apps/web/src/worker-runtime-viewport.ts`.
  - Stop treating missing sheets as successful empty snapshots.
  - Attach stable sheet id and revision fields to viewport output.
  - Share cell and style serialization with `worker-runtime-view-window.ts`.
- [ ] Modify `apps/web/src/worker-runtime.ts`.
  - Add a worker message for authoritative view-window reads.
  - Ensure view-window reads happen after pending mutation and recalc work.
- [ ] Modify `apps/web/src/projected-viewport-store.ts`.
  - Cache by `sheetId`, not sheet name.
  - Invalidate by revision and sheet id after structural changes.
  - Reject stale patches whose revision is older than the active document revision.

Commands:

```bash
pnpm exec vitest run apps/web/src/__tests__/projected-viewport-store.test.ts apps/web/src/__tests__/projected-viewport-patch-application.test.ts apps/web/src/__tests__/worker-workbook-app.test.tsx
```

### Phase 3: Add The Visible Commit Barrier

- [ ] Create `apps/web/src/workbook-visible-commit-barrier.ts`.
  - Accept expected document revision, calc revision, sheet id, viewport, and selection.
  - Wait for projected viewport update.
  - Wait for tile-scene update.
  - Wait for renderer visible batch acknowledgement.
  - Return authoritative and rendered `WorkbookViewWindow` snapshots.
- [ ] Add `apps/web/src/__tests__/workbook-visible-commit-barrier.test.ts`.
  - Cover success, stale projected viewport, stale render batch, missing sheet, and formula-error cases.
- [ ] Modify `apps/web/src/use-worker-workbook-app-state.tsx`.
  - Route assistant mutations and direct workbook commands through the barrier when visible verification is requested.
- [ ] Modify assistant context code.
  - `apps/web/src/workbook-agent-context.ts`
  - `apps/bilig/src/codex-app/workbook-agent-tools.ts`
  - `apps/bilig/src/codex-app/workbook-agent-dynamic-tool-handler.ts`
  - `apps/bilig/src/codex-app/workbook-agent-sheet-read-tools.ts`
- [ ] Add `apply_and_verify`.
  - Return applied revision.
  - Return recalculation status.
  - Return authoritative readback.
  - Return rendered readback.
  - Return formula issues grouped by severity and compatibility.
  - Return invariant status.

Commands:

```bash
pnpm exec vitest run apps/web/src/__tests__/workbook-visible-commit-barrier.test.ts apps/web/src/__tests__/workbook-agent-context.test.ts
```

### Phase 4: Migrate Structural State To Stable Sheet Identity

- [ ] Create `apps/web/src/workbook-sheet-identity.ts`.
  - Centralize sheet-id lookup, label resolution, URL sheet-name resolution, and not-found status creation.
- [ ] Modify `apps/web/src/projected-viewport-store.ts`.
  - Replace sheet-name cache keys with sheet-id keys.
  - Keep name-to-id mapping only at URL and display boundaries.
- [ ] Modify `apps/web/src/worker-viewport-tile-store.ts`.
  - Include sheet id in tile keys and invalidation events.
- [ ] Modify `packages/worker-transport/src/workbook-delta-v3.ts`.
  - Require sheet id for structural, value, style, axis, and freeze deltas.
- [ ] Modify structural and table assistant tools.
  - `apps/bilig/src/codex-app/workbook-agent-structural-tools.ts`
  - `apps/bilig/src/codex-app/workbook-agent-object-tools.ts`
  - Tables store `sheetId` and resolve current name at read time.
  - Sheet delete removes or invalidates dependent table metadata atomically.
- [ ] Add invariant tests.
  - Rename `Sheet3` to `Prepaid Template`.
  - Read context, sheets, workbook, and tables.
  - Assert no stale `Sheet3` reference.
  - Create tables after rename.
  - Assert no table points at a missing sheet.

Commands:

```bash
pnpm exec vitest run apps/web/src/__tests__/workbook-agent-context.test.ts apps/web/src/__tests__/worker-workbook-app-model.test.ts apps/web/src/__tests__/worker-workbook-app-model.fuzz.test.ts
```

### Phase 5: Fix Geometry And Selection Correctness

- [ ] Make `GridGeometrySnapshot` the only source for pointer-to-cell conversion.
  - Modify `packages/grid/src/gridGeometry.ts`.
  - Modify `packages/grid/src/useWorkbookGridGeometryRuntime.ts`.
  - Modify `packages/grid/src/useWorkbookGridPointerResolvers.ts`.
  - Modify `packages/grid/src/useWorkbookGridInteractionRuntime.ts`.
  - Modify `packages/grid/src/WorkbookGridSurface.tsx`.
- [ ] Remove duplicate coordinate math from pointer handlers, scroll handlers, editor placement, and overlays.
- [ ] Add a monotonically versioned geometry snapshot.
  - Pointer down, drag, context menu, editor open, and fill handle must use the same snapshot for a gesture.
  - If scroll changes during a gesture, the gesture receives a new snapshot through one controlled transition.
- [ ] Add tests for scrolled selection.
  - `packages/grid/src/__tests__/gridGeometry.test.ts`
  - `packages/grid/src/__tests__/useWorkbookGridPointerResolvers.test.tsx`
  - `packages/grid/src/__tests__/useWorkbookGridGeometryRuntime.test.tsx`
  - `e2e/tests/web-shell-selection.pw.ts`

Commands:

```bash
pnpm exec vitest run packages/grid/src/__tests__/gridGeometry.test.ts packages/grid/src/__tests__/useWorkbookGridPointerResolvers.test.tsx packages/grid/src/__tests__/useWorkbookGridGeometryRuntime.test.tsx
pnpm exec playwright test e2e/tests/web-shell-selection.pw.ts
```

### Phase 6: Add Assistant Semantic Tools

- [ ] Create `apps/bilig/src/codex-app/workbook-agent-view-window-tools.ts`.
  - `read_authoritative_range`
  - `read_rendered_range`
  - `read_rendered_selection`
  - `read_visible_window`
- [ ] Create `apps/bilig/src/codex-app/workbook-agent-semantic-template-tools.ts`.
  - Select by named table.
  - Select by current region.
  - Select by header label.
  - Select by used range.
  - Select by visible viewport.
  - Select relative to anchor labels such as "first data row under header Amount".
- [ ] Create `apps/bilig/src/codex-app/workbook-agent-verified-workflows.ts`.
  - `apply_and_verify`
  - `write_template_and_verify`
  - `format_range_and_verify`
  - `create_table_and_verify`
- [ ] Modify `apps/bilig/src/codex-app/workbook-agent-tools.ts`.
  - Register stable tools with explicit schemas.
  - Expose typed write values for string, number, boolean, date, blank, and formula.
- [ ] Modify `apps/bilig/src/codex-app/workbook-selector-resolver.ts`.
  - Resolve semantic selectors to sheet-id-backed ranges.
  - Return ambiguous selector errors with candidate ranges.
- [ ] Modify `apps/web/src/workbook-agent-tool-output.tsx`.
  - Render verification output compactly.
  - Separate actionable formula errors from compatibility warnings.

Commands:

```bash
pnpm exec vitest run apps/web/src/__tests__/workbook-agent-context.test.ts
pnpm exec vitest run apps/web/src/__tests__/worker-workbook-app.test.tsx
```

### Phase 7: Add Plugin-Grade Analytics Views

- [ ] Create `packages/workbook-domain/src/workbook-view-plugin.ts`.
  - Define draw, update, resize, restyle, save, restore, and clear lifecycle.
- [ ] Create `apps/web/src/workbook-view-plugin-registry.ts`.
  - Register grid as the first view consumer.
  - Keep plugin state serializable.
- [ ] Create `apps/web/src/workbook-columnar-view.ts`.
  - Convert `WorkbookViewWindow` ranges into typed columnar arrays for analytics and charts.
  - Preserve workbook value types and display formatting metadata.
- [ ] Create `apps/web/src/workbook-analytics-view.ts`.
  - Provide aggregation, filter, sort, and profile summaries over view windows.
  - Use typed columnar reads for chart and pivot foundations.
- [ ] Keep editing and formula semantics in the workbook engine.

Commands:

```bash
pnpm exec vitest run apps/web/src/__tests__/workbook-analytics-view.test.ts apps/web/src/__tests__/workbook-columnar-view.test.ts
```

### Phase 8: CI And Production Gates

- [ ] Run targeted transport tests.

```bash
pnpm exec vitest run packages/worker-transport/src/__tests__/workbook-view-window.test.ts packages/worker-transport/src/__tests__/viewport-patch.test.ts packages/worker-transport/src/__tests__/workbook-delta-v3.test.ts
```

- [ ] Run targeted web workbook tests.

```bash
pnpm exec vitest run apps/web/src/__tests__/projected-viewport-store.test.ts apps/web/src/__tests__/workbook-agent-context.test.ts apps/web/src/__tests__/worker-workbook-app.test.tsx
```

- [ ] Run targeted grid geometry tests.

```bash
pnpm exec vitest run packages/grid/src/__tests__/gridGeometry.test.ts packages/grid/src/__tests__/useWorkbookGridPointerResolvers.test.tsx packages/grid/src/__tests__/useWorkbookGridGeometryRuntime.test.tsx
```

- [ ] Run browser workbook tests.

```bash
pnpm exec playwright test e2e/tests/web-shell-selection.pw.ts e2e/tests/web-shell-scroll-performance.pw.ts e2e/tests/web-shell-typegpu.pw.ts
```

- [ ] Run repository gates.

```bash
pnpm typecheck
pnpm lint
pnpm run ci
```

## Browser Validation Protocol

Use the local app, not tests alone:

- Open `http://localhost:5173/?sheet=Prepaid+Template&cell=E46`.
- Confirm existing workbook data renders in the grid and formula bar.
- Scroll vertically to row 138 and click several cells; selected addresses must match visible cells.
- Scroll horizontally and click cells around columns beyond the initial viewport.
- Rename a sheet, create a table, and verify visible tabs, context, tables, and assistant readback all use the new sheet label.
- Use assistant `apply_and_verify` to write numeric values, date values, blanks, strings, and formulas.
- Confirm rendered readback matches authoritative readback.
- Confirm formula diagnostics show actionable errors separately from compatibility warnings.

## Risk Controls

- Keep `ViewportPatch` compatibility while introducing `WorkbookViewWindow`; migrate consumers incrementally.
- Add revision checks before deleting old sheet-name cache paths.
- Use sheet id internally and resolve sheet name only at URL, tab, and display boundaries.
- Keep TypeGPU V3 as the only primary renderer path.
- Treat empty viewport for existing data as a failing state.
- Keep tests focused and colocated with each migration step.
- Commit after large completed phases and run full CI on the committed tree before pushing.

## Definition Of Done

- Bilig can prove a workbook mutation from operation intent through visible rendered state.
- Assistant tools can read and verify active sheet, selection, authoritative state, rendered state, formula status, and invariants.
- Structural edits cannot leave stale context or orphan table metadata.
- Selection remains exact after scroll and viewport virtualization.
- Existing backend data cannot load as a blank successful grid.
- Analytics and future plugin surfaces can consume stable view windows without reaching into renderer internals.
- Targeted Vitest and Playwright suites pass.
- `pnpm typecheck`, `pnpm lint`, and `pnpm run ci` pass before merge.
