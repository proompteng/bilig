# Workbook Assistant Tools Production Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Make the embedded workbook assistant tools good enough that the assistant can use them on a real workbook, make and verify a safe edit, undo it, and honestly say it is happy with the tools.

**Architecture:** Preserve the current Bilig workbook engine, app-server transport, React workbook shell, and TypeGPU renderer. Add a first-class transaction loop around workbook tools: observe precisely, preview and diff, apply atomically where the engine supports it, verify authoritative and rendered state, and undo or report a precise incomplete state. Shared receipt, rendered-readback, large-range, selection, formula-diagnostic, and semantic-targeting contracts must be enforced by code and tests.

**Tech Stack:** TypeScript, ESM, pnpm monorepo, `apps/bilig`, `apps/web`, `packages/grid`, `packages/worker-transport`, Vitest, Playwright, local headed app at `http://localhost:5173`.

---

## Working Directory

Run all implementation from:

```bash
cd /Users/gregkonush/github.com/bilig
```

The main app-specific seams live under:

```bash
cd /Users/gregkonush/github.com/bilig/apps/bilig
```

Do not use temporary clones, detached worktrees, or outside folders for implementation. Stay on `main` unless the user explicitly requests another branch.

## Production Objective

The workbook assistant must stop treating broad tool coverage as reliability. The product bar is confidence after actual use:

- The assistant can inspect a real workbook.
- The assistant can preview a small safe edit.
- The assistant can apply the edit with deterministic status.
- The assistant can verify authoritative workbook state.
- The assistant can verify rendered or browser-visible state.
- The assistant can undo or restore the prior state.
- The assistant can explain whether the tools were sufficient without hiding stale rendered state, ambiguous apply status, selection uncertainty, or missing preview/apply/verify/undo.

## Current Known Problems To Fix

- Tool coverage is broad, but confidence remains weak.
- Authoritative workbook reads are useful, but rendered and visual verification can be stale, viewport-bound, truncated, or unavailable.
- The workflow is not a clean transaction loop: observe precisely, analyze, preview and diff, apply atomically, verify authoritative and rendered state, and undo when needed.
- Mutation tool receipts can be ambiguous and must not let the assistant treat queued or staged work as completed.
- Active sheet, selection, visible viewport, and rendered readback need stronger post-action confirmation.
- Formula diagnostics are too low-level when complex sheets wobble.
- The assistant needs actionable formula explanations, precedents, dependents, stale or recalculation status, and exact problematic formulas and ranges.
- Large workbook snapshots truncate too easily without a clean chunking and continuation path.
- Higher-level workbook-native actions are missing or too awkward: autofit, merge, unmerge, semantic table and range targeting, polished report and template operations, and clear visual verification helpers.

## Hard Acceptance Bar

The implementation is not complete until all of these are true:

1. A real headed local workbook assistant flow runs through the embedded assistant and asks a question equivalent to:

   ```text
   Use your workbook tools on this workbook, make a small safe edit, verify it, undo or restore it, and then tell me honestly: are you happy with the tools now?
   ```

   The captured answer must be materially "yes" and must not list remaining blockers around stale rendered state, unclear apply status, selection confidence, or missing preview/apply/verify/undo.

2. A mutation tool never returns ambiguous success. It clearly reports:

   - applied, staged, or queued status
   - head revision before and after
   - affected ranges
   - whether authoritative readback matched
   - whether rendered readback matched, or why it could not be proven
   - whether undo or rollback data exists

3. If rendered readback is stale or unavailable, the loop either refreshes, navigates, and captures it automatically, or marks verification incomplete with a concrete reason. It must never silently pass by substituting authoritative state for visual proof.

4. Large ranges have a chunked read and verify path so the assistant can verify an entire used range without losing confidence to truncation.

5. Preview, apply, verify, and undo are first-class workflow operations with tests, not an informal sequence of unrelated calls.

6. Focused tests, typecheck, lint, and full CI pass before completion:

   ```bash
   pnpm typecheck
   pnpm lint
   pnpm run ci
   ```

## Files To Inspect First

Start by reading the real current implementation in this checkout:

- `apps/bilig/src/codex-app/workbook-agent-service.ts`
- `apps/bilig/src/codex-app/workbook-agent-dynamic-tool-handler.ts`
- `apps/bilig/src/codex-app/workbook-agent-tools.ts`
- `apps/bilig/src/codex-app/workbook-agent-session-model.ts`
- `apps/web/src/use-workbook-agent-pane.tsx`
- `apps/web/src/use-workbook-app-panels.tsx`
- `apps/web/src/workbook-agent-context.ts`
- `apps/web/src/projected-viewport-store.ts`
- `apps/web/src/worker-runtime-viewport.ts`
- `packages/grid/src/gridGeometry.ts`
- `packages/grid/src/useWorkbookGridPointerResolvers.ts`
- workbook tool registry, schema, and test files under `apps/bilig/src/codex-app`
- browser and headed validation scripts under `e2e`, `scripts`, and `output/playwright`

Do not assume the failure is prompt wording. Fix the code path that produces ambiguous receipts, stale rendered reads, truncated verification, or unconfirmed browser state.

## Implementation Requirements

- Add or tighten focused tests first where practical.
- Preserve existing useful tools.
- Do not remove capabilities to make tests easier.
- Do not paper over stale rendered state by falling back to authoritative state and calling it visual verification.
- If a capability is genuinely blocked by current architecture, implement the smallest infrastructure needed, or make the user-facing receipt explicit about the exact blocking seam.
- Avoid oversized source files. If a source file approaches roughly 1000 lines, split into focused modules before adding more behavior.
- Use strict TypeScript, ESM imports, and explicit `.js` suffixes where local code uses them.
- Avoid `any`.
- Avoid floating promises.
- Keep receipts deterministic and machine-readable.

## Shared Receipt Contract

All mutation tools must return a consistent receipt shape, even if individual tools add domain-specific fields:

```ts
export interface WorkbookToolMutationReceipt {
  toolName: string;
  status: "applied" | "staged" | "queued" | "failed" | "verification_incomplete";
  revision: {
    before: number | null;
    after: number | null;
  };
  affectedRanges: Array<{
    sheetId: string;
    sheetName: string;
    range: string;
    kind: "values" | "formulas" | "formats" | "tables" | "objects" | "selection" | "sheet";
  }>;
  authoritativeReadback: {
    requested: boolean;
    matched: boolean | null;
    mismatches: WorkbookVerificationMismatch[];
  };
  renderedReadback: {
    requested: boolean;
    matched: boolean | null;
    stale: boolean;
    capturedRange: string | null;
    capturedRevision: number | null;
    capturedBatchId: string | null;
    missingCells: string[];
    mismatches: WorkbookVerificationMismatch[];
    incompleteReason: string | null;
    nextChunk: WorkbookVerificationChunk | null;
  };
  undo: {
    available: boolean;
    token: string | null;
    reasonUnavailable: string | null;
  };
  warnings: string[];
}
```

Any receipt with `status: "queued"` or `status: "staged"` is not complete. Assistant-facing tool output must say so plainly.

## First-Class Workflow Contract

Add one coherent workflow that wraps existing low-level operations:

```text
observe workbook state
-> build preview and diff
-> apply mutation atomically at a workbook revision where feasible
-> recalculate
-> refresh context
-> verify authoritative state
-> verify rendered state
-> return receipt with undo token
-> undo or restore when requested
```

Required tool-level workflows:

- `preview_workbook_mutation`
- `apply_workbook_mutation`
- `verify_workbook_mutation`
- `undo_workbook_mutation`
- `apply_and_verify_workbook_mutation`

If existing tool names already cover these concepts, keep the stable public name and route through the shared workflow internally.

## Task 1: Baseline Current Behavior

**Files:**

- Read: `apps/bilig/src/codex-app/workbook-agent-service.ts`
- Read: `apps/bilig/src/codex-app/workbook-agent-dynamic-tool-handler.ts`
- Read: `apps/bilig/src/codex-app/workbook-agent-tools.ts`
- Read: `apps/bilig/src/codex-app/workbook-agent-session-model.ts`
- Read: `apps/web/src/use-workbook-agent-pane.tsx`
- Read: `apps/web/src/use-workbook-app-panels.tsx`
- Read: `apps/web/src/workbook-agent-context.ts`
- Read: `apps/web/src/projected-viewport-store.ts`
- Read: `apps/web/src/worker-runtime-viewport.ts`
- Read: `e2e/tests`
- Write evidence under: `output/playwright`

- [ ] Run the smallest workbook-agent test slice first.

  ```bash
  pnpm exec vitest run apps/bilig/src/codex-app
  ```

- [ ] Start or reuse the local app.

  ```bash
  pnpm dev:web-local
  ```

- [ ] Reproduce the weak loop in a disposable workbook or document.
  - Inspect workbook.
  - Make a safe edit.
  - Capture tool receipts.
  - Capture whether the visible grid reflects the edit.
  - Capture whether any rendered readback is stale, truncated, viewport-bound, or unavailable.

- [ ] Save evidence under the repo's existing Playwright output convention.

  Example path:

  ```text
  output/playwright/workbook-assistant-tools-baseline-<timestamp>.json
  ```

## Task 2: Transaction Control Loop

**Files:**

- Modify: `apps/bilig/src/codex-app/workbook-agent-service.ts`
- Modify: `apps/bilig/src/codex-app/workbook-agent-dynamic-tool-handler.ts`
- Modify: `apps/bilig/src/codex-app/workbook-agent-tools.ts`
- Modify: `apps/bilig/src/codex-app/workbook-agent-session-model.ts`
- Create: `apps/bilig/src/codex-app/workbook-agent-transaction-workflow.ts`
- Create: `apps/bilig/src/codex-app/workbook-agent-mutation-receipt.ts`
- Test: `apps/bilig/src/codex-app/__tests__/workbook-agent-transaction-workflow.test.ts`

- [ ] Add failing tests for preview, apply, verify, and undo receipts.
- [ ] Implement `WorkbookToolMutationReceipt`.
- [ ] Implement preview that describes exact intended workbook mutations and affected ranges.
- [ ] Implement apply that records revision before and after.
- [ ] Implement verify that compares authoritative values, formulas, and styles.
- [ ] Implement undo token capture before apply when data needed for restoration is available.
- [ ] Implement undo that restores prior state or returns a deterministic unavailable reason before apply.
- [ ] Route existing mutation tools through the shared receipt path.
- [ ] Remove or quarantine any path where `queuedForTurnApply: true` can be interpreted as completed work.

Commands:

```bash
pnpm exec vitest run apps/bilig/src/codex-app/__tests__/workbook-agent-transaction-workflow.test.ts
```

## Task 3: Rendered Confidence

**Files:**

- Modify: `apps/web/src/workbook-agent-context.ts`
- Modify: `apps/web/src/projected-viewport-store.ts`
- Modify: `apps/web/src/worker-runtime-viewport.ts`
- Create: `apps/web/src/workbook-rendered-readback.ts`
- Create: `apps/web/src/__tests__/workbook-rendered-readback.test.ts`
- Modify: `apps/bilig/src/codex-app/workbook-agent-tools.ts`
- Modify: `apps/bilig/src/codex-app/workbook-agent-dynamic-tool-handler.ts`

- [ ] Make `read_rendered_selection`, `read_rendered_range`, and visible-range helpers revision-aware.
- [ ] Include captured revision, batch id, timestamp, requested range, captured range, stale flag, missing cells, mismatches, truncation, and next chunk in rendered readback output.
- [ ] Detect stale `batchId`, captured timestamp, and revision mismatch.
- [ ] For ranges outside the cached viewport, add automatic refresh, navigation, and capture where the browser architecture supports it.
- [ ] If automatic visual capture cannot prove a range, return `verification_incomplete` with a concrete reason.
- [ ] Add tests for stale rendered readback, viewport miss, successful refresh, and incomplete visual proof.

Commands:

```bash
pnpm exec vitest run apps/web/src/__tests__/workbook-rendered-readback.test.ts
pnpm exec vitest run apps/web/src/__tests__/workbook-agent-context.test.ts
```

## Task 4: Selection And Active Sheet Confirmation

**Files:**

- Modify: `apps/bilig/src/codex-app/workbook-agent-tools.ts`
- Modify: `apps/bilig/src/codex-app/workbook-agent-dynamic-tool-handler.ts`
- Modify: `apps/web/src/use-workbook-agent-pane.tsx`
- Modify: `apps/web/src/use-workbook-app-panels.tsx`
- Modify: `packages/grid/src/gridGeometry.ts`
- Modify: `packages/grid/src/useWorkbookGridPointerResolvers.ts`
- Test: `apps/web/src/__tests__/selection-persistence.test.ts`
- Test: `apps/web/src/__tests__/use-workbook-selection-actions.test.ts`
- Test: `e2e/tests/web-shell-selection.pw.ts`

- [ ] Make `set_active_sheet` verify workbook model state and browser-rendered state.
- [ ] Make `set_selection` verify workbook model state and browser-rendered state.
- [ ] Return a specific timeout if browser confirmation does not arrive.
- [ ] Do not treat browser timeout as success.
- [ ] Add sheet switch, selection, and rendered readback tests.
- [ ] Add scrolled selection tests where pointer coordinates are resolved only through current geometry snapshots.

Commands:

```bash
pnpm exec vitest run apps/web/src/__tests__/selection-persistence.test.ts apps/web/src/__tests__/use-workbook-selection-actions.test.ts
pnpm exec playwright test e2e/tests/web-shell-selection.pw.ts
```

## Task 5: Mutation Receipt Unification

**Files:**

- Modify: `apps/bilig/src/codex-app/workbook-agent-tools.ts`
- Modify: `apps/bilig/src/codex-app/workbook-agent-dynamic-tool-handler.ts`
- Modify: mutation-specific tool modules under `apps/bilig/src/codex-app`
- Test: `apps/bilig/src/codex-app/__tests__/workbook-agent-tools.test.ts`
- Test: `apps/bilig/src/codex-app/__tests__/workbook-agent-dynamic-tool-handler.test.ts`

- [ ] Route write tools through the shared receipt.
- [ ] Route format tools through the shared receipt.
- [ ] Route table tools through the shared receipt.
- [ ] Route chart tools through the shared receipt.
- [ ] Route protection tools through the shared receipt.
- [ ] Route validation tools through the shared receipt.
- [ ] Route comment tools through the shared receipt.
- [ ] Route shape and image tools through the shared receipt.
- [ ] Assert every mutation receipt includes revision, applied state, affected ranges, verification summary, warnings, and undo status.

Commands:

```bash
pnpm exec vitest run apps/bilig/src/codex-app/__tests__/workbook-agent-tools.test.ts apps/bilig/src/codex-app/__tests__/workbook-agent-dynamic-tool-handler.test.ts
```

## Task 6: Formula Debugging

**Files:**

- Modify: formula diagnostic tool modules under `apps/bilig/src/codex-app`
- Modify: `apps/bilig/src/codex-app/workbook-agent-tools.ts`
- Create: `apps/bilig/src/codex-app/workbook-agent-formula-diagnostics.ts`
- Test: `apps/bilig/src/codex-app/__tests__/workbook-agent-formula-diagnostics.test.ts`

- [ ] Improve `find_formula_issues`, `inspect_cell`, `trace_dependencies`, or wrappers so formula errors are actionable.
- [ ] Include exact formula text.
- [ ] Include displayed value.
- [ ] Include current error.
- [ ] Include direct precedents.
- [ ] Include direct dependents.
- [ ] Include stale or recalculation status.
- [ ] Include unsupported function, unsupported operator, and compatibility details.
- [ ] Include suggested next inspection ranges.
- [ ] Add tests for broken formulas.
- [ ] Add tests for hidden dependencies.
- [ ] Add tests for inconsistent copied formulas.
- [ ] Add tests for unsupported formulas.

Commands:

```bash
pnpm exec vitest run apps/bilig/src/codex-app/__tests__/workbook-agent-formula-diagnostics.test.ts
```

## Task 7: Large Ranges And Truncation

**Files:**

- Create: `apps/bilig/src/codex-app/workbook-agent-range-chunks.ts`
- Modify: `apps/bilig/src/codex-app/workbook-agent-tools.ts`
- Modify: `apps/bilig/src/codex-app/workbook-agent-dynamic-tool-handler.ts`
- Test: `apps/bilig/src/codex-app/__tests__/workbook-agent-range-chunks.test.ts`

- [ ] Add chunked read helpers for ranges larger than the current single-call display limit.
- [ ] Add chunked verify helpers for ranges larger than rendered viewport limits.
- [ ] Include truncation status in receipts.
- [ ] Include continuation token or next chunk plan in receipts.
- [ ] Verify an entire used range across multiple chunks.
- [ ] Add tests that exceed the current visible or 400-cell-style limits.

Commands:

```bash
pnpm exec vitest run apps/bilig/src/codex-app/__tests__/workbook-agent-range-chunks.test.ts
```

## Task 8: Higher-Level Workbook Helpers

**Files:**

- Create: `apps/bilig/src/codex-app/workbook-agent-semantic-targeting.ts`
- Create: `apps/bilig/src/codex-app/workbook-agent-polished-operations.ts`
- Modify: `apps/bilig/src/codex-app/workbook-agent-tools.ts`
- Test: `apps/bilig/src/codex-app/__tests__/workbook-agent-semantic-targeting.test.ts`
- Test: `apps/bilig/src/codex-app/__tests__/workbook-agent-polished-operations.test.ts`

- [ ] Add autofit columns and rows, or content-based best-effort sizing if exact autofit is not supported.
- [ ] Add merge and unmerge if absent.
- [ ] Add semantic targeting by table name.
- [ ] Add semantic targeting by header.
- [ ] Add semantic targeting by visible range.
- [ ] Add semantic targeting by current region.
- [ ] Add polished report and template helper operations for common financial workbook tasks.
- [ ] Add tests for each helper.

Commands:

```bash
pnpm exec vitest run apps/bilig/src/codex-app/__tests__/workbook-agent-semantic-targeting.test.ts apps/bilig/src/codex-app/__tests__/workbook-agent-polished-operations.test.ts
```

## Task 9: Headed Validation

**Files:**

- Create or modify a headed validation script under `scripts` or `e2e` following existing conventions.
- Save output under: `output/playwright`

- [ ] Use the actual local workbook UI.
- [ ] Run the embedded assistant flow:
  - inspect workbook
  - make a safe edit
  - preview, apply, and verify
  - verify rendered state
  - undo or restore
  - answer whether it is happy with the tools
- [ ] Save prompt, tool receipts, rendered proof, undo proof, and final assistant answer.
- [ ] If the final answer is not materially positive, keep fixing the blockers it names.

Expected evidence path format:

```text
output/playwright/workbook-assistant-tools-headed-validation-<timestamp>.json
```

## Task 10: Final Verification And Commit

**Files:**

- Commit all implementation and test files changed for this plan.

- [ ] Run relevant focused tests.

  ```bash
  pnpm exec vitest run apps/bilig/src/codex-app apps/web/src/__tests__/workbook-agent-context.test.ts
  pnpm exec vitest run apps/web/src/__tests__/selection-persistence.test.ts apps/web/src/__tests__/use-workbook-selection-actions.test.ts
  pnpm exec playwright test e2e/tests/web-shell-selection.pw.ts
  ```

- [ ] Run typecheck.

  ```bash
  pnpm typecheck
  ```

- [ ] Run lint.

  ```bash
  pnpm lint
  ```

- [ ] Run full CI.

  ```bash
  pnpm run ci
  ```

- [ ] Fix any real CI failure.
- [ ] Commit to `main` with a focused Conventional Commit.

  ```bash
  git status --short
  git add apps/bilig apps/web packages/grid packages/worker-transport docs output/playwright
  git commit -m "feat(workbook-agent): verify assistant tool transactions"
  ```

## Final Response Requirements For The Implementing Agent

The final response after implementation must include:

- Summary of problems fixed.
- Files changed.
- Tests and commands run with pass or fail status.
- Headed validation evidence path.
- The embedded assistant's final satisfaction answer.
- Remaining limitations only when they are truly architectural and not fixable in this pass.
- Whether full `pnpm run ci` passed.

## Non-Negotiable Product Standard

The implementation must earn confidence from the product behavior, not from optimistic receipts. A tool that only stages or queues work has not completed the workbook mutation. A rendered read that is stale, viewport-bound, or missing is not visual proof. A large range that truncates without continuation is not verified. The embedded assistant should be able to say it is happy with the tools because it has observed, edited, verified, rendered, and undone a real workbook flow successfully.
