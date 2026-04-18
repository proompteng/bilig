# Workbook Redo Rewrite Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Replace the flaky redo inference path with one authoritative actor-history model so redo remains correct across multi-step undo chains, redo-after-redo, and redo invalidation after new edits.

**Architecture:** Stop inferring redo eligibility from ad hoc “latest history row” predicates in separate client and server code paths. Instead, add one shared pure reducer that derives an actor’s undo/redo stacks from authoritative `workbook_change` rows in chronological order, then reuse that reducer in both the web app and the server mutator path.

**Tech Stack:** TypeScript, Vitest, Playwright, Fastify/Zero sync, React hooks, shared `packages/zero-sync`.

---

## File Structure

### New files

- `packages/zero-sync/src/workbook-history-state.ts`
  Pure shared reducer for deriving per-actor undo/redo stacks from authoritative workbook change rows.
- `packages/zero-sync/src/__tests__/workbook-history-state.test.ts`
  Shared reducer regression coverage for multi-undo/multi-redo and branch invalidation.
- `apps/bilig/src/zero/__tests__/workbook-history-selector.test.ts`
  Server-side selector tests proving the mutator path picks the same redo target as the shared reducer.

### Modified files

- `packages/zero-sync/src/index.ts`
  Export the shared history reducer types/functions.
- `apps/web/src/use-workbook-changes.ts`
  Return normalized raw rows plus rendered entries so history state is derived from authoritative rows, not UI-decorated entries.
- `apps/web/src/workbook-changes-model.ts`
  Replace heuristic `selectWorkbookHistoryState` logic with the shared reducer output.
- `apps/web/src/use-workbook-changes-pane.tsx`
  Drive undo/redo enablement from the shared history cursor.
- `apps/web/src/__tests__/workbook-changes.test.tsx`
  Add the reproduced failure: redo remains available after the first redo in a two-step undo chain.
- `apps/bilig/src/zero/workbook-change-store.ts`
  Replace `loadLatestUndoableWorkbookChange` / `loadLatestRedoableWorkbookChange` heuristics with an actor-history selector powered by the shared reducer.
- `apps/bilig/src/zero/server-mutators.ts`
  Route `workbook.undoLatestChange` / `workbook.redoLatestChange` through the new selector.
- `e2e/tests/web-shell.pw.ts`
  Add browser regressions for multi-step redo and branch invalidation after a fresh edit.

---

### Task 1: Capture the Real Failure Modes in Tests

**Files:**
- Modify: `apps/web/src/__tests__/workbook-changes.test.tsx`
- Create: `packages/zero-sync/src/__tests__/workbook-history-state.test.ts`
- Create: `apps/bilig/src/zero/__tests__/workbook-history-selector.test.ts`

- [ ] **Step 1: Add a failing web regression for two-step redo**

```tsx
it('keeps redo enabled after the first redo in a two-step undo chain', async () => {
  const changes = createMockZeroChangeHarness([
    {
      revision: 15,
      actorUserId: 'alex@example.com',
      clientMutationId: 'mutation-15',
      eventKind: 'redoChange',
      summary: 'Redid r14: Updated Sheet1!A1',
      sheetId: 1,
      sheetName: 'Sheet1',
      anchorAddress: 'A1',
      rangeJson: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' },
      undoBundleJson: {
        kind: 'engineOps',
        ops: [{ kind: 'clearCell', sheetName: 'Sheet1', address: 'A1' }],
      },
      revertedByRevision: null,
      revertsRevision: 14,
      createdAt: Date.parse('2026-04-18T08:15:00.000Z'),
    },
    {
      revision: 14,
      actorUserId: 'alex@example.com',
      clientMutationId: 'mutation-14',
      eventKind: 'revertChange',
      summary: 'Reverted r11: Updated Sheet1!A1',
      sheetId: 1,
      sheetName: 'Sheet1',
      anchorAddress: 'A1',
      rangeJson: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' },
      undoBundleJson: {
        kind: 'engineOps',
        ops: [{ kind: 'setCellValue', sheetName: 'Sheet1', address: 'A1', value: 'a1' }],
      },
      revertedByRevision: 15,
      revertsRevision: 11,
      createdAt: Date.parse('2026-04-18T08:14:00.000Z'),
    },
    {
      revision: 13,
      actorUserId: 'alex@example.com',
      clientMutationId: 'mutation-13',
      eventKind: 'revertChange',
      summary: 'Reverted r12: Updated Sheet1!B1',
      sheetId: 1,
      sheetName: 'Sheet1',
      anchorAddress: 'B1',
      rangeJson: { sheetName: 'Sheet1', startAddress: 'B1', endAddress: 'B1' },
      undoBundleJson: {
        kind: 'engineOps',
        ops: [{ kind: 'setCellValue', sheetName: 'Sheet1', address: 'B1', value: 'b1' }],
      },
      revertedByRevision: null,
      revertsRevision: 12,
      createdAt: Date.parse('2026-04-18T08:13:00.000Z'),
    },
    {
      revision: 12,
      actorUserId: 'alex@example.com',
      clientMutationId: 'mutation-12',
      eventKind: 'setCellValue',
      summary: 'Updated Sheet1!B1',
      sheetId: 1,
      sheetName: 'Sheet1',
      anchorAddress: 'B1',
      rangeJson: { sheetName: 'Sheet1', startAddress: 'B1', endAddress: 'B1' },
      undoBundleJson: {
        kind: 'engineOps',
        ops: [{ kind: 'clearCell', sheetName: 'Sheet1', address: 'B1' }],
      },
      revertedByRevision: 13,
      revertsRevision: null,
      createdAt: Date.parse('2026-04-18T08:12:00.000Z'),
    },
    {
      revision: 11,
      actorUserId: 'alex@example.com',
      clientMutationId: 'mutation-11',
      eventKind: 'setCellValue',
      summary: 'Updated Sheet1!A1',
      sheetId: 1,
      sheetName: 'Sheet1',
      anchorAddress: 'A1',
      rangeJson: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' },
      undoBundleJson: {
        kind: 'engineOps',
        ops: [{ kind: 'clearCell', sheetName: 'Sheet1', address: 'A1' }],
      },
      revertedByRevision: 14,
      revertsRevision: null,
      createdAt: Date.parse('2026-04-18T08:11:00.000Z'),
    },
  ]);

  const host = document.createElement('div');
  document.body.appendChild(host);
  const root = createRoot(host);

  await act(async () => {
    root.render(
      <ChangesHarness
        currentUserId="alex@example.com"
        documentId="doc-1"
        enabled
        onJump={() => {}}
        sheetNames={['Sheet1']}
        zero={changes.zero}
      />,
    );
  });

  expect(host.querySelector("[data-testid='workbook-can-redo']")?.textContent).toBe('true');

  await act(async () => {
    root.unmount();
  });
});
```

- [ ] **Step 2: Run the web test and verify it fails for the right reason**

Run:

```bash
pnpm exec vitest --run apps/web/src/__tests__/workbook-changes.test.tsx -t "keeps redo enabled after the first redo in a two-step undo chain"
```

Expected: FAIL because `selectWorkbookHistoryState()` returns `canRedo: false` once a `redoChange` row is the latest own history action.

- [ ] **Step 3: Add a failing shared reducer spec that defines the correct stack behavior**

```ts
import { describe, expect, it } from 'vitest';
import { deriveWorkbookActorHistoryState } from '../workbook-history-state.js';

describe('deriveWorkbookActorHistoryState', () => {
  it('preserves older redo entries after a newer redo is applied', () => {
    const state = deriveWorkbookActorHistoryState({
      actorUserId: 'alex@example.com',
      rows: [
        { revision: 11, actorUserId: 'alex@example.com', eventKind: 'setCellValue', undoBundleJson: { kind: 'engineOps', ops: [] }, revertedByRevision: 14, revertsRevision: null },
        { revision: 12, actorUserId: 'alex@example.com', eventKind: 'setCellValue', undoBundleJson: { kind: 'engineOps', ops: [] }, revertedByRevision: 13, revertsRevision: null },
        { revision: 13, actorUserId: 'alex@example.com', eventKind: 'revertChange', undoBundleJson: { kind: 'engineOps', ops: [] }, revertedByRevision: null, revertsRevision: 12 },
        { revision: 14, actorUserId: 'alex@example.com', eventKind: 'revertChange', undoBundleJson: { kind: 'engineOps', ops: [] }, revertedByRevision: 15, revertsRevision: 11 },
        { revision: 15, actorUserId: 'alex@example.com', eventKind: 'redoChange', undoBundleJson: { kind: 'engineOps', ops: [] }, revertedByRevision: null, revertsRevision: 14 },
      ],
    });

    expect(state.undoRevision).toBe(15);
    expect(state.redoRevision).toBe(13);
    expect(state.canUndo).toBe(true);
    expect(state.canRedo).toBe(true);
  });
});
```

- [ ] **Step 4: Run the shared reducer test and verify it fails**

Run:

```bash
pnpm exec vitest --run packages/zero-sync/src/__tests__/workbook-history-state.test.ts
```

Expected: FAIL because `workbook-history-state.ts` does not exist yet.

- [ ] **Step 5: Add a failing server selector spec for branch invalidation**

```ts
import { describe, expect, it } from 'vitest';
import { selectRedoableWorkbookChangeRevision } from '../workbook-history-selector.js';

describe('selectRedoableWorkbookChangeRevision', () => {
  it('clears redo after a fresh authored change is appended after an undo', () => {
    const revision = selectRedoableWorkbookChangeRevision({
      actorUserId: 'alex@example.com',
      rows: [
        { revision: 21, actorUserId: 'alex@example.com', eventKind: 'setCellValue', undoBundle: { kind: 'engineOps', ops: [] }, revertedByRevision: 22, revertsRevision: null },
        { revision: 22, actorUserId: 'alex@example.com', eventKind: 'revertChange', undoBundle: { kind: 'engineOps', ops: [] }, revertedByRevision: null, revertsRevision: 21 },
        { revision: 23, actorUserId: 'alex@example.com', eventKind: 'setCellValue', undoBundle: { kind: 'engineOps', ops: [] }, revertedByRevision: null, revertsRevision: null },
      ],
    });

    expect(revision).toBeNull();
  });
});
```

- [ ] **Step 6: Run the server selector test and verify it fails**

Run:

```bash
pnpm exec vitest --run apps/bilig/src/zero/__tests__/workbook-history-selector.test.ts
```

Expected: FAIL because the selector module does not exist yet.

---

### Task 2: Build the Shared Redo/Undo Reducer From Scratch

**Files:**
- Create: `packages/zero-sync/src/workbook-history-state.ts`
- Modify: `packages/zero-sync/src/index.ts`
- Test: `packages/zero-sync/src/__tests__/workbook-history-state.test.ts`

- [ ] **Step 1: Implement the shared actor-history types**

```ts
import type { WorkbookChangeUndoBundle } from './workbook-events.js';

export interface WorkbookHistoryStateRow {
  readonly revision: number;
  readonly actorUserId: string;
  readonly eventKind: string;
  readonly undoBundleJson: WorkbookChangeUndoBundle | null;
  readonly revertedByRevision: number | null;
  readonly revertsRevision: number | null;
}

export interface WorkbookActorHistoryState {
  readonly canUndo: boolean;
  readonly canRedo: boolean;
  readonly undoRevision: number | null;
  readonly redoRevision: number | null;
  readonly undoStack: readonly number[];
  readonly redoStack: readonly number[];
}
```

- [ ] **Step 2: Implement the reducer that derives stacks chronologically**

```ts
export function deriveWorkbookActorHistoryState(input: {
  readonly actorUserId: string;
  readonly rows: readonly WorkbookHistoryStateRow[];
}): WorkbookActorHistoryState {
  const ownRows = [...input.rows]
    .filter((row) => row.actorUserId === input.actorUserId && row.undoBundleJson !== null)
    .sort((left, right) => left.revision - right.revision);

  let undoStack: number[] = [];
  let redoStack: number[] = [];

  for (const row of ownRows) {
    if (row.eventKind === 'revertChange') {
      if (row.revertsRevision !== null) {
        undoStack = undoStack.filter((revision) => revision !== row.revertsRevision);
        redoStack = redoStack.filter((revision) => revision !== row.revision);
        if (row.revertedByRevision === null) {
          redoStack.push(row.revision);
        }
      }
      continue;
    }

    if (row.eventKind === 'redoChange') {
      if (row.revertsRevision !== null) {
        redoStack = redoStack.filter((revision) => revision !== row.revertsRevision);
        undoStack = undoStack.filter((revision) => revision !== row.revision);
        if (row.revertedByRevision === null) {
          undoStack.push(row.revision);
        }
      }
      continue;
    }

    undoStack = undoStack.filter((revision) => revision !== row.revision);
    undoStack.push(row.revision);
    redoStack = [];
  }

  return {
    canUndo: undoStack.length > 0,
    canRedo: redoStack.length > 0,
    undoRevision: undoStack.at(-1) ?? null,
    redoRevision: redoStack.at(-1) ?? null,
    undoStack,
    redoStack,
  };
}
```

- [ ] **Step 3: Export the reducer**

```ts
export {
  deriveWorkbookActorHistoryState,
  type WorkbookActorHistoryState,
  type WorkbookHistoryStateRow,
} from './workbook-history-state.js';
```

- [ ] **Step 4: Run the shared reducer tests and verify they pass**

Run:

```bash
pnpm exec vitest --run packages/zero-sync/src/__tests__/workbook-history-state.test.ts
```

Expected: PASS

- [ ] **Step 5: Commit the shared reducer slice**

```bash
git add \
  packages/zero-sync/src/workbook-history-state.ts \
  packages/zero-sync/src/index.ts \
  packages/zero-sync/src/__tests__/workbook-history-state.test.ts
git commit -m "fix(history): derive redo state from authoritative actor stacks"
```

---

### Task 3: Replace Web Redo Heuristics With the Shared Reducer

**Files:**
- Modify: `apps/web/src/use-workbook-changes.ts`
- Modify: `apps/web/src/workbook-changes-model.ts`
- Modify: `apps/web/src/use-workbook-changes-pane.tsx`
- Test: `apps/web/src/__tests__/workbook-changes.test.tsx`

- [ ] **Step 1: Return normalized rows alongside rendered entries**

```ts
export interface WorkbookChangesViewModel {
  readonly rows: readonly ReturnType<typeof normalizeWorkbookChangeRows>[number][];
  readonly entries: readonly WorkbookChangeEntry[];
}

export function useWorkbookChanges(/* existing args */): WorkbookChangesViewModel {
  // keep the live Zero subscription exactly as-is
  return useMemo(
    () => ({
      rows,
      entries: selectWorkbookChangeEntries({
        rows,
        knownSheetNames: sheetNames,
      }),
    }),
    [rows, sheetNames],
  );
}
```

- [ ] **Step 2: Replace `selectWorkbookHistoryState` with a thin adapter over the shared reducer**

```ts
import { deriveWorkbookActorHistoryState } from '@bilig/zero-sync';

export function selectWorkbookHistoryState(input: {
  readonly rows: readonly WorkbookChangeRow[];
  readonly currentUserId: string;
}): WorkbookHistoryState {
  const state = deriveWorkbookActorHistoryState({
    actorUserId: input.currentUserId,
    rows: input.rows.map((row) => ({
      revision: row.revision,
      actorUserId: row.actorUserId,
      eventKind: row.eventKind,
      undoBundleJson: row.undoBundleJson,
      revertedByRevision: row.revertedByRevision,
      revertsRevision: row.revertsRevision,
    })),
  });

  return {
    canUndo: state.canUndo,
    canRedo: state.canRedo,
    undoRevision: state.undoRevision,
    redoRevision: state.redoRevision,
  };
}
```

- [ ] **Step 3: Drive the pane from raw rows, not entry heuristics**

```ts
const changes = useWorkbookChanges({ documentId, sheetNames, zero, enabled });
const historyState = useMemo(
  () => selectWorkbookHistoryState({ rows: changes.rows, currentUserId }),
  [changes.rows, currentUserId],
);

const changesPanel = useMemo(
  () => <WorkbookChangesPanel changes={changes.entries} onJump={onJump} />,
  [changes.entries, onJump],
);
```

- [ ] **Step 4: Run the web history tests and verify they pass**

Run:

```bash
pnpm exec vitest --run apps/web/src/__tests__/workbook-changes.test.tsx
```

Expected: PASS, including the new multi-step redo case.

- [ ] **Step 5: Commit the web history slice**

```bash
git add \
  apps/web/src/use-workbook-changes.ts \
  apps/web/src/workbook-changes-model.ts \
  apps/web/src/use-workbook-changes-pane.tsx \
  apps/web/src/__tests__/workbook-changes.test.tsx
git commit -m "fix(web): use shared redo history state"
```

---

### Task 4: Replace Server Redo Target Selection With the Shared Reducer

**Files:**
- Create: `apps/bilig/src/zero/__tests__/workbook-history-selector.test.ts`
- Modify: `apps/bilig/src/zero/workbook-change-store.ts`
- Modify: `apps/bilig/src/zero/server-mutators.ts`

- [ ] **Step 1: Add a small server-side selector module inside the change store**

```ts
import { deriveWorkbookActorHistoryState } from '@bilig/zero-sync';

export async function selectLatestRedoableWorkbookChange(
  db: Queryable,
  input: { documentId: string; actorUserId: string },
): Promise<WorkbookChangeRecord | null> {
  const rows = await listWorkbookChangesForActor(db, input);
  const history = deriveWorkbookActorHistoryState({
    actorUserId: input.actorUserId,
    rows: rows.map((row) => ({
      revision: row.revision,
      actorUserId: row.actorUserId,
      eventKind: row.eventKind,
      undoBundleJson: row.undoBundle,
      revertedByRevision: row.revertedByRevision,
      revertsRevision: row.revertsRevision,
    })),
  });
  return history.redoRevision === null ? null : rows.find((row) => row.revision === history.redoRevision) ?? null;
}
```

- [ ] **Step 2: Replace the SQL “latest unreverted revert row” helpers**

```ts
export async function listWorkbookChangesForActor(
  db: Queryable,
  input: { documentId: string; actorUserId: string },
): Promise<WorkbookChangeRecord[]> {
  const result = await db.query<WorkbookChangeSelectRow>(
    `
      SELECT revision AS "revision",
             actor_user_id AS "actorUserId",
             client_mutation_id AS "clientMutationId",
             event_kind AS "eventKind",
             summary AS "summary",
             sheet_id AS "sheetId",
             sheet_name AS "sheetName",
             anchor_address AS "anchorAddress",
             range_json AS "rangeJson",
             undo_bundle_json AS "undoBundleJson",
             reverted_by_revision AS "revertedByRevision",
             reverts_revision AS "revertsRevision",
             created_at AS "createdAtUnixMs"
        FROM workbook_change
       WHERE workbook_id = $1
         AND actor_user_id = $2
       ORDER BY revision ASC
    `,
    [input.documentId, input.actorUserId],
  );

  return result.rows.flatMap((row) => {
    const record = normalizeWorkbookChangeRecord(row);
    return record ? [record] : [];
  });
}
```

- [ ] **Step 3: Route the mutators through the selector**

```ts
case 'workbook.undoLatestChange': {
  const parsed = undoLatestWorkbookChangeArgsSchema.parse(args);
  const targetChange = await selectLatestUndoableWorkbookChange(serverTx.dbTransaction.wrappedTransaction, {
    documentId: parsed.documentId,
    actorUserId: session?.userID ?? 'system',
  });
  if (!targetChange?.undoBundle) {
    throw new Error('No undoable workbook change was found');
  }
  // existing commitWorkbookHistoryMutation call unchanged
}

case 'workbook.redoLatestChange': {
  const parsed = redoLatestWorkbookChangeArgsSchema.parse(args);
  const targetChange = await selectLatestRedoableWorkbookChange(serverTx.dbTransaction.wrappedTransaction, {
    documentId: parsed.documentId,
    actorUserId: session?.userID ?? 'system',
  });
  if (!targetChange?.undoBundle) {
    throw new Error('No redoable workbook change was found');
  }
  // existing commitWorkbookHistoryMutation call unchanged
}
```

- [ ] **Step 4: Run the server selector tests**

Run:

```bash
pnpm exec vitest --run apps/bilig/src/zero/__tests__/workbook-history-selector.test.ts apps/bilig/src/zero/__tests__/workbook-change-store.test.ts
```

Expected: PASS

- [ ] **Step 5: Commit the server history slice**

```bash
git add \
  apps/bilig/src/zero/workbook-change-store.ts \
  apps/bilig/src/zero/server-mutators.ts \
  apps/bilig/src/zero/__tests__/workbook-history-selector.test.ts
git commit -m "fix(server): select redo targets from actor history stacks"
```

---

### Task 5: Add Browser Regressions for the Reproduced Histories

**Files:**
- Modify: `e2e/tests/web-shell.pw.ts`

- [ ] **Step 1: Add a multi-step redo Playwright regression**

```ts
test('web app preserves redo through a two-step undo chain', async ({ page }) => {
  await page.goto('/');
  await waitForWorkbookReady(page);

  const undoButton = page.getByRole('button', { name: 'Undo', exact: true });
  const redoButton = page.getByRole('button', { name: 'Redo', exact: true });
  const nameBox = page.getByTestId('name-box');
  const formulaInput = page.getByTestId('formula-input');

  await nameBox.fill('A1');
  await nameBox.press('Enter');
  await formulaInput.fill('alpha');
  await formulaInput.press('Enter');

  await nameBox.fill('B1');
  await nameBox.press('Enter');
  await formulaInput.fill('beta');
  await formulaInput.press('Enter');

  await undoButton.click();
  await undoButton.click();

  await expect(redoButton).toBeEnabled();
  await redoButton.click();
  await expect(redoButton).toBeEnabled();
  await redoButton.click();
  await expect(redoButton).toBeDisabled();
});
```

- [ ] **Step 2: Add branch invalidation coverage**

```ts
test('web app clears redo after a fresh edit following undo', async ({ page }) => {
  await page.goto('/');
  await waitForWorkbookReady(page);

  const undoButton = page.getByRole('button', { name: 'Undo', exact: true });
  const redoButton = page.getByRole('button', { name: 'Redo', exact: true });
  const nameBox = page.getByTestId('name-box');
  const formulaInput = page.getByTestId('formula-input');

  await nameBox.fill('A1');
  await nameBox.press('Enter');
  await formulaInput.fill('seed');
  await formulaInput.press('Enter');

  await undoButton.click();
  await expect(redoButton).toBeEnabled();

  await nameBox.fill('C1');
  await nameBox.press('Enter');
  await formulaInput.fill('branch');
  await formulaInput.press('Enter');

  await expect(redoButton).toBeDisabled();
});
```

- [ ] **Step 3: Run the browser regression file**

Run:

```bash
pnpm exec playwright test e2e/tests/web-shell.pw.ts --grep "redo"
```

Expected: PASS

- [ ] **Step 4: Commit the browser regressions**

```bash
git add e2e/tests/web-shell.pw.ts
git commit -m "test(e2e): lock redo history behavior"
```

---

### Task 6: Final Verification and Cleanup

**Files:**
- Modify only if verification exposes a real bug

- [ ] **Step 1: Run the focused redo suite**

Run:

```bash
pnpm exec vitest --run \
  packages/zero-sync/src/__tests__/workbook-history-state.test.ts \
  apps/web/src/__tests__/workbook-changes.test.tsx \
  apps/bilig/src/zero/__tests__/workbook-history-selector.test.ts \
  apps/bilig/src/zero/__tests__/workbook-change-store.test.ts
```

Expected: PASS

- [ ] **Step 2: Run typecheck and lint on touched files**

Run:

```bash
pnpm exec tsc -p tsconfig.json --pretty false --noEmit
pnpm exec oxlint --config .oxlintrc.json --type-aware --deny-warnings \
  packages/zero-sync/src/workbook-history-state.ts \
  packages/zero-sync/src/index.ts \
  apps/web/src/use-workbook-changes.ts \
  apps/web/src/workbook-changes-model.ts \
  apps/web/src/use-workbook-changes-pane.tsx \
  apps/web/src/__tests__/workbook-changes.test.tsx \
  apps/bilig/src/zero/workbook-change-store.ts \
  apps/bilig/src/zero/server-mutators.ts \
  apps/bilig/src/zero/__tests__/workbook-history-selector.test.ts \
  e2e/tests/web-shell.pw.ts
```

Expected: PASS

- [ ] **Step 3: Run the full CI gate on the committed tree**

Run:

```bash
pnpm run ci
```

Expected: PASS

- [ ] **Step 4: Final commit if verification forced follow-up fixes**

```bash
git add -A
git commit -m "fix(history): finalize redo rewrite verification"
```

---

## Spec Coverage Check

- Multi-step redo after multiple undos: covered by Task 1 and Task 5.
- Shared client/server correctness rule: covered by Task 2, Task 3, and Task 4.
- Branch invalidation after a fresh edit: covered by Task 1, Task 4, and Task 5.
- Authoritative, non-flaky redo enablement: covered by replacing heuristic selectors in both web and server paths.

## Placeholder Scan

- No `TODO`, `TBD`, or “handle appropriately” placeholders remain.
- Every task lists exact files.
- Every code step includes concrete code.
- Every verification step includes exact commands and expected outcomes.

## Type Consistency Check

- Shared reducer naming is consistent: `deriveWorkbookActorHistoryState`.
- Web adapter reads `rows`, not `entries`, for correctness.
- Server selector reads `WorkbookChangeRecord` rows and maps them into the shared reducer input shape.

---

Plan complete and saved to `/Users/gregkonush/github.com/bilig/docs/superpowers/plans/2026-04-18-workbook-redo-rewrite.md`. Two execution options:

**1. Subagent-Driven (recommended)** - I dispatch a fresh subagent per task, review between tasks, fast iteration

**2. Inline Execution** - Execute tasks in this session using executing-plans, batch execution with checkpoints

**Which approach?**
