import { describe, expect, it } from "vitest";
import type { PendingWorkbookMutation } from "../workbook-sync.js";
import {
  buildPendingMutationSummary,
  markJournalMutationsRebased,
  replaceJournalMutation,
  syncPendingMutationsFromJournal,
} from "../worker-runtime-pending-mutations.js";

function createMutation(overrides: Partial<PendingWorkbookMutation> = {}): PendingWorkbookMutation {
  return {
    id: "worker-doc:pending:1",
    localSeq: 1,
    baseRevision: 0,
    method: "setCellValue",
    args: ["Sheet1", "A1", 17],
    enqueuedAtUnixMs: 100,
    submittedAtUnixMs: null,
    lastAttemptedAtUnixMs: null,
    ackedAtUnixMs: null,
    rebasedAtUnixMs: null,
    failedAtUnixMs: null,
    attemptCount: 0,
    failureMessage: null,
    status: "local",
    ...overrides,
  };
}

describe("worker runtime pending mutations", () => {
  it("filters active journal mutations into a cloned pending list", () => {
    const active = createMutation();
    const acked = createMutation({
      id: "worker-doc:pending:2",
      status: "acked",
      ackedAtUnixMs: 200,
    });

    const pending = syncPendingMutationsFromJournal([active, acked]);

    expect(pending).toEqual([active]);
    expect(pending[0]).not.toBe(active);
  });

  it("replaces a journal mutation and refreshes the active pending list", () => {
    const local = createMutation();
    const submitted = createMutation({
      status: "submitted",
      submittedAtUnixMs: 150,
      attemptCount: 1,
      lastAttemptedAtUnixMs: 150,
    });

    const result = replaceJournalMutation([local], submitted);

    expect(result.mutationJournalEntries).toEqual([submitted]);
    expect(result.pendingMutations).toEqual([submitted]);
    expect(result.mutationJournalEntries[0]).not.toBe(submitted);
  });

  it("marks active journal mutations as rebased and summarizes failures", () => {
    const local = createMutation();
    const failed = createMutation({
      id: "worker-doc:pending:2",
      status: "failed",
      failedAtUnixMs: 180,
      failureMessage: "mutation rejected",
      attemptCount: 2,
    });

    const rebased = markJournalMutationsRebased([local, failed], 250);

    expect(rebased.updatedMutations).toEqual([
      expect.objectContaining({
        id: local.id,
        status: "rebased",
        rebasedAtUnixMs: 250,
      }),
      expect.objectContaining({
        id: failed.id,
        status: "failed",
        rebasedAtUnixMs: null,
      }),
    ]);
    expect(
      buildPendingMutationSummary(rebased.mutationJournalEntries, rebased.pendingMutations),
    ).toEqual({
      activeCount: 2,
      failedCount: 1,
      firstFailed: {
        id: failed.id,
        method: failed.method,
        failureMessage: "mutation rejected",
        attemptCount: 2,
      },
    });
  });
});
