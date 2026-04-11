import { describe, expect, it } from "vitest";
import {
  WorkbookLocalStoreLockedError,
  type WorkbookLocalStoreFactory,
} from "@bilig/storage-browser";
import type { PendingWorkbookMutation } from "../workbook-sync.js";
import { restoreBootstrapPersistence } from "../worker-runtime-bootstrap-persistence.js";

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

describe("worker runtime bootstrap persistence", () => {
  it("returns defaults when persistence is disabled", async () => {
    const result = await restoreBootstrapPersistence({
      persistState: false,
      documentId: "doc-1",
      localStoreFactory: {
        async open() {
          throw new Error("should not open");
        },
      },
    });

    expect(result).toEqual({
      localStore: null,
      restoredFromPersistence: false,
      restoredBootstrapState: null,
      authoritativeRevision: 0,
      appliedPendingLocalSeq: 0,
      mutationJournalEntries: [],
      pendingMutations: [],
      nextPendingMutationSeq: 1,
      localPersistenceMode: "ephemeral",
    });
  });

  it("restores bootstrap state and filters the pending mutation journal", async () => {
    const active = createMutation();
    const acked = createMutation({
      id: "worker-doc:pending:2",
      localSeq: 4,
      status: "acked",
      ackedAtUnixMs: 150,
    });
    const factory: WorkbookLocalStoreFactory = {
      async open() {
        return {
          async loadBootstrapState() {
            return {
              workbookName: "Budget",
              sheetNames: ["Sheet1"],
              materializedCellCount: 12,
              authoritativeRevision: 9,
              appliedPendingLocalSeq: 3,
            };
          },
          async loadState() {
            return null;
          },
          async persistProjectionState() {},
          async ingestAuthoritativeDelta() {},
          async listPendingMutations() {
            return [];
          },
          async listMutationJournalEntries() {
            return [active, acked, { localSeq: 8 }];
          },
          async appendPendingMutation() {},
          async updatePendingMutation() {},
          async removePendingMutation() {},
          readViewportProjection() {
            return null;
          },
          close() {},
        };
      },
    };

    const result = await restoreBootstrapPersistence({
      persistState: true,
      documentId: "doc-1",
      localStoreFactory: factory,
    });

    expect(result.restoredFromPersistence).toBe(true);
    expect(result.authoritativeRevision).toBe(9);
    expect(result.appliedPendingLocalSeq).toBe(3);
    expect(result.mutationJournalEntries).toEqual([active, acked]);
    expect(result.pendingMutations).toEqual([active]);
    expect(result.nextPendingMutationSeq).toBe(9);
  });

  it("treats a locked local store as unavailable", async () => {
    const result = await restoreBootstrapPersistence({
      persistState: true,
      documentId: "doc-1",
      localStoreFactory: {
        async open() {
          throw new WorkbookLocalStoreLockedError();
        },
      },
      lockRetryCount: 0,
      sleep: async () => {},
    });

    expect(result).toEqual({
      localStore: null,
      restoredFromPersistence: false,
      restoredBootstrapState: null,
      authoritativeRevision: 0,
      appliedPendingLocalSeq: 0,
      mutationJournalEntries: [],
      pendingMutations: [],
      nextPendingMutationSeq: 1,
      localPersistenceMode: "follower",
    });
  });

  it("retries transient lock conflicts before falling back to follower mode", async () => {
    let openCount = 0;
    const result = await restoreBootstrapPersistence({
      persistState: true,
      documentId: "doc-1",
      localStoreFactory: {
        async open() {
          openCount += 1;
          if (openCount < 3) {
            throw new WorkbookLocalStoreLockedError("locked");
          }
          return {
            async loadBootstrapState() {
              return null;
            },
            async loadState() {
              return null;
            },
            async persistProjectionState() {},
            async ingestAuthoritativeDelta() {},
            async listPendingMutations() {
              return [];
            },
            async listMutationJournalEntries() {
              return [];
            },
            async appendPendingMutation() {},
            async updatePendingMutation() {},
            async removePendingMutation() {},
            readViewportProjection() {
              return null;
            },
            close() {},
          };
        },
      },
      lockRetryCount: 2,
      sleep: async () => {},
    });

    expect(openCount).toBe(3);
    expect(result.localPersistenceMode).toBe("persistent");
    expect(result.localStore).not.toBeNull();
  });
});
