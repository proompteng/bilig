import { describe, expect, it } from "vitest";
import {
  EMPTY_RUNTIME_METRICS,
  buildWorkerRuntimeStateFromBootstrap,
  cloneWorkerRuntimeState,
  listOrderedSheetNames,
  withExternalSyncState,
} from "../worker-runtime-state.js";

describe("worker runtime state helpers", () => {
  it("orders sheet names by workbook order", () => {
    expect(
      listOrderedSheetNames({
        sheetsByName: new Map([
          ["c", { name: "Sheet3", order: 2 }],
          ["a", { name: "Sheet1", order: 0 }],
          ["b", { name: "Sheet2", order: 1 }],
        ]),
      }),
    ).toEqual(["Sheet1", "Sheet2", "Sheet3"]);
  });

  it("clones runtime state and applies external sync overrides without mutating the cache copy", () => {
    const cachedState = cloneWorkerRuntimeState({
      workbookName: "Book",
      sheetNames: ["Sheet1"],
      definedNames: [
        {
          name: "TaxRate",
          value: { kind: "cell-ref", sheetName: "Sheet1", address: "B2" },
        },
      ],
      metrics: EMPTY_RUNTIME_METRICS,
      syncState: "local",
      localPersistenceMode: "persistent",
    });

    const publicState = withExternalSyncState(cachedState, "syncing");

    expect(publicState.syncState).toBe("syncing");
    expect(cachedState.syncState).toBe("local");
    expect(publicState.localPersistenceMode).toBe("persistent");
    expect(publicState.metrics).not.toBe(cachedState.metrics);
    expect(publicState.definedNames).not.toBe(cachedState.definedNames);
    expect(publicState.definedNames).toEqual(cachedState.definedNames);
  });

  it("builds bootstrap runtime state with empty metrics and syncing status", () => {
    expect(
      buildWorkerRuntimeStateFromBootstrap({
        workbookName: "Book",
        sheetNames: ["Sheet1"],
        localPersistenceMode: "follower",
      }),
    ).toEqual({
      workbookName: "Book",
      sheetNames: ["Sheet1"],
      definedNames: [],
      metrics: EMPTY_RUNTIME_METRICS,
      syncState: "syncing",
      localPersistenceMode: "follower",
    });
  });
});
