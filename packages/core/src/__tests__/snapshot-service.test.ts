import { Effect, Exit } from "effect";
import { describe, expect, it } from "vitest";
import { FormulaTable } from "../formula-table.js";
import { createReplicaState } from "../replica-state.js";
import { createEngineSnapshotService } from "../engine/services/snapshot-service.js";
import { StringPool } from "../string-pool.js";
import { WorkbookStore } from "../workbook-store.js";

describe("EngineSnapshotService", () => {
  it("normalizes thrown import failures into tagged snapshot errors", () => {
    const workbook = new WorkbookStore("book");
    const service = createEngineSnapshotService({
      state: {
        workbook,
        strings: new StringPool(),
        formulas: new FormulaTable(workbook.cellStore),
        replicaState: createReplicaState("replica"),
        entityVersions: new Map(),
        sheetDeleteVersions: new Map(),
      },
      getCellByIndex: () => {
        throw new Error("unused");
      },
      resetWorkbook: () => {
        throw new Error("broken");
      },
      executeRestoreTransaction: () => {},
    });

    const exit = Effect.runSyncExit(
      service.importWorkbook({
        version: 1,
        workbook: { name: "book" },
        sheets: [],
      }),
    );

    expect(Exit.isFailure(exit)).toBe(true);
  });

  it("roundtrips replica tracking maps through the service boundary", () => {
    const workbook = new WorkbookStore("book");
    const state = {
      workbook,
      strings: new StringPool(),
      formulas: new FormulaTable(workbook.cellStore),
      replicaState: createReplicaState("replica"),
      entityVersions: new Map([["cell:1", { counter: 2, replicaId: "replica", batchId: "replica:2", opIndex: 0 }]]),
      sheetDeleteVersions: new Map([["Sheet1", { counter: 3, replicaId: "replica", batchId: "replica:3", opIndex: 0 }]]),
    };
    const service = createEngineSnapshotService({
      state,
      getCellByIndex: () => {
        throw new Error("unused");
      },
      resetWorkbook: () => {},
      executeRestoreTransaction: () => {},
    });

    const exported = Effect.runSync(service.exportReplica());
    state.entityVersions.clear();
    state.sheetDeleteVersions.clear();

    Effect.runSync(service.importReplica(exported));

    expect(state.entityVersions.get("cell:1")).toEqual(exported.entityVersions[0]?.order);
    expect(state.sheetDeleteVersions.get("Sheet1")).toEqual(
      exported.sheetDeleteVersions[0]?.order,
    );
  });
});
