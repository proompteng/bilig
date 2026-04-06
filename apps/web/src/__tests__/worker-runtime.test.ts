import { afterEach, describe, expect, it, vi } from "vitest";
import { SpreadsheetEngine } from "@bilig/core";
import type {
  WorkbookLocalAuthoritativeBase,
  WorkbookLocalMutationRecord,
  WorkbookLocalStoreFactory,
  WorkbookStoredState,
} from "@bilig/storage-browser";
import { WorkbookLocalStoreLockedError } from "@bilig/storage-browser";
import { ValueTag } from "@bilig/protocol";
import { decodeViewportPatch } from "@bilig/worker-transport";
import { collectChangedCellsBySheet, collectViewportCells } from "../worker-runtime-support.js";
import { WorkbookWorkerRuntime } from "../worker-runtime";

function cloneMutationRecord(mutation: WorkbookLocalMutationRecord): WorkbookLocalMutationRecord {
  const nextMutation = structuredClone(mutation);
  nextMutation.args = [...mutation.args];
  return nextMutation;
}

function createMemoryLocalStoreFactory(seed?: {
  state?: WorkbookStoredState | null;
  pendingMutations?: readonly WorkbookLocalMutationRecord[];
  onSaveState?: (state: WorkbookStoredState) => Promise<void> | void;
  authoritativeBase?: WorkbookLocalAuthoritativeBase | null;
}): WorkbookLocalStoreFactory {
  let currentState = seed?.state ? structuredClone(seed.state) : null;
  let currentPendingMutations = (seed?.pendingMutations ?? []).map(cloneMutationRecord);
  let currentAuthoritativeBase = seed?.authoritativeBase
    ? structuredClone(seed.authoritativeBase)
    : null;
  return {
    async open() {
      return {
        async loadState() {
          return currentState ? structuredClone(currentState) : null;
        },
        async saveState(state) {
          currentState = structuredClone(state);
          await seed?.onSaveState?.(state);
        },
        async listPendingMutations() {
          return currentPendingMutations.map(cloneMutationRecord);
        },
        async appendPendingMutation(mutation) {
          currentPendingMutations.push(cloneMutationRecord(mutation));
        },
        async updatePendingMutation(mutation) {
          currentPendingMutations = currentPendingMutations.map((entry) =>
            entry.id === mutation.id ? cloneMutationRecord(mutation) : entry,
          );
        },
        async removePendingMutation(id) {
          currentPendingMutations = currentPendingMutations.filter(
            (mutation) => mutation.id !== id,
          );
        },
        replaceAuthoritativeBase(base) {
          currentAuthoritativeBase = structuredClone(base);
        },
        readViewportBase(sheetName, viewport) {
          const authoritativeBase = currentAuthoritativeBase;
          if (!authoritativeBase) {
            return null;
          }
          const styles = authoritativeBase.styles.filter((style) => {
            return (
              style.id === "style-0" ||
              authoritativeBase.cellRenders.some((cell) => {
                return (
                  cell.sheetName === sheetName &&
                  cell.styleId === style.id &&
                  cell.rowNum >= viewport.rowStart &&
                  cell.rowNum <= viewport.rowEnd &&
                  cell.colNum >= viewport.colStart &&
                  cell.colNum <= viewport.colEnd
                );
              })
            );
          });
          return {
            sheetName,
            cells: authoritativeBase.cellRenders
              .filter((cell) => {
                return (
                  cell.sheetName === sheetName &&
                  cell.rowNum >= viewport.rowStart &&
                  cell.rowNum <= viewport.rowEnd &&
                  cell.colNum >= viewport.colStart &&
                  cell.colNum <= viewport.colEnd
                );
              })
              .map((cell) => {
                const input = authoritativeBase.cellInputs.find(
                  (entry) => entry.sheetName === cell.sheetName && entry.address === cell.address,
                );
                return {
                  row: cell.rowNum,
                  col: cell.colNum,
                  snapshot: {
                    sheetName: cell.sheetName,
                    address: cell.address,
                    value: structuredClone(cell.value),
                    flags: cell.flags,
                    version: cell.version,
                    styleId: cell.styleId,
                    numberFormatId: cell.numberFormatId,
                    input: input?.input,
                    formula: input?.formula,
                    format: input?.format,
                  },
                };
              }),
            rowAxisEntries: authoritativeBase.rowAxisEntries
              .filter((entry) => entry.sheetName === sheetName)
              .map((entry) => structuredClone(entry.entry)),
            columnAxisEntries: authoritativeBase.columnAxisEntries
              .filter((entry) => entry.sheetName === sheetName)
              .map((entry) => structuredClone(entry.entry)),
            styles: structuredClone(styles),
          };
        },
        close() {},
      };
    },
  };
}

describe("WorkbookWorkerRuntime", () => {
  afterEach(() => {
    vi.useRealTimers();
  });

  it("restores persisted workbook state and emits viewport patches for visible edits", async () => {
    const seedEngine = new SpreadsheetEngine({ workbookName: "phase3-doc", replicaId: "seed" });
    seedEngine.createSheet("Sheet1");
    seedEngine.setCellValue("Sheet1", "A1", 7);

    const localStoreFactory = createMemoryLocalStoreFactory({
      state: {
        snapshot: seedEngine.exportSnapshot(),
        replica: seedEngine.exportReplicaSnapshot(),
        authoritativeRevision: 0,
        appliedPendingLocalSeq: 0,
      },
    });

    const runtime = new WorkbookWorkerRuntime({ localStoreFactory });
    await runtime.bootstrap({
      documentId: "phase3-doc",
      replicaId: "browser:test",
      persistState: true,
    });

    const received = new Array<ReturnType<typeof decodeViewportPatch>>();
    runtime.subscribeViewportPatches(
      {
        sheetName: "Sheet1",
        rowStart: 0,
        rowEnd: 1,
        colStart: 0,
        colEnd: 1,
      },
      (bytes) => {
        received.push(decodeViewportPatch(bytes));
      },
    );

    expect(received[0]?.full).toBe(true);
    expect(received[0]?.cells.find((cell) => cell.snapshot.address === "A1")?.displayText).toBe(
      "7",
    );

    runtime.setCellFormula("Sheet1", "B1", "A1*2");

    expect(received).toHaveLength(2);
    expect(received[1]?.full).toBe(false);
    expect(received[1]?.cells).toHaveLength(1);
    expect(received[1]?.cells.find((cell) => cell.snapshot.address === "B1")?.displayText).toBe(
      "14",
    );
  });

  it("skips persistence restore when bootstrapped in ephemeral mode", async () => {
    const seedEngine = new SpreadsheetEngine({ workbookName: "phase3-doc", replicaId: "seed" });
    seedEngine.createSheet("Sheet1");
    seedEngine.setCellValue("Sheet1", "A1", 99);

    const localStoreFactory = createMemoryLocalStoreFactory({
      state: {
        snapshot: seedEngine.exportSnapshot(),
        replica: seedEngine.exportReplicaSnapshot(),
        authoritativeRevision: 0,
        appliedPendingLocalSeq: 0,
      },
    });

    const runtime = new WorkbookWorkerRuntime({ localStoreFactory });
    await runtime.bootstrap({
      documentId: "phase3-doc",
      replicaId: "browser:test",
      persistState: false,
    });

    expect(runtime.getCell("Sheet1", "A1").value).toEqual({ tag: ValueTag.Empty });
  });

  it("falls back to ephemeral runtime state when the local sqlite store is locked by another tab", async () => {
    const runtime = new WorkbookWorkerRuntime({
      localStoreFactory: {
        async open() {
          throw new WorkbookLocalStoreLockedError("locked");
        },
      },
    });

    const bootstrap = await runtime.bootstrap({
      documentId: "locked-doc",
      replicaId: "browser:test",
      persistState: true,
    });

    expect(bootstrap.restoredFromPersistence).toBe(false);
    expect(bootstrap.requiresAuthoritativeHydrate).toBe(false);
    expect(runtime.getCell("Sheet1", "A1").value).toEqual({ tag: ValueTag.Empty });
  });

  it("publishes viewport style dictionaries and stable style ids", async () => {
    const runtime = new WorkbookWorkerRuntime({
      localStoreFactory: createMemoryLocalStoreFactory(),
    });
    await runtime.bootstrap({
      documentId: "style-doc",
      replicaId: "browser:test",
      persistState: false,
    });

    const received = new Array<ReturnType<typeof decodeViewportPatch>>();
    runtime.subscribeViewportPatches(
      {
        sheetName: "Sheet1",
        rowStart: 0,
        rowEnd: 0,
        colStart: 0,
        colEnd: 0,
      },
      (bytes) => {
        received.push(decodeViewportPatch(bytes));
      },
    );

    runtime.setRangeStyle(
      { sheetName: "Sheet1", startAddress: "A1", endAddress: "A1" },
      { fill: { backgroundColor: "#336699" }, font: { family: "Fira Sans" } },
    );

    const patch = received.at(-1);
    expect(patch?.full).toBe(false);
    expect(patch?.styles).toHaveLength(1);
    expect(patch?.styles[0]).toMatchObject({
      fill: { backgroundColor: "#336699" },
      font: { family: "Fira Sans" },
    });
    expect(patch?.cells[0]?.styleId).toBe(patch?.styles[0]?.id);
  });

  it("builds the initial full viewport patch from the local authoritative base when it matches the projection", async () => {
    const seedEngine = new SpreadsheetEngine({ workbookName: "base-doc", replicaId: "seed" });
    seedEngine.createSheet("Sheet1");
    let viewportReadCount = 0;
    const runtime = new WorkbookWorkerRuntime({
      localStoreFactory: {
        async open() {
          return {
            async loadState() {
              return {
                snapshot: seedEngine.exportSnapshot(),
                replica: seedEngine.exportReplicaSnapshot(),
                authoritativeRevision: 0,
                appliedPendingLocalSeq: 0,
              };
            },
            async saveState() {},
            async listPendingMutations() {
              return [];
            },
            async appendPendingMutation() {},
            async updatePendingMutation() {},
            async removePendingMutation() {},
            replaceAuthoritativeBase() {},
            readViewportBase() {
              viewportReadCount += 1;
              return {
                sheetName: "Sheet1",
                cells: [
                  {
                    row: 0,
                    col: 0,
                    snapshot: {
                      sheetName: "Sheet1",
                      address: "A1",
                      value: { tag: ValueTag.Number, value: 42 },
                      flags: 0,
                      version: 1,
                    },
                  },
                ],
                rowAxisEntries: [],
                columnAxisEntries: [],
                styles: [{ id: "style-0" }],
              };
            },
            close() {},
          };
        },
      },
    });

    await runtime.bootstrap({
      documentId: "base-doc",
      replicaId: "browser:test",
      persistState: true,
    });

    const received = new Array<ReturnType<typeof decodeViewportPatch>>();
    runtime.subscribeViewportPatches(
      {
        sheetName: "Sheet1",
        rowStart: 0,
        rowEnd: 0,
        colStart: 0,
        colEnd: 0,
      },
      (bytes) => {
        received.push(decodeViewportPatch(bytes));
      },
    );

    expect(viewportReadCount).toBe(1);
    expect(received[0]?.cells[0]?.displayText).toBe("42");
  });

  it("patches only affected axis entries for column metadata edits", async () => {
    const runtime = new WorkbookWorkerRuntime({
      localStoreFactory: createMemoryLocalStoreFactory(),
    });
    await runtime.bootstrap({
      documentId: "axis-doc",
      replicaId: "browser:test",
      persistState: false,
    });

    const received = new Array<ReturnType<typeof decodeViewportPatch>>();
    runtime.subscribeViewportPatches(
      {
        sheetName: "Sheet1",
        rowStart: 0,
        rowEnd: 2,
        colStart: 0,
        colEnd: 3,
      },
      (bytes) => {
        received.push(decodeViewportPatch(bytes));
      },
    );

    runtime.updateColumnWidth("Sheet1", 1, 160);

    const patch = received.at(-1);
    expect(patch?.full).toBe(false);
    expect(patch?.cells).toHaveLength(0);
    expect(patch?.rows).toHaveLength(0);
    expect(patch?.columns).toEqual([{ index: 1, size: 160, hidden: false }]);
  });

  it("persists pending workbook mutations across bootstraps and removes them on ack", async () => {
    const localStoreFactory = createMemoryLocalStoreFactory();
    const runtime = new WorkbookWorkerRuntime({ localStoreFactory });
    await runtime.bootstrap({
      documentId: "pending-doc",
      replicaId: "browser:test",
      persistState: true,
    });

    const pending = await runtime.enqueuePendingMutation({
      method: "setCellValue",
      args: ["Sheet1", "A1", 17],
    });

    expect(runtime.listPendingMutations()).toEqual([pending]);

    const reloaded = new WorkbookWorkerRuntime({ localStoreFactory });
    await reloaded.bootstrap({
      documentId: "pending-doc",
      replicaId: "browser:reloaded",
      persistState: true,
    });

    expect(reloaded.listPendingMutations()).toEqual([pending]);

    await reloaded.ackPendingMutation(pending.id);
    expect(reloaded.listPendingMutations()).toEqual([]);

    const afterAck = new WorkbookWorkerRuntime({ localStoreFactory });
    await afterAck.bootstrap({
      documentId: "pending-doc",
      replicaId: "browser:after-ack",
      persistState: true,
    });

    expect(afterAck.listPendingMutations()).toEqual([]);
  });

  it("absorbs submitted pending mutations when authoritative events arrive", async () => {
    const localStoreFactory = createMemoryLocalStoreFactory();
    const runtime = new WorkbookWorkerRuntime({ localStoreFactory });
    await runtime.bootstrap({
      documentId: "authoritative-doc",
      replicaId: "browser:test",
      persistState: true,
    });

    const pending = await runtime.enqueuePendingMutation({
      method: "setCellValue",
      args: ["Sheet1", "A1", 17],
    });
    await runtime.markPendingMutationSubmitted(pending.id);

    expect(runtime.listPendingMutations()).toEqual([
      {
        ...pending,
        args: [...pending.args],
        submittedAtUnixMs: expect.any(Number),
        status: "submitted",
      },
    ]);

    await runtime.applyAuthoritativeEvents(
      [
        {
          revision: 1,
          clientMutationId: pending.id,
          payload: {
            kind: "setCellValue",
            sheetName: "Sheet1",
            address: "A1",
            value: 17,
          },
        },
      ],
      1,
    );

    expect(runtime.listPendingMutations()).toEqual([]);
    expect(runtime.getCell("Sheet1", "A1").value).toEqual({
      tag: ValueTag.Number,
      value: 17,
    });

    const reloaded = new WorkbookWorkerRuntime({ localStoreFactory });
    await reloaded.bootstrap({
      documentId: "authoritative-doc",
      replicaId: "browser:reloaded",
      persistState: true,
    });

    expect(reloaded.listPendingMutations()).toEqual([]);
    expect(reloaded.getCell("Sheet1", "A1").value).toEqual({
      tag: ValueTag.Number,
      value: 17,
    });
  });

  it("replays journaled mutations that were not yet captured in the persisted snapshot", async () => {
    const seedEngine = new SpreadsheetEngine({ workbookName: "journal-doc", replicaId: "seed" });
    seedEngine.createSheet("Sheet1");
    seedEngine.setCellValue("Sheet1", "A1", 5);

    const localStoreFactory = createMemoryLocalStoreFactory({
      state: {
        snapshot: seedEngine.exportSnapshot(),
        replica: seedEngine.exportReplicaSnapshot(),
        authoritativeRevision: 0,
        appliedPendingLocalSeq: 0,
      },
      pendingMutations: [
        {
          id: "journal-doc:pending:1",
          localSeq: 1,
          baseRevision: 0,
          method: "setCellValue",
          args: ["Sheet1", "A1", 17],
          enqueuedAtUnixMs: 1,
          submittedAtUnixMs: null,
          status: "pending",
        },
      ],
    });

    const runtime = new WorkbookWorkerRuntime({ localStoreFactory });
    await runtime.bootstrap({
      documentId: "journal-doc",
      replicaId: "browser:test",
      persistState: true,
    });

    expect(runtime.getCell("Sheet1", "A1").value).toEqual({
      tag: ValueTag.Number,
      value: 17,
    });
    expect(runtime.listPendingMutations()).toHaveLength(1);
  });

  it("rebases authoritative snapshots by replaying pending local mutations", async () => {
    const runtime = new WorkbookWorkerRuntime({
      localStoreFactory: createMemoryLocalStoreFactory(),
    });
    await runtime.bootstrap({
      documentId: "rebase-doc",
      replicaId: "browser:test",
      persistState: true,
    });

    runtime.setCellValue("Sheet1", "A1", 17);
    await runtime.enqueuePendingMutation({
      method: "setCellValue",
      args: ["Sheet1", "A1", 17],
    });

    await runtime.rebaseToSnapshot(
      {
        version: 1,
        workbook: { name: "rebase-doc" },
        sheets: [
          {
            name: "Sheet1",
            order: 0,
            cells: [{ address: "A1", value: 5 }],
          },
        ],
      },
      3,
    );

    expect(runtime.getAuthoritativeRevision()).toBe(3);
    expect(runtime.getCell("Sheet1", "A1").value).toEqual({
      tag: ValueTag.Number,
      value: 17,
    });
    expect(runtime.listPendingMutations()).toHaveLength(1);
  });

  it("skips unrelated viewport subscriptions when an edit is outside their sheet or region", async () => {
    const runtime = new WorkbookWorkerRuntime({
      localStoreFactory: createMemoryLocalStoreFactory(),
    });
    await runtime.bootstrap({
      documentId: "fanout-doc",
      replicaId: "browser:test",
      persistState: false,
    });

    runtime.renderCommit([{ kind: "upsertSheet", name: "Sheet2", order: 1 }]);

    const primary = new Array<ReturnType<typeof decodeViewportPatch>>();
    const offsheet = new Array<ReturnType<typeof decodeViewportPatch>>();
    const offregion = new Array<ReturnType<typeof decodeViewportPatch>>();

    runtime.subscribeViewportPatches(
      {
        sheetName: "Sheet1",
        rowStart: 0,
        rowEnd: 2,
        colStart: 0,
        colEnd: 2,
      },
      (bytes) => {
        primary.push(decodeViewportPatch(bytes));
      },
    );

    runtime.subscribeViewportPatches(
      {
        sheetName: "Sheet2",
        rowStart: 0,
        rowEnd: 2,
        colStart: 0,
        colEnd: 2,
      },
      (bytes) => {
        offsheet.push(decodeViewportPatch(bytes));
      },
    );

    runtime.subscribeViewportPatches(
      {
        sheetName: "Sheet1",
        rowStart: 10,
        rowEnd: 12,
        colStart: 10,
        colEnd: 12,
      },
      (bytes) => {
        offregion.push(decodeViewportPatch(bytes));
      },
    );

    expect(primary).toHaveLength(1);
    expect(offsheet).toHaveLength(1);
    expect(offregion).toHaveLength(1);

    runtime.setCellValue("Sheet1", "A1", 123);

    expect(primary).toHaveLength(2);
    expect(primary[1]?.cells[0]?.snapshot.address).toBe("A1");
    expect(offsheet).toHaveLength(1);
    expect(offregion).toHaveLength(1);
  });

  it("builds viewport patches only for subscriptions on impacted sheets", async () => {
    const runtime = new WorkbookWorkerRuntime({
      localStoreFactory: createMemoryLocalStoreFactory(),
    });
    await runtime.bootstrap({
      documentId: "sheet-index-doc",
      replicaId: "browser:test",
      persistState: false,
    });

    runtime.renderCommit([{ kind: "upsertSheet", name: "Sheet2", order: 1 }]);

    runtime.subscribeViewportPatches(
      {
        sheetName: "Sheet1",
        rowStart: 0,
        rowEnd: 2,
        colStart: 0,
        colEnd: 2,
      },
      () => {},
    );
    runtime.subscribeViewportPatches(
      {
        sheetName: "Sheet2",
        rowStart: 0,
        rowEnd: 2,
        colStart: 0,
        colEnd: 2,
      },
      () => {},
    );
    runtime.subscribeViewportPatches(
      {
        sheetName: "Sheet2",
        rowStart: 10,
        rowEnd: 12,
        colStart: 10,
        colEnd: 12,
      },
      () => {},
    );

    const originalBuildViewportPatch = runtime["buildViewportPatch"];
    if (typeof originalBuildViewportPatch !== "function") {
      throw new Error("Expected buildViewportPatch method");
    }

    let buildViewportPatchCalls = 0;
    runtime["buildViewportPatch"] = (...args: unknown[]) => {
      buildViewportPatchCalls += 1;
      return Reflect.apply(originalBuildViewportPatch, runtime, args);
    };

    runtime.setCellValue("Sheet1", "A1", 321);

    expect(buildViewportPatchCalls).toBe(1);
    runtime["buildViewportPatch"] = originalBuildViewportPatch;
  });

  it("dedupes changed viewport cells against invalidated range expansion", async () => {
    const runtime = new WorkbookWorkerRuntime({
      localStoreFactory: createMemoryLocalStoreFactory(),
    });
    await runtime.bootstrap({
      documentId: "range-dedupe-doc",
      replicaId: "browser:test",
      persistState: false,
    });

    const cells = collectViewportCells(
      {
        sheetName: "Sheet1",
        rowStart: 0,
        rowEnd: 1,
        colStart: 0,
        colEnd: 1,
      },
      {
        addresses: new Set(["A1"]),
        positions: [{ address: "A1", row: 0, col: 0 }],
      },
      [{ rowStart: 0, rowEnd: 0, colStart: 0, colEnd: 1 }],
    );

    expect(cells).toEqual([
      { address: "A1", row: 0, col: 0 },
      { address: "B1", row: 0, col: 1 },
    ]);
  });

  it("collects changed cells without qualified address string round-trips", async () => {
    const runtime = new WorkbookWorkerRuntime({
      localStoreFactory: createMemoryLocalStoreFactory(),
    });
    await runtime.bootstrap({
      documentId: "cell-store-impact-doc",
      replicaId: "browser:test",
      persistState: false,
    });

    runtime.setCellValue("Sheet1", "A1", 7);

    const engine = runtime["engine"];
    if (!engine || !engine.workbook) {
      throw new Error("Expected bootstrapped engine");
    }

    engine.workbook.getQualifiedAddress = () => {
      throw new Error("collectChangedCellsBySheet should not use getQualifiedAddress");
    };

    const impacts = collectChangedCellsBySheet(engine, [0]);

    expect(impacts.get("Sheet1")?.positions).toEqual([{ address: "A1", row: 0, col: 0 }]);
  });

  it("does not rewrite authoritative persistence for projected-only edit bursts", async () => {
    const saveState = vi.fn(async () => {});
    const runtime = new WorkbookWorkerRuntime({
      localStoreFactory: createMemoryLocalStoreFactory({
        onSaveState: saveState,
      }),
    });

    await runtime.bootstrap({
      documentId: "perf-doc",
      replicaId: "browser:test",
      persistState: true,
    });

    expect(saveState).toHaveBeenCalledTimes(1);

    runtime.setCellValue("Sheet1", "A1", 1);
    runtime.setCellValue("Sheet1", "A2", 2);
    runtime.setCellValue("Sheet1", "A3", 3);

    expect(saveState).toHaveBeenCalledTimes(1);
  });

  it("reuses exported snapshots until the workbook changes", async () => {
    const runtime = new WorkbookWorkerRuntime({
      localStoreFactory: createMemoryLocalStoreFactory(),
    });
    await runtime.bootstrap({
      documentId: "snapshot-cache-doc",
      replicaId: "browser:test",
      persistState: false,
    });

    const first = runtime.exportSnapshot();
    const second = runtime.exportSnapshot();

    expect(second).toBe(first);

    runtime.setCellValue("Sheet1", "A1", 42);

    const third = runtime.exportSnapshot();
    const fourth = runtime.exportSnapshot();

    expect(third).not.toBe(first);
    expect(third.sheets[0]?.cells).toContainEqual(
      expect.objectContaining({ address: "A1", value: 42 }),
    );
    expect(fourth).toBe(third);
  });
});
