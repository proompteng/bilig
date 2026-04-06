import { afterEach, describe, expect, it, vi } from "vitest";
import { SpreadsheetEngine } from "@bilig/core";
import { formatAddress } from "@bilig/formula";
import type {
  WorkbookLocalAuthoritativeDelta,
  WorkbookLocalAuthoritativeBase,
  WorkbookLocalMutationRecord,
  WorkbookLocalProjectionOverlay,
  WorkbookLocalStoreFactory,
  WorkbookStoredState,
} from "@bilig/storage-browser";
import { WorkbookLocalStoreLockedError } from "@bilig/storage-browser";
import { ValueTag } from "@bilig/protocol";
import { decodeViewportPatch } from "@bilig/worker-transport";
import { buildWorkbookLocalAuthoritativeBase } from "../worker-local-base.js";
import { collectChangedCellsBySheet, collectViewportCells } from "../worker-runtime-support.js";
import { WorkbookWorkerRuntime } from "../worker-runtime";

function cloneMutationRecord(mutation: WorkbookLocalMutationRecord): WorkbookLocalMutationRecord {
  const nextMutation = structuredClone(mutation);
  nextMutation.args = [...mutation.args];
  return nextMutation;
}

function buildViewportFromAuthoritativeBase(input: {
  authoritativeBase: WorkbookLocalAuthoritativeBase;
  sheetName: string;
  viewport: {
    rowStart: number;
    rowEnd: number;
    colStart: number;
    colEnd: number;
  };
}) {
  const { authoritativeBase, sheetName, viewport } = input;
  const sheet = authoritativeBase.sheets.find((entry) => entry.name === sheetName);
  if (!sheet) {
    throw new Error(`Missing authoritative sheet ${sheetName}`);
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
    sheetId: sheet.sheetId,
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
        const inputRecord = authoritativeBase.cellInputs.find(
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
            input: inputRecord?.input,
            formula: inputRecord?.formula,
            format: inputRecord?.format,
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
}

function mergeViewportWithProjectionOverlay(input: {
  baseViewport: ReturnType<typeof buildViewportFromAuthoritativeBase>;
  projectionOverlay: WorkbookLocalProjectionOverlay | null;
  sheetName: string;
  viewport: {
    rowStart: number;
    rowEnd: number;
    colStart: number;
    colEnd: number;
  };
}) {
  const { baseViewport, projectionOverlay, sheetName, viewport } = input;
  if (!projectionOverlay) {
    return baseViewport;
  }

  const cells = new Map(baseViewport.cells.map((cell) => [cell.snapshot.address, cell]));
  projectionOverlay.cells
    .filter((cell) => {
      return (
        cell.sheetName === sheetName &&
        cell.rowNum >= viewport.rowStart &&
        cell.rowNum <= viewport.rowEnd &&
        cell.colNum >= viewport.colStart &&
        cell.colNum <= viewport.colEnd
      );
    })
    .forEach((cell) => {
      cells.set(cell.address, {
        row: cell.rowNum,
        col: cell.colNum,
        snapshot: {
          sheetName: cell.sheetName,
          address: cell.address,
          value: structuredClone(cell.value),
          flags: cell.flags,
          version: cell.version,
          input: cell.input,
          formula: cell.formula,
          format: cell.format,
          styleId: cell.styleId,
          numberFormatId: cell.numberFormatId,
        },
      });
    });

  const rowAxisEntries = new Map(baseViewport.rowAxisEntries.map((entry) => [entry.index, entry]));
  projectionOverlay.rowAxisEntries
    .filter((entry) => entry.sheetName === sheetName)
    .forEach((entry) => {
      rowAxisEntries.set(entry.entry.index, structuredClone(entry.entry));
    });

  const columnAxisEntries = new Map(
    baseViewport.columnAxisEntries.map((entry) => [entry.index, entry]),
  );
  projectionOverlay.columnAxisEntries
    .filter((entry) => entry.sheetName === sheetName)
    .forEach((entry) => {
      columnAxisEntries.set(entry.entry.index, structuredClone(entry.entry));
    });

  const styles = new Map(baseViewport.styles.map((style) => [style.id, style]));
  projectionOverlay.styles.forEach((style) => {
    styles.set(style.id, structuredClone(style));
  });

  return {
    ...baseViewport,
    cells: [...cells.values()].toSorted(
      (left, right) => left.row - right.row || left.col - right.col,
    ),
    rowAxisEntries: [...rowAxisEntries.values()].toSorted(
      (left, right) => left.index - right.index,
    ),
    columnAxisEntries: [...columnAxisEntries.values()].toSorted(
      (left, right) => left.index - right.index,
    ),
    styles: [...styles.values()],
  };
}

function mergeAuthoritativeBaseDelta(input: {
  currentBase: WorkbookLocalAuthoritativeBase | null;
  authoritativeDelta: WorkbookLocalAuthoritativeDelta;
}): WorkbookLocalAuthoritativeBase {
  const { currentBase, authoritativeDelta } = input;
  if (authoritativeDelta.replaceAll || currentBase === null) {
    return structuredClone(authoritativeDelta.base);
  }

  const replacedSheetIds = new Set(authoritativeDelta.replacedSheetIds);
  return {
    sheets: [
      ...currentBase.sheets.filter((sheet) => !replacedSheetIds.has(sheet.sheetId)),
      ...structuredClone(authoritativeDelta.base.sheets),
    ].toSorted((left, right) => left.sortOrder - right.sortOrder),
    cellInputs: [
      ...currentBase.cellInputs.filter((cell) => !replacedSheetIds.has(cell.sheetId)),
      ...structuredClone(authoritativeDelta.base.cellInputs),
    ],
    cellRenders: [
      ...currentBase.cellRenders.filter((cell) => !replacedSheetIds.has(cell.sheetId)),
      ...structuredClone(authoritativeDelta.base.cellRenders),
    ],
    rowAxisEntries: [
      ...currentBase.rowAxisEntries.filter((entry) => !replacedSheetIds.has(entry.sheetId)),
      ...structuredClone(authoritativeDelta.base.rowAxisEntries),
    ],
    columnAxisEntries: [
      ...currentBase.columnAxisEntries.filter((entry) => !replacedSheetIds.has(entry.sheetId)),
      ...structuredClone(authoritativeDelta.base.columnAxisEntries),
    ],
    styles: structuredClone(authoritativeDelta.base.styles),
  };
}

function createMemoryLocalStoreFactory(seed?: {
  state?: WorkbookStoredState | null;
  pendingMutations?: readonly WorkbookLocalMutationRecord[];
  onPersistProjectionState?: (state: WorkbookStoredState) => Promise<void> | void;
  onIngestAuthoritativeDelta?: (
    state: WorkbookStoredState,
    delta: WorkbookLocalAuthoritativeDelta,
  ) => Promise<void> | void;
  onReadViewportProjection?: (
    sheetName: string,
    viewport: {
      rowStart: number;
      rowEnd: number;
      colStart: number;
      colEnd: number;
    },
  ) => void;
  authoritativeBase?: WorkbookLocalAuthoritativeBase | null;
  projectionOverlay?: WorkbookLocalProjectionOverlay | null;
}): WorkbookLocalStoreFactory {
  let currentState = seed?.state ? structuredClone(seed.state) : null;
  let currentPendingMutations = (seed?.pendingMutations ?? []).map(cloneMutationRecord);
  let currentAuthoritativeBase = seed?.authoritativeBase
    ? structuredClone(seed.authoritativeBase)
    : null;
  let currentProjectionOverlay = seed?.projectionOverlay
    ? structuredClone(seed.projectionOverlay)
    : null;
  return {
    async open() {
      return {
        async loadState() {
          return currentState ? structuredClone(currentState) : null;
        },
        async persistProjectionState(input) {
          currentState = structuredClone(input.state);
          currentAuthoritativeBase = structuredClone(input.authoritativeBase);
          currentProjectionOverlay = structuredClone(input.projectionOverlay);
          await seed?.onPersistProjectionState?.(input.state);
        },
        async ingestAuthoritativeDelta(input) {
          currentState = structuredClone(input.state);
          currentAuthoritativeBase = mergeAuthoritativeBaseDelta({
            currentBase: currentAuthoritativeBase,
            authoritativeDelta: input.authoritativeDelta,
          });
          currentProjectionOverlay = structuredClone(input.projectionOverlay);
          if ((input.removePendingMutationIds?.length ?? 0) > 0) {
            const removedIds = new Set(input.removePendingMutationIds);
            currentPendingMutations = currentPendingMutations.filter(
              (mutation) => !removedIds.has(mutation.id),
            );
          }
          await seed?.onIngestAuthoritativeDelta?.(input.state, input.authoritativeDelta);
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
        readViewportProjection(sheetName, viewport) {
          const authoritativeBase = currentAuthoritativeBase;
          if (!authoritativeBase) {
            return null;
          }
          seed?.onReadViewportProjection?.(sheetName, viewport);
          return mergeViewportWithProjectionOverlay({
            baseViewport: buildViewportFromAuthoritativeBase({
              authoritativeBase,
              sheetName,
              viewport,
            }),
            projectionOverlay: currentProjectionOverlay,
            sheetName,
            viewport,
          });
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

  it("builds the initial full viewport patch from the local projection store when it matches the worker projection", async () => {
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
            async persistProjectionState() {},
            async ingestAuthoritativeDelta() {},
            async listPendingMutations() {
              return [];
            },
            async appendPendingMutation() {},
            async updatePendingMutation() {},
            async removePendingMutation() {},
            readViewportProjection() {
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

  it("reads initial local full patches through 128x32 worker tiles instead of a single wide viewport query", async () => {
    const seedEngine = new SpreadsheetEngine({ workbookName: "tile-doc", replicaId: "seed" });
    seedEngine.createSheet("Sheet1");
    seedEngine.setCellValue("Sheet1", "A1", 1);
    seedEngine.setCellValue("Sheet1", formatAddress(0, 128), 2);
    seedEngine.setCellValue("Sheet1", formatAddress(32, 0), 3);
    seedEngine.setCellValue("Sheet1", formatAddress(32, 128), 4);

    const viewportReads: Array<{
      rowStart: number;
      rowEnd: number;
      colStart: number;
      colEnd: number;
    }> = [];
    const runtime = new WorkbookWorkerRuntime({
      localStoreFactory: createMemoryLocalStoreFactory({
        state: {
          snapshot: seedEngine.exportSnapshot(),
          replica: seedEngine.exportReplicaSnapshot(),
          authoritativeRevision: 0,
          appliedPendingLocalSeq: 0,
        },
        authoritativeBase: buildWorkbookLocalAuthoritativeBase(seedEngine),
        onReadViewportProjection(_sheetName, viewport) {
          viewportReads.push({ ...viewport });
        },
      }),
    });

    await runtime.bootstrap({
      documentId: "tile-doc",
      replicaId: "browser:test",
      persistState: true,
    });

    const received = new Array<ReturnType<typeof decodeViewportPatch>>();
    runtime.subscribeViewportPatches(
      {
        sheetName: "Sheet1",
        rowStart: 0,
        rowEnd: 40,
        colStart: 0,
        colEnd: 140,
      },
      (bytes) => {
        received.push(decodeViewportPatch(bytes));
      },
    );

    expect(viewportReads).toEqual([
      { rowStart: 0, rowEnd: 31, colStart: 0, colEnd: 127 },
      { rowStart: 0, rowEnd: 31, colStart: 128, colEnd: 255 },
      { rowStart: 32, rowEnd: 63, colStart: 0, colEnd: 127 },
      { rowStart: 32, rowEnd: 63, colStart: 128, colEnd: 255 },
    ]);
    expect(received[0]?.cells.map((cell) => cell.displayText).toSorted()).toEqual([
      "1",
      "2",
      "3",
      "4",
    ]);
  });

  it("restores pending local projection overlays from persistence across bootstrap", async () => {
    const seedEngine = new SpreadsheetEngine({ workbookName: "overlay-doc", replicaId: "seed" });
    seedEngine.createSheet("Sheet1");
    seedEngine.setCellValue("Sheet1", "A1", 5);
    let viewportReadCount = 0;

    const runtime = new WorkbookWorkerRuntime({
      localStoreFactory: createMemoryLocalStoreFactory({
        state: {
          snapshot: seedEngine.exportSnapshot(),
          replica: seedEngine.exportReplicaSnapshot(),
          authoritativeRevision: 0,
          appliedPendingLocalSeq: 1,
        },
        pendingMutations: [
          {
            id: "overlay-doc:pending:1",
            localSeq: 1,
            baseRevision: 0,
            method: "setCellValue",
            args: ["Sheet1", "A1", 17],
            enqueuedAtUnixMs: 1,
            submittedAtUnixMs: null,
            status: "pending",
          },
        ],
        onReadViewportProjection() {
          viewportReadCount += 1;
        },
        authoritativeBase: buildWorkbookLocalAuthoritativeBase(seedEngine),
        projectionOverlay: {
          cells: [
            {
              sheetId: 1,
              sheetName: "Sheet1",
              address: "A1",
              rowNum: 0,
              colNum: 0,
              value: { tag: ValueTag.Number, value: 17 },
              flags: 0,
              version: 2,
              input: 17,
              formula: undefined,
              format: undefined,
              styleId: undefined,
              numberFormatId: undefined,
            },
          ],
          rowAxisEntries: [],
          columnAxisEntries: [],
          styles: [],
        },
      }),
    });

    const bootstrap = await runtime.bootstrap({
      documentId: "overlay-doc",
      replicaId: "browser:test",
      persistState: true,
    });

    expect(bootstrap.restoredFromPersistence).toBe(true);
    expect(bootstrap.requiresAuthoritativeHydrate).toBe(false);
    expect(runtime.getCell("Sheet1", "A1").value).toEqual({
      tag: ValueTag.Number,
      value: 17,
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
    expect(received[0]?.cells[0]?.displayText).toBe("17");
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

  it("ingests narrow authoritative event batches through delta persistence", async () => {
    const persistProjectionState = vi.fn(async () => {});
    const ingestAuthoritativeDelta = vi.fn(async () => {});
    const localStoreFactory = createMemoryLocalStoreFactory({
      onPersistProjectionState: persistProjectionState,
      onIngestAuthoritativeDelta: ingestAuthoritativeDelta,
    });
    const runtime = new WorkbookWorkerRuntime({ localStoreFactory });
    await runtime.bootstrap({
      documentId: "authoritative-delta-doc",
      replicaId: "browser:test",
      persistState: true,
    });

    expect(persistProjectionState).toHaveBeenCalledTimes(1);
    expect(ingestAuthoritativeDelta).toHaveBeenCalledTimes(0);

    await runtime.applyAuthoritativeEvents(
      [
        {
          revision: 1,
          clientMutationId: null,
          payload: {
            kind: "setCellValue",
            sheetName: "Sheet1",
            address: "B2",
            value: 23,
          },
        },
      ],
      1,
    );

    expect(persistProjectionState).toHaveBeenCalledTimes(1);
    expect(ingestAuthoritativeDelta).toHaveBeenCalledTimes(1);

    const reloaded = new WorkbookWorkerRuntime({ localStoreFactory });
    await reloaded.bootstrap({
      documentId: "authoritative-delta-doc",
      replicaId: "browser:reloaded",
      persistState: true,
    });

    expect(reloaded.getCell("Sheet1", "B2").value).toEqual({
      tag: ValueTag.Number,
      value: 23,
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
    const persistProjectionState = vi.fn(async () => {});
    const runtime = new WorkbookWorkerRuntime({
      localStoreFactory: createMemoryLocalStoreFactory({
        onPersistProjectionState: persistProjectionState,
      }),
    });

    await runtime.bootstrap({
      documentId: "perf-doc",
      replicaId: "browser:test",
      persistState: true,
    });

    expect(persistProjectionState).toHaveBeenCalledTimes(1);

    runtime.setCellValue("Sheet1", "A1", 1);
    runtime.setCellValue("Sheet1", "A2", 2);
    runtime.setCellValue("Sheet1", "A3", 3);

    expect(persistProjectionState).toHaveBeenCalledTimes(1);
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
