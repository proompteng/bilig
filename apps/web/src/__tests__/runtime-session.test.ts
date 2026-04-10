import { MessageChannel } from "node:worker_threads";
import { afterEach, describe, expect, it, vi } from "vitest";
import { createWorkerEngineHost } from "@bilig/worker-transport";
import { formatAddress } from "@bilig/formula";
import type {
  WorkbookLocalMutationRecord,
  WorkbookLocalStoreFactory,
} from "@bilig/storage-browser";
import { SpreadsheetEngine } from "@bilig/core";
import { ErrorCode, isWorkbookSnapshot, ValueTag, type WorkbookSnapshot } from "@bilig/protocol";
import { applyPendingWorkbookMutationToEngine } from "../worker-runtime-mutation-replay.js";
import { WorkbookWorkerRuntime } from "../worker-runtime.js";
import { createWorkerRuntimeSessionController } from "../runtime-session.js";

type TestLocalStore = Awaited<ReturnType<WorkbookLocalStoreFactory["open"]>>;
type TestStoredState = Awaited<ReturnType<TestLocalStore["loadState"]>>;
type TestAxisEntry = {
  id: string;
  index: number;
  size?: number;
  hidden?: boolean;
};
type TestCellSnapshot = {
  sheetName: string;
  address: string;
  value: {
    tag: ValueTag;
    value?: boolean | number | string;
    code?: number;
    stringId?: number;
  };
  flags: number;
  version: number;
  input?: boolean | number | string | null;
  formula?: string;
  format?: string;
  styleId?: string;
  numberFormatId?: string;
};
type TestSpreadsheetEngine = {
  ready(): Promise<void>;
  importSnapshot(snapshot: WorkbookSnapshot): void;
  getCell(sheetName: string, address: string): TestCellSnapshot;
  getRowAxisEntries(sheetName: string): readonly TestAxisEntry[];
  getColumnAxisEntries(sheetName: string): readonly TestAxisEntry[];
  workbook: {
    getSheet(sheetName: string): { id: number } | undefined;
  };
};

function cloneMutationRecord(mutation: WorkbookLocalMutationRecord): WorkbookLocalMutationRecord {
  const nextMutation = structuredClone(mutation);
  nextMutation.args = [...mutation.args];
  return nextMutation;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function createMemoryLocalStoreFactory(seed?: {
  state?: TestStoredState;
  pendingMutations?: readonly WorkbookLocalMutationRecord[];
}): WorkbookLocalStoreFactory {
  let currentState = seed?.state ? structuredClone(seed.state) : null;
  let currentMutationJournal = (seed?.pendingMutations ?? []).map(cloneMutationRecord);
  let currentEngine: TestSpreadsheetEngine | null = null;
  const loadCurrentEngine = async (): Promise<TestSpreadsheetEngine | null> => {
    if (!currentState) {
      currentEngine = null;
      return null;
    }
    currentEngine = new SpreadsheetEngine({
      workbookName: "derived",
      replicaId: "derived",
    });
    await currentEngine.ready();
    if (isWorkbookSnapshot(currentState.snapshot)) {
      currentEngine.importSnapshot(currentState.snapshot);
    }
    const engine = currentEngine;
    currentMutationJournal
      .filter((mutation) => mutation.status !== "acked")
      .forEach((mutation) => {
        applyPendingWorkbookMutationToEngine(engine, mutation);
      });
    return engine;
  };
  return {
    async open() {
      await loadCurrentEngine();
      return {
        async loadBootstrapState() {
          const workbookName =
            isRecord(currentState?.snapshot) &&
            isRecord(currentState.snapshot["workbook"]) &&
            typeof currentState.snapshot["workbook"]["name"] === "string"
              ? currentState.snapshot["workbook"]["name"]
              : "Sheet1";
          const sheetNames =
            isRecord(currentState?.snapshot) && Array.isArray(currentState.snapshot["sheets"])
              ? currentState.snapshot["sheets"].flatMap((sheet) =>
                  isRecord(sheet) && typeof sheet["name"] === "string" ? [sheet["name"]] : [],
                )
              : [];
          const state = currentState
            ? {
                workbookName,
                sheetNames,
                materializedCellCount:
                  isRecord(currentState.snapshot) && Array.isArray(currentState.snapshot["sheets"])
                    ? currentState.snapshot["sheets"].reduce((count, sheet) => {
                        if (!isRecord(sheet) || !Array.isArray(sheet["cells"])) {
                          return count;
                        }
                        return count + sheet["cells"].length;
                      }, 0)
                    : 0,
                authoritativeRevision: currentState.authoritativeRevision,
                appliedPendingLocalSeq: currentState.appliedPendingLocalSeq,
              }
            : null;
          return state ? structuredClone(state) : null;
        },
        async loadState() {
          return currentState ? structuredClone(currentState) : null;
        },
        async persistProjectionState(input) {
          currentState = structuredClone(input.state);
          await loadCurrentEngine();
        },
        async ingestAuthoritativeDelta(input) {
          currentState = structuredClone(input.state);
          if ((input.removePendingMutationIds?.length ?? 0) > 0) {
            const removedIds = new Set(input.removePendingMutationIds);
            currentMutationJournal = currentMutationJournal.map((mutation) =>
              removedIds.has(mutation.id)
                ? {
                    ...cloneMutationRecord(mutation),
                    ackedAtUnixMs: Date.now(),
                    status: "acked",
                  }
                : mutation,
            );
          }
          await loadCurrentEngine();
        },
        async listPendingMutations() {
          return currentMutationJournal
            .filter((mutation) => mutation.status !== "acked")
            .map(cloneMutationRecord);
        },
        async listMutationJournalEntries() {
          return currentMutationJournal.map(cloneMutationRecord);
        },
        async appendPendingMutation(mutation) {
          currentMutationJournal.push(cloneMutationRecord(mutation));
          await loadCurrentEngine();
        },
        async updatePendingMutation(mutation) {
          currentMutationJournal = currentMutationJournal.map((entry) =>
            entry.id === mutation.id ? cloneMutationRecord(mutation) : entry,
          );
          await loadCurrentEngine();
        },
        async removePendingMutation(id) {
          currentMutationJournal = currentMutationJournal.filter((mutation) => mutation.id !== id);
          await loadCurrentEngine();
        },
        readViewportProjection(sheetName, viewport) {
          if (!currentEngine) {
            return null;
          }
          const sheet = currentEngine.workbook.getSheet(sheetName);
          if (!sheet) {
            return null;
          }
          const cells = [];
          for (let row = viewport.rowStart; row <= viewport.rowEnd; row += 1) {
            for (let col = viewport.colStart; col <= viewport.colEnd; col += 1) {
              const address = formatAddress(row, col);
              const snapshot = currentEngine.getCell(sheetName, address);
              if (
                snapshot.value.tag === ValueTag.Empty &&
                snapshot.input == null &&
                !snapshot.formula
              ) {
                continue;
              }
              cells.push({ row, col, snapshot });
            }
          }
          return {
            sheetId: sheet.id,
            sheetName,
            cells,
            rowAxisEntries: currentEngine.getRowAxisEntries(sheetName),
            columnAxisEntries: currentEngine.getColumnAxisEntries(sheetName),
            styles: [{ id: "style-0" }],
          };
        },
        close() {},
      };
    },
  };
}

function createSnapshot(cells: readonly { address: string; value: number }[]): WorkbookSnapshot {
  return {
    version: 1,
    workbook: { name: "phase0-doc" },
    sheets: [
      {
        name: "Sheet1",
        order: 0,
        cells: cells.map((cell) => ({
          address: cell.address,
          value: cell.value,
        })),
      },
    ],
  };
}

function createMockWorkerPort(runtime: WorkbookWorkerRuntime) {
  const channel = new MessageChannel();
  const host = createWorkerEngineHost(runtime, channel.port1);
  return {
    postMessage(message) {
      channel.port2.postMessage(message);
    },
    addEventListener(type, listener) {
      channel.port2.addEventListener(type, listener);
    },
    removeEventListener(type, listener) {
      channel.port2.removeEventListener(type, listener);
    },
    start() {
      channel.port2.start();
    },
    terminate() {
      host.dispose();
      channel.port1.close();
      channel.port2.close();
    },
  };
}

interface MockZeroView<T> {
  zero: {
    materialize(): {
      readonly data: T;
      addListener(listener: (value: T) => void): () => void;
      destroy(): void;
    };
  };
  emit(value: T): void;
}

function createMockZeroView<T>(initialValue: T): MockZeroView<T> {
  let currentValue = initialValue;
  const listeners = new Set<(value: T) => void>();
  const view = {
    get data() {
      return currentValue;
    },
    addListener(listener: (value: T) => void) {
      listeners.add(listener);
      return () => {
        listeners.delete(listener);
      };
    },
    destroy() {},
  };

  return {
    zero: {
      materialize() {
        return view;
      },
    },
    emit(value: T) {
      currentValue = value;
      listeners.forEach((listener) => listener(value));
    },
  };
}

function createSequencedZeroViews(...views: Array<{ zero: { materialize(): unknown } }>) {
  let nextIndex = 0;
  return {
    zero: {
      materialize() {
        const next = views[nextIndex++];
        if (!next) {
          throw new Error("Unexpected Zero materialize call");
        }
        return next.zero.materialize();
      },
    },
  };
}

function expectPendingMutationId(value: unknown): string {
  if (!isRecord(value) || typeof value["id"] !== "string") {
    throw new Error("Expected a pending mutation payload with an id");
  }
  return value["id"];
}

function expectPendingMutationSummaries(
  value: unknown,
): Array<{ id: string; status: "local" | "submitted" | "rebased" | "failed" }> {
  if (!Array.isArray(value)) {
    throw new Error("Expected a pending mutation list");
  }
  return value.flatMap((entry) => {
    if (
      !isRecord(entry) ||
      typeof entry["id"] !== "string" ||
      (entry["status"] !== "local" &&
        entry["status"] !== "submitted" &&
        entry["status"] !== "rebased" &&
        entry["status"] !== "failed")
    ) {
      throw new Error("Expected a pending mutation summary");
    }
    return [
      {
        id: entry["id"],
        status: entry["status"],
      },
    ];
  });
}

describe("createWorkerRuntimeSessionController", () => {
  afterEach(() => {
    vi.restoreAllMocks();
  });

  it("replays structural insert and delete mutations into the local projection engine", async () => {
    const engine = new SpreadsheetEngine({
      workbookName: "phase0-doc",
      replicaId: "browser:test",
    });
    await engine.ready();
    engine.updateRowMetadata("Sheet1", 1, 1, 30, false);
    engine.updateColumnMetadata("Sheet1", 3, 1, 140, false);

    applyPendingWorkbookMutationToEngine(engine, {
      id: "pending-1",
      localSeq: 1,
      baseRevision: 0,
      enqueuedAtUnixMs: 1,
      submittedAtUnixMs: null,
      lastAttemptedAtUnixMs: null,
      ackedAtUnixMs: null,
      rebasedAtUnixMs: null,
      failedAtUnixMs: null,
      attemptCount: 0,
      failureMessage: null,
      status: "local",
      method: "insertRows",
      args: ["Sheet1", 1, 2],
    });
    applyPendingWorkbookMutationToEngine(engine, {
      id: "pending-2",
      localSeq: 2,
      baseRevision: 0,
      enqueuedAtUnixMs: 2,
      submittedAtUnixMs: null,
      lastAttemptedAtUnixMs: null,
      ackedAtUnixMs: null,
      rebasedAtUnixMs: null,
      failedAtUnixMs: null,
      attemptCount: 0,
      failureMessage: null,
      status: "local",
      method: "deleteColumns",
      args: ["Sheet1", 3, 1],
    });

    expect(engine.getRowAxisEntries("Sheet1")).toEqual([
      { id: "row-2", index: 1 },
      { id: "row-3", index: 2 },
      { id: "row-1", index: 3, size: 30, hidden: false },
    ]);
    expect(engine.getColumnAxisEntries("Sheet1")).toEqual([]);
  });

  it("hydrates the mounted worker session from the latest authoritative snapshot", async () => {
    const runtime = new WorkbookWorkerRuntime({
      localStoreFactory: createMemoryLocalStoreFactory(),
    });
    const runtimeStates: string[] = [];
    const selections: string[] = [];
    const phases: string[] = [];

    const controller = await createWorkerRuntimeSessionController(
      {
        documentId: "phase0-doc",
        replicaId: "browser:test",
        persistState: false,
        initialSelection: { sheetName: "Sheet1", address: "B2" },
        createWorker: () => createMockWorkerPort(runtime),
        fetchImpl: vi.fn(async () => {
          return new Response(JSON.stringify(createSnapshot([{ address: "B2", value: 41 }])), {
            status: 200,
            headers: {
              "content-type": "application/vnd.bilig.workbook+json",
            },
          });
        }),
      },
      {
        onRuntimeState(runtimeState) {
          runtimeStates.push(runtimeState.workbookName);
        },
        onSelection(selection) {
          selections.push(`${selection.sheetName}!${selection.address}`);
        },
        onError(message) {
          throw new Error(message);
        },
        onPhase(phase) {
          phases.push(phase);
        },
      },
    );

    expect(runtimeStates.at(-1)).toBe("phase0-doc");
    expect(selections.at(-1)).toBe("Sheet1!B2");
    expect(phases).toEqual(["hydratingLocal", "syncing", "steady"]);
    expect(controller.handle.viewportStore.getCell("Sheet1", "B2").value).toEqual({
      tag: ValueTag.Number,
      value: 41,
    });

    controller.dispose();
  });

  it("keeps persisted local state instead of rehydrating over it", async () => {
    const seedEngine = new SpreadsheetEngine({ workbookName: "phase0-doc", replicaId: "seed" });
    seedEngine.createSheet("Sheet1");
    seedEngine.setCellValue("Sheet1", "A1", 99);

    const runtime = new WorkbookWorkerRuntime({
      localStoreFactory: createMemoryLocalStoreFactory({
        state: {
          snapshot: seedEngine.exportSnapshot(),
          replica: seedEngine.exportReplicaSnapshot(),
          authoritativeRevision: 0,
          appliedPendingLocalSeq: 0,
        },
      }),
    });
    const fetchImpl = vi.fn(async () => {
      return new Response(JSON.stringify(createSnapshot([{ address: "A1", value: 1 }])), {
        status: 200,
        headers: {
          "content-type": "application/vnd.bilig.workbook+json",
        },
      });
    });

    const controller = await createWorkerRuntimeSessionController(
      {
        documentId: "phase0-doc",
        replicaId: "browser:test",
        persistState: true,
        initialSelection: { sheetName: "Sheet1", address: "A1" },
        createWorker: () => createMockWorkerPort(runtime),
        fetchImpl,
      },
      {
        onRuntimeState() {},
        onSelection() {},
        onError(message) {
          throw new Error(message);
        },
      },
    );

    expect(fetchImpl).not.toHaveBeenCalled();
    expect(controller.handle.viewportStore.getCell("Sheet1", "A1").value).toEqual({
      tag: ValueTag.Number,
      value: 99,
    });

    controller.dispose();
  });

  it("reports startup perf milestones for a persisted local-ready bootstrap", async () => {
    const seedEngine = new SpreadsheetEngine({ workbookName: "phase0-doc", replicaId: "seed" });
    seedEngine.createSheet("Sheet1");
    seedEngine.setCellValue("Sheet1", "A1", 99);

    const runtime = new WorkbookWorkerRuntime({
      localStoreFactory: createMemoryLocalStoreFactory({
        state: {
          snapshot: seedEngine.exportSnapshot(),
          replica: seedEngine.exportReplicaSnapshot(),
          authoritativeRevision: 0,
          appliedPendingLocalSeq: 0,
        },
      }),
    });
    const perfSession = {
      scope: "phase0-doc:perf",
      markShellMounted: vi.fn(),
      noteBootstrapResult: vi.fn(),
      markFirstAuthoritativePatchVisible: vi.fn(),
      markFirstSelectionVisible: vi.fn(),
    };

    const controller = await createWorkerRuntimeSessionController(
      {
        documentId: "phase0-doc",
        replicaId: "browser:test",
        persistState: true,
        initialSelection: { sheetName: "Sheet1", address: "A1" },
        createWorker: () => createMockWorkerPort(runtime),
        fetchImpl: vi.fn(async () => {
          throw new Error("persisted local-ready bootstrap should not fetch");
        }),
        perfSession,
      },
      {
        onRuntimeState() {},
        onSelection() {},
        onError(message) {
          throw new Error(message);
        },
      },
    );

    expect(perfSession.noteBootstrapResult).toHaveBeenCalledWith(
      expect.objectContaining({
        restoredFromPersistence: true,
        requiresAuthoritativeHydrate: false,
      }),
    );
    expect(perfSession.markFirstSelectionVisible).toHaveBeenCalledTimes(1);
    expect(perfSession.markFirstAuthoritativePatchVisible).not.toHaveBeenCalled();

    controller.dispose();
  });

  it("marks the first authoritative patch when bootstrap hydrate installs a snapshot", async () => {
    const runtime = new WorkbookWorkerRuntime({
      localStoreFactory: createMemoryLocalStoreFactory(),
    });
    const perfSession = {
      scope: "phase0-doc:hydrate",
      markShellMounted: vi.fn(),
      noteBootstrapResult: vi.fn(),
      markFirstAuthoritativePatchVisible: vi.fn(),
      markFirstSelectionVisible: vi.fn(),
    };

    const controller = await createWorkerRuntimeSessionController(
      {
        documentId: "phase0-doc",
        replicaId: "browser:test",
        persistState: true,
        initialSelection: { sheetName: "Sheet1", address: "A1" },
        createWorker: () => createMockWorkerPort(runtime),
        fetchImpl: vi.fn(
          async () =>
            new Response(JSON.stringify(createSnapshot([{ address: "B2", value: 42 }])), {
              status: 200,
              headers: { "content-type": "application/json" },
            }),
        ),
        perfSession,
      },
      {
        onRuntimeState() {},
        onSelection() {},
        onError(message) {
          throw new Error(message);
        },
      },
    );

    expect(perfSession.noteBootstrapResult).toHaveBeenCalledWith(
      expect.objectContaining({
        restoredFromPersistence: false,
      }),
    );
    expect(perfSession.markFirstAuthoritativePatchVisible).toHaveBeenCalledTimes(1);
    expect(controller.runtimeState.sheetNames).toEqual(["Sheet1"]);

    controller.dispose();
  });

  it("keeps restored local pending projection state instead of blocking on snapshot hydrate", async () => {
    const seedEngine = new SpreadsheetEngine({ workbookName: "phase0-doc", replicaId: "seed" });
    seedEngine.createSheet("Sheet1");
    seedEngine.setCellValue("Sheet1", "A1", 5);

    const runtime = new WorkbookWorkerRuntime({
      localStoreFactory: createMemoryLocalStoreFactory({
        state: {
          snapshot: seedEngine.exportSnapshot(),
          replica: seedEngine.exportReplicaSnapshot(),
          authoritativeRevision: 7,
          appliedPendingLocalSeq: 1,
        },
        pendingMutations: [
          {
            id: "phase0-doc:pending:1",
            localSeq: 1,
            baseRevision: 7,
            method: "setCellValue",
            args: ["Sheet1", "A1", 17],
            enqueuedAtUnixMs: 1,
            submittedAtUnixMs: null,
            lastAttemptedAtUnixMs: null,
            ackedAtUnixMs: null,
            rebasedAtUnixMs: null,
            failedAtUnixMs: null,
            attemptCount: 0,
            failureMessage: null,
            status: "local",
          },
        ],
      }),
    });
    const fetchImpl = vi.fn(async () => {
      return new Response(JSON.stringify(createSnapshot([{ address: "A1", value: 1 }])), {
        status: 200,
        headers: {
          "content-type": "application/vnd.bilig.workbook+json",
        },
      });
    });

    const controller = await createWorkerRuntimeSessionController(
      {
        documentId: "phase0-doc",
        replicaId: "browser:test",
        persistState: true,
        initialSelection: { sheetName: "Sheet1", address: "A1" },
        createWorker: () => createMockWorkerPort(runtime),
        fetchImpl,
      },
      {
        onRuntimeState() {},
        onSelection() {},
        onError(message) {
          throw new Error(message);
        },
      },
    );

    expect(fetchImpl).not.toHaveBeenCalled();
    expect(controller.handle.viewportStore.getCell("Sheet1", "A1").value).toEqual({
      tag: ValueTag.Number,
      value: 17,
    });

    controller.dispose();
  });

  it("rebases authoritative revisions through the worker and replays pending local mutations", async () => {
    const runtime = new WorkbookWorkerRuntime({
      localStoreFactory: createMemoryLocalStoreFactory(),
    });
    const workbookView = createMockZeroView<unknown>({
      headRevision: 0,
      calculatedRevision: 0,
    });
    const zero = createSequencedZeroViews(workbookView);
    const phases: string[] = [];
    const perfSession = {
      scope: "phase0-doc:reconcile",
      markShellMounted: vi.fn(),
      noteBootstrapResult: vi.fn(),
      markFirstAuthoritativePatchVisible: vi.fn(),
      markFirstReconcileStarted: vi.fn(),
      markFirstReconcileSettled: vi.fn(),
      markFirstSelectionVisible: vi.fn(),
    };
    const fetchImpl = vi.fn<typeof fetch>(async (input) => {
      const url =
        typeof input === "string" ? input : input instanceof URL ? input.toString() : input.url;
      if (url.includes("/snapshot/latest")) {
        return new Response(null, { status: 404 });
      }
      if (url.includes("/events?afterRevision=0")) {
        return new Response(
          JSON.stringify({
            afterRevision: 0,
            headRevision: 2,
            calculatedRevision: 2,
            events: [
              {
                revision: 2,
                clientMutationId: null,
                payload: {
                  kind: "setCellValue",
                  sheetName: "Sheet1",
                  address: "A1",
                  value: 5,
                },
              },
            ],
          }),
          {
            status: 200,
            headers: {
              "content-type": "application/json",
            },
          },
        );
      }
      throw new Error(`Unexpected fetch: ${url}`);
    });

    const controller = await createWorkerRuntimeSessionController(
      {
        documentId: "phase0-doc",
        replicaId: "browser:test",
        persistState: true,
        initialSelection: { sheetName: "Sheet1", address: "A1" },
        createWorker: () => createMockWorkerPort(runtime),
        zero: zero.zero,
        fetchImpl,
        perfSession,
      },
      {
        onRuntimeState() {},
        onSelection() {},
        onError(message) {
          throw new Error(message);
        },
        onPhase(phase) {
          phases.push(phase);
        },
      },
    );

    await controller.invoke("enqueuePendingMutation", {
      method: "setCellValue",
      args: ["Sheet1", "A1", 17],
    });

    workbookView.emit({
      headRevision: 2,
      calculatedRevision: 2,
    });

    await vi.waitFor(() => {
      expect(fetchImpl).toHaveBeenCalledTimes(2);
      expect(controller.handle.viewportStore.getCell("Sheet1", "A1").value).toEqual({
        tag: ValueTag.Number,
        value: 17,
      });
    });
    expect(phases).toContain("reconciling");
    expect(perfSession.markFirstReconcileStarted).toHaveBeenCalledTimes(1);
    expect(perfSession.markFirstReconcileSettled).toHaveBeenCalledTimes(1);
    expect(perfSession.markFirstAuthoritativePatchVisible).toHaveBeenCalledTimes(1);

    controller.dispose();
  });

  it("enters recovering when authoritative rebase falls back to a snapshot", async () => {
    const runtime = new WorkbookWorkerRuntime({
      localStoreFactory: createMemoryLocalStoreFactory(),
    });
    const workbookView = createMockZeroView<unknown>({
      headRevision: 0,
      calculatedRevision: 0,
    });
    const zero = createSequencedZeroViews(workbookView);
    const phases: string[] = [];
    const fetchImpl = vi.fn<typeof fetch>(async (input) => {
      const url =
        typeof input === "string" ? input : input instanceof URL ? input.toString() : input.url;
      if (url.includes("/snapshot/latest")) {
        return new Response(JSON.stringify(createSnapshot([{ address: "A1", value: 33 }])), {
          status: 200,
          headers: {
            "content-type": "application/vnd.bilig.workbook+json",
          },
        });
      }
      if (url.includes("/events?afterRevision=0")) {
        return new Response(
          JSON.stringify({
            afterRevision: 0,
            headRevision: 2,
            calculatedRevision: 2,
            events: [],
          }),
          {
            status: 200,
            headers: {
              "content-type": "application/json",
            },
          },
        );
      }
      throw new Error(`Unexpected fetch: ${url}`);
    });

    const controller = await createWorkerRuntimeSessionController(
      {
        documentId: "phase0-doc",
        replicaId: "browser:test",
        persistState: false,
        initialSelection: { sheetName: "Sheet1", address: "A1" },
        createWorker: () => createMockWorkerPort(runtime),
        zero: zero.zero,
        fetchImpl,
      },
      {
        onRuntimeState() {},
        onSelection() {},
        onError(message) {
          throw new Error(message);
        },
        onPhase(phase) {
          phases.push(phase);
        },
      },
    );

    workbookView.emit({
      headRevision: 2,
      calculatedRevision: 2,
    });

    await vi.waitFor(() => {
      expect(fetchImpl).toHaveBeenCalledTimes(3);
    });
    expect(phases).toContain("recovering");
    expect(controller.handle.viewportStore.getCell("Sheet1", "A1").value).toEqual({
      tag: ValueTag.Number,
      value: 33,
    });

    controller.dispose();
  });

  it("retries failed authoritative rebases without losing submitted pending ops", async () => {
    const runtime = new WorkbookWorkerRuntime({
      localStoreFactory: createMemoryLocalStoreFactory(),
    });
    const workbookView = createMockZeroView<unknown>({
      headRevision: 0,
      calculatedRevision: 0,
    });
    const zero = createSequencedZeroViews(workbookView);
    const phases: string[] = [];
    const errors: string[] = [];
    let pendingMutationId = "";
    let eventFetchCount = 0;
    const fetchImpl = vi.fn<typeof fetch>(async (input) => {
      const url =
        typeof input === "string" ? input : input instanceof URL ? input.toString() : input.url;
      if (url.includes("/snapshot/latest")) {
        return new Response(null, { status: 404 });
      }
      if (url.includes("/events?afterRevision=0")) {
        eventFetchCount += 1;
        if (eventFetchCount === 1) {
          return new Response("upstream error", { status: 500 });
        }
        return new Response(
          JSON.stringify({
            afterRevision: 0,
            headRevision: 1,
            calculatedRevision: 1,
            events: [
              {
                revision: 1,
                clientMutationId: pendingMutationId,
                payload: {
                  kind: "setCellValue",
                  sheetName: "Sheet1",
                  address: "A1",
                  value: 17,
                },
              },
            ],
          }),
          {
            status: 200,
            headers: {
              "content-type": "application/json",
            },
          },
        );
      }
      throw new Error(`Unexpected fetch: ${url}`);
    });

    const controller = await createWorkerRuntimeSessionController(
      {
        documentId: "phase0-doc",
        replicaId: "browser:test",
        persistState: true,
        initialSelection: { sheetName: "Sheet1", address: "A1" },
        createWorker: () => createMockWorkerPort(runtime),
        zero: zero.zero,
        fetchImpl,
      },
      {
        onRuntimeState() {},
        onSelection() {},
        onError(message) {
          errors.push(message);
        },
        onPhase(phase) {
          phases.push(phase);
        },
      },
    );

    pendingMutationId = expectPendingMutationId(
      await controller.invoke("enqueuePendingMutation", {
        method: "setCellValue",
        args: ["Sheet1", "A1", 17],
      }),
    );
    await controller.invoke("markPendingMutationSubmitted", pendingMutationId);

    workbookView.emit({
      headRevision: 1,
      calculatedRevision: 1,
    });

    await vi.waitFor(() => {
      expect(errors).toContain("Failed to load authoritative events (500)");
    });
    expect(expectPendingMutationSummaries(await controller.invoke("listPendingMutations"))).toEqual(
      [
        expect.objectContaining({
          id: pendingMutationId,
          status: "submitted",
        }),
      ],
    );
    expect(controller.handle.viewportStore.getCell("Sheet1", "A1").value).toEqual({
      tag: ValueTag.Number,
      value: 17,
    });

    workbookView.emit({
      headRevision: 1,
      calculatedRevision: 1,
    });

    await vi.waitFor(async () => {
      expect(fetchImpl).toHaveBeenCalledTimes(3);
      expect(
        expectPendingMutationSummaries(await controller.invoke("listPendingMutations")),
      ).toEqual([]);
      expect(controller.handle.viewportStore.getCell("Sheet1", "A1").value).toEqual({
        tag: ValueTag.Number,
        value: 17,
      });
    });
    expect(phases.filter((phase) => phase === "reconciling")).toHaveLength(2);

    controller.dispose();
  });

  it("applies local mutations through worker viewport patches", async () => {
    const runtime = new WorkbookWorkerRuntime({
      localStoreFactory: createMemoryLocalStoreFactory(),
    });
    const controller = await createWorkerRuntimeSessionController(
      {
        documentId: "phase0-doc",
        replicaId: "browser:test",
        persistState: false,
        initialSelection: { sheetName: "Sheet1", address: "A1" },
        createWorker: () => createMockWorkerPort(runtime),
        fetchImpl: vi.fn(async () => new Response(null, { status: 404 })),
      },
      {
        onRuntimeState() {},
        onSelection() {},
        onError(message) {
          throw new Error(message);
        },
      },
    );

    const unsubscribe = controller.subscribeViewport(
      "Sheet1",
      { rowStart: 0, rowEnd: 0, colStart: 0, colEnd: 0 },
      () => {},
    );

    await controller.invoke("setCellValue", "Sheet1", "A1", 123);

    await vi.waitFor(() => {
      expect(controller.handle.viewportStore.getCell("Sheet1", "A1").value).toEqual({
        tag: ValueTag.Number,
        value: 123,
      });
    });

    unsubscribe();
    controller.dispose();
  });

  it("applies invalid local formulas through pending worker viewport patches", async () => {
    const runtime = new WorkbookWorkerRuntime({
      localStoreFactory: createMemoryLocalStoreFactory(),
    });
    const controller = await createWorkerRuntimeSessionController(
      {
        documentId: "phase0-doc",
        replicaId: "browser:test",
        persistState: false,
        initialSelection: { sheetName: "Sheet1", address: "A1" },
        createWorker: () => createMockWorkerPort(runtime),
        fetchImpl: vi.fn(async () => new Response(null, { status: 404 })),
      },
      {
        onRuntimeState() {},
        onSelection() {},
        onError(message) {
          throw new Error(message);
        },
      },
    );

    const unsubscribe = controller.subscribeViewport(
      "Sheet1",
      { rowStart: 0, rowEnd: 0, colStart: 0, colEnd: 0 },
      () => {},
    );

    await controller.invoke("enqueuePendingMutation", {
      method: "setCellFormula",
      args: ["Sheet1", "A1", "1+"],
    });

    await vi.waitFor(() => {
      expect(controller.handle.viewportStore.getCell("Sheet1", "A1")).toMatchObject({
        value: {
          tag: ValueTag.Error,
          code: ErrorCode.Value,
        },
      });
    });

    unsubscribe();
    controller.dispose();
  });

  it("preserves case-insensitive string comparison results for pending local formulas", async () => {
    const runtime = new WorkbookWorkerRuntime({
      localStoreFactory: createMemoryLocalStoreFactory(),
    });
    const controller = await createWorkerRuntimeSessionController(
      {
        documentId: "phase0-doc",
        replicaId: "browser:test",
        persistState: false,
        initialSelection: { sheetName: "Sheet1", address: "A1" },
        createWorker: () => createMockWorkerPort(runtime),
        fetchImpl: vi.fn(async () => new Response(null, { status: 404 })),
      },
      {
        onRuntimeState() {},
        onSelection() {},
        onError(message) {
          throw new Error(message);
        },
      },
    );

    const unsubscribe = controller.subscribeViewport(
      "Sheet1",
      { rowStart: 0, rowEnd: 1, colStart: 0, colEnd: 0 },
      () => {},
    );

    await controller.invoke("enqueuePendingMutation", {
      method: "setCellValue",
      args: ["Sheet1", "A1", "hello"],
    });
    await controller.setSelection({ sheetName: "Sheet1", address: "A2" });
    await controller.invoke("enqueuePendingMutation", {
      method: "setCellFormula",
      args: ["Sheet1", "A2", 'A1="HELLO"'],
    });

    await vi.waitFor(() => {
      expect(controller.handle.viewportStore.getCell("Sheet1", "A2")).toMatchObject({
        value: {
          tag: ValueTag.Boolean,
          value: true,
        },
        formula: 'A1="HELLO"',
      });
    });

    unsubscribe();
    controller.dispose();
  });

  it("treats a 204 latest snapshot response as a cold-start miss", async () => {
    const runtime = new WorkbookWorkerRuntime({
      localStoreFactory: createMemoryLocalStoreFactory(),
    });
    const errors: string[] = [];
    const fetchImpl = vi.fn<typeof fetch>(async (input) => {
      const url =
        typeof input === "string" ? input : input instanceof URL ? input.toString() : input.url;
      if (url.includes("/snapshot/latest")) {
        return new Response(null, { status: 204 });
      }
      throw new Error(`Unexpected fetch: ${url}`);
    });

    const controller = await createWorkerRuntimeSessionController(
      {
        documentId: "phase0-doc",
        replicaId: "browser:test",
        persistState: false,
        initialSelection: { sheetName: "Sheet1", address: "A1" },
        createWorker: () => createMockWorkerPort(runtime),
        fetchImpl,
      },
      {
        onRuntimeState() {},
        onSelection() {},
        onError(message) {
          errors.push(message);
        },
      },
    );

    expect(fetchImpl).toHaveBeenCalledTimes(1);
    expect(errors).toEqual([]);
    expect(controller.handle.viewportStore.getCell("Sheet1", "A1").value).toEqual({
      tag: ValueTag.Empty,
    });

    controller.dispose();
  });

  it("applies remote cell and style changes through authoritative event rebases", async () => {
    const runtime = new WorkbookWorkerRuntime({
      localStoreFactory: createMemoryLocalStoreFactory(),
    });
    const workbookView = createMockZeroView<unknown>({
      headRevision: 0,
      calculatedRevision: 0,
    });
    const zero = createSequencedZeroViews(workbookView);
    const fetchImpl = vi.fn<typeof fetch>(async (input) => {
      const url =
        typeof input === "string" ? input : input instanceof URL ? input.toString() : input.url;
      if (url.includes("/snapshot/latest")) {
        return new Response(null, { status: 404 });
      }
      if (url.includes("/events?afterRevision=0")) {
        return new Response(
          JSON.stringify({
            afterRevision: 0,
            headRevision: 2,
            calculatedRevision: 2,
            events: [
              {
                revision: 1,
                clientMutationId: null,
                payload: {
                  kind: "setCellValue",
                  sheetName: "Sheet1",
                  address: "A1",
                  value: 8,
                },
              },
              {
                revision: 2,
                clientMutationId: null,
                payload: {
                  kind: "setRangeStyle",
                  range: {
                    sheetName: "Sheet1",
                    startAddress: "A1",
                    endAddress: "A1",
                  },
                  patch: {
                    fill: { backgroundColor: "#c9daf8" },
                  },
                },
              },
            ],
          }),
          {
            status: 200,
            headers: {
              "content-type": "application/json",
            },
          },
        );
      }
      throw new Error(`Unexpected fetch: ${url}`);
    });

    const controller = await createWorkerRuntimeSessionController(
      {
        documentId: "phase0-doc",
        replicaId: "browser:test",
        persistState: false,
        initialSelection: { sheetName: "Sheet1", address: "A1" },
        createWorker: () => createMockWorkerPort(runtime),
        zero: zero.zero,
        fetchImpl,
      },
      {
        onRuntimeState() {},
        onSelection() {},
        onError(message) {
          throw new Error(message);
        },
      },
    );

    const unsubscribe = controller.subscribeViewport(
      "Sheet1",
      { rowStart: 0, rowEnd: 0, colStart: 0, colEnd: 0 },
      () => {},
    );

    workbookView.emit({
      headRevision: 2,
      calculatedRevision: 2,
    });

    await vi.waitFor(() => {
      const snapshot = controller.handle.viewportStore.getCell("Sheet1", "A1");
      expect(snapshot.value).toEqual({
        tag: ValueTag.Number,
        value: 8,
      });
      expect(controller.handle.viewportStore.getCellStyle(snapshot.styleId)).toMatchObject({
        fill: { backgroundColor: "#c9daf8" },
      });
    });

    unsubscribe();
    controller.dispose();
  });

  it("does not subscribe render viewports directly to Zero tile queries", async () => {
    const runtime = new WorkbookWorkerRuntime({
      localStoreFactory: createMemoryLocalStoreFactory(),
    });
    const workbookView = createMockZeroView<unknown>({
      headRevision: 0,
      calculatedRevision: 0,
    });
    let materializeCount = 0;
    const zero = {
      materialize() {
        materializeCount += 1;
        if (materializeCount > 1) {
          throw new Error("Unexpected Zero tile materialize call");
        }
        return workbookView.zero.materialize();
      },
    };

    const controller = await createWorkerRuntimeSessionController(
      {
        documentId: "phase0-doc",
        replicaId: "browser:test",
        persistState: false,
        initialSelection: { sheetName: "Sheet1", address: "A1" },
        createWorker: () => createMockWorkerPort(runtime),
        zero,
        fetchImpl: vi.fn(async () => new Response(null, { status: 404 })),
      },
      {
        onRuntimeState() {},
        onSelection() {},
        onError(message) {
          throw new Error(message);
        },
      },
    );

    expect(materializeCount).toBe(1);

    const unsubscribe = controller.subscribeViewport(
      "Sheet1",
      { rowStart: 0, rowEnd: 0, colStart: 0, colEnd: 0 },
      () => {},
    );

    await controller.invoke("setCellValue", "Sheet1", "A1", 12);

    await vi.waitFor(() => {
      expect(controller.handle.viewportStore.getCell("Sheet1", "A1").value).toEqual({
        tag: ValueTag.Number,
        value: 12,
      });
    });

    expect(materializeCount).toBe(1);

    unsubscribe();
    controller.dispose();
  });
});
