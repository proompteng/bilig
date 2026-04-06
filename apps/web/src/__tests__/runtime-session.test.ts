import { MessageChannel } from "node:worker_threads";
import { afterEach, describe, expect, it, vi } from "vitest";
import { createWorkerEngineHost } from "@bilig/worker-transport";
import type {
  WorkbookLocalMutationRecord,
  WorkbookLocalStoreFactory,
  WorkbookStoredState,
} from "@bilig/storage-browser";
import { SpreadsheetEngine } from "@bilig/core";
import { ValueTag, type WorkbookSnapshot } from "@bilig/protocol";
import { WorkbookWorkerRuntime } from "../worker-runtime.js";
import { createWorkerRuntimeSessionController } from "../runtime-session.js";

function cloneMutationRecord(mutation: WorkbookLocalMutationRecord): WorkbookLocalMutationRecord {
  const nextMutation = structuredClone(mutation);
  nextMutation.args = [...mutation.args];
  return nextMutation;
}

function createMemoryLocalStoreFactory(seed?: {
  state?: WorkbookStoredState | null;
  pendingMutations?: readonly WorkbookLocalMutationRecord[];
}): WorkbookLocalStoreFactory {
  let currentState = seed?.state ? structuredClone(seed.state) : null;
  let currentPendingMutations = (seed?.pendingMutations ?? []).map(cloneMutationRecord);
  return {
    async open() {
      return {
        async loadState() {
          return currentState ? structuredClone(currentState) : null;
        },
        async saveState(state) {
          currentState = structuredClone(state);
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

describe("createWorkerRuntimeSessionController", () => {
  afterEach(() => {
    vi.restoreAllMocks();
  });

  it("hydrates the mounted worker session from the latest authoritative snapshot", async () => {
    const runtime = new WorkbookWorkerRuntime({
      localStoreFactory: createMemoryLocalStoreFactory(),
    });
    const runtimeStates: string[] = [];
    const selections: string[] = [];

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
      },
    );

    expect(runtimeStates.at(-1)).toBe("phase0-doc");
    expect(selections.at(-1)).toBe("Sheet1!B2");
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

  it("rebases authoritative revisions through the worker and replays pending local mutations", async () => {
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
      },
      {
        onRuntimeState() {},
        onSelection() {},
        onError(message) {
          throw new Error(message);
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
