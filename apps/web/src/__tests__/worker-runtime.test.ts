import { describe, expect, it } from "vitest";
import { SpreadsheetEngine } from "@bilig/core";
import type { BrowserPersistence } from "@bilig/storage-browser";
import type { WorkbookSnapshot } from "@bilig/protocol";
import { ValueTag } from "@bilig/protocol";
import { decodeViewportPatch } from "@bilig/worker-transport";
import type { EngineSyncClient } from "@bilig/core";
import type { EngineOpBatch } from "@bilig/workbook-domain";
import { WorkbookWorkerRuntime } from "../worker-runtime";

function createMemoryPersistence(seed: Record<string, unknown> = {}): BrowserPersistence {
  const store = new Map<string, string>(
    Object.entries(seed).map(([key, value]) => [key, JSON.stringify(value)]),
  );
  return {
    async loadJson<T>(key: string, parser: (value: unknown) => T | null): Promise<T | null> {
      const raw = store.get(key);
      if (!raw) {
        return null;
      }
      return parser(JSON.parse(raw) as unknown);
    },
    async saveJson(key: string, value: unknown): Promise<void> {
      store.set(key, JSON.stringify(value));
    },
    async remove(key: string): Promise<void> {
      store.delete(key);
    },
  };
}

describe("WorkbookWorkerRuntime", () => {
  it("restores persisted workbook state and emits viewport patches for visible edits", async () => {
    const seedEngine = new SpreadsheetEngine({ workbookName: "phase3-doc", replicaId: "seed" });
    seedEngine.createSheet("Sheet1");
    seedEngine.setCellValue("Sheet1", "A1", 7);

    const persistence = createMemoryPersistence({
      "bilig:web:phase3-doc:runtime": {
        snapshot: seedEngine.exportSnapshot(),
        replica: seedEngine.exportReplicaSnapshot(),
      },
    });

    const runtime = new WorkbookWorkerRuntime({ persistence });
    await runtime.bootstrap({
      documentId: "phase3-doc",
      replicaId: "browser:test",
      baseUrl: null,
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
    expect(received[1]?.cells.find((cell) => cell.snapshot.address === "B1")?.displayText).toBe(
      "14",
    );
  });

  it("connects sync in the worker runtime and applies remote batches through viewport patches", async () => {
    const syncHooks: {
      applyRemoteBatch?: (batch: EngineOpBatch) => void;
      disconnect?: () => Promise<void>;
    } = {};

    const createSyncClient = (): EngineSyncClient => ({
      async connect(handlers) {
        syncHooks.applyRemoteBatch = (batch) => {
          handlers.applyRemoteBatch(batch);
        };
        handlers.setState("live");
        syncHooks.disconnect = async () => {
          handlers.setState("local-only");
        };
        return {
          send() {},
          disconnect: syncHooks.disconnect,
        };
      },
    });

    const runtime = new WorkbookWorkerRuntime({
      persistence: createMemoryPersistence(),
      createSyncClient,
    });

    await runtime.bootstrap({
      documentId: "sync-doc",
      replicaId: "browser:test",
      baseUrl: "http://127.0.0.1:4381",
      persistState: true,
    });
    expect(runtime.getRuntimeState().syncState).toBe("live");

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

    expect(syncHooks.applyRemoteBatch).toBeDefined();
    if (!syncHooks.applyRemoteBatch) {
      throw new Error("Expected sync client to provide applyRemoteBatch");
    }
    syncHooks.applyRemoteBatch({
      id: "server-1",
      replicaId: "server",
      clock: { counter: 1 },
      ops: [{ kind: "setCellValue", sheetName: "Sheet1", address: "A1", value: 42 }],
    });

    expect(runtime.getCell("Sheet1", "A1").value).toEqual({ tag: ValueTag.Number, value: 42 });
    expect(received.at(-1)?.cells.find((cell) => cell.snapshot.address === "A1")?.displayText).toBe(
      "42",
    );

    expect(syncHooks.disconnect).toBeDefined();
    if (!syncHooks.disconnect) {
      throw new Error("Expected sync client to provide disconnect");
    }
    await syncHooks.disconnect();
    expect(runtime.getRuntimeState().syncState).toBe("local-only");
  });

  it("skips persistence restore when bootstrapped in ephemeral mode", async () => {
    const seedEngine = new SpreadsheetEngine({ workbookName: "phase3-doc", replicaId: "seed" });
    seedEngine.createSheet("Sheet1");
    seedEngine.setCellValue("Sheet1", "A1", 99);

    const persistence = createMemoryPersistence({
      "bilig:web:phase3-doc:runtime": {
        snapshot: seedEngine.exportSnapshot(),
        replica: seedEngine.exportReplicaSnapshot(),
      },
    });

    const runtime = new WorkbookWorkerRuntime({ persistence });
    await runtime.bootstrap({
      documentId: "phase3-doc",
      replicaId: "browser:test",
      baseUrl: null,
      persistState: false,
    });

    expect(runtime.getCell("Sheet1", "A1").value).toEqual({ tag: ValueTag.Empty });
  });

  it("bootstraps from the latest server snapshot before rendering the default sheet", async () => {
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: { name: "Imported" },
      sheets: [
        {
          name: "Sheet1",
          order: 0,
          cells: [{ address: "A1", value: 55 }],
        },
      ],
    };

    const fetchMock = async () =>
      new Response(JSON.stringify(snapshot), {
        status: 200,
        headers: { "x-bilig-snapshot-cursor": "3" },
      });

    const originalFetch = globalThis.fetch;
    Object.defineProperty(globalThis, "fetch", {
      configurable: true,
      value: fetchMock,
    });

    const seenInitialCursor: Array<number | undefined> = [];
    const runtime = new WorkbookWorkerRuntime({
      persistence: createMemoryPersistence(),
      createSyncClient: (options): EngineSyncClient => ({
        async connect(handlers) {
          seenInitialCursor.push(options.initialServerCursor);
          handlers.setState("live");
          return {
            send() {},
            disconnect() {
              handlers.setState("local-only");
            },
          };
        },
      }),
    });

    try {
      await runtime.bootstrap({
        documentId: "server-doc",
        replicaId: "browser:test",
        baseUrl: "http://127.0.0.1:4381",
        persistState: false,
      });
    } finally {
      Object.defineProperty(globalThis, "fetch", {
        configurable: true,
        value: originalFetch,
      });
    }

    expect(runtime.getCell("Sheet1", "A1").value).toEqual({ tag: ValueTag.Number, value: 55 });
    expect(seenInitialCursor).toEqual([3]);
  });

  it("applies live remote snapshots from the sync client and emits viewport patches", async () => {
    let applyRemoteSnapshot: ((snapshot: WorkbookSnapshot) => void) | undefined;
    const runtime = new WorkbookWorkerRuntime({
      persistence: createMemoryPersistence(),
      createSyncClient: (): EngineSyncClient => ({
        async connect(handlers) {
          applyRemoteSnapshot = (snapshot) => {
            handlers.applyRemoteSnapshot?.(snapshot);
          };
          handlers.setState("live");
          return {
            send() {},
            disconnect() {
              handlers.setState("local-only");
            },
          };
        },
      }),
    });

    await runtime.bootstrap({
      documentId: "sync-snapshot-doc",
      replicaId: "browser:test",
      baseUrl: "http://127.0.0.1:4381",
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

    if (!applyRemoteSnapshot) {
      throw new Error("Expected sync client to provide applyRemoteSnapshot");
    }
    applyRemoteSnapshot({
      version: 1,
      workbook: { name: "Imported live" },
      sheets: [
        {
          name: "Sheet1",
          order: 0,
          cells: [{ address: "A1", value: 88 }],
        },
      ],
    });

    expect(runtime.getCell("Sheet1", "A1").value).toEqual({ tag: ValueTag.Number, value: 88 });
    expect(received.at(-1)?.cells.find((cell) => cell.snapshot.address === "A1")?.displayText).toBe(
      "88",
    );
  });

  it("publishes viewport style dictionaries and stable style ids", async () => {
    const runtime = new WorkbookWorkerRuntime({ persistence: createMemoryPersistence() });
    await runtime.bootstrap({
      documentId: "style-doc",
      replicaId: "browser:test",
      baseUrl: null,
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
    expect(patch?.styles).toHaveLength(1);
    expect(patch?.styles[0]).toMatchObject({
      fill: { backgroundColor: "#336699" },
      font: { family: "Fira Sans" },
    });
    expect(patch?.cells[0]?.styleId).toBe(patch?.styles[0]?.id);
  });
});
