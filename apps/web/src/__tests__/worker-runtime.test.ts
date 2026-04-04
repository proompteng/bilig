import { afterEach, describe, expect, it, vi } from "vitest";
import { SpreadsheetEngine } from "@bilig/core";
import type { BrowserPersistence } from "@bilig/storage-browser";
import { ValueTag } from "@bilig/protocol";
import { decodeViewportPatch } from "@bilig/worker-transport";
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
  afterEach(() => {
    vi.useRealTimers();
  });

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
      persistState: false,
    });

    expect(runtime.getCell("Sheet1", "A1").value).toEqual({ tag: ValueTag.Empty });
  });

  it("publishes viewport style dictionaries and stable style ids", async () => {
    const runtime = new WorkbookWorkerRuntime({ persistence: createMemoryPersistence() });
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

  it("patches only affected axis entries for column metadata edits", async () => {
    const runtime = new WorkbookWorkerRuntime({ persistence: createMemoryPersistence() });
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

  it("skips unrelated viewport subscriptions when an edit is outside their sheet or region", async () => {
    const runtime = new WorkbookWorkerRuntime({ persistence: createMemoryPersistence() });
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
    const runtime = new WorkbookWorkerRuntime({ persistence: createMemoryPersistence() });
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
    const runtime = new WorkbookWorkerRuntime({ persistence: createMemoryPersistence() });
    await runtime.bootstrap({
      documentId: "range-dedupe-doc",
      replicaId: "browser:test",
      persistState: false,
    });

    const collectViewportCells = runtime["collectViewportCells"];
    if (typeof collectViewportCells !== "function") {
      throw new Error("Expected collectViewportCells method");
    }

    const cells = Reflect.apply(collectViewportCells, runtime, [
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
    ]) as Array<{ address: string; row: number; col: number }>;

    expect(cells).toEqual([
      { address: "A1", row: 0, col: 0 },
      { address: "B1", row: 0, col: 1 },
    ]);
  });

  it("collects changed cells without qualified address string round-trips", async () => {
    const runtime = new WorkbookWorkerRuntime({ persistence: createMemoryPersistence() });
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

    const collectChangedCellsBySheet = runtime["collectChangedCellsBySheet"];
    if (typeof collectChangedCellsBySheet !== "function") {
      throw new Error("Expected collectChangedCellsBySheet method");
    }

    const impacts = Reflect.apply(collectChangedCellsBySheet, runtime, [[0]]) as Map<
      string,
      { positions: Array<{ address: string; row: number; col: number }> }
    >;

    expect(impacts.get("Sheet1")?.positions).toEqual([{ address: "A1", row: 0, col: 0 }]);
  });

  it("coalesces persistence saves across edit bursts", async () => {
    vi.useFakeTimers();
    const saveJson = vi.fn(async () => {});
    const runtime = new WorkbookWorkerRuntime({
      persistence: {
        async loadJson() {
          return null;
        },
        saveJson,
        async remove() {},
      },
    });

    await runtime.bootstrap({
      documentId: "perf-doc",
      replicaId: "browser:test",
      persistState: true,
    });

    expect(saveJson).toHaveBeenCalledTimes(1);

    runtime.setCellValue("Sheet1", "A1", 1);
    runtime.setCellValue("Sheet1", "A2", 2);
    runtime.setCellValue("Sheet1", "A3", 3);

    expect(saveJson).toHaveBeenCalledTimes(1);

    await vi.advanceTimersByTimeAsync(120);

    expect(saveJson).toHaveBeenCalledTimes(2);
  });

  it("reuses exported snapshots until the workbook changes", async () => {
    const runtime = new WorkbookWorkerRuntime({ persistence: createMemoryPersistence() });
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
