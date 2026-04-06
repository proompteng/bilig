import { MessageChannel } from "node:worker_threads";
import { afterEach, describe, expect, it, vi } from "vitest";
import { createWorkerEngineHost } from "@bilig/worker-transport";
import type { BrowserPersistence } from "@bilig/storage-browser";
import { SpreadsheetEngine } from "@bilig/core";
import { ValueTag, type WorkbookSnapshot } from "@bilig/protocol";
import { WorkbookWorkerRuntime } from "../worker-runtime.js";
import { createWorkerRuntimeSessionController } from "../runtime-session.js";

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
    const runtime = new WorkbookWorkerRuntime({ persistence: createMemoryPersistence() });
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
    expect(controller.handle.cache.getCell("Sheet1", "B2").value).toEqual({
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
      persistence: createMemoryPersistence({
        "bilig:web:phase0-doc:runtime": {
          snapshot: seedEngine.exportSnapshot(),
          replica: seedEngine.exportReplicaSnapshot(),
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
    expect(controller.handle.cache.getCell("Sheet1", "A1").value).toEqual({
      tag: ValueTag.Number,
      value: 99,
    });

    controller.dispose();
  });

  it("applies local mutations through worker viewport patches", async () => {
    const runtime = new WorkbookWorkerRuntime({ persistence: createMemoryPersistence() });
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
      expect(controller.handle.cache.getCell("Sheet1", "A1").value).toEqual({
        tag: ValueTag.Number,
        value: 123,
      });
    });

    unsubscribe();
    controller.dispose();
  });

  it("projects live Zero cell and style updates into the mounted cache", async () => {
    const runtime = new WorkbookWorkerRuntime({ persistence: createMemoryPersistence() });
    const workbookView = createMockZeroView<unknown>({
      headRevision: 0,
      calculatedRevision: 0,
    });
    const stylesView = createMockZeroView<readonly unknown[]>([]);
    const formatsView = createMockZeroView<readonly unknown[]>([]);
    const selectedSourceTileView = createMockZeroView<readonly unknown[]>([]);
    const selectedEvalTileView = createMockZeroView<readonly unknown[]>([]);
    const selectedRowMetadataView = createMockZeroView<readonly unknown[]>([]);
    const selectedColumnMetadataView = createMockZeroView<readonly unknown[]>([]);
    const viewportSourceView = createMockZeroView<readonly unknown[]>([]);
    const viewportEvalView = createMockZeroView<readonly unknown[]>([]);
    const rowMetadataView = createMockZeroView<readonly unknown[]>([]);
    const columnMetadataView = createMockZeroView<readonly unknown[]>([]);
    const zero = createSequencedZeroViews(
      workbookView,
      stylesView,
      formatsView,
      selectedSourceTileView,
      selectedEvalTileView,
      selectedRowMetadataView,
      selectedColumnMetadataView,
      viewportSourceView,
      viewportEvalView,
      rowMetadataView,
      columnMetadataView,
    );

    const controller = await createWorkerRuntimeSessionController(
      {
        documentId: "phase0-doc",
        replicaId: "browser:test",
        persistState: false,
        initialSelection: { sheetName: "Sheet1", address: "A1" },
        createWorker: () => createMockWorkerPort(runtime),
        zero: zero.zero,
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

    selectedSourceTileView.emit([
      {
        sheetName: "Sheet1",
        address: "A1",
        inputValue: 5,
        styleId: "style-live",
      },
    ]);
    selectedEvalTileView.emit([
      {
        sheetName: "Sheet1",
        address: "A1",
        value: { tag: ValueTag.Number, value: 5 },
        flags: 0,
        version: 1,
        styleId: "style-live",
      },
    ]);

    await vi.waitFor(() => {
      expect(controller.handle.cache.getCell("Sheet1", "A1").value).toEqual({
        tag: ValueTag.Number,
        value: 5,
      });
    });

    const unsubscribe = controller.subscribeViewport(
      "Sheet1",
      { rowStart: 0, rowEnd: 0, colStart: 0, colEnd: 0 },
      () => {},
    );

    viewportEvalView.emit([
      {
        sheetName: "Sheet1",
        address: "A1",
        value: { tag: ValueTag.Number, value: 8 },
        flags: 0,
        version: 2,
        styleId: "style-live",
      },
    ]);
    stylesView.emit([
      {
        styleId: "style-live",
        styleJson: {
          fill: { backgroundColor: "#c9daf8" },
        },
      },
    ]);

    await vi.waitFor(() => {
      expect(controller.handle.cache.getCell("Sheet1", "A1").value).toEqual({
        tag: ValueTag.Number,
        value: 8,
      });
      expect(controller.handle.cache.getCellStyle("style-live")).toMatchObject({
        fill: { backgroundColor: "#c9daf8" },
      });
    });

    unsubscribe();
    controller.dispose();
  });

  it("does not let stale remote cell_eval overwrite a newer local formula result", async () => {
    const runtime = new WorkbookWorkerRuntime({ persistence: createMemoryPersistence() });
    const workbookView = createMockZeroView<unknown>({
      headRevision: 2,
      calculatedRevision: 1,
    });
    const stylesView = createMockZeroView<readonly unknown[]>([]);
    const formatsView = createMockZeroView<readonly unknown[]>([]);
    const initialSelectedSourceTileView = createMockZeroView<readonly unknown[]>([]);
    const initialSelectedEvalTileView = createMockZeroView<readonly unknown[]>([]);
    const initialSelectedRowTileView = createMockZeroView<readonly unknown[]>([]);
    const initialSelectedColumnTileView = createMockZeroView<readonly unknown[]>([]);
    const selectedSourceTileView = createMockZeroView<readonly unknown[]>([]);
    const selectedEvalTileView = createMockZeroView<readonly unknown[]>([]);
    const selectedRowTileView = createMockZeroView<readonly unknown[]>([]);
    const selectedColumnTileView = createMockZeroView<readonly unknown[]>([]);
    const zero = createSequencedZeroViews(
      workbookView,
      stylesView,
      formatsView,
      initialSelectedSourceTileView,
      initialSelectedEvalTileView,
      initialSelectedRowTileView,
      initialSelectedColumnTileView,
      selectedSourceTileView,
      selectedEvalTileView,
      selectedRowTileView,
      selectedColumnTileView,
    );

    const controller = await createWorkerRuntimeSessionController(
      {
        documentId: "phase0-doc",
        replicaId: "browser:test",
        persistState: false,
        initialSelection: { sheetName: "Sheet1", address: "A1" },
        createWorker: () => createMockWorkerPort(runtime),
        zero: zero.zero,
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

    await controller.invoke("setCellValue", "Sheet1", "A1", "hello");
    await controller.invoke("setCellFormula", "Sheet1", "A2", 'A1="HELLO"');
    await controller.setSelection({ sheetName: "Sheet1", address: "A2" });
    await expect(
      controller.setSelection({ sheetName: "Sheet1", address: "A2" }),
    ).resolves.toBeUndefined();

    await vi.waitFor(() => {
      expect(controller.handle.cache.getCell("Sheet1", "A2").value).toEqual({
        tag: ValueTag.Boolean,
        value: true,
      });
    });

    selectedSourceTileView.emit([
      {
        sheetName: "Sheet1",
        address: "A2",
        formula: 'A1="HELLO"',
      },
    ]);
    selectedEvalTileView.emit([
      {
        sheetName: "Sheet1",
        address: "A2",
        value: { tag: ValueTag.Boolean, value: false },
        flags: 0,
        version: 1,
      },
    ]);

    expect(controller.handle.cache.getCell("Sheet1", "A2").value).toEqual({
      tag: ValueTag.Boolean,
      value: true,
    });

    selectedEvalTileView.emit([
      {
        sheetName: "Sheet1",
        address: "A2",
        value: { tag: ValueTag.Boolean, value: true },
        flags: 0,
        version: 2,
      },
    ]);
    workbookView.emit({
      headRevision: 2,
      calculatedRevision: 2,
    });

    await vi.waitFor(() => {
      expect(controller.handle.cache.getCell("Sheet1", "A2").value).toEqual({
        tag: ValueTag.Boolean,
        value: true,
      });
    });

    controller.dispose();
  });
});
