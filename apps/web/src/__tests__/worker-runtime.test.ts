import { describe, expect, it } from "vitest";
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
    expect(patch?.styles).toHaveLength(1);
    expect(patch?.styles[0]).toMatchObject({
      fill: { backgroundColor: "#336699" },
      font: { family: "Fira Sans" },
    });
    expect(patch?.cells[0]?.styleId).toBe(patch?.styles[0]?.id);
  });
});
