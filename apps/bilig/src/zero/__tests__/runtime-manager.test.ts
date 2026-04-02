import { SpreadsheetEngine } from "@bilig/core";
import { ValueTag } from "@bilig/protocol";
import { describe, expect, it } from "vitest";
import { WorkbookRuntimeManager } from "../../workbook-runtime/runtime-manager.js";
import { buildWorkbookSourceProjectionFromEngine } from "../projection.js";
import type { Queryable, WorkbookRuntimeMetadata, WorkbookRuntimeState } from "../store.js";

const noopDb: Queryable = {
  async query() {
    return { rows: [] };
  },
};

async function createRuntimeState(
  workbookName: string,
  mutate?: (engine: SpreadsheetEngine) => void,
): Promise<WorkbookRuntimeState> {
  const engine = new SpreadsheetEngine({
    workbookName,
    replicaId: `runtime-test:${workbookName}`,
  });
  await engine.ready();
  mutate?.(engine);
  return {
    snapshot: engine.exportSnapshot(),
    replicaSnapshot: engine.exportReplicaSnapshot(),
    headRevision: 0,
    calculatedRevision: 0,
    ownerUserId: "owner-1",
  };
}

describe("WorkbookRuntimeManager", () => {
  it("reuses a warm runtime while the workbook revision stays current", async () => {
    let metadata: WorkbookRuntimeMetadata = {
      headRevision: 0,
      calculatedRevision: 0,
      ownerUserId: "owner-1",
    };
    const initialState = await createRuntimeState("doc-1", (engine) => {
      engine.setCellValue("Sheet1", "A1", 7);
    });
    let loadStateCalls = 0;

    const manager = new WorkbookRuntimeManager({
      loadMetadata: async () => metadata,
      loadState: async () => {
        loadStateCalls += 1;
        return initialState;
      },
    });

    const runtime = await manager.loadRuntime(noopDb, "doc-1");
    expect(loadStateCalls).toBe(1);
    expect(runtime.engine.getCell("Sheet1", "A1").value).toEqual({
      tag: ValueTag.Number,
      value: 7,
    });

    runtime.engine.setCellValue("Sheet1", "A1", 42);
    metadata = {
      headRevision: 1,
      calculatedRevision: 0,
      ownerUserId: "owner-1",
    };
    manager.commitMutation("doc-1", {
      projectionCommit: {
        kind: "replace",
        projection: buildWorkbookSourceProjectionFromEngine("doc-1", runtime.engine, {
          revision: 1,
          calculatedRevision: 0,
          ownerUserId: "owner-1",
          updatedBy: "owner-1",
          updatedAt: "2026-04-02T08:00:00.000Z",
        }),
      },
      headRevision: 1,
      calculatedRevision: 0,
      ownerUserId: "owner-1",
    });

    const reused = await manager.loadRuntime(noopDb, "doc-1");
    expect(reused).toBe(runtime);
    expect(loadStateCalls).toBe(1);
    expect(reused.engine.getCell("Sheet1", "A1").value).toEqual({
      tag: ValueTag.Number,
      value: 42,
    });
  });

  it("reloads from durable state when the cached revision falls behind", async () => {
    let metadata: WorkbookRuntimeMetadata = {
      headRevision: 0,
      calculatedRevision: 0,
      ownerUserId: "owner-1",
    };
    let state = await createRuntimeState("doc-2", (engine) => {
      engine.setCellValue("Sheet1", "A1", 1);
    });
    let loadStateCalls = 0;

    const manager = new WorkbookRuntimeManager({
      loadMetadata: async () => metadata,
      loadState: async () => {
        loadStateCalls += 1;
        return state;
      },
    });

    const first = await manager.loadRuntime(noopDb, "doc-2");
    expect(first.engine.getCell("Sheet1", "A1").value).toEqual({
      tag: ValueTag.Number,
      value: 1,
    });

    state = {
      ...(await createRuntimeState("doc-2", (engine) => {
        engine.setCellValue("Sheet1", "B2", 9);
      })),
      headRevision: 3,
      calculatedRevision: 2,
      ownerUserId: "owner-2",
    };
    metadata = {
      headRevision: 3,
      calculatedRevision: 2,
      ownerUserId: "owner-2",
    };

    const reloaded = await manager.loadRuntime(noopDb, "doc-2");
    expect(reloaded).not.toBe(first);
    expect(loadStateCalls).toBe(2);
    expect(reloaded.engine.getCell("Sheet1", "B2").value).toEqual({
      tag: ValueTag.Number,
      value: 9,
    });
  });

  it("serializes same-document work while allowing tasks to complete in order", async () => {
    const manager = new WorkbookRuntimeManager();
    const steps: string[] = [];

    const first = manager.runExclusive("doc-3", async () => {
      steps.push("first:start");
      await new Promise((resolve) => {
        setTimeout(resolve, 10);
      });
      steps.push("first:end");
    });

    const second = manager.runExclusive("doc-3", async () => {
      steps.push("second:start");
      steps.push("second:end");
    });

    await Promise.all([first, second]);
    expect(steps).toEqual(["first:start", "first:end", "second:start", "second:end"]);
  });
});
