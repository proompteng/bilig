import { performance } from "node:perf_hooks";
import { SpreadsheetEngine } from "../packages/core/src/engine.js";
import { createMemoryWorkbookLocalStoreFactory } from "../packages/storage-browser/src/index.js";
import { decodeViewportPatch } from "../packages/worker-transport/src/index.js";
import { buildWorkbookSnapshot } from "../packages/benchmarks/src/generate-workbook.js";
import {
  measureMemory,
  sampleMemory,
  type MemoryMeasurement,
} from "../packages/benchmarks/src/metrics.js";
import { buildWorkbookLocalAuthoritativeBase } from "../apps/web/src/worker-local-base.js";
import { buildWorkbookLocalProjectionOverlay } from "../apps/web/src/worker-local-overlay.js";
import { WorkbookWorkerRuntime } from "../apps/web/src/worker-runtime.js";

interface WorkerWarmStartBenchmarkResult {
  scenario: "worker-warm-start";
  materializedCells: number;
  elapsedMs: number;
  viewportCellCount: number;
  memory: MemoryMeasurement;
}

interface WorkerVisibleEditBenchmarkResult {
  scenario: "worker-visible-edit";
  materializedCells: number;
  visiblePatchMs: number;
  commitMs: number;
}

async function seedWorkerLocalStore(documentId: string, materializedCells: number) {
  const localStoreFactory = createMemoryWorkbookLocalStoreFactory();
  const seedEngine = new SpreadsheetEngine({ workbookName: documentId, replicaId: "seed" });
  await seedEngine.ready();
  seedEngine.importSnapshot(buildWorkbookSnapshot(materializedCells));

  const store = await localStoreFactory.open(documentId);
  await store.persistProjectionState({
    state: {
      snapshot: seedEngine.exportSnapshot(),
      replica: seedEngine.exportReplicaSnapshot(),
      authoritativeRevision: 0,
      appliedPendingLocalSeq: 0,
    },
    authoritativeBase: buildWorkbookLocalAuthoritativeBase(seedEngine),
    projectionOverlay: buildWorkbookLocalProjectionOverlay({
      authoritativeEngine: seedEngine,
      projectionEngine: seedEngine,
    }),
  });
  store.close();

  return localStoreFactory;
}

export async function runWorkerWarmStartBenchmark(
  materializedCells = 100_000,
): Promise<WorkerWarmStartBenchmarkResult> {
  const documentId = `worker-warm-start-${materializedCells}-${Date.now()}-${Math.random().toString(36).slice(2)}`;
  const localStoreFactory = await seedWorkerLocalStore(documentId, materializedCells);
  const runtime = new WorkbookWorkerRuntime({ localStoreFactory });

  const memoryBefore = sampleMemory();
  const started = performance.now();
  await runtime.bootstrap({
    documentId,
    replicaId: "browser:test",
    persistState: true,
  });

  let viewportCellCount = 0;
  runtime.subscribeViewportPatches(
    {
      sheetName: "Sheet1",
      rowStart: 0,
      rowEnd: 39,
      colStart: 0,
      colEnd: 1,
    },
    (bytes) => {
      viewportCellCount = decodeViewportPatch(bytes).cells.length;
    },
  );

  const elapsedMs = performance.now() - started;
  const memoryAfter = sampleMemory();
  runtime.dispose();

  return {
    scenario: "worker-warm-start",
    materializedCells,
    elapsedMs,
    viewportCellCount,
    memory: measureMemory(memoryBefore, memoryAfter),
  };
}

export async function runWorkerVisibleEditBenchmark(
  materializedCells = 10_000,
): Promise<WorkerVisibleEditBenchmarkResult> {
  const documentId = `worker-visible-edit-${materializedCells}-${Date.now()}-${Math.random().toString(36).slice(2)}`;
  const localStoreFactory = await seedWorkerLocalStore(documentId, materializedCells);
  const runtime = new WorkbookWorkerRuntime({ localStoreFactory });

  await runtime.bootstrap({
    documentId,
    replicaId: "browser:test",
    persistState: true,
  });

  let firstPatch = true;
  let visiblePatchStarted = 0;
  let resolveVisiblePatch: ((value: number) => void) | null = null;
  const visiblePatchPromise = new Promise<number>((resolve) => {
    resolveVisiblePatch = resolve;
  });

  runtime.subscribeViewportPatches(
    {
      sheetName: "Sheet1",
      rowStart: 0,
      rowEnd: 39,
      colStart: 0,
      colEnd: 1,
    },
    (bytes) => {
      const patch = decodeViewportPatch(bytes);
      if (firstPatch) {
        firstPatch = false;
        return;
      }
      if (
        patch.cells.some((cell) => cell.snapshot.address === "A1" && cell.displayText === "999")
      ) {
        resolveVisiblePatch?.(performance.now() - visiblePatchStarted);
        resolveVisiblePatch = null;
      }
    },
  );

  visiblePatchStarted = performance.now();
  const mutationStarted = visiblePatchStarted;
  const mutationPromise = runtime.enqueuePendingMutation({
    method: "setCellValue",
    args: ["Sheet1", "A1", 999],
  });
  const visiblePatchMs = await visiblePatchPromise;
  await mutationPromise;
  const commitMs = performance.now() - mutationStarted;
  runtime.dispose();

  return {
    scenario: "worker-visible-edit",
    materializedCells,
    visiblePatchMs,
    commitMs,
  };
}
