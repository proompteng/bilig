import { performance } from "node:perf_hooks";
import { SpreadsheetEngine } from "../packages/core/src/engine.js";
import { formatAddress } from "../packages/formula/src/addressing.js";
import {
  buildWorkbookBenchmarkCorpus,
  isWorkbookBenchmarkCorpusId,
  type WorkbookBenchmarkCorpusFamily,
  type WorkbookBenchmarkCorpusId,
} from "../packages/benchmarks/src/workbook-corpus.js";
import { createMemoryWorkbookLocalStoreFactory } from "../packages/storage-browser/src/index.js";
import { decodeViewportPatch } from "../packages/worker-transport/src/index.js";
import { buildWorkbookSnapshot } from "../packages/benchmarks/src/generate-workbook.js";
import {
  measureMemory,
  sampleMemory,
  type MemoryMeasurement,
} from "../packages/benchmarks/src/metrics.js";
import type { AuthoritativeWorkbookEventRecord } from "../packages/zero-sync/src/index.js";
import { buildWorkbookLocalAuthoritativeBase } from "../apps/web/src/worker-local-base.js";
import { buildWorkbookLocalProjectionOverlay } from "../apps/web/src/worker-local-overlay.js";
import { WorkbookWorkerRuntime } from "../apps/web/src/worker-runtime.js";

interface WorkerWarmStartBenchmarkResult {
  scenario: "worker-warm-start";
  materializedCells: number;
  corpusCaseId: WorkbookBenchmarkCorpusId | null;
  corpusFamily: WorkbookBenchmarkCorpusFamily | null;
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

interface WorkerReconnectCatchUpBenchmarkResult {
  scenario: "worker-reconnect-catch-up";
  materializedCells: number;
  pendingMutationCount: number;
  rebaseMs: number;
  submitDrainMs: number;
  ackMs: number;
  catchUpMs: number;
  finalPendingMutationCount: number;
}

interface PersistedWorkerSeed {
  readonly documentId: string;
  readonly localStoreFactory: ReturnType<typeof createMemoryWorkbookLocalStoreFactory>;
  readonly materializedCells: number;
  readonly corpusCaseId: WorkbookBenchmarkCorpusId | null;
  readonly corpusFamily: WorkbookBenchmarkCorpusFamily | null;
  readonly viewport: {
    readonly sheetName: string;
    readonly rowStart: number;
    readonly rowEnd: number;
    readonly colStart: number;
    readonly colEnd: number;
  };
}

async function runSequentially<T>(
  items: readonly T[],
  task: (item: T, index: number) => Promise<void>,
): Promise<void> {
  await items.reduce<Promise<void>>((previous, item, index) => {
    return previous.then(() => task(item, index));
  }, Promise.resolve());
}

function buildSetCellValueEvent(input: {
  revision: number;
  address: string;
  value: number;
  clientMutationId?: string | null;
}): AuthoritativeWorkbookEventRecord {
  return {
    revision: input.revision,
    clientMutationId: input.clientMutationId ?? null,
    payload: {
      kind: "setCellValue",
      sheetName: "Sheet1",
      address: input.address,
      value: input.value,
    },
  };
}

function hasNumericCellValue(
  runtime: WorkbookWorkerRuntime,
  address: string,
  value: number,
): boolean {
  const snapshotValue = runtime.getCell("Sheet1", address).value;
  return (
    "value" in snapshotValue &&
    typeof snapshotValue.value === "number" &&
    snapshotValue.value === value
  );
}

async function seedWorkerLocalStore(
  documentId: string,
  materializedCells: number,
): Promise<PersistedWorkerSeed> {
  const viewport = {
    sheetName: "Sheet1",
    rowStart: 0,
    rowEnd: 39,
    colStart: 0,
    colEnd: 1,
  } as const;
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

  return {
    documentId,
    localStoreFactory,
    materializedCells,
    corpusCaseId: null,
    corpusFamily: null,
    viewport,
  };
}

async function seedWorkerLocalStoreFromCorpus(
  documentId: string,
  corpusId: WorkbookBenchmarkCorpusId,
): Promise<PersistedWorkerSeed> {
  const corpus = buildWorkbookBenchmarkCorpus(corpusId);
  const localStoreFactory = createMemoryWorkbookLocalStoreFactory();
  const seedEngine = new SpreadsheetEngine({ workbookName: documentId, replicaId: "seed" });
  await seedEngine.ready();
  seedEngine.importSnapshot(corpus.snapshot);

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

  return {
    documentId,
    localStoreFactory,
    materializedCells: corpus.materializedCellCount,
    corpusCaseId: corpus.id,
    corpusFamily: corpus.family,
    viewport: corpus.primaryViewport,
  };
}

export async function runWorkerWarmStartBenchmark(
  input: number | WorkbookBenchmarkCorpusId = 100_000,
): Promise<WorkerWarmStartBenchmarkResult> {
  const seed =
    typeof input === "string"
      ? await seedWorkerLocalStoreFromCorpus(
          `worker-warm-start-${input}-${Date.now()}-${Math.random().toString(36).slice(2)}`,
          input,
        )
      : await seedWorkerLocalStore(
          `worker-warm-start-${String(input)}-${Date.now()}-${Math.random().toString(36).slice(2)}`,
          input,
        );
  const runtime = new WorkbookWorkerRuntime({ localStoreFactory: seed.localStoreFactory });

  const memoryBefore = sampleMemory();
  const started = performance.now();
  await runtime.bootstrap({
    documentId: seed.documentId,
    replicaId: "browser:test",
    persistState: true,
  });

  let viewportCellCount = 0;
  runtime.subscribeViewportPatches(seed.viewport, (bytes) => {
    viewportCellCount = decodeViewportPatch(bytes).cells.length;
  });

  const elapsedMs = performance.now() - started;
  const memoryAfter = sampleMemory();
  runtime.dispose();

  return {
    scenario: "worker-warm-start",
    materializedCells: seed.materializedCells,
    corpusCaseId: seed.corpusCaseId,
    corpusFamily: seed.corpusFamily,
    elapsedMs,
    viewportCellCount,
    memory: measureMemory(memoryBefore, memoryAfter),
  };
}

export async function runWorkerVisibleEditBenchmark(
  materializedCells = 10_000,
): Promise<WorkerVisibleEditBenchmarkResult> {
  const documentId = `worker-visible-edit-${materializedCells}-${Date.now()}-${Math.random().toString(36).slice(2)}`;
  const seed = await seedWorkerLocalStore(documentId, materializedCells);
  const runtime = new WorkbookWorkerRuntime({ localStoreFactory: seed.localStoreFactory });

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

export async function runWorkerReconnectCatchUpBenchmark(
  materializedCells = 10_000,
  pendingMutationCount = 100,
): Promise<WorkerReconnectCatchUpBenchmarkResult> {
  const documentId = `worker-reconnect-${materializedCells}-${pendingMutationCount}-${Date.now()}-${Math.random().toString(36).slice(2)}`;
  const seed = await seedWorkerLocalStore(documentId, materializedCells);
  const runtime = new WorkbookWorkerRuntime({ localStoreFactory: seed.localStoreFactory });

  await runtime.bootstrap({
    documentId,
    replicaId: "browser:test",
    persistState: true,
  });

  const pendingMutations = new Array<Awaited<ReturnType<typeof runtime.enqueuePendingMutation>>>();
  const pendingIndexes = Array.from({ length: pendingMutationCount }, (_, index) => index);
  await runSequentially(pendingIndexes, async (index) => {
    const localValue = 200_000 + index;
    const pending = await runtime.enqueuePendingMutation({
      method: "setCellValue",
      args: ["Sheet1", formatAddress(index, 0), localValue],
    });
    pendingMutations.push(pending);
  });

  runtime.subscribeViewportPatches(
    {
      sheetName: "Sheet1",
      rowStart: 0,
      rowEnd: 127,
      colStart: 0,
      colEnd: 1,
    },
    () => {},
  );

  const remoteEvents = pendingIndexes.map((index) =>
    buildSetCellValueEvent({
      revision: index + 1,
      address: formatAddress(index, 1),
      value: 300_000 + index,
    }),
  );

  const catchUpStarted = performance.now();
  const rebaseStarted = catchUpStarted;
  await runtime.applyAuthoritativeEvents(remoteEvents, remoteEvents.length);
  const rebaseMs = performance.now() - rebaseStarted;

  if (runtime.listPendingMutations().length !== pendingMutationCount) {
    throw new Error("Reconnect rebase dropped pending mutations before authoritative ack");
  }
  if (!hasNumericCellValue(runtime, "A1", 200_000)) {
    throw new Error("Reconnect rebase lost the first local pending value");
  }
  if (
    !hasNumericCellValue(
      runtime,
      formatAddress(pendingMutationCount - 1, 0),
      200_000 + pendingMutationCount - 1,
    )
  ) {
    throw new Error("Reconnect rebase lost the trailing local pending value");
  }
  if (!hasNumericCellValue(runtime, "B1", 300_000)) {
    throw new Error("Reconnect rebase failed to apply the first authoritative drift value");
  }
  if (
    !hasNumericCellValue(
      runtime,
      formatAddress(pendingMutationCount - 1, 1),
      300_000 + pendingMutationCount - 1,
    )
  ) {
    throw new Error("Reconnect rebase failed to apply the trailing authoritative drift value");
  }

  const submitDrainStarted = performance.now();
  await runSequentially(pendingMutations, async (mutation) => {
    await runtime.markPendingMutationSubmitted(mutation.id);
  });
  const submitDrainMs = performance.now() - submitDrainStarted;

  const ackEvents = pendingMutations.map((mutation, index) =>
    buildSetCellValueEvent({
      revision: remoteEvents.length + index + 1,
      address: formatAddress(index, 0),
      value: 200_000 + index,
      clientMutationId: mutation.id,
    }),
  );
  const ackStarted = performance.now();
  await runtime.applyAuthoritativeEvents(ackEvents, remoteEvents.length + ackEvents.length);
  const ackMs = performance.now() - ackStarted;
  const catchUpMs = performance.now() - catchUpStarted;

  const finalPendingMutationCount = runtime.listPendingMutations().length;
  runtime.dispose();

  if (finalPendingMutationCount !== 0) {
    throw new Error("Reconnect catch-up did not absorb all pending mutations");
  }

  return {
    scenario: "worker-reconnect-catch-up",
    materializedCells,
    pendingMutationCount,
    rebaseMs,
    submitDrainMs,
    ackMs,
    catchUpMs,
    finalPendingMutationCount,
  };
}

if (import.meta.url === `file://${process.argv[1]}`) {
  const benchmark = process.argv[2] ?? "warm-start";

  if (benchmark !== "warm-start") {
    throw new Error(`Unknown worker benchmark: ${benchmark}`);
  }

  const rawInput = process.argv[3] ?? "100000";
  const input: number | WorkbookBenchmarkCorpusId = /^\d+$/.test(rawInput)
    ? Number.parseInt(rawInput, 10)
    : isWorkbookBenchmarkCorpusId(rawInput)
      ? rawInput
      : (() => {
          throw new Error(`Unknown workbook benchmark corpus: ${rawInput}`);
        })();

  console.log(JSON.stringify(await runWorkerWarmStartBenchmark(input), null, 2));
}
