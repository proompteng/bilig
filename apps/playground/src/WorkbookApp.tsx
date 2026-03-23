import { startTransition, useCallback, useEffect, useMemo, useRef, useState } from "react";
import { SpreadsheetEngine, type CommitOp } from "@bilig/core";
import type { EngineOpBatch } from "@bilig/crdt";
import { createWorkbookRendererRoot } from "@bilig/renderer";
import {
  DependencyInspector,
  MetricsPanel,
  ReplicaPanel,
  WorkbookView,
  type EditMovement,
  type EditSelectionBehavior,
  useCell,
  useMetrics,
  useSelection,
} from "@bilig/grid";
import { formatAddress, parseCellAddress } from "@bilig/formula";
import type { LiteralInput, SyncState } from "@bilig/protocol";
import { MAX_COLS, MAX_ROWS, ValueTag, formatErrorCode } from "@bilig/protocol";
import { compactRelayEntries, type RelayEntry } from "./relay-queue.js";
import {
  PLAYGROUND_PRESETS,
  loadPlaygroundPreset,
  type PlaygroundPresetDefinition,
  type PlaygroundPresetId,
} from "./playgroundPresets.js";
import { loadPersistedJson, removePersistedJson, savePersistedJson } from "./browserPersistence.js";
import { useRemoteSpreadsheetSync } from "./useRemoteSpreadsheetSync.js";

const PRIMARY_STORAGE_KEY = "bilig:playground:primary";
const MIRROR_STORAGE_KEY = "bilig:playground:mirror";
const RELAY_STORAGE_KEY = "bilig:playground:relay";

interface PersistedReplicaState {
  snapshot: ReturnType<SpreadsheetEngine["exportSnapshot"]>;
  replica: ReturnType<SpreadsheetEngine["exportReplicaSnapshot"]>;
}

interface PersistedRelayState {
  syncPaused: boolean;
  queue: RelayEntry[];
}

type EditingMode = "idle" | "cell" | "formula";
export type WorkbookAppVariant = "playground" | "product";

export interface WorkbookAppProps {
  variant?: WorkbookAppVariant;
}

type ParsedEditorInput =
  | { kind: "clear" }
  | { kind: "formula"; formula: string }
  | { kind: "value"; value: LiteralInput };

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function isPersistedReplicaState(value: unknown): value is PersistedReplicaState {
  return (
    isRecord(value) &&
    isRecord(value["snapshot"]) &&
    Array.isArray(value["snapshot"]["sheets"]) &&
    isRecord(value["replica"]) &&
    isRecord(value["replica"]["replica"]) &&
    typeof value["replica"]["replica"]["replicaId"] === "string" &&
    Array.isArray(value["replica"]["entityVersions"]) &&
    Array.isArray(value["replica"]["sheetDeleteVersions"])
  );
}

function isPersistedRelayState(value: unknown): value is PersistedRelayState {
  return (
    isRecord(value) && typeof value["syncPaused"] === "boolean" && Array.isArray(value["queue"])
  );
}

function parsePersistedReplicaState(value: unknown): PersistedReplicaState | null {
  return isPersistedReplicaState(value) ? value : null;
}

function parsePersistedRelayState(value: unknown): PersistedRelayState | null {
  return isPersistedRelayState(value) ? value : null;
}

function toEditorValue(cell: ReturnType<typeof useCell>) {
  if (cell.formula) {
    return `=${cell.formula}`;
  }
  if (cell.input === null || cell.input === undefined) {
    return toResolvedValue(cell);
  }
  if (typeof cell.input === "boolean") {
    return cell.input ? "TRUE" : "FALSE";
  }
  return String(cell.input);
}

function toResolvedValue(cell: ReturnType<typeof useCell>) {
  switch (cell.value.tag) {
    case ValueTag.Number:
      return String(cell.value.value);
    case ValueTag.Boolean:
      return cell.value.value ? "TRUE" : "FALSE";
    case ValueTag.String:
      return cell.value.value;
    case ValueTag.Error:
      return formatErrorCode(cell.value.code);
    case ValueTag.Empty:
      return "";
  }
}

function parseEditorInput(rawValue: string): ParsedEditorInput {
  const normalized = rawValue.trim();
  if (normalized.startsWith("=")) {
    return { kind: "formula", formula: normalized.slice(1) };
  }
  if (normalized === "") {
    return { kind: "clear" };
  }
  if (normalized === "TRUE" || normalized === "FALSE") {
    return { kind: "value", value: normalized === "TRUE" };
  }
  const numeric = Number(normalized);
  if (!Number.isNaN(numeric) && /^-?\d+(\.\d+)?$/.test(normalized)) {
    return { kind: "value", value: numeric };
  }
  return { kind: "value", value: normalized };
}

function formatSyncStateLabel(state: SyncState): string {
  switch (state) {
    case "live":
      return "Live";
    case "syncing":
      return "Syncing";
    case "local-only":
      return "Local";
    case "behind":
      return "Behind";
    case "reconnecting":
      return "Reconnecting";
  }
  const exhaustiveState: never = state;
  throw new Error(`Unsupported sync state: ${String(exhaustiveState)}`);
}

function applyParsedInput(
  engine: SpreadsheetEngine,
  sheetName: string,
  address: string,
  parsed: ParsedEditorInput,
): void {
  if (parsed.kind === "formula") {
    engine.setCellFormula(sheetName, address, parsed.formula);
    return;
  }
  if (parsed.kind === "clear") {
    engine.clearCell(sheetName, address);
    return;
  }
  engine.setCellValue(sheetName, address, parsed.value);
}

function toCommitOp(sheetName: string, address: string, rawValue: string): CommitOp {
  const parsed = parseEditorInput(rawValue);
  if (parsed.kind === "formula") {
    return { kind: "upsertCell", sheetName, addr: address, formula: parsed.formula };
  }
  if (parsed.kind === "clear") {
    return { kind: "deleteCell", sheetName, addr: address };
  }
  return { kind: "upsertCell", sheetName, addr: address, value: parsed.value };
}

function clampSelectionMovement(
  address: string,
  sheetName: string,
  movement: EditMovement,
): string {
  const parsed = parseCellAddress(address, sheetName);
  const nextRow = Math.min(MAX_ROWS - 1, Math.max(0, parsed.row + movement[1]));
  const nextCol = Math.min(MAX_COLS - 1, Math.max(0, parsed.col + movement[0]));
  return formatAddress(nextRow, nextCol);
}

function parseSelectionTarget(
  input: string,
  fallbackSheet: string,
): { sheetName: string; address: string } | null {
  const trimmed = input.trim();
  if (trimmed.length === 0) {
    return null;
  }

  const bangIndex = trimmed.lastIndexOf("!");
  const nextSheetName = bangIndex === -1 ? fallbackSheet : trimmed.slice(0, bangIndex);
  const nextAddress = bangIndex === -1 ? trimmed : trimmed.slice(bangIndex + 1);

  try {
    const parsed = parseCellAddress(nextAddress.toUpperCase(), nextSheetName || fallbackSheet);
    return {
      sheetName: nextSheetName || fallbackSheet,
      address: formatAddress(parsed.row, parsed.col),
    };
  } catch {
    return null;
  }
}

async function waitForTaskCycles(count = 2): Promise<void> {
  await advanceTaskCycles(count);
}

async function advanceTaskCycles(remaining: number): Promise<void> {
  if (remaining <= 0) {
    return;
  }
  await new Promise<void>((resolve) => {
    window.setTimeout(() => resolve(), 0);
  });
  await advanceTaskCycles(remaining - 1);
}

async function settleRendererBoundary(operation: Promise<void>): Promise<void> {
  let capturedError: unknown = null;
  void operation.catch((error) => {
    capturedError = error;
  });
  await waitForTaskCycles();
  if (capturedError) {
    throw capturedError instanceof Error
      ? capturedError
      : new Error(
          typeof capturedError === "string" ? capturedError : JSON.stringify(capturedError),
        );
  }
}

export function WorkbookApp({ variant = "playground" }: WorkbookAppProps) {
  const isProductShell = variant === "product";
  const runtimeConfig = useMemo(() => {
    const defaultDocumentId = import.meta.env["VITE_BILIG_DOCUMENT_ID"] ?? "bilig-demo";
    const defaultLocalServerUrl =
      import.meta.env["VITE_BILIG_LOCAL_SERVER_URL"] ?? "http://127.0.0.1:4381";
    if (!isProductShell) {
      return {
        documentId: defaultDocumentId,
        localServerUrl: defaultLocalServerUrl,
      };
    }
    const searchParams = new URLSearchParams(window.location.search);
    return {
      documentId: searchParams.get("document") ?? defaultDocumentId,
      localServerUrl: searchParams.get("server") ?? defaultLocalServerUrl,
    };
  }, [isProductShell]);
  const documentId = runtimeConfig.documentId;
  const localServerUrl = runtimeConfig.localServerUrl;
  const replicaId = useMemo(
    () => (isProductShell ? `browser:${Math.random().toString(36).slice(2)}` : "playground"),
    [isProductShell],
  );
  const engine = useMemo(
    () => new SpreadsheetEngine({ workbookName: documentId, replicaId }),
    [documentId, replicaId],
  );
  const mirrorEngine = useMemo(
    () => new SpreadsheetEngine({ workbookName: documentId, replicaId: "replica-beta" }),
    [documentId],
  );
  const rendererRoot = useMemo(() => createWorkbookRendererRoot(engine), [engine]);
  const selection = useSelection(engine);
  const selectCell = selection.select;
  const selectedAddr = selection.address ?? "A1";
  const selectedCell = useCell(engine, selection.sheetName, selectedAddr);
  const mirroredSelectedCell = useCell(mirrorEngine, selection.sheetName, selectedAddr);
  const metrics = useMetrics(engine);
  const mirrorMetrics = useMetrics(mirrorEngine);
  const [editorValue, setEditorValue] = useState("");
  const [editorSelectionBehavior, setEditorSelectionBehavior] =
    useState<EditSelectionBehavior>("select-all");
  const [editingMode, setEditingMode] = useState<EditingMode>("idle");
  const [selectionLabel, setSelectionLabel] = useState(selectedAddr);
  const [replicationReady, setReplicationReady] = useState(false);
  const [syncPaused, setSyncPaused] = useState(false);
  const [relayQueue, setRelayQueue] = useState<RelayEntry[]>([]);
  const [activePresetId, setActivePresetId] = useState<PlaygroundPresetId | null>("starter");
  const [loadingPresetId, setLoadingPresetId] = useState<PlaygroundPresetId | null>(null);
  const [presetError, setPresetError] = useState<string | null>(null);
  const syncLatencyMs = 120;
  const relayQueueRef = useRef<RelayEntry[]>([]);
  const relayTimerRef = useRef<number | null>(null);
  const captureRelayBatchesRef = useRef(false);
  const syncMirrorFromPrimary = useCallback(() => {
    mirrorEngine.importSnapshot(engine.exportSnapshot());
    const replicaSnapshot = engine.exportReplicaSnapshot();
    mirrorEngine.importReplicaSnapshot({
      ...replicaSnapshot,
      replica: {
        ...replicaSnapshot.replica,
        replicaId: mirrorEngine.replica.replicaId,
      },
    });
  }, [engine, mirrorEngine]);
  const sheetNames = [...engine.workbook.sheetsByName.values()]
    .toSorted((left, right) => left.order - right.order)
    .map((sheet) => sheet.name);
  const pendingSyncCount = syncPaused ? 0 : relayQueue.length;
  const queuedSyncCount = syncPaused ? relayQueue.length : 0;
  const selectedPreset = activePresetId
    ? (PLAYGROUND_PRESETS.find((preset) => preset.id === activePresetId) ?? null)
    : null;

  const resolvedValue = toResolvedValue(selectedCell);
  const mirroredValue = toResolvedValue(mirroredSelectedCell);
  const isEditing = editingMode !== "idle";
  const isEditingCell = editingMode === "cell";
  const visibleEditorValue = isEditing ? editorValue : toEditorValue(selectedCell);
  const remoteSyncState = useRemoteSpreadsheetSync({
    enabled: isProductShell && replicationReady,
    engine,
    documentId,
    replicaId,
    baseUrl: isProductShell ? localServerUrl : null,
  });

  const loadPreset = useCallback(
    async (presetId: PlaygroundPresetId, syncMirror = true) => {
      startTransition(() => {
        setLoadingPresetId(presetId);
        setPresetError(null);
      });
      await waitForTaskCycles(1);
      const previousCaptureState = captureRelayBatchesRef.current;
      captureRelayBatchesRef.current = false;

      try {
        const preset = await loadPlaygroundPreset(presetId);
        relayQueueRef.current = [];
        setRelayQueue([]);
        setEditingMode("idle");

        if (preset.kind === "renderer") {
          await settleRendererBoundary(rendererRoot.render(preset.element));
        } else {
          await settleRendererBoundary(rendererRoot.unmount());
          engine.importSnapshot(preset.snapshot);
        }

        if (syncMirror) {
          syncMirrorFromPrimary();
          captureRelayBatchesRef.current = true;
        }
        selectCell(preset.defaultSheet, preset.defaultAddress);
        setActivePresetId(presetId);
      } catch (error) {
        captureRelayBatchesRef.current = previousCaptureState;
        setPresetError(error instanceof Error ? error.message : String(error));
      } finally {
        setLoadingPresetId(null);
      }
    },
    [engine, rendererRoot, selectCell, syncMirrorFromPrimary],
  );

  useEffect(() => {
    relayQueueRef.current = relayQueue;
  }, [relayQueue]);

  useEffect(() => {
    let cancelled = false;

    void Promise.all([engine.ready(), mirrorEngine.ready()]).then(async () => {
      if (isProductShell) {
        if (engine.workbook.sheetsByName.size === 0) {
          engine.createSheet("Sheet1");
        }
        if (!cancelled) {
          setReplicationReady(true);
        }
        return undefined;
      }

      const resetWorkspace = new URLSearchParams(window.location.search).get("reset") === "1";
      if (resetWorkspace) {
        await Promise.all([
          removePersistedJson(PRIMARY_STORAGE_KEY),
          removePersistedJson(MIRROR_STORAGE_KEY),
          removePersistedJson(RELAY_STORAGE_KEY),
        ]);
        window.history.replaceState(null, "", window.location.pathname);
        relayQueueRef.current = [];
        if (!cancelled) {
          setRelayQueue([]);
          setSyncPaused(false);
        }
        await loadPreset("starter", false);
        syncMirrorFromPrimary();
        if (!cancelled) {
          captureRelayBatchesRef.current = true;
          setReplicationReady(true);
        }
        return undefined;
      }

      const [primaryPersisted, mirrorPersisted, relayPersisted] = await Promise.all([
        loadPersistedJson(PRIMARY_STORAGE_KEY, parsePersistedReplicaState),
        loadPersistedJson(MIRROR_STORAGE_KEY, parsePersistedReplicaState),
        loadPersistedJson(RELAY_STORAGE_KEY, parsePersistedRelayState),
      ]);

      if (primaryPersisted) {
        await settleRendererBoundary(rendererRoot.unmount());
        engine.importSnapshot(primaryPersisted.snapshot);
        engine.importReplicaSnapshot(primaryPersisted.replica);
        setActivePresetId(null);
      } else {
        await loadPreset("starter", false);
      }

      if (mirrorPersisted) {
        mirrorEngine.importSnapshot(mirrorPersisted.snapshot);
        mirrorEngine.importReplicaSnapshot(mirrorPersisted.replica);
      } else {
        syncMirrorFromPrimary();
      }

      if (relayPersisted) {
        relayQueueRef.current = relayPersisted.queue;
        if (!cancelled) {
          setRelayQueue(relayPersisted.queue);
          setSyncPaused(relayPersisted.syncPaused);
        }
      }

      if (!cancelled) {
        captureRelayBatchesRef.current = true;
        setReplicationReady(true);
      }
      return undefined;
    });

    return () => {
      cancelled = true;
      if (relayTimerRef.current !== null) {
        window.clearTimeout(relayTimerRef.current);
        relayTimerRef.current = null;
      }
      relayQueueRef.current = [];
      void rendererRoot.unmount();
    };
  }, [engine, isProductShell, loadPreset, mirrorEngine, rendererRoot, syncMirrorFromPrimary]);

  useEffect(() => {
    if (!replicationReady || isProductShell) {
      return;
    }

    const persistPrimary = () => {
      const persisted: PersistedReplicaState = {
        snapshot: engine.exportSnapshot(),
        replica: engine.exportReplicaSnapshot(),
      };
      void savePersistedJson(PRIMARY_STORAGE_KEY, persisted);
    };

    const persistMirror = () => {
      const persisted: PersistedReplicaState = {
        snapshot: mirrorEngine.exportSnapshot(),
        replica: mirrorEngine.exportReplicaSnapshot(),
      };
      void savePersistedJson(MIRROR_STORAGE_KEY, persisted);
    };

    persistPrimary();
    persistMirror();

    const unsubscribePrimary = engine.subscribe(() => persistPrimary());
    const unsubscribeMirror = mirrorEngine.subscribe(() => persistMirror());

    return () => {
      unsubscribePrimary();
      unsubscribeMirror();
    };
  }, [engine, isProductShell, mirrorEngine, replicationReady]);

  useEffect(() => {
    if (isProductShell) {
      return;
    }

    const enqueueBatch = (target: RelayEntry["target"], batch: EngineOpBatch) => {
      if (!captureRelayBatchesRef.current) {
        return;
      }
      setRelayQueue((current) =>
        compactRelayEntries([
          ...current,
          {
            target,
            batch,
            deliverAt: Date.now() + syncLatencyMs,
          },
        ]),
      );
    };

    const unsubscribeLocal = engine.subscribeBatches((batch) => enqueueBatch("mirror", batch));
    const unsubscribeMirror = mirrorEngine.subscribeBatches((batch) =>
      enqueueBatch("primary", batch),
    );

    return () => {
      unsubscribeLocal();
      unsubscribeMirror();
    };
  }, [engine, isProductShell, mirrorEngine]);

  useEffect(() => {
    if (!replicationReady || isProductShell) {
      return;
    }

    const persistedRelayState: PersistedRelayState = {
      syncPaused,
      queue: relayQueue,
    };
    void savePersistedJson(RELAY_STORAGE_KEY, persistedRelayState);
  }, [isProductShell, relayQueue, replicationReady, syncPaused]);

  useEffect(() => {
    if (!replicationReady || isProductShell) {
      return;
    }

    if (relayTimerRef.current !== null) {
      window.clearTimeout(relayTimerRef.current);
      relayTimerRef.current = null;
    }

    if (syncPaused || relayQueue.length === 0) {
      return;
    }

    const nextDeliverAt = relayQueue.reduce(
      (earliest, entry) => Math.min(earliest, entry.deliverAt),
      Number.POSITIVE_INFINITY,
    );
    const delay = Math.max(0, nextDeliverAt - Date.now());

    relayTimerRef.current = window.setTimeout(() => {
      relayTimerRef.current = null;
      const now = Date.now();
      const due: RelayEntry[] = [];
      const pending: RelayEntry[] = [];
      relayQueueRef.current.forEach((entry) => {
        if (entry.deliverAt <= now) {
          due.push(entry);
          return;
        }
        pending.push(entry);
      });
      relayQueueRef.current = pending;
      setRelayQueue(pending);
      due.forEach(({ target, batch }) => {
        (target === "mirror" ? mirrorEngine : engine).applyRemoteBatch(batch);
      });
    }, delay);

    return () => {
      if (relayTimerRef.current !== null) {
        window.clearTimeout(relayTimerRef.current);
        relayTimerRef.current = null;
      }
    };
  }, [engine, isProductShell, mirrorEngine, relayQueue, replicationReady, syncPaused]);

  useEffect(() => {
    if (sheetNames.length === 0) {
      return;
    }
    if (!sheetNames.includes(selection.sheetName)) {
      selection.select(sheetNames[0]!, "A1");
    }
  }, [selection, sheetNames]);

  useEffect(() => {
    if (isEditing) {
      return;
    }
    setEditorValue("");
    setEditorSelectionBehavior("select-all");
  }, [isEditing, selectedAddr, selection.sheetName]);

  const dependencySnapshot = engine.explainCell(selection.sheetName, selectedAddr);

  const beginEditing = useCallback(
    (
      seed?: string,
      selectionBehavior: EditSelectionBehavior = "select-all",
      mode: Exclude<EditingMode, "idle"> = "cell",
    ) => {
      setEditorValue(seed ?? toEditorValue(engine.getCell(selection.sheetName, selectedAddr)));
      setEditorSelectionBehavior(selectionBehavior);
      setEditingMode(mode);
    },
    [engine, selectedAddr, selection.sheetName],
  );

  const commitEditor = useCallback(
    (movement?: EditMovement) => {
      const nextValue = isEditing ? editorValue : toEditorValue(selectedCell);
      applyParsedInput(engine, selection.sheetName, selectedAddr, parseEditorInput(nextValue));
      setEditingMode("idle");
      setEditorSelectionBehavior("select-all");
      if (movement) {
        selection.select(
          selection.sheetName,
          clampSelectionMovement(selectedAddr, selection.sheetName, movement),
        );
      }
    },
    [editorValue, engine, isEditing, selectedAddr, selectedCell, selection],
  );

  const cancelEditor = useCallback(() => {
    setEditorValue(toEditorValue(selectedCell));
    setEditorSelectionBehavior("select-all");
    setEditingMode("idle");
  }, [selectedCell]);

  const clearSelectedCell = useCallback(() => {
    engine.clearCell(selection.sheetName, selectedAddr);
    setEditorValue("");
    setEditingMode("idle");
  }, [engine, selectedAddr, selection.sheetName]);

  const pasteIntoSelection = useCallback(
    (startAddr: string, values: readonly (readonly string[])[]) => {
      const start = parseCellAddress(startAddr, selection.sheetName);
      const ops: CommitOp[] = [];
      values.forEach((rowValues, rowOffset) => {
        rowValues.forEach((cellValue, colOffset) => {
          const address = formatAddress(start.row + rowOffset, start.col + colOffset);
          ops.push(toCommitOp(selection.sheetName, address, cellValue));
        });
      });
      if (ops.length > 0) {
        engine.renderCommit(ops);
      }
      setEditorSelectionBehavior("select-all");
      setEditingMode("idle");
    },
    [engine, selection.sheetName],
  );

  const fillSelectionRange = useCallback(
    (
      sourceStartAddr: string,
      sourceEndAddr: string,
      targetStartAddr: string,
      targetEndAddr: string,
    ) => {
      engine.fillRange(
        {
          sheetName: selection.sheetName,
          startAddress: sourceStartAddr,
          endAddress: sourceEndAddr,
        },
        {
          sheetName: selection.sheetName,
          startAddress: targetStartAddr,
          endAddress: targetEndAddr,
        },
      );
      setEditingMode("idle");
    },
    [engine, selection.sheetName],
  );

  const copySelectionRange = useCallback(
    (
      sourceStartAddr: string,
      sourceEndAddr: string,
      targetStartAddr: string,
      targetEndAddr: string,
    ) => {
      engine.copyRange(
        {
          sheetName: selection.sheetName,
          startAddress: sourceStartAddr,
          endAddress: sourceEndAddr,
        },
        {
          sheetName: selection.sheetName,
          startAddress: targetStartAddr,
          endAddress: targetEndAddr,
        },
      );
      setEditingMode("idle");
    },
    [engine, selection.sheetName],
  );

  const selectAddress = useCallback(
    (nextSheetName: string, nextAddress: string) => {
      if (editingMode === "formula") {
        const nextValue = editorValue;
        applyParsedInput(engine, selection.sheetName, selectedAddr, parseEditorInput(nextValue));
        setEditingMode("idle");
      }
      selection.select(nextSheetName, nextAddress);
    },
    [editingMode, editorValue, engine, selectedAddr, selection],
  );

  const resetWorkspace = () => {
    void Promise.all([
      removePersistedJson(PRIMARY_STORAGE_KEY),
      removePersistedJson(MIRROR_STORAGE_KEY),
      removePersistedJson(RELAY_STORAGE_KEY),
    ]).finally(() => {
      window.location.reload();
    });
  };

  const toggleSync = useCallback(() => {
    if (syncPaused) {
      const queuedEntries = [...relayQueueRef.current].toSorted(
        (left, right) => left.deliverAt - right.deliverAt,
      );
      queuedEntries.forEach(({ target, batch }) => {
        (target === "mirror" ? mirrorEngine : engine).applyRemoteBatch(batch);
      });
      syncMirrorFromPrimary();
      relayQueueRef.current = [];
      setRelayQueue([]);
      setSyncPaused(false);
      return;
    }

    setSyncPaused(true);
  }, [engine, mirrorEngine, syncMirrorFromPrimary, syncPaused]);

  const statusBar = isProductShell ? (
    <>
      <span data-testid="status-mode">{formatSyncStateLabel(remoteSyncState)}</span>
      <span data-testid="status-selection">
        {selection.sheetName}!{selectionLabel}
      </span>
      <span data-testid="status-sync">{isEditing ? "Editing" : "Ready"}</span>
    </>
  ) : (
    <>
      <span data-testid="status-active-preset">
        {selectedPreset?.label ?? "Restored workspace"}
      </span>
      <span data-testid="metric-js">Fallback {metrics.jsFormulaCount.toLocaleString()}</span>
      <span data-testid="metric-wasm">WASM {metrics.wasmFormulaCount.toLocaleString()}</span>
      <span data-testid="metric-recalc">Recalc {metrics.recalcMs.toFixed(2)} ms</span>
    </>
  );

  const ribbon = isProductShell ? null : (
    <div className="ribbon-strip">
      <div className="ribbon-brand">
        <p className="eyebrow">bilig playground</p>
        <strong>Excel-like shell on top of the local-first engine</strong>
      </div>
      <div className="preset-strip" data-testid="preset-strip">
        {PLAYGROUND_PRESETS.map((preset: PlaygroundPresetDefinition) => (
          <button
            className={preset.id === activePresetId ? "preset-chip active" : "preset-chip"}
            data-testid={`preset-${preset.id}`}
            disabled={loadingPresetId !== null}
            key={preset.id}
            onClick={() => void loadPreset(preset.id)}
            type="button"
          >
            {preset.label}
          </button>
        ))}
      </div>
      <div className="ribbon-actions">
        <button className="ghost-button" onClick={() => void loadPreset("starter")} type="button">
          Rebuild starter
        </button>
        <button className="ghost-button" onClick={resetWorkspace} type="button">
          Reset local workspace
        </button>
      </div>
    </div>
  );

  const sidebar = isProductShell ? null : (
    <>
      <MetricsPanel metrics={metrics} />
      <ReplicaPanel
        latencyMs={syncLatencyMs}
        localReplicaId={engine.replica.replicaId}
        onToggleSync={toggleSync}
        pendingSyncCount={pendingSyncCount}
        queuedSyncCount={queuedSyncCount}
        remoteMetrics={mirrorMetrics}
        remoteReplicaId={mirrorEngine.replica.replicaId}
        remoteValue={mirroredValue}
        selectedLabel={`${selection.sheetName}!${selectedAddr}`}
        syncPaused={syncPaused}
      />
      <DependencyInspector snapshot={dependencySnapshot} />
    </>
  );

  return (
    <div className={isProductShell ? "app-shell app-shell-product" : "app-shell"}>
      {loadingPresetId ? (
        <div className="loading-banner" data-testid="preset-loading">
          Loading{" "}
          {PLAYGROUND_PRESETS.find((preset) => preset.id === loadingPresetId)?.label ??
            loadingPresetId}
          ...
        </div>
      ) : null}
      {presetError ? (
        <div className="error-banner" data-testid="preset-error">
          {presetError}
        </div>
      ) : null}
      {!isProductShell && !replicationReady ? (
        <div className="loading-banner" data-testid="replication-loading">
          Preparing local-first mirror...
        </div>
      ) : (
        <WorkbookView
          editorValue={visibleEditorValue}
          editorSelectionBehavior={editorSelectionBehavior}
          engine={engine}
          isEditing={isEditing}
          isEditingCell={isEditingCell}
          onAddressCommit={(input) => {
            const nextTarget = parseSelectionTarget(input, selection.sheetName);
            if (nextTarget) {
              selectAddress(nextTarget.sheetName, nextTarget.address);
            }
          }}
          onBeginEdit={beginEditing}
          onBeginFormulaEdit={(seed?: string) => beginEditing(seed, "select-all", "formula")}
          onCancelEdit={cancelEditor}
          onClearCell={clearSelectedCell}
          onCommitEdit={commitEditor}
          onCopyRange={copySelectionRange}
          onEditorChange={(next) => {
            setEditorValue(next);
            setEditingMode((current) => (current === "idle" ? "cell" : current));
          }}
          onFillRange={fillSelectionRange}
          onPaste={pasteIntoSelection}
          onSelectionLabelChange={setSelectionLabel}
          onSelect={(addr) => selectAddress(selection.sheetName, addr)}
          onSelectSheet={(sheetName) => selectAddress(sheetName, "A1")}
          resolvedValue={resolvedValue}
          ribbon={ribbon}
          selectedAddr={selectedAddr}
          sheetName={selection.sheetName}
          sheetNames={sheetNames}
          sidebar={sidebar}
          statusBar={statusBar}
          variant={variant}
          workbookName={engine.workbook.workbookName}
        />
      )}
    </div>
  );
}
