import React, { startTransition, useCallback, useEffect, useMemo, useRef, useState } from "react";
import { SpreadsheetEngine, type CommitOp } from "@bilig/core";
import type { EngineOpBatch } from "@bilig/crdt";
import { createWorkbookRendererRoot } from "@bilig/renderer";
import {
  DependencyInspector,
  MetricsPanel,
  ReplicaPanel,
  WorkbookView,
  type EditMovement,
  useCell,
  useMetrics,
  useSelection
} from "@bilig/grid";
import { formatAddress, parseCellAddress } from "@bilig/formula";
import type { LiteralInput } from "@bilig/protocol";
import { MAX_COLS, MAX_ROWS } from "@bilig/protocol";
import { compactRelayEntries, type RelayEntry } from "./relay-queue.js";
import {
  PLAYGROUND_PRESETS,
  loadPlaygroundPreset,
  type PlaygroundPresetDefinition,
  type PlaygroundPresetId
} from "./playgroundPresets.js";

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

type ParsedEditorInput =
  | { kind: "clear" }
  | { kind: "formula"; formula: string }
  | { kind: "value"; value: LiteralInput };

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
    case 1:
      return String(cell.value.value);
    case 2:
      return cell.value.value ? "TRUE" : "FALSE";
    case 3:
      return cell.value.value;
    case 4:
      return `#${cell.value.code}`;
    default:
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

function applyParsedInput(engine: SpreadsheetEngine, sheetName: string, address: string, parsed: ParsedEditorInput): void {
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

function clampSelectionMovement(address: string, sheetName: string, movement: EditMovement): string {
  const parsed = parseCellAddress(address, sheetName);
  const nextRow = Math.min(MAX_ROWS - 1, Math.max(0, parsed.row + movement[1]));
  const nextCol = Math.min(MAX_COLS - 1, Math.max(0, parsed.col + movement[0]));
  return formatAddress(nextRow, nextCol);
}

function parseSelectionTarget(input: string, fallbackSheet: string): { sheetName: string; address: string } | null {
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
      address: formatAddress(parsed.row, parsed.col)
    };
  } catch {
    return null;
  }
}

async function waitForTaskCycles(count = 2): Promise<void> {
  for (let remaining = count; remaining > 0; remaining -= 1) {
    await new Promise<void>((resolve) => {
      window.setTimeout(() => resolve(), 0);
    });
  }
}

async function settleRendererBoundary(operation: Promise<void>): Promise<void> {
  let capturedError: unknown = null;
  void operation.catch((error) => {
    capturedError = error;
  });
  await waitForTaskCycles();
  if (capturedError) {
    throw capturedError instanceof Error ? capturedError : new Error(String(capturedError));
  }
}

export function App() {
  const engine = useMemo(() => new SpreadsheetEngine({ workbookName: "bilig-demo", replicaId: "playground" }), []);
  const mirrorEngine = useMemo(() => new SpreadsheetEngine({ workbookName: "bilig-demo", replicaId: "replica-beta" }), []);
  const rendererRoot = useMemo(() => createWorkbookRendererRoot(engine), [engine]);
  const selection = useSelection(engine);
  const selectCell = selection.select;
  const selectedAddr = selection.address ?? "A1";
  const selectedCell = useCell(engine, selection.sheetName, selectedAddr);
  const mirroredSelectedCell = useCell(mirrorEngine, selection.sheetName, selectedAddr);
  const metrics = useMetrics(engine);
  const mirrorMetrics = useMetrics(mirrorEngine);
  const [editorValue, setEditorValue] = useState("");
  const [isEditingCell, setIsEditingCell] = useState(false);
  const [replicationReady, setReplicationReady] = useState(false);
  const [syncPaused, setSyncPaused] = useState(false);
  const [relayQueue, setRelayQueue] = useState<RelayEntry[]>([]);
  const [activePresetId, setActivePresetId] = useState<PlaygroundPresetId | null>("starter");
  const [loadingPresetId, setLoadingPresetId] = useState<PlaygroundPresetId | null>(null);
  const [presetError, setPresetError] = useState<string | null>(null);
  const syncLatencyMs = 120;
  const relayQueueRef = useRef<RelayEntry[]>([]);
  const relayTimerRef = useRef<number | null>(null);
  const sheetNames = [...engine.workbook.sheetsByName.values()]
    .sort((left, right) => left.order - right.order)
    .map((sheet) => sheet.name);
  const pendingSyncCount = syncPaused ? 0 : relayQueue.length;
  const queuedSyncCount = syncPaused ? relayQueue.length : 0;
  const selectedPreset = activePresetId
    ? PLAYGROUND_PRESETS.find((preset) => preset.id === activePresetId) ?? null
    : null;

  const resolvedValue = toResolvedValue(selectedCell);
  const mirroredValue = toResolvedValue(mirroredSelectedCell);
  const visibleEditorValue = isEditingCell ? editorValue : toEditorValue(selectedCell);

  const loadPreset = useCallback(
    async (presetId: PlaygroundPresetId) => {
      startTransition(() => {
        setLoadingPresetId(presetId);
        setPresetError(null);
      });
      await waitForTaskCycles(1);

      try {
        const preset = await loadPlaygroundPreset(presetId);
        relayQueueRef.current = [];
        setRelayQueue([]);
        setIsEditingCell(false);

        if (preset.kind === "renderer") {
          await settleRendererBoundary(rendererRoot.render(preset.element));
        } else {
          await settleRendererBoundary(rendererRoot.unmount());
          engine.importSnapshot(preset.snapshot);
        }

        mirrorEngine.importSnapshot(engine.exportSnapshot());
        selectCell(preset.defaultSheet, preset.defaultAddress);
        setActivePresetId(presetId);
      } catch (error) {
        setPresetError(error instanceof Error ? error.message : String(error));
      } finally {
        setLoadingPresetId(null);
      }
    },
    [engine, mirrorEngine, rendererRoot, selectCell]
  );

  useEffect(() => {
    relayQueueRef.current = relayQueue;
  }, [relayQueue]);

  useEffect(() => {
    let cancelled = false;

    void Promise.all([engine.ready(), mirrorEngine.ready()]).then(async () => {
      const resetWorkspace = new URLSearchParams(window.location.search).get("reset") === "1";
      if (resetWorkspace) {
        window.localStorage.removeItem(PRIMARY_STORAGE_KEY);
        window.localStorage.removeItem(MIRROR_STORAGE_KEY);
        window.localStorage.removeItem(RELAY_STORAGE_KEY);
        window.history.replaceState(null, "", window.location.pathname);
        relayQueueRef.current = [];
        if (!cancelled) {
          setRelayQueue([]);
          setSyncPaused(false);
        }
        await loadPreset("starter");
        mirrorEngine.importSnapshot(engine.exportSnapshot());
        if (!cancelled) {
          setReplicationReady(true);
        }
        return;
      }

      const primaryPersisted = window.localStorage.getItem(PRIMARY_STORAGE_KEY);
      const mirrorPersisted = window.localStorage.getItem(MIRROR_STORAGE_KEY);
      const relayPersisted = window.localStorage.getItem(RELAY_STORAGE_KEY);

      if (primaryPersisted) {
        const restored = JSON.parse(primaryPersisted) as PersistedReplicaState;
        await settleRendererBoundary(rendererRoot.unmount());
        engine.importSnapshot(restored.snapshot);
        engine.importReplicaSnapshot(restored.replica);
        setActivePresetId(null);
      } else {
        await loadPreset("starter");
      }

      if (mirrorPersisted) {
        const restoredMirror = JSON.parse(mirrorPersisted) as PersistedReplicaState;
        mirrorEngine.importSnapshot(restoredMirror.snapshot);
        mirrorEngine.importReplicaSnapshot(restoredMirror.replica);
      } else {
        mirrorEngine.importSnapshot(engine.exportSnapshot());
      }

      if (relayPersisted) {
        const restoredRelay = JSON.parse(relayPersisted) as PersistedRelayState;
        relayQueueRef.current = restoredRelay.queue;
        if (!cancelled) {
          setRelayQueue(restoredRelay.queue);
          setSyncPaused(restoredRelay.syncPaused);
        }
      }

      if (!cancelled) {
        setReplicationReady(true);
      }
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
  }, [engine, loadPreset, mirrorEngine, rendererRoot]);

  useEffect(() => {
    if (!replicationReady) {
      return;
    }

    const persistPrimary = () => {
      const persisted: PersistedReplicaState = {
        snapshot: engine.exportSnapshot(),
        replica: engine.exportReplicaSnapshot()
      };
      window.localStorage.setItem(PRIMARY_STORAGE_KEY, JSON.stringify(persisted));
    };

    const persistMirror = () => {
      const persisted: PersistedReplicaState = {
        snapshot: mirrorEngine.exportSnapshot(),
        replica: mirrorEngine.exportReplicaSnapshot()
      };
      window.localStorage.setItem(MIRROR_STORAGE_KEY, JSON.stringify(persisted));
    };

    persistPrimary();
    persistMirror();

    const unsubscribePrimary = engine.subscribe(() => persistPrimary());
    const unsubscribeMirror = mirrorEngine.subscribe(() => persistMirror());

    return () => {
      unsubscribePrimary();
      unsubscribeMirror();
    };
  }, [engine, mirrorEngine, replicationReady]);

  useEffect(() => {
    if (!replicationReady) {
      return;
    }

    const enqueueBatch = (target: RelayEntry["target"], batch: EngineOpBatch) => {
      setRelayQueue((current) =>
        compactRelayEntries([
          ...current,
          {
            target,
            batch,
            deliverAt: Date.now() + syncLatencyMs
          }
        ])
      );
    };

    const unsubscribeLocal = engine.subscribeBatches((batch) => enqueueBatch("mirror", batch));
    const unsubscribeMirror = mirrorEngine.subscribeBatches((batch) => enqueueBatch("primary", batch));

    return () => {
      unsubscribeLocal();
      unsubscribeMirror();
    };
  }, [engine, mirrorEngine, replicationReady]);

  useEffect(() => {
    if (!replicationReady) {
      return;
    }

    const persistedRelayState: PersistedRelayState = {
      syncPaused,
      queue: relayQueue
    };
    window.localStorage.setItem(RELAY_STORAGE_KEY, JSON.stringify(persistedRelayState));
  }, [relayQueue, replicationReady, syncPaused]);

  useEffect(() => {
    if (!replicationReady) {
      return;
    }

    if (relayTimerRef.current !== null) {
      window.clearTimeout(relayTimerRef.current);
      relayTimerRef.current = null;
    }

    if (syncPaused || relayQueue.length === 0) {
      return;
    }

    const nextDeliverAt = relayQueue.reduce((earliest, entry) => Math.min(earliest, entry.deliverAt), Number.POSITIVE_INFINITY);
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
  }, [engine, mirrorEngine, relayQueue, replicationReady, syncPaused]);

  useEffect(() => {
    if (sheetNames.length === 0) {
      return;
    }
    if (!sheetNames.includes(selection.sheetName)) {
      selection.select(sheetNames[0]!, "A1");
    }
  }, [selection, sheetNames]);

  useEffect(() => {
    setIsEditingCell(false);
    setEditorValue("");
  }, [selectedAddr, selection.sheetName]);

  const dependencySnapshot = engine.explainCell(selection.sheetName, selectedAddr);

  const beginEditing = useCallback(
    (seed?: string) => {
      setEditorValue(seed ?? toEditorValue(engine.getCell(selection.sheetName, selectedAddr)));
      setIsEditingCell(true);
    },
    [engine, selectedAddr, selection.sheetName]
  );

  const commitEditor = useCallback(
    (movement?: EditMovement) => {
      const nextValue = isEditingCell ? editorValue : toEditorValue(selectedCell);
      applyParsedInput(engine, selection.sheetName, selectedAddr, parseEditorInput(nextValue));
      setIsEditingCell(false);
      if (movement) {
        selection.select(selection.sheetName, clampSelectionMovement(selectedAddr, selection.sheetName, movement));
      }
    },
    [editorValue, engine, isEditingCell, selectedAddr, selectedCell, selection]
  );

  const cancelEditor = useCallback(() => {
    setEditorValue(toEditorValue(selectedCell));
    setIsEditingCell(false);
  }, [selectedCell]);

  const clearSelectedCell = useCallback(() => {
    engine.clearCell(selection.sheetName, selectedAddr);
    setEditorValue("");
    setIsEditingCell(false);
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
      setIsEditingCell(false);
    },
    [engine, selection.sheetName]
  );

  const resetWorkspace = () => {
    window.localStorage.removeItem(PRIMARY_STORAGE_KEY);
    window.localStorage.removeItem(MIRROR_STORAGE_KEY);
    window.localStorage.removeItem(RELAY_STORAGE_KEY);
    window.location.reload();
  };

  const toggleSync = useCallback(() => {
    if (syncPaused) {
      const queuedEntries = [...relayQueueRef.current].sort((left, right) => left.deliverAt - right.deliverAt);
      queuedEntries.forEach(({ target, batch }) => {
        (target === "mirror" ? mirrorEngine : engine).applyRemoteBatch(batch);
      });
      mirrorEngine.importSnapshot(engine.exportSnapshot());
      relayQueueRef.current = [];
      setRelayQueue([]);
      setSyncPaused(false);
      return;
    }

    setSyncPaused(true);
  }, [engine, mirrorEngine, syncPaused]);

  const statusBar = (
    <>
      <span data-testid="status-active-preset">{selectedPreset?.label ?? "Restored workspace"}</span>
      <span data-testid="metric-js">JS {metrics.jsFormulaCount.toLocaleString()}</span>
      <span data-testid="metric-wasm">WASM {metrics.wasmFormulaCount.toLocaleString()}</span>
      <span data-testid="metric-recalc">Recalc {metrics.recalcMs.toFixed(2)} ms</span>
    </>
  );

  const ribbon = (
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

  const sidebar = (
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
    <div className="app-shell">
      {loadingPresetId ? (
        <div className="loading-banner" data-testid="preset-loading">
          Loading {PLAYGROUND_PRESETS.find((preset) => preset.id === loadingPresetId)?.label ?? loadingPresetId}...
        </div>
      ) : null}
      {presetError ? (
        <div className="error-banner" data-testid="preset-error">
          {presetError}
        </div>
      ) : null}
      <WorkbookView
        editorValue={visibleEditorValue}
        engine={engine}
        isEditingCell={isEditingCell}
        onAddressCommit={(input) => {
          const nextTarget = parseSelectionTarget(input, selection.sheetName);
          if (nextTarget) {
            selection.select(nextTarget.sheetName, nextTarget.address);
          }
        }}
        onBeginEdit={beginEditing}
        onCancelEdit={cancelEditor}
        onClearCell={clearSelectedCell}
        onCommitEdit={commitEditor}
        onEditorChange={(next) => {
          setEditorValue(next);
          setIsEditingCell(true);
        }}
        onPaste={pasteIntoSelection}
        onSelect={(addr) => selection.select(selection.sheetName, addr)}
        onSelectSheet={(sheetName) => selection.select(sheetName, "A1")}
        resolvedValue={resolvedValue}
        ribbon={ribbon}
        selectedAddr={selectedAddr}
        sheetName={selection.sheetName}
        sheetNames={sheetNames}
        sidebar={sidebar}
        statusBar={statusBar}
        workbookName={engine.workbook.workbookName}
      />
    </div>
  );
}
