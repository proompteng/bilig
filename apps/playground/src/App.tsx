import React, { useEffect, useMemo, useRef, useState } from "react";
import { SpreadsheetEngine } from "@bilig/core";
import type { EngineOpBatch } from "@bilig/crdt";
import { createWorkbookRendererRoot } from "@bilig/renderer";
import {
  CellEditorOverlay,
  DependencyInspector,
  FormulaBar,
  MetricsPanel,
  ReplicaPanel,
  WorkbookView,
  useCell,
  useMetrics,
  useSelection
} from "@bilig/grid";
import { buildDemoWorkbook } from "./demoWorkbook.js";
import { compactRelayEntries, type RelayEntry } from "./relay-queue.js";
import { renderSnapshotWorkbook } from "./snapshotWorkbook.js";

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

export function App() {
  const engine = useMemo(() => new SpreadsheetEngine({ workbookName: "bilig-demo", replicaId: "playground" }), []);
  const mirrorEngine = useMemo(() => new SpreadsheetEngine({ workbookName: "bilig-demo", replicaId: "replica-beta" }), []);
  const rendererRoot = useMemo(() => createWorkbookRendererRoot(engine), [engine]);
  const selection = useSelection(engine);
  const selectedAddr = selection.address ?? "A1";
  const selectedCell = useCell(engine, selection.sheetName, selectedAddr);
  const mirroredSelectedCell = useCell(mirrorEngine, selection.sheetName, selectedAddr);
  const metrics = useMetrics(engine);
  const mirrorMetrics = useMetrics(mirrorEngine);
  const [editorValue, setEditorValue] = useState("");
  const [replicationReady, setReplicationReady] = useState(false);
  const [syncPaused, setSyncPaused] = useState(false);
  const [relayQueue, setRelayQueue] = useState<RelayEntry[]>([]);
  const syncLatencyMs = 120;
  const relayQueueRef = useRef<RelayEntry[]>([]);
  const relayTimerRef = useRef<number | null>(null);
  const sheetNames = [...engine.workbook.sheetsByName.values()]
    .sort((left, right) => left.order - right.order)
    .map((sheet) => sheet.name);
  const pendingSyncCount = syncPaused ? 0 : relayQueue.length;
  const queuedSyncCount = syncPaused ? relayQueue.length : 0;

  useEffect(() => {
    relayQueueRef.current = relayQueue;
  }, [relayQueue]);

  useEffect(() => {
    let cancelled = false;

    void Promise.all([engine.ready(), mirrorEngine.ready()]).then(async () => {
      const primaryPersisted = window.localStorage.getItem(PRIMARY_STORAGE_KEY);
      const mirrorPersisted = window.localStorage.getItem(MIRROR_STORAGE_KEY);
      const relayPersisted = window.localStorage.getItem(RELAY_STORAGE_KEY);

      if (primaryPersisted) {
        const restored = JSON.parse(primaryPersisted) as PersistedReplicaState;
        await rendererRoot.render(renderSnapshotWorkbook(restored.snapshot));
        engine.importReplicaSnapshot(restored.replica);
      } else {
        await rendererRoot.render(buildDemoWorkbook());
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
  }, [engine, mirrorEngine, rendererRoot]);

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
    if (!replicationReady) return;

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
    if (selectedCell.formula) {
      setEditorValue(`=${selectedCell.formula}`);
      return;
    }
    if (selectedCell.value.tag === 1) {
      setEditorValue(String(selectedCell.value.value));
      return;
    }
    if (selectedCell.value.tag === 2) {
      setEditorValue(selectedCell.value.value ? "TRUE" : "FALSE");
      return;
    }
    if (selectedCell.value.tag === 3) {
      setEditorValue(selectedCell.value.value);
      return;
    }
    setEditorValue("");
  }, [selectedCell]);

  useEffect(() => {
    if (sheetNames.length === 0) return;
    if (!sheetNames.includes(selection.sheetName)) {
      selection.select(sheetNames[0]!, "A1");
    }
  }, [selection, sheetNames]);

  const dependencySnapshot = engine.explainCell(selection.sheetName, selectedAddr);
  const resolvedValue =
    selectedCell.value.tag === 1
      ? String(selectedCell.value.value)
      : selectedCell.value.tag === 2
        ? String(selectedCell.value.value)
        : selectedCell.value.tag === 3
          ? selectedCell.value.value
          : selectedCell.value.tag === 4
            ? `#${selectedCell.value.code}`
            : "";
  const mirroredValue =
    mirroredSelectedCell.value.tag === 1
      ? String(mirroredSelectedCell.value.value)
      : mirroredSelectedCell.value.tag === 2
        ? String(mirroredSelectedCell.value.value)
        : mirroredSelectedCell.value.tag === 3
          ? mirroredSelectedCell.value.value
          : mirroredSelectedCell.value.tag === 4
            ? `#${mirroredSelectedCell.value.code}`
            : "";

  const commitEditor = () => {
    const normalized = editorValue.trim();
    if (normalized.startsWith("=")) {
      engine.setCellFormula(selection.sheetName, selectedAddr, normalized.slice(1));
      return;
    }
    if (normalized === "") {
      engine.clearCell(selection.sheetName, selectedAddr);
      return;
    }
    if (normalized === "TRUE" || normalized === "FALSE") {
      engine.setCellValue(selection.sheetName, selectedAddr, normalized === "TRUE");
      return;
    }
    const numeric = Number(normalized);
    if (!Number.isNaN(numeric) && /^-?\d+(\.\d+)?$/.test(normalized)) {
      engine.setCellValue(selection.sheetName, selectedAddr, numeric);
      return;
    }
    engine.setCellValue(selection.sheetName, selectedAddr, normalized);
  };

  const resetWorkspace = () => {
    window.localStorage.removeItem(PRIMARY_STORAGE_KEY);
    window.localStorage.removeItem(MIRROR_STORAGE_KEY);
    window.localStorage.removeItem(RELAY_STORAGE_KEY);
    window.location.reload();
  };

  return (
    <div className="app-shell">
      <header className="hero">
        <div className="hero-copy">
          <p className="eyebrow">bilig</p>
          <h1>Custom reconciler playground for a local-first spreadsheet engine</h1>
          <p className="lede">
            React declares workbook structure, the engine owns recalculation and CRDT-ready ops, and AssemblyScript/WASM
            handles the numeric fast path.
          </p>
        </div>
        <button className="hero-reset" onClick={resetWorkspace} type="button">
          Reset local workspace
        </button>
      </header>

      <FormulaBar
        label={`${selection.sheetName}!${selectedAddr}`}
        value={editorValue}
        onChange={setEditorValue}
        onClear={() => {
          engine.clearCell(selection.sheetName, selectedAddr);
          setEditorValue("");
        }}
        onCommit={commitEditor}
      />

      <main className="workspace">
        <WorkbookView
          engine={engine}
          sheetNames={sheetNames}
          workbookName={engine.workbook.workbookName}
          sheetName={selection.sheetName}
          selectedAddr={selectedAddr}
          onSelectSheet={(sheetName) => selection.select(sheetName, selectedAddr)}
          onSelect={(addr) => selection.select(selection.sheetName, addr)}
        />
        <aside className="sidebar">
          <CellEditorOverlay
            label={`${selection.sheetName}!${selectedAddr}`}
            resolvedValue={resolvedValue}
            value={editorValue}
            onChange={setEditorValue}
            onClear={() => {
              engine.clearCell(selection.sheetName, selectedAddr);
              setEditorValue("");
            }}
            onCommit={commitEditor}
          />
          <MetricsPanel metrics={metrics} />
          <ReplicaPanel
            latencyMs={syncLatencyMs}
            localReplicaId={engine.replica.replicaId}
            onToggleSync={() => setSyncPaused((current) => !current)}
            pendingSyncCount={pendingSyncCount}
            queuedSyncCount={queuedSyncCount}
            remoteMetrics={mirrorMetrics}
            remoteReplicaId={mirrorEngine.replica.replicaId}
            remoteValue={mirroredValue}
            selectedLabel={`${selection.sheetName}!${selectedAddr}`}
            syncPaused={syncPaused}
          />
          <DependencyInspector snapshot={dependencySnapshot} />
        </aside>
      </main>
    </div>
  );
}
