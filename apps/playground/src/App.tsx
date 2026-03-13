import React, { useEffect, useMemo, useRef, useState } from "react";
import { SpreadsheetEngine } from "@bilig/core";
import type { EngineOpBatch } from "@bilig/crdt";
import { buildDemoWorkbook } from "./demoWorkbook.js";
import { createWorkbookRendererRoot } from "./reconciler/index.js";
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
} from "./ui/index.js";

export function App() {
  const engine = useMemo(() => new SpreadsheetEngine({ workbookName: "bilig-demo", replicaId: "playground" }), []);
  const mirrorEngine = useMemo(() => new SpreadsheetEngine({ workbookName: "bilig-demo", replicaId: "replica-beta" }), []);
  const rendererRoot = useMemo(() => createWorkbookRendererRoot(engine), [engine]);
  const selection = useSelection("Sheet1", "A1");
  const selectedCell = useCell(engine, selection.sheetName, selection.address);
  const mirroredSelectedCell = useCell(mirrorEngine, selection.sheetName, selection.address);
  const metrics = useMetrics(engine);
  const mirrorMetrics = useMetrics(mirrorEngine);
  const [editorValue, setEditorValue] = useState("");
  const [replicationReady, setReplicationReady] = useState(false);
  const [syncPaused, setSyncPaused] = useState(false);
  const [pendingSyncCount, setPendingSyncCount] = useState(0);
  const [queuedSyncCount, setQueuedSyncCount] = useState(0);
  const syncLatencyMs = 120;
  const queuedBatchesRef = useRef<Array<{ target: "primary" | "mirror"; batch: EngineOpBatch }>>([]);
  const pendingTimersRef = useRef(new Set<number>());
  const sheetNames = [...engine.workbook.sheetsByName.values()]
    .sort((left, right) => left.order - right.order)
    .map((sheet) => sheet.name);

  useEffect(() => {
    let cancelled = false;

    void Promise.all([engine.ready(), mirrorEngine.ready()]).then(async () => {
      await rendererRoot.render(buildDemoWorkbook());
      mirrorEngine.importSnapshot(engine.exportSnapshot());
      if (!cancelled) {
        setReplicationReady(true);
      }
    });

    return () => {
      cancelled = true;
      pendingTimersRef.current.forEach((timer) => window.clearTimeout(timer));
      pendingTimersRef.current.clear();
      queuedBatchesRef.current = [];
      void rendererRoot.unmount();
    };
  }, [engine, mirrorEngine, rendererRoot]);

  useEffect(() => {
    if (!replicationReady) return;

    const scheduleBatch = (target: SpreadsheetEngine, batch: EngineOpBatch) => {
      setPendingSyncCount((current) => current + 1);
      const timer = window.setTimeout(() => {
        pendingTimersRef.current.delete(timer);
        target.applyRemoteBatch(batch);
        setPendingSyncCount((current) => Math.max(0, current - 1));
      }, syncLatencyMs);
      pendingTimersRef.current.add(timer);
    };

    const routeBatch = (target: "primary" | "mirror", batch: EngineOpBatch) => {
      if (syncPaused) {
        queuedBatchesRef.current.push({ target, batch });
        setQueuedSyncCount(queuedBatchesRef.current.length);
        return;
      }
      scheduleBatch(target === "mirror" ? mirrorEngine : engine, batch);
    };

    const unsubscribeLocal = engine.subscribeBatches((batch) => routeBatch("mirror", batch));
    const unsubscribeMirror = mirrorEngine.subscribeBatches((batch) => routeBatch("primary", batch));

    return () => {
      unsubscribeLocal();
      unsubscribeMirror();
    };
  }, [engine, mirrorEngine, replicationReady, syncPaused]);

  useEffect(() => {
    if (syncPaused || queuedBatchesRef.current.length === 0) {
      return;
    }

    const queuedBatches = queuedBatchesRef.current.splice(0);
    setQueuedSyncCount(0);
    queuedBatches.forEach(({ target, batch }) => {
      setPendingSyncCount((current) => current + 1);
      const timer = window.setTimeout(() => {
        pendingTimersRef.current.delete(timer);
        (target === "mirror" ? mirrorEngine : engine).applyRemoteBatch(batch);
        setPendingSyncCount((current) => Math.max(0, current - 1));
      }, syncLatencyMs);
      pendingTimersRef.current.add(timer);
    });
  }, [engine, mirrorEngine, syncPaused]);

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

  const dependencySnapshot = engine.getDependencies(selection.sheetName, selection.address);
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
      engine.setCellFormula(selection.sheetName, selection.address, normalized.slice(1));
      return;
    }
    if (normalized === "") {
      engine.clearCell(selection.sheetName, selection.address);
      return;
    }
    if (normalized === "TRUE" || normalized === "FALSE") {
      engine.setCellValue(selection.sheetName, selection.address, normalized === "TRUE");
      return;
    }
    const numeric = Number(normalized);
    if (!Number.isNaN(numeric) && /^-?\d+(\.\d+)?$/.test(normalized)) {
      engine.setCellValue(selection.sheetName, selection.address, numeric);
      return;
    }
    engine.setCellValue(selection.sheetName, selection.address, normalized);
  };

  return (
    <div className="app-shell">
      <header className="hero">
        <div>
          <p className="eyebrow">bilig</p>
          <h1>Custom reconciler playground for a local-first spreadsheet engine</h1>
          <p className="lede">
            React declares workbook structure, the engine owns recalculation and CRDT-ready ops, and AssemblyScript/WASM
            handles the numeric fast path.
          </p>
        </div>
      </header>

      <FormulaBar
        label={`${selection.sheetName}!${selection.address}`}
        value={editorValue}
        onChange={setEditorValue}
        onClear={() => {
          engine.clearCell(selection.sheetName, selection.address);
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
          selectedAddr={selection.address}
          onSelectSheet={(sheetName) => selection.select(sheetName, selection.address)}
          onSelect={(addr) => selection.select(selection.sheetName, addr)}
        />
        <aside className="sidebar">
          <CellEditorOverlay
            label={`${selection.sheetName}!${selection.address}`}
            resolvedValue={resolvedValue}
            value={editorValue}
            onChange={setEditorValue}
            onClear={() => {
              engine.clearCell(selection.sheetName, selection.address);
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
            selectedLabel={`${selection.sheetName}!${selection.address}`}
            syncPaused={syncPaused}
          />
          <DependencyInspector snapshot={dependencySnapshot} />
        </aside>
      </main>
    </div>
  );
}
