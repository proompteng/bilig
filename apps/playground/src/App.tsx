import React, { useEffect, useMemo, useState } from "react";
import { SpreadsheetEngine } from "@bilig/core";
import { buildDemoWorkbook } from "./demoWorkbook.js";
import { createWorkbookRendererRoot } from "./reconciler/index.js";
import {
  CellEditorOverlay,
  DependencyInspector,
  FormulaBar,
  MetricsPanel,
  WorkbookView,
  useCell,
  useMetrics,
  useSelection
} from "./ui/index.js";

export function App() {
  const engine = useMemo(() => new SpreadsheetEngine({ workbookName: "bilig-demo", replicaId: "playground" }), []);
  const rendererRoot = useMemo(() => createWorkbookRendererRoot(engine), [engine]);
  const selection = useSelection("Sheet1", "A1");
  const selectedCell = useCell(engine, selection.sheetName, selection.address);
  const metrics = useMetrics(engine);
  const [editorValue, setEditorValue] = useState("");
  const sheetNames = [...engine.workbook.sheetsByName.values()]
    .sort((left, right) => left.order - right.order)
    .map((sheet) => sheet.name);

  useEffect(() => {
    void engine.ready().then(() => rendererRoot.render(buildDemoWorkbook()));
    return () => {
      void rendererRoot.unmount();
    };
  }, [engine, rendererRoot]);

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
          <DependencyInspector snapshot={dependencySnapshot} />
        </aside>
      </main>
    </div>
  );
}
