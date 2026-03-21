import { FormulaMode, ValueTag, formatErrorCode, type ExplainCellSnapshot } from "@bilig/protocol";

function formatValue(snapshot: ExplainCellSnapshot): string {
  switch (snapshot.value.tag) {
    case ValueTag.Empty:
      return "∅";
    case ValueTag.Number:
      return String(snapshot.value.value);
    case ValueTag.Boolean:
      return snapshot.value.value ? "TRUE" : "FALSE";
    case ValueTag.String:
      return snapshot.value.value;
    case ValueTag.Error:
      return formatErrorCode(snapshot.value.code);
  }
}

export function DependencyInspector({ snapshot }: { snapshot: ExplainCellSnapshot }) {
  return (
    <div className="panel dependency-panel">
      <h3>Cell Inspector</h3>
      <dl className="dependency-meta">
        <div><dt>Value</dt><dd>{formatValue(snapshot)}</dd></div>
        <div><dt>Format</dt><dd>{snapshot.format ?? "—"}</dd></div>
        <div><dt>Version</dt><dd>{snapshot.version}</dd></div>
        <div><dt>Mode</dt><dd>{snapshot.mode === FormulaMode.WasmFastPath ? "WASM fast path" : snapshot.mode === FormulaMode.JsOnly ? "JS evaluator" : "Literal"}</dd></div>
        <div><dt>Cycle</dt><dd>{snapshot.inCycle ? "Yes" : "No"}</dd></div>
        <div><dt>Topo rank</dt><dd>{snapshot.topoRank ?? "—"}</dd></div>
      </dl>
      <strong>Formula</strong>
      <p className="inspector-formula">{snapshot.formula ? `=${snapshot.formula}` : "Literal cell"}</p>
      <strong>Precedents</strong>
      <ul>
        {snapshot.directPrecedents.length === 0 ? <li className="empty-state">No precedents</li> : null}
        {snapshot.directPrecedents.map((value) => <li key={`p-${value}`}>{value}</li>)}
      </ul>
      <strong>Dependents</strong>
      <ul>
        {snapshot.directDependents.length === 0 ? <li className="empty-state">No dependents</li> : null}
        {snapshot.directDependents.map((value) => <li key={`d-${value}`}>{value}</li>)}
      </ul>
    </div>
  );
}
