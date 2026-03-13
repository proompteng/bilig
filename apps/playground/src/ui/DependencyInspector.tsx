import React from "react";
import type { DependencySnapshot } from "@bilig/protocol";

export function DependencyInspector({ snapshot }: { snapshot: DependencySnapshot }) {
  return (
    <div className="panel dependency-panel">
      <h3>Dependencies</h3>
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
