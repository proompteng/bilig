import React from "react";
import type { RecalcMetrics } from "@bilig/protocol";

export function MetricsPanel({ metrics }: { metrics: RecalcMetrics }) {
  return (
    <div className="panel metrics-panel">
      <h3>Recalc Metrics</h3>
      <dl>
        <div><dt>Batch</dt><dd>{metrics.batchId}</dd></div>
        <div><dt>Dirty</dt><dd>{metrics.dirtyFormulaCount}</dd></div>
        <div><dt>JS</dt><dd>{metrics.jsFormulaCount}</dd></div>
        <div><dt>WASM</dt><dd>{metrics.wasmFormulaCount}</dd></div>
        <div><dt>Recalc ms</dt><dd>{metrics.recalcMs.toFixed(2)}</dd></div>
      </dl>
    </div>
  );
}
