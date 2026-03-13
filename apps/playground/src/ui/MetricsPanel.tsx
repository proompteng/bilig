import React from "react";
import type { RecalcMetrics } from "@bilig/protocol";

export function MetricsPanel({ metrics }: { metrics: RecalcMetrics }) {
  return (
    <div className="panel metrics-panel" data-testid="metrics-panel">
      <h3>Recalc Metrics</h3>
      <dl>
        <div><dt>Batch</dt><dd data-testid="metric-batch">{metrics.batchId}</dd></div>
        <div><dt>Dirty</dt><dd data-testid="metric-dirty">{metrics.dirtyFormulaCount}</dd></div>
        <div><dt>JS</dt><dd data-testid="metric-js">{metrics.jsFormulaCount}</dd></div>
        <div><dt>WASM</dt><dd data-testid="metric-wasm">{metrics.wasmFormulaCount}</dd></div>
        <div><dt>Recalc ms</dt><dd data-testid="metric-recalc-ms">{metrics.recalcMs.toFixed(2)}</dd></div>
      </dl>
    </div>
  );
}
