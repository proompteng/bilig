import React from "react";
import type { RecalcMetrics } from "@bilig/protocol";

interface ReplicaPanelProps {
  localReplicaId: string;
  remoteReplicaId: string;
  selectedLabel: string;
  remoteValue: string;
  syncPaused: boolean;
  pendingSyncCount: number;
  queuedSyncCount: number;
  latencyMs: number;
  remoteMetrics: RecalcMetrics;
  onToggleSync(): void;
}

export function ReplicaPanel({
  localReplicaId,
  remoteReplicaId,
  selectedLabel,
  remoteValue,
  syncPaused,
  pendingSyncCount,
  queuedSyncCount,
  latencyMs,
  remoteMetrics,
  onToggleSync
}: ReplicaPanelProps) {
  return (
    <div className="panel replica-panel" data-testid="replica-panel">
      <div className="replica-panel-header">
        <div>
          <p className="panel-eyebrow">Local-First Mirror</p>
          <h3>{localReplicaId} → {remoteReplicaId}</h3>
        </div>
        <button className={syncPaused ? "ghost-button" : ""} onClick={onToggleSync} type="button">
          {syncPaused ? "Resume sync" : "Pause sync"}
        </button>
      </div>
      <dl className="replica-stats">
        <div><dt>Status</dt><dd data-testid="replica-status">{syncPaused ? "Paused" : "Live"}</dd></div>
        <div><dt>Latency</dt><dd>{latencyMs} ms</dd></div>
        <div><dt>In flight</dt><dd data-testid="replica-pending">{pendingSyncCount}</dd></div>
        <div><dt>Queued</dt><dd data-testid="replica-queued">{queuedSyncCount}</dd></div>
        <div><dt>Replica batch</dt><dd>{remoteMetrics.batchId}</dd></div>
      </dl>
      <div className="replica-selected">
        <strong>{selectedLabel}</strong>
        <span data-testid="replica-value">{remoteValue || "∅"}</span>
      </div>
    </div>
  );
}
