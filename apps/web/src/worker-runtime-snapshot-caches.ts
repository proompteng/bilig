import type { EngineReplicaSnapshot } from "@bilig/core";
import type { WorkbookSnapshot } from "@bilig/protocol";

export class WorkerRuntimeSnapshotCaches {
  private projectionSnapshot: WorkbookSnapshot | null = null;
  private projectionSnapshotDirty = true;
  private authoritativeSnapshot: WorkbookSnapshot | null = null;
  private authoritativeSnapshotDirty = true;
  private authoritativeReplica: EngineReplicaSnapshot | null = null;
  private authoritativeReplicaDirty = true;

  reset(): void {
    this.projectionSnapshot = null;
    this.projectionSnapshotDirty = true;
    this.authoritativeSnapshot = null;
    this.authoritativeSnapshotDirty = true;
    this.authoritativeReplica = null;
    this.authoritativeReplicaDirty = true;
  }

  invalidateProjectionSnapshot(): void {
    this.projectionSnapshotDirty = true;
  }

  getProjectionSnapshot(exportSnapshot: () => WorkbookSnapshot): WorkbookSnapshot {
    if (this.projectionSnapshot && !this.projectionSnapshotDirty) {
      return this.projectionSnapshot;
    }
    this.projectionSnapshot = exportSnapshot();
    this.projectionSnapshotDirty = false;
    return this.projectionSnapshot;
  }

  getReadyAuthoritativeSnapshot(): WorkbookSnapshot | null {
    return this.authoritativeSnapshotDirty ? null : this.authoritativeSnapshot;
  }

  installAuthoritativeState(
    snapshot: WorkbookSnapshot | null,
    replica: EngineReplicaSnapshot | null,
  ): void {
    this.authoritativeSnapshot = snapshot;
    this.authoritativeSnapshotDirty = snapshot === null;
    this.authoritativeReplica = replica;
    this.authoritativeReplicaDirty = replica === null;
  }

  invalidateAuthoritativeState(): void {
    this.authoritativeSnapshotDirty = true;
    this.authoritativeReplicaDirty = true;
  }

  resolveAuthoritativeState(input: {
    exportSnapshot: (() => WorkbookSnapshot) | null;
    exportReplica: (() => EngineReplicaSnapshot) | null;
  }): {
    snapshot: WorkbookSnapshot | null;
    replica: EngineReplicaSnapshot | null;
  } {
    if (this.authoritativeSnapshotDirty && input.exportSnapshot) {
      this.storeAuthoritativeSnapshot(input.exportSnapshot());
    }
    if (this.authoritativeReplicaDirty && input.exportReplica) {
      this.storeAuthoritativeReplica(input.exportReplica());
    }
    return {
      snapshot: this.authoritativeSnapshot,
      replica: this.authoritativeReplica,
    };
  }

  storeAuthoritativeSnapshot(snapshot: WorkbookSnapshot): WorkbookSnapshot {
    this.authoritativeSnapshot = snapshot;
    this.authoritativeSnapshotDirty = false;
    return snapshot;
  }

  getAuthoritativeSnapshot(input: {
    canReuseProjectionState: boolean;
    exportProjectionSnapshot: () => WorkbookSnapshot;
    exportAuthoritativeSnapshot: () => WorkbookSnapshot;
  }): WorkbookSnapshot {
    if (this.authoritativeSnapshot && !this.authoritativeSnapshotDirty) {
      return this.authoritativeSnapshot;
    }
    return this.storeAuthoritativeSnapshot(
      input.canReuseProjectionState
        ? input.exportProjectionSnapshot()
        : input.exportAuthoritativeSnapshot(),
    );
  }

  storeAuthoritativeReplica(replica: EngineReplicaSnapshot): EngineReplicaSnapshot {
    this.authoritativeReplica = replica;
    this.authoritativeReplicaDirty = false;
    return replica;
  }

  getAuthoritativeReplica(input: {
    canReuseProjectionState: boolean;
    exportProjectionReplica: () => EngineReplicaSnapshot;
    exportAuthoritativeReplica: () => EngineReplicaSnapshot;
  }): EngineReplicaSnapshot {
    if (this.authoritativeReplica && !this.authoritativeReplicaDirty) {
      return this.authoritativeReplica;
    }
    return this.storeAuthoritativeReplica(
      input.canReuseProjectionState
        ? input.exportProjectionReplica()
        : input.exportAuthoritativeReplica(),
    );
  }
}
