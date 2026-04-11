import type { WorkbookLocalStore } from "@bilig/storage-browser";
import type { PendingWorkbookMutation, PendingWorkbookMutationInput } from "./workbook-sync.js";
import { isPendingWorkbookMutationInput } from "./workbook-sync.js";
import { clonePendingWorkbookMutation } from "./workbook-mutation-journal.js";
import {
  ackAbsorbedMutations,
  ackMutationInJournal,
  createRuntimePendingMutation,
  failMutationInJournal,
  markMutationSubmittedInJournal,
  recordMutationAttemptInJournal,
  retryMutationInJournal,
} from "./worker-runtime-mutation-actions.js";
import { applyPendingWorkbookMutationToEngine } from "./worker-runtime-mutation-replay.js";
import {
  buildPendingMutationSummary,
  markJournalMutationsRebased,
  syncPendingMutationsFromJournal,
  type WorkerRuntimePendingMutationSummarySnapshot,
} from "./worker-runtime-pending-mutations.js";
import type { WorkerEngine } from "./worker-runtime-support.js";

type PersistableMutationJournalStore = Pick<
  WorkbookLocalStore,
  "appendPendingMutation" | "updatePendingMutation"
>;

export interface WorkerRuntimeMutationJournalBootstrapState {
  readonly mutationJournalEntries: readonly PendingWorkbookMutation[];
  readonly nextPendingMutationSeq: number;
}

export class WorkerRuntimeMutationJournal {
  private mutationJournalEntries: PendingWorkbookMutation[] = [];
  private pendingMutations: PendingWorkbookMutation[] = [];
  private nextPendingMutationSeq = 1;

  constructor(
    private readonly deps: {
      readonly getDocumentId: () => string;
      readonly getAuthoritativeRevision: () => number;
      readonly getLocalStore: () => PersistableMutationJournalStore | null;
      readonly getProjectionEngine: () => Promise<WorkerEngine>;
      readonly markProjectionDivergedFromLocalStore: () => void;
      readonly queuePersist: () => Promise<void>;
      readonly now?: () => number;
    },
  ) {}

  reset(): void {
    this.mutationJournalEntries = [];
    this.pendingMutations = [];
    this.nextPendingMutationSeq = 1;
  }

  restoreFromBootstrap(state: WorkerRuntimeMutationJournalBootstrapState): void {
    this.mutationJournalEntries = state.mutationJournalEntries.map(clonePendingWorkbookMutation);
    this.pendingMutations = syncPendingMutationsFromJournal(this.mutationJournalEntries);
    this.nextPendingMutationSeq = state.nextPendingMutationSeq;
  }

  listPendingMutations(): PendingWorkbookMutation[] {
    return this.pendingMutations.map(clonePendingWorkbookMutation);
  }

  listMutationJournalEntries(): PendingWorkbookMutation[] {
    return this.mutationJournalEntries.map(clonePendingWorkbookMutation);
  }

  getPendingMutationCount(): number {
    return this.pendingMutations.length;
  }

  getAppliedPendingLocalSeq(): number {
    return this.pendingMutations.at(-1)?.localSeq ?? 0;
  }

  buildPendingMutationSummary(): WorkerRuntimePendingMutationSummarySnapshot {
    return buildPendingMutationSummary(this.mutationJournalEntries, this.pendingMutations);
  }

  async enqueuePendingMutation(
    input: PendingWorkbookMutationInput,
  ): Promise<PendingWorkbookMutation> {
    if (!isPendingWorkbookMutationInput(input)) {
      throw new Error("Invalid pending workbook mutation");
    }
    const nextMutation = createRuntimePendingMutation({
      documentId: this.deps.getDocumentId(),
      localSeq: this.nextPendingMutationSeq++,
      authoritativeRevision: this.deps.getAuthoritativeRevision(),
      input,
      enqueuedAtUnixMs: this.now(),
    });
    this.mutationJournalEntries.push(nextMutation);
    this.pendingMutations = syncPendingMutationsFromJournal(this.mutationJournalEntries);
    this.deps.markProjectionDivergedFromLocalStore();
    const localStore = this.deps.getLocalStore();
    if (localStore) {
      await localStore.appendPendingMutation(clonePendingWorkbookMutation(nextMutation));
    }
    applyPendingWorkbookMutationToEngine(await this.deps.getProjectionEngine(), nextMutation);
    await this.deps.queuePersist();
    return clonePendingWorkbookMutation(nextMutation);
  }

  async markPendingMutationSubmitted(id: string): Promise<void> {
    const result = markMutationSubmittedInJournal({
      mutationJournalEntries: this.mutationJournalEntries,
      id,
      submittedAtUnixMs: this.now(),
    });
    if (!result) {
      return;
    }
    this.updateState(result.mutationJournalEntries, result.pendingMutations);
    await this.persistUpdatedMutation(result.updatedMutation);
  }

  async ackPendingMutation(id: string): Promise<void> {
    const result = ackMutationInJournal({
      mutationJournalEntries: this.mutationJournalEntries,
      id,
      ackedAtUnixMs: this.now(),
    });
    if (!result) {
      return;
    }
    this.updateState(result.mutationJournalEntries, result.pendingMutations);
    this.deps.markProjectionDivergedFromLocalStore();
    await this.persistUpdatedMutation(result.updatedMutation);
    if (this.deps.getLocalStore()) {
      await this.deps.queuePersist();
    }
  }

  async recordPendingMutationAttempt(id: string): Promise<void> {
    const result = recordMutationAttemptInJournal({
      mutationJournalEntries: this.mutationJournalEntries,
      id,
      attemptedAtUnixMs: this.now(),
    });
    if (!result) {
      return;
    }
    this.updateState(result.mutationJournalEntries, result.pendingMutations);
    await this.persistUpdatedMutation(result.updatedMutation);
  }

  async markPendingMutationFailed(id: string, failureMessage: string): Promise<void> {
    const result = failMutationInJournal({
      mutationJournalEntries: this.mutationJournalEntries,
      id,
      failureMessage,
      failedAtUnixMs: this.now(),
    });
    if (!result) {
      return;
    }
    this.updateState(result.mutationJournalEntries, result.pendingMutations);
    this.deps.markProjectionDivergedFromLocalStore();
    await this.persistUpdatedMutation(result.updatedMutation);
    if (this.deps.getLocalStore()) {
      await this.deps.queuePersist();
    }
  }

  async retryPendingMutation(id: string): Promise<void> {
    const result = retryMutationInJournal({
      mutationJournalEntries: this.mutationJournalEntries,
      id,
    });
    if (!result) {
      return;
    }
    this.updateState(result.mutationJournalEntries, result.pendingMutations);
    this.deps.markProjectionDivergedFromLocalStore();
    await this.persistUpdatedMutation(result.updatedMutation);
    if (this.deps.getLocalStore()) {
      await this.deps.queuePersist();
    }
  }

  ackAbsorbedMutations(absorbedMutationIds: ReadonlySet<string>): void {
    if (absorbedMutationIds.size === 0) {
      return;
    }
    const result = ackAbsorbedMutations({
      mutationJournalEntries: this.mutationJournalEntries,
      absorbedMutationIds,
      ackedAtUnixMs: this.now(),
    });
    this.updateState(result.mutationJournalEntries, result.pendingMutations);
  }

  async markRemainingJournalMutationsRebased(rebasedAtUnixMs = this.now()): Promise<void> {
    const nextState = markJournalMutationsRebased(this.mutationJournalEntries, rebasedAtUnixMs);
    this.updateState(nextState.mutationJournalEntries, nextState.pendingMutations);
    const localStore = this.deps.getLocalStore();
    if (!localStore) {
      return;
    }
    await Promise.all(
      nextState.updatedMutations.map((mutation) =>
        localStore.updatePendingMutation(clonePendingWorkbookMutation(mutation)),
      ),
    );
  }

  private now(): number {
    return this.deps.now?.() ?? Date.now();
  }

  private updateState(
    mutationJournalEntries: readonly PendingWorkbookMutation[],
    pendingMutations: readonly PendingWorkbookMutation[],
  ): void {
    this.mutationJournalEntries = mutationJournalEntries.map(clonePendingWorkbookMutation);
    this.pendingMutations = pendingMutations.map(clonePendingWorkbookMutation);
  }

  private async persistUpdatedMutation(updatedMutation: PendingWorkbookMutation): Promise<void> {
    const localStore = this.deps.getLocalStore();
    if (!localStore) {
      return;
    }
    await localStore.updatePendingMutation(clonePendingWorkbookMutation(updatedMutation));
  }
}
