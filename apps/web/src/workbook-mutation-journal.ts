import { Data, Effect } from "effect";
import type { PendingWorkbookMutation } from "./workbook-sync.js";

export class MutationJournalTransitionError extends Data.TaggedError(
  "MutationJournalTransitionError",
)<{
  readonly message: string;
}> {}

export function clonePendingWorkbookMutation(
  mutation: PendingWorkbookMutation,
): PendingWorkbookMutation {
  return {
    ...mutation,
    args: [...mutation.args],
  };
}

export function isActivePendingWorkbookMutation(mutation: PendingWorkbookMutation): boolean {
  return mutation.status !== "acked";
}

export function isPendingWorkbookMutationReadyForSubmission(
  mutation: PendingWorkbookMutation,
): boolean {
  return (
    (mutation.status === "local" || mutation.status === "rebased") &&
    mutation.submittedAtUnixMs === null
  );
}

function transitionMutation(
  mutation: PendingWorkbookMutation,
  transition: () => PendingWorkbookMutation,
): Effect.Effect<PendingWorkbookMutation, MutationJournalTransitionError> {
  return Effect.try({
    try: transition,
    catch: (error) =>
      error instanceof MutationJournalTransitionError
        ? error
        : new MutationJournalTransitionError({
            message: error instanceof Error ? error.message : String(error),
          }),
  });
}

export function recordPendingWorkbookMutationAttempt(
  mutation: PendingWorkbookMutation,
  attemptedAtUnixMs: number,
): Effect.Effect<PendingWorkbookMutation, MutationJournalTransitionError> {
  return transitionMutation(mutation, () => {
    if (mutation.status === "submitted" || mutation.status === "acked") {
      throw new MutationJournalTransitionError({
        message: `Cannot record an additional submission attempt for ${mutation.status} mutation ${mutation.id}`,
      });
    }
    const nextStatus =
      mutation.status === "failed"
        ? mutation.rebasedAtUnixMs === null
          ? "local"
          : "rebased"
        : mutation.status;
    return clonePendingWorkbookMutation({
      ...mutation,
      status: nextStatus,
      attemptCount: mutation.attemptCount + 1,
      lastAttemptedAtUnixMs: attemptedAtUnixMs,
      failedAtUnixMs: null,
      failureMessage: null,
    });
  });
}

export function markPendingWorkbookMutationSubmitted(
  mutation: PendingWorkbookMutation,
  submittedAtUnixMs: number,
): Effect.Effect<PendingWorkbookMutation, MutationJournalTransitionError> {
  return transitionMutation(mutation, () => {
    if (
      mutation.status !== "local" &&
      mutation.status !== "rebased" &&
      mutation.status !== "failed"
    ) {
      throw new MutationJournalTransitionError({
        message: `Cannot submit ${mutation.status} mutation ${mutation.id}`,
      });
    }
    return clonePendingWorkbookMutation({
      ...mutation,
      status: "submitted",
      submittedAtUnixMs,
      failedAtUnixMs: null,
      failureMessage: null,
    });
  });
}

export function markPendingWorkbookMutationRebased(
  mutation: PendingWorkbookMutation,
  rebasedAtUnixMs: number,
): Effect.Effect<PendingWorkbookMutation, MutationJournalTransitionError> {
  return transitionMutation(mutation, () => {
    if (mutation.status === "acked" || mutation.status === "failed") {
      if (mutation.status === "failed") {
        return clonePendingWorkbookMutation(mutation);
      }
      throw new MutationJournalTransitionError({
        message: `Cannot rebase acked mutation ${mutation.id}`,
      });
    }
    return clonePendingWorkbookMutation({
      ...mutation,
      status: mutation.status === "local" ? "rebased" : mutation.status,
      rebasedAtUnixMs,
    });
  });
}

export function markPendingWorkbookMutationFailed(
  mutation: PendingWorkbookMutation,
  failureMessage: string,
  failedAtUnixMs: number,
): Effect.Effect<PendingWorkbookMutation, MutationJournalTransitionError> {
  return transitionMutation(mutation, () => {
    if (mutation.status === "acked") {
      throw new MutationJournalTransitionError({
        message: `Cannot fail acked mutation ${mutation.id}`,
      });
    }
    return clonePendingWorkbookMutation({
      ...mutation,
      status: "failed",
      failedAtUnixMs,
      failureMessage,
    });
  });
}

export function queuePendingWorkbookMutationRetry(
  mutation: PendingWorkbookMutation,
): Effect.Effect<PendingWorkbookMutation, MutationJournalTransitionError> {
  return transitionMutation(mutation, () => {
    if (mutation.status !== "failed") {
      throw new MutationJournalTransitionError({
        message: `Cannot retry ${mutation.status} mutation ${mutation.id}`,
      });
    }
    return clonePendingWorkbookMutation({
      ...mutation,
      status: mutation.rebasedAtUnixMs === null ? "local" : "rebased",
      submittedAtUnixMs: null,
      failedAtUnixMs: null,
      failureMessage: null,
    });
  });
}

export function markPendingWorkbookMutationAcked(
  mutation: PendingWorkbookMutation,
  ackedAtUnixMs: number,
): Effect.Effect<PendingWorkbookMutation, MutationJournalTransitionError> {
  return transitionMutation(mutation, () => {
    if (mutation.status === "acked") {
      throw new MutationJournalTransitionError({
        message: `Mutation ${mutation.id} has already been acked`,
      });
    }
    return clonePendingWorkbookMutation({
      ...mutation,
      status: "acked",
      ackedAtUnixMs,
      failedAtUnixMs: null,
      failureMessage: null,
    });
  });
}
