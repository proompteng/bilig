import { Effect } from "effect";
import type { ProjectionOverlayScope } from "./worker-local-overlay.js";
import { createEmptyProjectionOverlayScope } from "./worker-local-overlay.js";

export function resolveProjectionOverlayScopeForPersist(args: {
  projectionOverlayScope: ProjectionOverlayScope | null;
  pendingMutationCount: number;
}): ProjectionOverlayScope | null {
  if (args.projectionOverlayScope) {
    return args.projectionOverlayScope;
  }
  return args.pendingMutationCount === 0 ? createEmptyProjectionOverlayScope() : null;
}

export class WorkerRuntimePersistCoordinator<
  TLocalStore,
  TAuthoritativeEngine,
  TProjectionEngine,
  TPersistedState,
> {
  private flushInFlight: Promise<void> | null = null;
  private requestedPersistVersion = 0;
  private flushedPersistVersion = 0;

  constructor(
    private readonly options: {
      canPersistState: () => boolean;
      getLocalStore: () => TLocalStore | null;
      getAuthoritativeEngine: () => Promise<TAuthoritativeEngine>;
      getProjectionEngine: () => Promise<TProjectionEngine>;
      buildPersistedState: (input: {
        authoritativeEngine: TAuthoritativeEngine;
        projectionEngine: TProjectionEngine;
      }) => TPersistedState;
      getProjectionOverlayScope: () => ProjectionOverlayScope | null;
      saveState: (input: {
        localStore: TLocalStore;
        state: TPersistedState;
        authoritativeEngine: TAuthoritativeEngine;
        projectionEngine: TProjectionEngine;
        projectionOverlayScope: ProjectionOverlayScope | null;
      }) => Promise<void>;
      markProjectionMatchesLocalStore: () => void;
    },
  ) {}

  reset(): void {
    this.flushInFlight = null;
    this.requestedPersistVersion = 0;
    this.flushedPersistVersion = 0;
  }

  async queuePersist(): Promise<void> {
    if (!this.options.canPersistState()) {
      return;
    }
    this.requestedPersistVersion += 1;
    if (this.flushInFlight) {
      await this.flushInFlight;
      return;
    }
    const flushPromise = this.flush();
    this.flushInFlight = flushPromise;
    try {
      await flushPromise;
    } finally {
      if (this.flushInFlight === flushPromise) {
        this.flushInFlight = null;
      }
    }
  }

  private async flush(): Promise<void> {
    await Effect.runPromise(
      Effect.gen(this, function* () {
        while (
          this.flushedPersistVersion < this.requestedPersistVersion &&
          this.options.canPersistState()
        ) {
          const targetVersion = this.requestedPersistVersion;
          const authoritativeEngine = yield* Effect.promise(() =>
            this.options.getAuthoritativeEngine(),
          );
          const projectionEngine = yield* Effect.promise(() => this.options.getProjectionEngine());
          const state = this.options.buildPersistedState({
            authoritativeEngine,
            projectionEngine,
          });

          const localStore = this.options.getLocalStore();
          if (!localStore) {
            return;
          }

          yield* Effect.promise(() =>
            this.options.saveState({
              localStore,
              state,
              authoritativeEngine,
              projectionEngine,
              projectionOverlayScope: this.options.getProjectionOverlayScope(),
            }),
          );
          this.options.markProjectionMatchesLocalStore();
          this.flushedPersistVersion = targetVersion;
        }
      }),
    );
  }
}
