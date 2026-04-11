import type { SyncState } from "@bilig/protocol";
import { assign, fromCallback, sendTo, setup } from "xstate";
import type { WorkbookWorkerStateSnapshot } from "./worker-runtime.js";
import type { ZeroConnectionState } from "./worker-workbook-app-model.js";
import {
  createWorkerRuntimeSessionController,
  type CreateWorkerRuntimeSessionInput,
  type WorkerHandle,
  type WorkerRuntimeSelection,
  type WorkerRuntimeSessionController,
  type WorkerRuntimeSessionPhase,
} from "./runtime-session.js";

type ConnectionStateName = ZeroConnectionState["name"];

interface WorkerRuntimeMachineContext {
  readonly sessionInput: WorkerRuntimeMachineInput;
  readonly persistState: boolean;
  readonly controller: WorkerRuntimeSessionController | null;
  readonly handle: WorkerHandle | null;
  readonly runtimeState: WorkbookWorkerStateSnapshot | null;
  readonly selection: WorkerRuntimeSelection;
  readonly connectionStateName: ConnectionStateName;
  readonly error: string | null;
}

type WorkerRuntimeMachineEvent =
  | { type: "retry"; persistState?: boolean }
  | { type: "error.clear" }
  | { type: "selection.changed"; selection: WorkerRuntimeSelection }
  | { type: "connection.changed"; connectionStateName: ConnectionStateName }
  | { type: "session.ready"; controller: WorkerRuntimeSessionController }
  | { type: "session.runtime"; runtimeState: WorkbookWorkerStateSnapshot }
  | { type: "session.selection"; selection: WorkerRuntimeSelection }
  | { type: "session.phase"; phase: WorkerRuntimeSessionPhase }
  | { type: "session.error"; message: string }
  | { type: "session.failed"; message: string };

export interface WorkerRuntimeMachineInput extends CreateWorkerRuntimeSessionInput {
  readonly connectionStateName?: ConnectionStateName;
  readonly createSession?: (
    input: CreateWorkerRuntimeSessionInput,
    callbacks: Parameters<typeof createWorkerRuntimeSessionController>[1],
  ) => Promise<WorkerRuntimeSessionController>;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function isWorkbookWorkerStateSnapshotValue(value: unknown): value is WorkbookWorkerStateSnapshot {
  return (
    isRecord(value) &&
    typeof value["workbookName"] === "string" &&
    Array.isArray(value["sheetNames"]) &&
    isRecord(value["metrics"]) &&
    typeof value["syncState"] === "string"
  );
}

function initialConnectionStateName(input: WorkerRuntimeMachineInput): ConnectionStateName {
  return input.connectionStateName ?? (input.zero ? "connecting" : "closed");
}

function mapConnectionStateToRuntimeSyncState(
  connectionStateName: ConnectionStateName,
  hasZero: boolean,
): SyncState | null {
  if (!hasZero) {
    return "local-only";
  }
  switch (connectionStateName) {
    case "connected":
      return "live";
    case "connecting":
      return "syncing";
    case "disconnected":
      return "reconnecting";
    case "needs-auth":
    case "error":
    case "closed":
      return "local-only";
  }
}

function resolveSteadySubstate(input: {
  hasZero: boolean;
  connectionStateName: ConnectionStateName;
}): "localReady" | "live" | "syncing" | "offline" {
  if (!input.hasZero) {
    return "localReady";
  }
  switch (input.connectionStateName) {
    case "connected":
      return "live";
    case "connecting":
      return "syncing";
    case "disconnected":
    case "needs-auth":
    case "error":
    case "closed":
      return "offline";
  }
}

function buildSessionCreateInput(
  input: WorkerRuntimeMachineInput,
): CreateWorkerRuntimeSessionInput {
  return {
    documentId: input.documentId,
    replicaId: input.replicaId,
    persistState: input.persistState,
    initialSelection: input.initialSelection,
    ...(input.perfSession ? { perfSession: input.perfSession } : {}),
    ...(input.zero ? { zero: input.zero } : {}),
    ...(input.fetchImpl ? { fetchImpl: input.fetchImpl } : {}),
    ...(input.createWorker ? { createWorker: input.createWorker } : {}),
  };
}

export function createWorkerRuntimeMachine() {
  const runtimeSessionActor = fromCallback(
    ({
      sendBack,
      receive,
      input,
    }: {
      sendBack: (event: WorkerRuntimeMachineEvent) => void;
      receive: (listener: (event: WorkerRuntimeMachineEvent) => void) => void;
      input: WorkerRuntimeMachineInput;
    }) => {
      const createSession = input.createSession ?? createWorkerRuntimeSessionController;
      let controller: WorkerRuntimeSessionController | null = null;
      let disposed = false;
      let pendingSelection = input.initialSelection;
      let pendingConnectionStateName = initialConnectionStateName(input);

      const applyExternalSyncState = async (): Promise<void> => {
        if (!controller) {
          return;
        }
        const value = await controller.invoke(
          "setExternalSyncState",
          mapConnectionStateToRuntimeSyncState(pendingConnectionStateName, Boolean(input.zero)),
        );
        if (!disposed && isWorkbookWorkerStateSnapshotValue(value)) {
          sendBack({ type: "session.runtime", runtimeState: value });
        }
      };

      receive((event) => {
        if (event.type === "selection.changed") {
          pendingSelection = event.selection;
          if (!controller) {
            return;
          }
          void controller.setSelection(event.selection).catch((error: unknown) => {
            if (!disposed) {
              sendBack({
                type: "session.error",
                message: error instanceof Error ? error.message : String(error),
              });
            }
          });
          return;
        }
        if (event.type === "connection.changed") {
          pendingConnectionStateName = event.connectionStateName;
          if (!controller) {
            return;
          }
          void applyExternalSyncState().catch((error: unknown) => {
            if (!disposed) {
              sendBack({
                type: "session.error",
                message: error instanceof Error ? error.message : String(error),
              });
            }
          });
        }
      });

      void createSession(
        buildSessionCreateInput({
          ...input,
          initialSelection: pendingSelection,
        }),
        {
          onRuntimeState(runtimeState) {
            sendBack({ type: "session.runtime", runtimeState });
          },
          onSelection(selection) {
            pendingSelection = selection;
            sendBack({ type: "session.selection", selection });
          },
          onPhase(phase) {
            sendBack({ type: "session.phase", phase });
          },
          onError(message) {
            sendBack({ type: "session.error", message });
          },
        },
      )
        .then((createdController) => {
          if (disposed) {
            createdController.dispose();
            return undefined;
          }
          controller = createdController;
          sendBack({ type: "session.ready", controller: createdController });
          void applyExternalSyncState().catch((error: unknown) => {
            if (!disposed) {
              sendBack({
                type: "session.error",
                message: error instanceof Error ? error.message : String(error),
              });
            }
          });
          if (
            pendingSelection.sheetName !== createdController.selection.sheetName ||
            pendingSelection.address !== createdController.selection.address
          ) {
            void createdController.setSelection(pendingSelection).catch((error: unknown) => {
              if (!disposed) {
                sendBack({
                  type: "session.error",
                  message: error instanceof Error ? error.message : String(error),
                });
              }
            });
          }
          return undefined;
        })
        .catch((error: unknown) => {
          if (!disposed) {
            sendBack({
              type: "session.failed",
              message: error instanceof Error ? error.message : String(error),
            });
          }
          return undefined;
        });

      return () => {
        disposed = true;
        controller?.dispose();
      };
    },
  );

  return setup<
    WorkerRuntimeMachineContext,
    WorkerRuntimeMachineEvent,
    {
      readonly runtimeSession: typeof runtimeSessionActor;
    },
    {},
    {},
    {},
    never,
    string,
    WorkerRuntimeMachineInput
  >({
    actors: {
      runtimeSession: runtimeSessionActor,
    },
  }).createMachine({
    id: "workerRuntime",
    initial: "active",
    context: ({ input }) => ({
      sessionInput: input,
      persistState: input.persistState,
      controller: null,
      handle: null,
      runtimeState: null,
      selection: input.initialSelection,
      connectionStateName: initialConnectionStateName(input),
      error: null,
    }),
    states: {
      active: {
        invoke: {
          id: "runtimeSession",
          src: "runtimeSession",
          input: ({ context }) => ({
            ...context.sessionInput,
            persistState: context.persistState,
            initialSelection: context.selection,
            connectionStateName: context.connectionStateName,
          }),
        },
        on: {
          retry: {
            target: "#workerRuntime.active",
            reenter: true,
            actions: assign({
              persistState: ({ context, event }) => event["persistState"] ?? context.persistState,
              handle: () => null,
              controller: () => null,
              runtimeState: () => null,
              error: () => null,
            }),
          },
          "error.clear": {
            actions: assign({
              error: () => null,
            }),
          },
          "selection.changed": {
            actions: [
              assign({
                selection: ({ event }) => event["selection"],
              }),
              sendTo("runtimeSession", ({ event }) => event),
            ],
          },
          "connection.changed": {
            actions: [
              assign({
                connectionStateName: ({ event }) => event["connectionStateName"],
              }),
              sendTo("runtimeSession", ({ event }) => event),
            ],
          },
          "session.runtime": {
            actions: assign({
              runtimeState: ({ event }) => event["runtimeState"],
            }),
          },
          "session.selection": {
            actions: assign({
              selection: ({ event }) => event["selection"],
            }),
          },
          "session.error": {
            actions: assign({
              error: ({ event }) => event["message"],
            }),
          },
          "session.failed": {
            target: "failed",
            actions: assign({
              error: ({ event }) => event["message"],
              controller: () => null,
              handle: () => null,
              runtimeState: () => null,
            }),
          },
          "session.ready": [
            {
              guard: ({ context }) =>
                resolveSteadySubstate({
                  hasZero: Boolean(context.sessionInput.zero),
                  connectionStateName: context.connectionStateName,
                }) === "live",
              target: ".live",
              actions: assign({
                handle: ({ event }) => event["controller"].handle,
                controller: ({ event }) => event["controller"],
                runtimeState: ({ event }) => event["controller"].runtimeState,
                selection: ({ event }) => event["controller"].selection,
                error: () => null,
              }),
            },
            {
              guard: ({ context }) =>
                resolveSteadySubstate({
                  hasZero: Boolean(context.sessionInput.zero),
                  connectionStateName: context.connectionStateName,
                }) === "syncing",
              target: ".syncing",
              actions: assign({
                handle: ({ event }) => event["controller"].handle,
                controller: ({ event }) => event["controller"],
                runtimeState: ({ event }) => event["controller"].runtimeState,
                selection: ({ event }) => event["controller"].selection,
                error: () => null,
              }),
            },
            {
              guard: ({ context }) =>
                resolveSteadySubstate({
                  hasZero: Boolean(context.sessionInput.zero),
                  connectionStateName: context.connectionStateName,
                }) === "offline",
              target: ".offline",
              actions: assign({
                handle: ({ event }) => event["controller"].handle,
                controller: ({ event }) => event["controller"],
                runtimeState: ({ event }) => event["controller"].runtimeState,
                selection: ({ event }) => event["controller"].selection,
                error: () => null,
              }),
            },
            {
              target: ".localReady",
              actions: assign({
                handle: ({ event }) => event["controller"].handle,
                controller: ({ event }) => event["controller"],
                runtimeState: ({ event }) => event["controller"].runtimeState,
                selection: ({ event }) => event["controller"].selection,
                error: () => null,
              }),
            },
          ],
          "session.phase": [
            {
              guard: ({ event }) => event["phase"] === "hydratingLocal",
              target: ".hydratingLocal",
            },
            {
              guard: ({ event }) => event["phase"] === "syncing",
              target: ".syncing",
            },
            {
              guard: ({ event }) => event["phase"] === "reconciling",
              target: ".reconciling",
            },
            {
              guard: ({ event }) => event["phase"] === "recovering",
              target: ".recovering",
            },
            {
              guard: ({ context, event }) =>
                event["phase"] === "steady" &&
                resolveSteadySubstate({
                  hasZero: Boolean(context.sessionInput.zero),
                  connectionStateName: context.connectionStateName,
                }) === "live",
              target: ".live",
            },
            {
              guard: ({ context, event }) =>
                event["phase"] === "steady" &&
                resolveSteadySubstate({
                  hasZero: Boolean(context.sessionInput.zero),
                  connectionStateName: context.connectionStateName,
                }) === "syncing",
              target: ".syncing",
            },
            {
              guard: ({ context, event }) =>
                event["phase"] === "steady" &&
                resolveSteadySubstate({
                  hasZero: Boolean(context.sessionInput.zero),
                  connectionStateName: context.connectionStateName,
                }) === "offline",
              target: ".offline",
            },
            {
              guard: ({ event }) => event["phase"] === "steady",
              target: ".localReady",
            },
          ],
        },
        initial: "booting",
        states: {
          booting: {},
          hydratingLocal: {},
          syncing: {
            on: {
              "connection.changed": [
                {
                  guard: ({ context, event }) =>
                    context.controller !== null &&
                    resolveSteadySubstate({
                      hasZero: Boolean(context.sessionInput.zero),
                      connectionStateName: event["connectionStateName"],
                    }) === "live",
                  target: "#workerRuntime.active.live",
                  actions: [
                    assign({
                      connectionStateName: ({ event }) => event["connectionStateName"],
                    }),
                    sendTo("runtimeSession", ({ event }) => event),
                  ],
                },
                {
                  guard: ({ context, event }) =>
                    context.controller !== null &&
                    resolveSteadySubstate({
                      hasZero: Boolean(context.sessionInput.zero),
                      connectionStateName: event["connectionStateName"],
                    }) === "offline",
                  target: "#workerRuntime.active.offline",
                  actions: [
                    assign({
                      connectionStateName: ({ event }) => event["connectionStateName"],
                    }),
                    sendTo("runtimeSession", ({ event }) => event),
                  ],
                },
                {
                  guard: ({ context, event }) =>
                    context.controller !== null &&
                    resolveSteadySubstate({
                      hasZero: Boolean(context.sessionInput.zero),
                      connectionStateName: event["connectionStateName"],
                    }) === "localReady",
                  target: "#workerRuntime.active.localReady",
                  actions: [
                    assign({
                      connectionStateName: ({ event }) => event["connectionStateName"],
                    }),
                    sendTo("runtimeSession", ({ event }) => event),
                  ],
                },
                {
                  actions: [
                    assign({
                      connectionStateName: ({ event }) => event["connectionStateName"],
                    }),
                    sendTo("runtimeSession", ({ event }) => event),
                  ],
                },
              ],
            },
          },
          localReady: {
            on: {
              "connection.changed": [
                {
                  guard: ({ event }) => event["connectionStateName"] === "connected",
                  target: "#workerRuntime.active.live",
                  actions: [
                    assign({
                      connectionStateName: ({ event }) => event["connectionStateName"],
                    }),
                    sendTo("runtimeSession", ({ event }) => event),
                  ],
                },
                {
                  actions: [
                    assign({
                      connectionStateName: ({ event }) => event["connectionStateName"],
                    }),
                    sendTo("runtimeSession", ({ event }) => event),
                  ],
                },
              ],
            },
          },
          live: {
            on: {
              "connection.changed": [
                {
                  guard: ({ event }) => event["connectionStateName"] === "connected",
                  actions: [
                    assign({
                      connectionStateName: ({ event }) => event["connectionStateName"],
                    }),
                    sendTo("runtimeSession", ({ event }) => event),
                  ],
                },
                {
                  guard: ({ event }) => event["connectionStateName"] === "connecting",
                  target: "#workerRuntime.active.syncing",
                  actions: [
                    assign({
                      connectionStateName: ({ event }) => event["connectionStateName"],
                    }),
                    sendTo("runtimeSession", ({ event }) => event),
                  ],
                },
                {
                  target: "#workerRuntime.active.offline",
                  actions: [
                    assign({
                      connectionStateName: ({ event }) => event["connectionStateName"],
                    }),
                    sendTo("runtimeSession", ({ event }) => event),
                  ],
                },
              ],
            },
          },
          offline: {
            on: {
              "connection.changed": [
                {
                  guard: ({ event }) => event["connectionStateName"] === "connected",
                  target: "#workerRuntime.active.live",
                  actions: [
                    assign({
                      connectionStateName: ({ event }) => event["connectionStateName"],
                    }),
                    sendTo("runtimeSession", ({ event }) => event),
                  ],
                },
                {
                  guard: ({ event }) => event["connectionStateName"] === "connecting",
                  target: "#workerRuntime.active.syncing",
                  actions: [
                    assign({
                      connectionStateName: ({ event }) => event["connectionStateName"],
                    }),
                    sendTo("runtimeSession", ({ event }) => event),
                  ],
                },
                {
                  actions: [
                    assign({
                      connectionStateName: ({ event }) => event["connectionStateName"],
                    }),
                    sendTo("runtimeSession", ({ event }) => event),
                  ],
                },
              ],
            },
          },
          reconciling: {},
          recovering: {},
        },
      },
      failed: {
        on: {
          "error.clear": {
            actions: assign({
              error: () => null,
            }),
          },
          retry: {
            target: "active",
            actions: assign({
              persistState: ({ context, event }) => event["persistState"] ?? context.persistState,
              handle: () => null,
              controller: () => null,
              runtimeState: () => null,
              error: () => null,
            }),
          },
        },
      },
    },
  });
}
