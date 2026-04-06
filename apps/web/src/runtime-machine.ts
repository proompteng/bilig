import { assign, fromCallback, sendTo, setup } from "xstate";
import type { WorkbookWorkerStateSnapshot } from "./worker-runtime.js";
import {
  createWorkerRuntimeSessionController,
  type CreateWorkerRuntimeSessionInput,
  type WorkerHandle,
  type WorkerRuntimeSelection,
  type WorkerRuntimeSessionController,
} from "./runtime-session.js";
import type { ZeroWorkbookBridgeState } from "./zero/ZeroWorkbookBridge.js";

interface WorkerRuntimeMachineContext {
  readonly sessionInput: WorkerRuntimeMachineInput;
  readonly controller: WorkerRuntimeSessionController | null;
  readonly handle: WorkerHandle | null;
  readonly runtimeState: WorkbookWorkerStateSnapshot | null;
  readonly bridgeState: ZeroWorkbookBridgeState | null;
  readonly selection: WorkerRuntimeSelection;
  readonly error: string | null;
}

type WorkerRuntimeMachineEvent =
  | { type: "retry" }
  | { type: "selection.changed"; selection: WorkerRuntimeSelection }
  | {
      type: "session.ready";
      controller: WorkerRuntimeSessionController;
    }
  | { type: "session.runtime"; runtimeState: WorkbookWorkerStateSnapshot }
  | { type: "session.bridge"; bridgeState: ZeroWorkbookBridgeState | null }
  | { type: "session.selection"; selection: WorkerRuntimeSelection }
  | { type: "session.error"; message: string }
  | { type: "session.failed"; message: string };

export interface WorkerRuntimeMachineInput extends CreateWorkerRuntimeSessionInput {
  readonly createSession?: (
    input: CreateWorkerRuntimeSessionInput,
    callbacks: Parameters<typeof createWorkerRuntimeSessionController>[1],
  ) => Promise<WorkerRuntimeSessionController>;
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

      receive((event) => {
        if (event.type !== "selection.changed") {
          return;
        }
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
      });

      void createSession(
        {
          ...input,
          initialSelection: pendingSelection,
        },
        {
          onRuntimeState(runtimeState) {
            sendBack({ type: "session.runtime", runtimeState });
          },
          onBridgeState(bridgeState) {
            sendBack({ type: "session.bridge", bridgeState });
          },
          onSelection(selection) {
            pendingSelection = selection;
            sendBack({ type: "session.selection", selection });
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
      controller: null,
      handle: null,
      runtimeState: null,
      bridgeState: null,
      selection: input.initialSelection,
      error: null,
    }),
    states: {
      active: {
        invoke: {
          id: "runtimeSession",
          src: "runtimeSession",
          input: ({ context }) => ({
            ...context.sessionInput,
            initialSelection: context.selection,
          }),
        },
        on: {
          "selection.changed": {
            actions: [
              assign({
                selection: ({ event }) => event.selection,
              }),
              sendTo("runtimeSession", ({ event }) => event),
            ],
          },
          "session.runtime": {
            actions: assign({
              runtimeState: ({ event }) => event.runtimeState,
            }),
          },
          "session.bridge": {
            actions: assign({
              bridgeState: ({ event }) => event.bridgeState,
            }),
          },
          "session.selection": {
            actions: assign({
              selection: ({ event }) => event.selection,
            }),
          },
          "session.error": {
            actions: assign({
              error: ({ event }) => event.message,
            }),
          },
          "session.failed": {
            target: "failed",
            actions: assign({
              error: ({ event }) => event.message,
              controller: () => null,
              handle: () => null,
              runtimeState: () => null,
              bridgeState: () => null,
            }),
          },
        },
        initial: "booting",
        states: {
          booting: {
            on: {
              "session.ready": {
                target: "ready",
                actions: assign({
                  handle: ({ event }) => event.controller.handle,
                  controller: ({ event }) => event.controller,
                  runtimeState: ({ event }) => event.controller.runtimeState,
                  bridgeState: ({ event }) => event.controller.bridgeState,
                  selection: ({ event }) => event.controller.selection,
                  error: () => null,
                }),
              },
            },
          },
          ready: {},
        },
      },
      failed: {
        on: {
          retry: {
            target: "active",
            actions: assign({
              handle: () => null,
              controller: () => null,
              runtimeState: () => null,
              bridgeState: () => null,
              error: () => null,
            }),
          },
        },
      },
    },
  });
}
