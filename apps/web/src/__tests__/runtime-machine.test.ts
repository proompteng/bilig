import { describe, expect, it, vi } from "vitest";
import { createActor } from "xstate";
import { createWorkerRuntimeMachine } from "../runtime-machine.js";
import { ProjectedViewportStore } from "../projected-viewport-store.js";
import type {
  CreateWorkerRuntimeSessionInput,
  WorkerHandle,
  WorkerRuntimeSessionCallbacks,
  WorkerRuntimeSessionController,
} from "../runtime-session.js";

function createWorkerHandle(): WorkerHandle {
  return {
    viewportStore: new ProjectedViewportStore(),
  };
}

function createController(
  selection = { sheetName: "Sheet1", address: "A1" },
): WorkerRuntimeSessionController {
  return {
    handle: createWorkerHandle(),
    runtimeState: {
      workbookName: "bilig-demo",
      sheetNames: ["Sheet1"],
      metrics: {
        batchId: 0,
        changedInputCount: 0,
        dirtyFormulaCount: 0,
        wasmFormulaCount: 0,
        jsFormulaCount: 0,
        rangeNodeVisits: 0,
        recalcMs: 0,
        compileMs: 0,
      },
      syncState: "local-only",
    },
    selection,
    invoke: vi.fn(async () => undefined),
    setSelection: vi.fn(async () => undefined),
    subscribeViewport: () => () => {},
    dispose: vi.fn(),
  };
}

describe("worker runtime machine", () => {
  it("boots into ready and forwards selection changes to the session controller", async () => {
    const controller = createController();
    const createSession = vi.fn(
      async (
        _input: CreateWorkerRuntimeSessionInput,
        callbacks: WorkerRuntimeSessionCallbacks,
      ): Promise<WorkerRuntimeSessionController> => {
        callbacks.onRuntimeState(controller.runtimeState);
        return controller;
      },
    );

    const actor = createActor(createWorkerRuntimeMachine(), {
      input: {
        documentId: "book-1",
        replicaId: "browser:test",
        persistState: true,
        initialSelection: { sheetName: "Sheet1", address: "A1" },
        createSession,
      },
    });

    actor.start();
    await vi.waitFor(() => {
      expect(actor.getSnapshot().matches({ active: "ready" })).toBe(true);
    });

    const nextSelection = { sheetName: "Sheet1", address: "C3" } as const;
    actor.send({ type: "selection.changed", selection: nextSelection });

    await vi.waitFor(() => {
      expect(controller.setSelection).toHaveBeenCalledWith(nextSelection);
    });

    expect(actor.getSnapshot().context.selection).toEqual(nextSelection);
    actor.stop();
  });

  it("transitions to failed when session startup rejects and recovers on retry", async () => {
    const controller = createController();
    const createSession = vi
      .fn<
        (
          input: CreateWorkerRuntimeSessionInput,
          callbacks: WorkerRuntimeSessionCallbacks,
        ) => Promise<WorkerRuntimeSessionController>
      >()
      .mockRejectedValueOnce(new Error("bootstrap failed"))
      .mockResolvedValueOnce(controller);

    const actor = createActor(createWorkerRuntimeMachine(), {
      input: {
        documentId: "book-1",
        replicaId: "browser:test",
        persistState: true,
        initialSelection: { sheetName: "Sheet1", address: "A1" },
        createSession,
      },
    });

    actor.start();

    await vi.waitFor(() => {
      expect(actor.getSnapshot().matches("failed")).toBe(true);
    });
    expect(actor.getSnapshot().context.error).toBe("bootstrap failed");

    actor.send({ type: "retry" });

    await vi.waitFor(() => {
      expect(actor.getSnapshot().matches({ active: "ready" })).toBe(true);
    });

    actor.stop();
  });
});
