import { SpreadsheetEngine } from "@bilig/core";
import { beforeEach, describe, expect, it, vi } from "vitest";
import type { WorkbookContainer } from "../host-config.js";

const createContainer = vi.fn();
const updateContainer = vi.fn();

vi.mock("../host-config.js", () => ({
  WorkbookReconciler: {
    createContainer,
    updateContainer,
  },
}));

describe("renderer compat", () => {
  beforeEach(() => {
    createContainer.mockReset();
    updateContainer.mockReset();
    vi.resetModules();
  });

  async function makeContainer(workbookName: string): Promise<WorkbookContainer> {
    const engine = new SpreadsheetEngine({ workbookName });
    await engine.ready();
    return {
      engine,
      root: null,
      pendingOps: [],
      shouldSyncSheetOrders: false,
      lastError: null,
    };
  }

  it("normalizes non-Error renderer failures into Error instances", async () => {
    createContainer.mockImplementation(
      (
        container: { lastError: Error | null },
        _tag: unknown,
        _hydrate: unknown,
        _strict: unknown,
        _concurrentOverride: unknown,
        _identifierPrefix: unknown,
        onCaughtError: (error: unknown) => void,
      ) => {
        onCaughtError("boom");
        return { kind: "root", container };
      },
    );

    const { createFiberRoot } = await import("../compat.js");
    const container = await makeContainer("renderer-compat-non-error");

    expect(createFiberRoot(container)).toEqual({ kind: "root", container });
    expect(container.lastError).toBeInstanceOf(Error);
    expect(container.lastError?.message).toBe("boom");
  });

  it("preserves Error instances raised by the reconciler callbacks", async () => {
    const originalError = new Error("renderer failed");
    createContainer.mockImplementation(
      (
        container: { lastError: Error | null },
        _tag: unknown,
        _hydrate: unknown,
        _strict: unknown,
        _concurrentOverride: unknown,
        _identifierPrefix: unknown,
        _onUncaughtError: (error: unknown) => void,
        onCaughtError: (error: unknown) => void,
      ) => {
        onCaughtError(originalError);
        return { kind: "root", container };
      },
    );

    const { createFiberRoot } = await import("../compat.js");
    const container = await makeContainer("renderer-compat-error");

    createFiberRoot(container);
    expect(container.lastError).toBe(originalError);
  });
});
