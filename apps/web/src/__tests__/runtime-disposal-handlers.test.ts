import { describe, expect, it, vi } from "vitest";

import { registerRuntimeDisposalHandlers } from "../runtime-disposal-handlers.js";

describe("registerRuntimeDisposalHandlers", () => {
  it("disposes the latest controller on pagehide and beforeunload", () => {
    const listeners = new Map<string, () => void>();
    const target = {
      addEventListener(type: "beforeunload" | "pagehide", listener: () => void) {
        listeners.set(type, listener);
      },
      removeEventListener(type: "beforeunload" | "pagehide", listener: () => void) {
        if (listeners.get(type) === listener) {
          listeners.delete(type);
        }
      },
    };
    const firstDispose = vi.fn();
    const secondDispose = vi.fn();
    let controller: { dispose(): void } | null = { dispose: firstDispose };

    const cleanup = registerRuntimeDisposalHandlers({
      getController: () => controller,
      target,
    });

    listeners.get("pagehide")?.();
    expect(firstDispose).toHaveBeenCalledTimes(1);

    controller = { dispose: secondDispose };
    listeners.get("beforeunload")?.();
    expect(secondDispose).toHaveBeenCalledTimes(1);

    cleanup();
    expect(listeners.size).toBe(0);
  });
});
