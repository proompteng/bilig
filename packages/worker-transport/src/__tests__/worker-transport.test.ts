import { MessageChannel } from "node:worker_threads";

import { describe, expect, it } from "vitest";

import type { EngineEvent } from "@bilig/protocol";

import { createWorkerEngineClient, createWorkerEngineHost } from "../index.js";

async function waitFor(predicate: () => boolean, attempts = 20): Promise<void> {
  for (let attempt = 0; attempt < attempts; attempt += 1) {
    if (predicate()) {
      return;
    }
    await new Promise((resolve) => setTimeout(resolve, 0));
  }
  throw new Error("Timed out waiting for worker transport condition");
}

describe("worker transport", () => {
  it("invokes engine methods across a message channel", async () => {
    const channel = new MessageChannel();
    const host = createWorkerEngineHost({
      async ready() {
        return;
      },
      add(left: number, right: number) {
        return left + right;
      }
    }, channel.port1);

    const client = createWorkerEngineClient({ port: channel.port2 });

    await expect(client.ready()).resolves.toBeUndefined();
    await expect(client.invoke("add", 2, 5)).resolves.toBe(7);

    client.dispose();
    host.dispose();
  });

  it("relays subscriptions back to the client", async () => {
    const channel = new MessageChannel();
    const eventListeners = new Set<(event: EngineEvent) => void>();
    const host = createWorkerEngineHost({
      subscribe(listener: (event: EngineEvent) => void) {
        eventListeners.add(listener);
        return () => eventListeners.delete(listener);
      }
    }, channel.port1);

    const client = createWorkerEngineClient({ port: channel.port2 });
    const received: EngineEvent[] = [];
    const unsubscribe = client.subscribe((event) => {
      received.push(event);
    });

    await waitFor(() => eventListeners.size === 1);

    eventListeners.forEach((listener) => {
      listener({
        kind: "batch",
        changedCellIndices: Uint32Array.from([1, 2]),
        metrics: {
          batchId: 1,
          changedInputCount: 1,
          dirtyFormulaCount: 0,
          wasmFormulaCount: 0,
          jsFormulaCount: 0,
          rangeNodeVisits: 0,
          recalcMs: 0,
          compileMs: 0
        }
      });
    });

    await waitFor(() => received.length === 1);

    expect(received).toHaveLength(1);
    unsubscribe();
    client.dispose();
    host.dispose();
  });
});
