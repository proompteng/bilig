import { describe, expect, it } from "vitest";
import { createActor } from "xstate";
import { createBootstrapMachine } from "../index.js";

describe("@bilig/actors bootstrap machine", () => {
  it("reaches ready after loading config and session", async () => {
    const machine = createBootstrapMachine<{ apiBaseUrl: string }, { authToken: string }>();
    const actor = createActor(machine, {
      input: {
        loadConfig: async () => ({ apiBaseUrl: "http://127.0.0.1:4321" }),
        loadSession: async () => ({ authToken: "token-123" }),
      },
    });

    const done = new Promise<void>((resolve, reject) => {
      const subscription = actor.subscribe((snapshot) => {
        if (snapshot.matches("ready")) {
          subscription.unsubscribe();
          resolve();
          return;
        }
        if (snapshot.matches("failed")) {
          subscription.unsubscribe();
          reject(new Error(snapshot.context.error ?? "bootstrap failed"));
        }
      });
    });

    actor.start();
    await done;

    expect(actor.getSnapshot().context.config).toEqual({
      apiBaseUrl: "http://127.0.0.1:4321",
    });
    expect(actor.getSnapshot().context.session).toEqual({
      authToken: "token-123",
    });
  });

  it("supports retry after a failed config load", async () => {
    let attempts = 0;
    const machine = createBootstrapMachine<{ apiBaseUrl: string }, { authToken: string }>();
    const actor = createActor(machine, {
      input: {
        loadConfig: async () => {
          attempts += 1;
          if (attempts === 1) {
            throw new Error("temporary failure");
          }
          return { apiBaseUrl: "http://127.0.0.1:4321" };
        },
        loadSession: async () => ({ authToken: "token-123" }),
      },
    });

    actor.start();
    await new Promise<void>((resolve) => {
      const subscription = actor.subscribe((snapshot) => {
        if (snapshot.matches("failed")) {
          subscription.unsubscribe();
          resolve();
        }
      });
    });

    actor.send({ type: "retry" });
    await new Promise<void>((resolve, reject) => {
      const subscription = actor.subscribe((snapshot) => {
        if (snapshot.matches("ready")) {
          subscription.unsubscribe();
          resolve();
          return;
        }
        if (snapshot.matches("failed") && attempts > 1) {
          subscription.unsubscribe();
          reject(new Error(snapshot.context.error ?? "retry failed"));
        }
      });
    });

    expect(attempts).toBe(2);
  });
});
