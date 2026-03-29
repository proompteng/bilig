import { createActor } from "xstate";
import { describe, expect, it } from "vitest";

import { createDocumentSupervisorMachine } from "../index.js";

describe("createDocumentSupervisorMachine", () => {
  it("tracks active lifecycle, cursors, and browser subscribers", () => {
    const actor = createActor(createDocumentSupervisorMachine("book-1"));
    actor.start();

    actor.send({ type: "browser.attached" });
    actor.send({ type: "cursor.updated", cursor: 12 });
    actor.send({ type: "snapshot.updated", cursor: 8 });
    actor.send({ type: "operation.recorded", operation: "openBrowserSession" });
    actor.send({ type: "browser.detached" });

    const snapshot = actor.getSnapshot();
    expect(snapshot.matches("active")).toBe(true);
    expect(snapshot.context.documentId).toBe("book-1");
    expect(snapshot.context.browserSubscriberCount).toBe(0);
    expect(snapshot.context.lastKnownCursor).toBe(12);
    expect(snapshot.context.lastSnapshotCursor).toBe(8);
    expect(snapshot.context.lastOperation).toBe("openBrowserSession");
  });

  it("enters degraded and recovers on reset", () => {
    const actor = createActor(createDocumentSupervisorMachine("book-2"));
    actor.start();

    actor.send({ type: "error.raised", message: "boom" });
    expect(actor.getSnapshot().matches("degraded")).toBe(true);
    expect(actor.getSnapshot().context.lastError).toBe("boom");

    actor.send({ type: "reset" });
    expect(actor.getSnapshot().matches("active")).toBe(true);
    expect(actor.getSnapshot().context.lastError).toBeNull();
  });
});
