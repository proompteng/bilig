import { describe, expect, it } from "vitest";

import { createInMemoryDocumentPersistence } from "../index.js";

describe("storage-server", () => {
  it("assigns monotonic cursors and lists after a cursor", async () => {
    const persistence = createInMemoryDocumentPersistence();

    await persistence.batches.append("book-1", {
      id: "a:1",
      replicaId: "a",
      clock: { counter: 1 },
      ops: []
    });
    await persistence.batches.append("book-1", {
      id: "a:2",
      replicaId: "a",
      clock: { counter: 2 },
      ops: []
    });

    const entries = await persistence.batches.listAfter("book-1", 1);
    expect(entries.map((entry) => entry.cursor)).toEqual([2]);
  });

  it("tracks owner leases and presence", async () => {
    const persistence = createInMemoryDocumentPersistence();

    await expect(persistence.ownership.claim("book-1", "svc-a", Date.now() + 1000)).resolves.toBe(true);
    await expect(persistence.ownership.claim("book-1", "svc-b", Date.now() + 1000)).resolves.toBe(false);

    await persistence.presence.join("book-1", "sess-1");
    await persistence.presence.join("book-1", "sess-2");
    expect(await persistence.presence.sessions("book-1")).toEqual(["sess-1", "sess-2"]);
  });
});
