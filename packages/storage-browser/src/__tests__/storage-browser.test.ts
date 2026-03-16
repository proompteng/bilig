import { describe, expect, it } from "vitest";

import { createBrowserPersistence } from "../index.js";

describe("storage-browser", () => {
  it("creates a persistence facade", async () => {
    const persistence = createBrowserPersistence({ databaseName: "spec", storeName: "state" });
    expect(typeof persistence.loadJson).toBe("function");
    expect(typeof persistence.saveJson).toBe("function");
    expect(typeof persistence.remove).toBe("function");
  });
});
