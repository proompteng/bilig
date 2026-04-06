import { describe, expect, it } from "vitest";

import { createBrowserMetadataStore } from "../index.js";

describe("storage-browser", () => {
  it("creates a persistence facade", async () => {
    const persistence = createBrowserMetadataStore({ databaseName: "spec", storeName: "state" });
    expect(typeof persistence.loadJson).toBe("function");
    expect(typeof persistence.saveJson).toBe("function");
    expect(typeof persistence.remove).toBe("function");
  });
});
