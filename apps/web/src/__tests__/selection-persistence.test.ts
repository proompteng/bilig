import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { loadPersistedSelection, persistSelection } from "../selection-persistence.js";

describe("selection persistence", () => {
  const storage = new Map<string, string>();

  beforeEach(() => {
    storage.clear();
    vi.stubGlobal("window", {
      localStorage: {
        getItem(key: string) {
          return storage.get(key) ?? null;
        },
        setItem(key: string, value: string) {
          storage.set(key, value);
        },
        clear() {
          storage.clear();
        },
      },
    });
  });

  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it("falls back to Sheet1!A1 when nothing is stored", () => {
    expect(loadPersistedSelection("book-1")).toEqual({
      sheetName: "Sheet1",
      address: "A1",
    });
  });

  it("restores the last stored sheet selection for a document", () => {
    persistSelection("book-1", { sheetName: "Sheet3", address: "G22" });

    expect(loadPersistedSelection("book-1")).toEqual({
      sheetName: "Sheet3",
      address: "G22",
    });
  });

  it("ignores invalid stored values", () => {
    storage.set("bilig:selection:book-1", '{"sheetName":"","address":42}');

    expect(loadPersistedSelection("book-1")).toEqual({
      sheetName: "Sheet1",
      address: "A1",
    });
  });
});
