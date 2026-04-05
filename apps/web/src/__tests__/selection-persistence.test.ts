import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { loadPersistedSelection, persistSelection } from "../selection-persistence.js";

describe("selection persistence", () => {
  const storage = new Map<string, string>();
  const replaceState = vi.fn();

  beforeEach(() => {
    storage.clear();
    replaceState.mockReset();
    vi.stubGlobal("window", {
      history: {
        replaceState,
        state: { from: "test" },
      },
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
      location: new URL("https://bilig.test/"),
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

  it("prefers a URL-backed sheet selection over local storage", () => {
    storage.set("bilig:selection:book-1", JSON.stringify({ sheetName: "Sheet3", address: "G22" }));
    vi.stubGlobal("window", {
      history: {
        replaceState,
        state: { from: "test" },
      },
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
      location: new URL("https://bilig.test/?sheet=Sheet7"),
    });

    expect(loadPersistedSelection("book-1")).toEqual({
      sheetName: "Sheet7",
      address: "A1",
    });
  });

  it("reuses the stored address when the URL sheet matches it", () => {
    storage.set("bilig:selection:book-1", JSON.stringify({ sheetName: "Sheet7", address: "G22" }));
    vi.stubGlobal("window", {
      history: {
        replaceState,
        state: { from: "test" },
      },
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
      location: new URL("https://bilig.test/?sheet=Sheet7"),
    });

    expect(loadPersistedSelection("book-1")).toEqual({
      sheetName: "Sheet7",
      address: "G22",
    });
  });

  it("writes only the sheet into the URL state", () => {
    persistSelection("book-1", { sheetName: "Sheet7", address: "b12" });

    expect(replaceState).toHaveBeenCalledTimes(1);
    const [, , nextUrl] = replaceState.mock.calls[0];
    expect(String(nextUrl)).toBe("https://bilig.test/?sheet=Sheet7");
    expect(storage.get("bilig:selection:book-1")).toBe(
      JSON.stringify({ sheetName: "Sheet7", address: "B12" }),
    );
  });
});
