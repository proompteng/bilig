import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

const installOpfsSAHPoolVfs = vi.fn();
const sqlite3InitModule = vi.fn(async () => ({
  installOpfsSAHPoolVfs,
}));

vi.mock("@sqlite.org/sqlite-wasm", () => ({
  default: sqlite3InitModule,
}));

describe("opfs workbook local store", () => {
  beforeEach(() => {
    installOpfsSAHPoolVfs.mockReset();
    sqlite3InitModule.mockClear();
    delete (globalThis as { document?: Document }).document;
    Object.defineProperty(globalThis, "navigator", {
      configurable: true,
      value: {
        storage: {
          getDirectory: vi.fn(),
        },
      },
    });
  });

  afterEach(() => {
    vi.resetModules();
  });

  it("caches access-handle lock failures instead of reinstalling the sahpool vfs", async () => {
    const accessHandleError = Object.assign(
      new Error(
        "Failed to execute 'createSyncAccessHandle' on 'FileSystemFileHandle': Access Handles cannot be created if there is another open Access Handle or Writable stream associated with the same file.",
      ),
      { name: "NoModificationAllowedError" },
    );
    installOpfsSAHPoolVfs.mockRejectedValue(accessHandleError);

    const { WorkbookLocalStoreLockedError, createOpfsWorkbookLocalStoreFactory } =
      await import("../index.js");
    const factory = createOpfsWorkbookLocalStoreFactory();

    await expect(factory.open("doc-a")).rejects.toMatchObject({
      name: WorkbookLocalStoreLockedError.name,
      message: "Workbook local store is locked by another tab for doc-a",
    });
    await expect(factory.open("doc-b")).rejects.toMatchObject({
      name: WorkbookLocalStoreLockedError.name,
      message: "Workbook local store is locked by another tab for doc-b",
    });

    expect(sqlite3InitModule).toHaveBeenCalledTimes(1);
    expect(installOpfsSAHPoolVfs).toHaveBeenCalledTimes(1);
    expect(installOpfsSAHPoolVfs).toHaveBeenCalledWith(
      expect.objectContaining({
        verbosity: 0,
      }),
    );
  });

  it("treats removeVfs NoModificationAllowedError failures as access-handle lock conflicts", async () => {
    const removeVfsError = Object.assign(
      new Error(
        "Failed to execute 'removeEntry' on 'FileSystemDirectoryHandle': An attempt was made to modify an object where modifications are not allowed.",
      ),
      { name: "NoModificationAllowedError" },
    );
    installOpfsSAHPoolVfs.mockRejectedValue(removeVfsError);

    const { WorkbookLocalStoreLockedError, createOpfsWorkbookLocalStoreFactory } =
      await import("../index.js");
    const factory = createOpfsWorkbookLocalStoreFactory();

    await expect(factory.open("doc-remove-vfs")).rejects.toMatchObject({
      name: WorkbookLocalStoreLockedError.name,
      message: "Workbook local store is locked by another tab for doc-remove-vfs",
    });

    expect(sqlite3InitModule).toHaveBeenCalledTimes(1);
    expect(installOpfsSAHPoolVfs).toHaveBeenCalledTimes(1);
  });

  it("treats error-like removeVfs lock failures as access-handle conflicts", async () => {
    installOpfsSAHPoolVfs.mockRejectedValue({
      name: "NoModificationAllowedError",
      message:
        "Failed to execute 'removeEntry' on 'FileSystemDirectoryHandle': An attempt was made to modify an object where modifications are not allowed.",
    });

    const { WorkbookLocalStoreLockedError, createOpfsWorkbookLocalStoreFactory } =
      await import("../index.js");
    const factory = createOpfsWorkbookLocalStoreFactory();

    await expect(factory.open("doc-remove-vfs-like")).rejects.toMatchObject({
      name: WorkbookLocalStoreLockedError.name,
      message: "Workbook local store is locked by another tab for doc-remove-vfs-like",
    });
  });
});
