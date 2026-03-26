import { SpreadsheetEngine } from "@bilig/core";
import { describe, expect, it, vi } from "vitest";

const createFiberRoot = vi.fn(() => ({ kind: "fiber-root" }));
const updateFiberRoot = vi.fn(() => {
  throw new Error("update failed");
});

vi.mock("../compat.js", () => ({
  createFiberRoot,
  updateFiberRoot,
}));

describe("renderer root error handling", () => {
  it("rejects unmount when the compat layer throws during update", async () => {
    const { createWorkbookRendererRoot } = await import("../renderer-root.js");
    const engine = new SpreadsheetEngine({ workbookName: "renderer-root-unmount-error" });
    await engine.ready();
    const root = createWorkbookRendererRoot(engine);

    await expect(root.unmount()).rejects.toThrow("update failed");
    expect(updateFiberRoot).toHaveBeenCalledWith(
      { kind: "fiber-root" },
      null,
      expect.any(Function),
    );
  });
});
