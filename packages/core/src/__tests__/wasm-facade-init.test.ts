import { afterEach, describe, expect, it, vi } from "vitest";

afterEach(() => {
  vi.resetModules();
  vi.doUnmock("@bilig/wasm-kernel");
  vi.clearAllMocks();
});

describe("WasmKernelFacade init failures", () => {
  it("keeps the facade unready when kernel creation fails", async () => {
    vi.doMock("@bilig/wasm-kernel", () => ({
      createKernel: vi.fn(async () => {
        throw new Error("kernel init failed");
      }),
    }));

    const { WasmKernelFacade } = await import("../wasm-facade.js");
    const facade = new WasmKernelFacade();

    await expect(facade.init()).resolves.toBeUndefined();
    expect(facade.ready).toBe(false);
  });
});
