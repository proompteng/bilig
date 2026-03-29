// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from "vitest";
import { resolveRuntimeConfig } from "../runtime-config";
import { resolveZeroCacheUrl } from "../zero-connection";

const BASE_CONFIG = {
  apiBaseUrl: "http://127.0.0.1:4321",
  zeroCacheUrl: "http://127.0.0.1:4848",
  defaultDocumentId: "bilig-demo",
  persistState: true,
  zeroViewportBridge: true,
} as const;

describe("resolveRuntimeConfig", () => {
  afterEach(() => {
    window.history.replaceState({}, "", "/");
    vi.unstubAllGlobals();
  });

  it("forces zero viewport bridge off for direct server mode", () => {
    window.history.replaceState(
      {},
      "",
      "/?document=multiplayer-debug&server=http://127.0.0.1:4381",
    );

    expect(resolveRuntimeConfig(BASE_CONFIG)).toMatchObject({
      documentId: "multiplayer-debug",
      baseUrl: "http://127.0.0.1:4381",
      zeroViewportBridge: false,
    });
  });

  it("still allows the explicit bridge override without a direct server", () => {
    window.history.replaceState({}, "", "/?zeroViewportBridge=on");
    vi.stubGlobal("navigator", { webdriver: false });

    expect(resolveRuntimeConfig(BASE_CONFIG).zeroViewportBridge).toBe(true);
  });

  it("keeps the zero viewport bridge enabled under webdriver when no direct server is present", () => {
    window.history.replaceState({}, "", "/?document=playwright-doc");
    vi.stubGlobal("navigator", { webdriver: true });

    expect(resolveRuntimeConfig(BASE_CONFIG)).toMatchObject({
      documentId: "playwright-doc",
      baseUrl: null,
      zeroViewportBridge: true,
    });
  });

  it("resolves relative zero cache URLs against the current origin", () => {
    expect(resolveZeroCacheUrl("/zero", "http://127.0.0.1:4180")).toBe(
      "http://127.0.0.1:4180/zero",
    );
  });
});
