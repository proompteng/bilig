import { describe, expect, it } from "vitest";
import webViteConfig, { crossOriginIsolationHeaders } from "../../vite.config";

describe("web vite config", () => {
  it("serves dev and preview with cross-origin isolation headers", () => {
    expect(webViteConfig.server?.headers).toEqual(crossOriginIsolationHeaders);
    expect(webViteConfig.preview?.headers).toEqual(crossOriginIsolationHeaders);
  });
});
