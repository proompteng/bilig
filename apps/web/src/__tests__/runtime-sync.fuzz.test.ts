import { describe, it } from "vitest";
import fc from "fast-check";
import { runProperty } from "@bilig/test-fuzz";
import {
  applyRuntimeSyncAction,
  assertRuntimeSyncState,
  createRuntimeSyncHarness,
  runtimeSyncActionArbitrary,
} from "./runtime-sync-fuzz-helpers.js";

describe("runtime sync fuzz", () => {
  it("keeps runtime projection and pending journal state coherent across reconnect-style local and authoritative interleavings", async () => {
    await runProperty({
      suite: "web/runtime-sync/reconnect-convergence",
      arbitrary: fc.array(runtimeSyncActionArbitrary, { minLength: 4, maxLength: 20 }),
      predicate: async (actions) => {
        const { runtime, model } = await createRuntimeSyncHarness();
        try {
          await actions.reduce<Promise<void>>(async (previous, action) => {
            await previous;
            await applyRuntimeSyncAction(runtime, model, action);
            assertRuntimeSyncState(runtime, model);
          }, Promise.resolve());
        } finally {
          runtime.dispose();
        }
      },
    });
  });
});
