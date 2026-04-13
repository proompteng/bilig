import { describe, expect, it } from "vitest";
import fc from "fast-check";
import type { WorkbookSnapshot } from "@bilig/protocol";
import { SpreadsheetEngine } from "../engine.js";
import { runProperty } from "@bilig/test-fuzz";
import {
  assertSnapshotInvariants,
  createEngineSeedSnapshot,
  engineSeedNameArbitrary,
  normalizeSnapshotForSemanticComparison,
} from "./engine-fuzz-helpers.js";
import {
  applyEngineSemanticActionAndCaptureResult,
  exportEngineSemanticReplaySnapshot,
  projectMetadataSnapshot,
  snapshotSemanticActionArbitrary,
} from "./engine-fuzz-metadata-helpers.js";

function expectSemanticSnapshot(actual: WorkbookSnapshot, expected: WorkbookSnapshot): void {
  expect(normalizeSnapshotForSemanticComparison(actual)).toEqual(
    normalizeSnapshotForSemanticComparison(expected),
  );
}

async function roundtripSnapshot(
  snapshot: WorkbookSnapshot,
  suffix: string,
): Promise<WorkbookSnapshot> {
  const restored = new SpreadsheetEngine({
    workbookName: snapshot.workbook.name,
    replicaId: `snapshot-${suffix}`,
  });
  await restored.ready();
  restored.importSnapshot(structuredClone(snapshot));
  return restored.exportSnapshot();
}

describe("engine snapshot fuzz", () => {
  it("preserves semantic parity across repeated snapshot restore cycles", async () => {
    await runProperty({
      suite: "core/snapshot/restore-roundtrip-parity",
      arbitrary: fc.record({
        seedName: engineSeedNameArbitrary,
        actions: fc.array(snapshotSemanticActionArbitrary, { minLength: 4, maxLength: 18 }),
      }),
      predicate: async ({ seedName, actions }) => {
        const seedSnapshot = await createEngineSeedSnapshot(
          seedName,
          `fuzz-core-snapshot-${seedName}`,
        );
        const engine = new SpreadsheetEngine({
          workbookName: seedSnapshot.workbook.name,
          replicaId: `snapshot-primary-${seedName}`,
        });
        await engine.ready();
        engine.importSnapshot(structuredClone(seedSnapshot));

        const accepted = await actions.reduce<Promise<Array<(typeof actions)[number]>>>(
          async (acceptedPromise, action) => {
            const acceptedActions = await acceptedPromise;
            const result = applyEngineSemanticActionAndCaptureResult(engine, action);
            assertSnapshotInvariants(result.after);
            if (result.accepted) {
              acceptedActions.push(action);
            }
            return acceptedActions;
          },
          Promise.resolve([]),
        );

        const finalSnapshot = engine.exportSnapshot();
        assertSnapshotInvariants(finalSnapshot);
        const expectedSnapshot = await exportEngineSemanticReplaySnapshot(seedSnapshot, accepted);
        expectSemanticSnapshot(finalSnapshot, expectedSnapshot);
        expect(projectMetadataSnapshot(finalSnapshot)).toEqual(
          projectMetadataSnapshot(expectedSnapshot),
        );

        const restoredOnce = await roundtripSnapshot(finalSnapshot, `${seedName}-once`);
        const restoredTwice = await roundtripSnapshot(restoredOnce, `${seedName}-twice`);
        expectSemanticSnapshot(restoredOnce, finalSnapshot);
        expectSemanticSnapshot(restoredTwice, finalSnapshot);
        expect(projectMetadataSnapshot(restoredTwice)).toEqual(
          projectMetadataSnapshot(finalSnapshot),
        );
      },
    });
  });
});
