import { describe, expect, it } from "vitest";
import fc from "fast-check";
import { isWorkbookSnapshot, type WorkbookSnapshot } from "@bilig/protocol";
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
  projectMetadataSnapshot,
  snapshotSemanticActionArbitrary,
} from "./engine-fuzz-metadata-helpers.js";

describe("snapshot wire parity fuzz", () => {
  it("preserves workbook semantics and metadata across JSON wire serialization boundaries", async () => {
    await runProperty({
      suite: "core/import-export/snapshot-wire-parity",
      arbitrary: fc.record({
        seedName: engineSeedNameArbitrary,
        actions: fc.array(snapshotSemanticActionArbitrary, { minLength: 4, maxLength: 18 }),
      }),
      predicate: async ({ seedName, actions }) => {
        const seedSnapshot = await createEngineSeedSnapshot(
          seedName,
          `fuzz-snapshot-wire-${seedName}`,
        );
        const engine = new SpreadsheetEngine({
          workbookName: seedSnapshot.workbook.name,
          replicaId: `snapshot-wire-${seedName}`,
        });
        await engine.ready();
        engine.importSnapshot(structuredClone(seedSnapshot));

        actions.forEach((action) => {
          const result = applyEngineSemanticActionAndCaptureResult(engine, action);
          assertSnapshotInvariants(result.after);
        });

        const finalSnapshot = engine.exportSnapshot();
        assertSnapshotInvariants(finalSnapshot);
        const restoredSnapshot = await roundtripViaWire(finalSnapshot, seedName);

        expectSemanticSnapshot(restoredSnapshot, finalSnapshot);
        expect(projectMetadataSnapshot(restoredSnapshot)).toEqual(
          projectMetadataSnapshot(finalSnapshot),
        );
      },
    });
  });
});

function expectSemanticSnapshot(actual: WorkbookSnapshot, expected: WorkbookSnapshot): void {
  expect(normalizeSnapshotForSemanticComparison(actual)).toEqual(
    normalizeSnapshotForSemanticComparison(expected),
  );
}

async function roundtripViaWire(
  snapshot: WorkbookSnapshot,
  suffix: string,
): Promise<WorkbookSnapshot> {
  const restored = new SpreadsheetEngine({
    workbookName: snapshot.workbook.name,
    replicaId: `snapshot-wire-restored-${suffix}`,
  });
  await restored.ready();
  const wireSnapshot = JSON.parse(JSON.stringify(snapshot)) as unknown;
  if (!isWorkbookSnapshot(wireSnapshot)) {
    throw new Error("Snapshot wire roundtrip produced an invalid workbook snapshot payload");
  }
  restored.importSnapshot(wireSnapshot);
  return restored.exportSnapshot();
}
