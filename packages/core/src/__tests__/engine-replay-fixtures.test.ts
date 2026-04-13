import { describe, expect, it } from "vitest";
import { SpreadsheetEngine } from "../engine.js";
import {
  applyReplayCommand,
  createEngineSeedSnapshot,
  normalizeSnapshotForSemanticComparison,
} from "./engine-fuzz-helpers.js";
import { loadEngineReplayFixtures } from "./engine-fuzz-replay-fixtures.js";

describe("engine replay fixtures", () => {
  for (const fixture of loadEngineReplayFixtures()) {
    it(`replays ${fixture.name}`, async () => {
      const seedSnapshot = await createEngineSeedSnapshot(fixture.seed, fixture.name);
      const engine = new SpreadsheetEngine({
        workbookName: seedSnapshot.workbook.name,
        replicaId: `fixture-${fixture.name}`,
      });
      await engine.ready();
      engine.importSnapshot(structuredClone(seedSnapshot));

      for (const step of fixture.steps) {
        applyReplayCommand(engine, step.command);
        if (!step.expect) {
          continue;
        }
        const expectedSnapshot = step.expect.kind === "seed" ? seedSnapshot : step.expect.snapshot;
        if (!expectedSnapshot) {
          throw new Error(`Fixture ${fixture.name} has an expectation with no snapshot payload`);
        }
        expect(normalizeSnapshotForSemanticComparison(engine.exportSnapshot())).toEqual(
          normalizeSnapshotForSemanticComparison(expectedSnapshot),
        );
      }
    });
  }
});
