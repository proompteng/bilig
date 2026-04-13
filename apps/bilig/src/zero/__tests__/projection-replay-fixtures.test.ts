import { readdirSync } from "node:fs";
import { extname } from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";
import { SpreadsheetEngine } from "@bilig/core";
import { loadReplayFixture } from "@bilig/test-fuzz";
import {
  applyProjectionAction,
  projectProjectionFromEngine,
  projectProjectionFromSnapshot,
  type ProjectionAction,
} from "./projection-fuzz-helpers.js";

const fixturesDir = fileURLToPath(new URL("./fixtures/fuzz-replays", import.meta.url));

describe("projection replay fixtures", () => {
  for (const fixture of loadProjectionReplayFixtures()) {
    it(`replays ${fixture.name}`, async () => {
      const engine = new SpreadsheetEngine({
        workbookName: `projection-replay-${fixture.name}`,
        replicaId: `projection-replay-${fixture.name}`,
      });
      await engine.ready();
      fixture.actions.forEach((action) => {
        applyProjectionAction(engine, action);
      });
      expect(projectProjectionFromEngine(engine)).toEqual(projectProjectionFromSnapshot(engine));
    });
  }
});

type ProjectionReplayFixture = {
  name: string;
  actions: ProjectionAction[];
};

function loadProjectionReplayFixtures(): ProjectionReplayFixture[] {
  return readdirSync(fixturesDir)
    .filter((fileName) => extname(fileName) === ".json")
    .toSorted((left, right) => left.localeCompare(right))
    .map((fileName) => {
      const fixture = loadReplayFixture(`${fixturesDir}/${fileName}`);
      if (!Array.isArray(fixture.counterexample)) {
        throw new Error(`Projection replay fixture ${fileName} is missing counterexample actions`);
      }
      return {
        name: fileName.replace(/\.json$/u, ""),
        actions: fixture.counterexample.map((action) => parseProjectionAction(action, fileName)),
      };
    });
}

function parseProjectionAction(value: unknown, fileName: string): ProjectionAction {
  if (!isRecord(value) || typeof value["kind"] !== "string") {
    throw new Error(`Invalid projection replay action in ${fileName}`);
  }
  const kind = value["kind"];
  switch (kind) {
    case "value":
      if (typeof value["address"] !== "string") {
        break;
      }
      return { kind, address: value["address"], value: parseLiteralInput(value["value"]) };
    case "formula":
      if (typeof value["address"] !== "string" || typeof value["formula"] !== "string") {
        break;
      }
      return { kind, address: value["address"], formula: value["formula"] };
    case "style":
      return {
        kind,
        range: parseRange(value["range"], fileName),
        patch: parsePatch(value["patch"], fileName),
      };
    case "format":
      return {
        kind,
        range: parseRange(value["range"], fileName),
        format: parseFormat(value["format"], fileName),
      };
    case "insertRows":
    case "deleteRows":
    case "insertColumns":
    case "deleteColumns":
      if (typeof value["start"] === "number" && typeof value["count"] === "number") {
        return { kind, start: value["start"], count: value["count"] };
      }
      break;
  }
  throw new Error(`Invalid projection replay action in ${fileName}: ${JSON.stringify(value)}`);
}

function parseLiteralInput(value: unknown) {
  if (
    value === null ||
    typeof value === "number" ||
    typeof value === "boolean" ||
    typeof value === "string"
  ) {
    return value;
  }
  throw new Error(`Invalid projection replay literal value: ${JSON.stringify(value)}`);
}

function parseRange(value: unknown, fileName: string) {
  if (
    !isRecord(value) ||
    typeof value["sheetName"] !== "string" ||
    typeof value["startAddress"] !== "string" ||
    typeof value["endAddress"] !== "string"
  ) {
    throw new Error(`Invalid range in ${fileName}`);
  }
  return {
    sheetName: value["sheetName"],
    startAddress: value["startAddress"],
    endAddress: value["endAddress"],
  };
}

function parsePatch(value: unknown, fileName: string) {
  if (!isRecord(value)) {
    throw new Error(`Invalid style patch in ${fileName}`);
  }
  return value;
}

function parseFormat(value: unknown, fileName: string) {
  if (typeof value === "string") {
    return value;
  }
  if (isRecord(value)) {
    return value;
  }
  throw new Error(`Invalid format patch in ${fileName}`);
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}
