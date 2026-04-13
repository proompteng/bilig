import { readdirSync, readFileSync } from "node:fs";
import { extname } from "node:path";
import { fileURLToPath } from "node:url";
import {
  engineSeedNames,
  type EngineReplayCommand,
  type EngineSeedName,
} from "./engine-fuzz-helpers.js";

export interface EngineReplayExpectation {
  kind: "seed" | "snapshot";
  snapshot?: unknown;
}

export interface EngineReplayStep {
  command: EngineReplayCommand;
  expect?: EngineReplayExpectation;
}

export interface EngineReplayFixture {
  name: string;
  seed: EngineSeedName;
  steps: EngineReplayStep[];
}

const fixturesDir = fileURLToPath(new URL("./fixtures/fuzz-replays", import.meta.url));

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
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

function parseExpectation(value: unknown, fileName: string): EngineReplayExpectation | undefined {
  if (value === undefined) {
    return undefined;
  }
  if (!isRecord(value) || (value["kind"] !== "seed" && value["kind"] !== "snapshot")) {
    throw new Error(`Invalid engine fuzz replay expectation in ${fileName}`);
  }
  return {
    kind: value["kind"],
    snapshot: value["snapshot"],
  };
}

function parseCommand(value: unknown, fileName: string): EngineReplayCommand {
  if (!isRecord(value) || typeof value["kind"] !== "string") {
    throw new Error(`Invalid engine fuzz replay command in ${fileName}`);
  }
  const kind = value["kind"];
  switch (kind) {
    case "undo":
    case "redo":
      return { kind };
    case "formula":
      if (typeof value["address"] !== "string" || typeof value["formula"] !== "string") {
        throw new Error(`Invalid formula replay command in ${fileName}`);
      }
      return {
        kind,
        address: value["address"],
        formula: value["formula"],
      };
    case "insertRows":
    case "deleteRows":
    case "insertColumns":
    case "deleteColumns":
      if (typeof value["start"] !== "number" || typeof value["count"] !== "number") {
        throw new Error(`Invalid structural replay command in ${fileName}`);
      }
      return {
        kind,
        start: value["start"],
        count: value["count"],
      };
    case "format":
      if (typeof value["format"] !== "string") {
        throw new Error(`Invalid format replay command in ${fileName}`);
      }
      return {
        kind,
        range: parseRange(value["range"], fileName),
        format: value["format"],
      };
    default:
      throw new Error(`Unsupported engine fuzz replay command kind in ${fileName}: ${kind}`);
  }
}

function parseStep(value: unknown, fileName: string): EngineReplayStep {
  if (!isRecord(value)) {
    throw new Error(`Invalid engine fuzz replay step in ${fileName}`);
  }
  return {
    command: parseCommand(value["command"], fileName),
    expect: parseExpectation(value["expect"], fileName),
  };
}

function parseFixture(fileName: string): EngineReplayFixture {
  const parsed = JSON.parse(readFileSync(`${fixturesDir}/${fileName}`, "utf8")) as unknown;
  if (
    !isRecord(parsed) ||
    typeof parsed["name"] !== "string" ||
    typeof parsed["seed"] !== "string" ||
    !Array.isArray(parsed["steps"])
  ) {
    throw new Error(`Invalid engine fuzz replay fixture: ${fileName}`);
  }
  if (!engineSeedNames.some((seedName) => seedName === parsed["seed"])) {
    throw new Error(`Unknown engine fuzz replay seed in ${fileName}: ${String(parsed["seed"])}`);
  }
  return {
    name: parsed["name"],
    seed: parsed["seed"],
    steps: parsed["steps"].map((step) => parseStep(step, fileName)),
  };
}

export function loadEngineReplayFixtures(): EngineReplayFixture[] {
  return readdirSync(fixturesDir)
    .filter((fileName) => extname(fileName) === ".json")
    .toSorted((left, right) => left.localeCompare(right))
    .map(parseFixture);
}
