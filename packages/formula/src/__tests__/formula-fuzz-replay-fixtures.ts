import { readdirSync, readFileSync } from "node:fs";
import { extname } from "node:path";
import { fileURLToPath } from "node:url";

export interface FormulaReplayTranslationFixture {
  rowDelta: number;
  colDelta: number;
  translated: string;
  restored: string;
}

export interface FormulaReplayRenameFixture {
  oldSheetName: string;
  newSheetName: string;
  renamed: string;
  restored: string;
}

export interface FormulaReplayStructuralFixture {
  ownerSheetName: string;
  targetSheetName: string;
  transform: {
    kind: "insert" | "delete";
    axis: "row" | "column";
    start: number;
    count: number;
  };
  rewritten: string;
  reversed?: {
    kind: "insert" | "delete";
    axis: "row" | "column";
    start: number;
    count: number;
    restored: string;
  };
}

export interface FormulaReplayEvaluationFixture {
  expected:
    | { kind: "number"; value: number }
    | { kind: "string"; value: string }
    | { kind: "boolean"; value: boolean };
}

export interface FormulaReplayFixture {
  name: string;
  suite: string;
  guarantee: string;
  origin: string;
  source: string;
  canonical: string;
  translation?: FormulaReplayTranslationFixture;
  rename?: FormulaReplayRenameFixture;
  structural?: FormulaReplayStructuralFixture;
  evaluation?: FormulaReplayEvaluationFixture;
}

const fixturesDir = fileURLToPath(new URL("./fixtures/fuzz-replays", import.meta.url));

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function parseTranslation(
  value: unknown,
  fileName: string,
): FormulaReplayTranslationFixture | undefined {
  if (value === undefined) {
    return undefined;
  }
  if (
    !isRecord(value) ||
    typeof value["rowDelta"] !== "number" ||
    typeof value["colDelta"] !== "number" ||
    typeof value["translated"] !== "string" ||
    typeof value["restored"] !== "string"
  ) {
    throw new Error(`Invalid formula replay translation block in ${fileName}`);
  }
  return {
    rowDelta: value["rowDelta"],
    colDelta: value["colDelta"],
    translated: value["translated"],
    restored: value["restored"],
  };
}

function parseRename(value: unknown, fileName: string): FormulaReplayRenameFixture | undefined {
  if (value === undefined) {
    return undefined;
  }
  if (
    !isRecord(value) ||
    typeof value["oldSheetName"] !== "string" ||
    typeof value["newSheetName"] !== "string" ||
    typeof value["renamed"] !== "string" ||
    typeof value["restored"] !== "string"
  ) {
    throw new Error(`Invalid formula replay rename block in ${fileName}`);
  }
  return {
    oldSheetName: value["oldSheetName"],
    newSheetName: value["newSheetName"],
    renamed: value["renamed"],
    restored: value["restored"],
  };
}

function parseStructural(
  value: unknown,
  fileName: string,
): FormulaReplayStructuralFixture | undefined {
  if (value === undefined) {
    return undefined;
  }
  if (
    !isRecord(value) ||
    typeof value["ownerSheetName"] !== "string" ||
    typeof value["targetSheetName"] !== "string" ||
    !isRecord(value["transform"]) ||
    (value["transform"]["kind"] !== "insert" && value["transform"]["kind"] !== "delete") ||
    (value["transform"]["axis"] !== "row" && value["transform"]["axis"] !== "column") ||
    typeof value["transform"]["start"] !== "number" ||
    typeof value["transform"]["count"] !== "number" ||
    typeof value["rewritten"] !== "string"
  ) {
    throw new Error(`Invalid formula replay structural block in ${fileName}`);
  }
  let reversed: FormulaReplayStructuralFixture["reversed"];
  if (value["reversed"] !== undefined) {
    const reverse = value["reversed"];
    if (
      !isRecord(reverse) ||
      (reverse["kind"] !== "insert" && reverse["kind"] !== "delete") ||
      (reverse["axis"] !== "row" && reverse["axis"] !== "column") ||
      typeof reverse["start"] !== "number" ||
      typeof reverse["count"] !== "number" ||
      typeof reverse["restored"] !== "string"
    ) {
      throw new Error(`Invalid formula replay structural reverse block in ${fileName}`);
    }
    reversed = {
      kind: reverse["kind"],
      axis: reverse["axis"],
      start: reverse["start"],
      count: reverse["count"],
      restored: reverse["restored"],
    };
  }
  return {
    ownerSheetName: value["ownerSheetName"],
    targetSheetName: value["targetSheetName"],
    transform: {
      kind: value["transform"]["kind"],
      axis: value["transform"]["axis"],
      start: value["transform"]["start"],
      count: value["transform"]["count"],
    },
    rewritten: value["rewritten"],
    reversed,
  };
}

function parseEvaluation(
  value: unknown,
  fileName: string,
): FormulaReplayEvaluationFixture | undefined {
  if (value === undefined) {
    return undefined;
  }
  if (
    !isRecord(value) ||
    !isRecord(value["expected"]) ||
    typeof value["expected"]["kind"] !== "string"
  ) {
    throw new Error(`Invalid formula replay evaluation block in ${fileName}`);
  }
  const expected = value["expected"];
  switch (expected["kind"]) {
    case "number":
      if (typeof expected["value"] !== "number") {
        throw new Error(`Invalid numeric formula replay evaluation block in ${fileName}`);
      }
      return { expected: { kind: "number", value: expected["value"] } };
    case "string":
      if (typeof expected["value"] !== "string") {
        throw new Error(`Invalid string formula replay evaluation block in ${fileName}`);
      }
      return { expected: { kind: "string", value: expected["value"] } };
    case "boolean":
      if (typeof expected["value"] !== "boolean") {
        throw new Error(`Invalid boolean formula replay evaluation block in ${fileName}`);
      }
      return { expected: { kind: "boolean", value: expected["value"] } };
    default:
      throw new Error(`Unsupported formula replay evaluation kind in ${fileName}`);
  }
}

function parseFixture(fileName: string): FormulaReplayFixture {
  const parsed = JSON.parse(readFileSync(`${fixturesDir}/${fileName}`, "utf8")) as unknown;
  if (
    !isRecord(parsed) ||
    typeof parsed["name"] !== "string" ||
    typeof parsed["suite"] !== "string" ||
    typeof parsed["guarantee"] !== "string" ||
    typeof parsed["origin"] !== "string" ||
    typeof parsed["source"] !== "string" ||
    typeof parsed["canonical"] !== "string"
  ) {
    throw new Error(`Invalid formula replay fixture: ${fileName}`);
  }
  return {
    name: parsed["name"],
    suite: parsed["suite"],
    guarantee: parsed["guarantee"],
    origin: parsed["origin"],
    source: parsed["source"],
    canonical: parsed["canonical"],
    translation: parseTranslation(parsed["translation"], fileName),
    rename: parseRename(parsed["rename"], fileName),
    structural: parseStructural(parsed["structural"], fileName),
    evaluation: parseEvaluation(parsed["evaluation"], fileName),
  };
}

export function loadFormulaReplayFixtures(): FormulaReplayFixture[] {
  return readdirSync(fixturesDir)
    .filter((fileName) => extname(fileName) === ".json")
    .toSorted((left, right) => left.localeCompare(right))
    .map(parseFixture);
}
