import type { LiteralInput } from "@bilig/protocol";
import type { CommitOp, EngineReplicaSnapshot } from "./engine.js";

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function isLiteralInput(value: unknown): value is LiteralInput {
  return (
    value === null ||
    typeof value === "boolean" ||
    typeof value === "number" ||
    typeof value === "string"
  );
}

export function isEngineReplicaSnapshot(value: unknown): value is EngineReplicaSnapshot {
  return (
    isRecord(value) &&
    isRecord(value["replica"]) &&
    Array.isArray(value["entityVersions"]) &&
    Array.isArray(value["sheetDeleteVersions"])
  );
}

export function isCommitOp(value: unknown): value is CommitOp {
  if (!isRecord(value) || typeof value["kind"] !== "string") {
    return false;
  }
  switch (value["kind"]) {
    case "upsertWorkbook":
      return typeof value["name"] === "string";
    case "upsertSheet":
      return (
        typeof value["name"] === "string" &&
        (value["order"] === undefined || typeof value["order"] === "number")
      );
    case "renameSheet":
      return typeof value["oldName"] === "string" && typeof value["newName"] === "string";
    case "deleteSheet":
      return typeof value["name"] === "string";
    case "upsertCell":
      return (
        typeof value["sheetName"] === "string" &&
        typeof value["addr"] === "string" &&
        (isLiteralInput(value["value"]) ||
          typeof value["formula"] === "string" ||
          typeof value["format"] === "string")
      );
    case "deleteCell":
      return typeof value["sheetName"] === "string" && typeof value["addr"] === "string";
    default:
      return false;
  }
}

export function isCommitOps(value: unknown): value is CommitOp[] {
  return Array.isArray(value) && value.every((entry) => isCommitOp(entry));
}
