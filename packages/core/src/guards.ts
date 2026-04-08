import type { EngineReplicaSnapshot } from "./engine.js";

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

export function isEngineReplicaSnapshot(value: unknown): value is EngineReplicaSnapshot {
  return (
    isRecord(value) &&
    isRecord(value["replica"]) &&
    Array.isArray(value["entityVersions"]) &&
    Array.isArray(value["sheetDeleteVersions"])
  );
}
