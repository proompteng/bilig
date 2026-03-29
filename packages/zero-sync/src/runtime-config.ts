export interface BiligRuntimeConfig {
  zeroCacheUrl: string;
  defaultDocumentId: string;
  persistState: boolean;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function requireNonEmptyString(value: unknown, field: string): string {
  if (typeof value === "string" && value.length > 0) {
    return value;
  }
  throw new Error(`Runtime config field ${field} must be a non-empty string`);
}

function requireBoolean(value: unknown, field: string): boolean {
  if (typeof value === "boolean") {
    return value;
  }
  throw new Error(`Runtime config field ${field} must be a boolean`);
}

export function parseRuntimeConfig(value: unknown): BiligRuntimeConfig {
  if (!isRecord(value)) {
    throw new Error("Runtime config payload must be an object");
  }

  return {
    zeroCacheUrl: requireNonEmptyString(value["zeroCacheUrl"], "zeroCacheUrl"),
    defaultDocumentId: requireNonEmptyString(value["defaultDocumentId"], "defaultDocumentId"),
    persistState: requireBoolean(value["persistState"], "persistState"),
  };
}

export async function loadRuntimeConfig(
  fetchImpl: typeof fetch = fetch,
): Promise<BiligRuntimeConfig> {
  const response = await fetchImpl("/runtime-config.json", {
    headers: {
      accept: "application/json",
    },
  });

  if (!response.ok) {
    throw new Error(`Failed to load runtime config (${response.status})`);
  }

  return parseRuntimeConfig(await response.json());
}
